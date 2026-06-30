import re
import sys
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import xml.etree.ElementTree as ET

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# =========================================================
# CONFIG
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
PASTA_XML = BASE_DIR / "entrada" / "xml"
PASTA_OC = BASE_DIR / "entrada" / "oc"
PASTA_RESULTADO = BASE_DIR / "resultado"
ARQUIVO_RESULTADO = PASTA_RESULTADO / "resultado.xlsx"
NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

# =========================================================
# UTILITÁRIOS
# =========================================================
def garantir_pastas():
    PASTA_RESULTADO.mkdir(parents=True, exist_ok=True)

def arquivo_vazio(caminho: Path) -> bool:
    return (not caminho.exists()) or caminho.stat().st_size == 0

def normalize_spaces(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "")).strip()

def normalizar_texto(texto: str) -> str:
    texto = normalize_spaces(texto).upper()
    texto = re.sub(r"[^A-Z0-9 ]", "", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto

def limpar_cnpj(texto: str) -> str:
    return re.sub(r"\D", "", str(texto or ""))

def raiz_cnpj(cnpj: str) -> str:
    c = limpar_cnpj(cnpj)
    return c[:8] if len(c) >= 8 else c

def numero(valor, default=0.0):
    try:
        if valor is None:
            return default
        if isinstance(valor, (int, float)):
            return float(valor)

        s = str(valor).strip().replace("R$", "")
        if not s:
            return default

        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        else:
            s = s.replace(",", ".")

        return float(s)
    except Exception:
        return default

def parse_data(valor):
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor

    s = str(valor).strip()
    if not s:
        return None

    for fmt in (
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d/%m/%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass

    m = re.match(r"^(\d{4}-\d{2}-\d{2})", s)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y-%m-%d")
        except Exception:
            pass

    return None

def data_str(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y") if isinstance(dt, datetime) else ""

def comparar_texto_simples(a: str, b: str) -> bool:
    na = normalizar_texto(a)
    nb = normalizar_texto(b)
    if not na or not nb:
        return False
    return na == nb or na in nb or nb in na

def comparar_nome_fornecedor(a: str, b: str) -> bool:
    na = normalizar_texto(a)
    nb = normalizar_texto(b)
    if not na or not nb:
        return False

    # aliases simples para reduzir falso negativo no caso MEDERI x MDR
    aliases = [
        ("MDR", "MEDERI"),
    ]
    for x, y in aliases:
        na = na.replace(x, y)
        nb = nb.replace(x, y)

    return na == nb or na in nb or nb in na

def normalizar_mod_frete(codigo):
    mapa = {
        "0": "Por conta do emitente",
        "1": "Por conta do destinatário/remetente",
        "2": "Por conta de terceiros",
        "9": "Sem frete",
    }
    return mapa.get(str(codigo), str(codigo or ""))

def texto_xml(elem, xpath, default=""):
    try:
        achado = elem.find(xpath, NS)
        return normalize_spaces(achado.text) if achado is not None and achado.text is not None else default
    except Exception:
        return default

def buscar_regex(texto: str, padrao: str, grupo=1, flags=re.I | re.S, default=""):
    m = re.search(padrao, texto, flags=flags)
    return normalize_spaces(m.group(grupo)) if m else default

# =========================================================
# DOC ORIGEM NA NF
# =========================================================
def extrair_doc_origem_da_obs(observacao: str) -> str:
    if not observacao:
        return ""

    padroes = [
        r"Doc\.?\s*Origem\s*[:\-]?\s*([0-9]+)",
        r"Doc\.?\s*Orig\.?\s*[:\-]?\s*([0-9]+)",
        r"Pedido\s+interno\s*[:\-]?\s*([0-9]+)",
        r"N[ºo]?\s*OC\(s\)\s*[:\-]?\s*([0-9]+)",
        r"N[ºo]?\s*OC\s*[:\-]?\s*([0-9]+)",
        r"Pedido\s*[:\-]?\s*([0-9]+)",
    ]

    for p in padroes:
        m = re.search(p, observacao, flags=re.IGNORECASE)
        if m:
            return m.group(1)
    return ""

# =========================================================
# XML NF-e
# =========================================================
def extrair_xml(caminho_xml: Path):
    if arquivo_vazio(caminho_xml):
        raise ValueError("Arquivo XML vazio ou inexistente")

    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
    except ET.ParseError as e:
        raise ValueError(f"Arquivo XML inválido: {e}")

    numero_nota = texto_xml(root, ".//nfe:ide/nfe:nNF")
    natureza_operacao = texto_xml(root, ".//nfe:ide/nfe:natOp")
    data_emissao_raw = texto_xml(root, ".//nfe:ide/nfe:dhEmi")
    data_saida_raw = texto_xml(root, ".//nfe:ide/nfe:dhSaiEnt")
    data_programada_raw = texto_xml(root, ".//nfe:ide/nfe:dPrevEntrega")

    emitente_nome = texto_xml(root, ".//nfe:emit/nfe:xNome")
    emitente_cnpj = limpar_cnpj(texto_xml(root, ".//nfe:emit/nfe:CNPJ"))

    destinatario_nome = texto_xml(root, ".//nfe:dest/nfe:xNome")
    destinatario_cnpj = limpar_cnpj(texto_xml(root, ".//nfe:dest/nfe:CNPJ"))
    destinatario_endereco = texto_xml(root, ".//nfe:dest/nfe:enderDest/nfe:xLgr")
    destinatario_numero = texto_xml(root, ".//nfe:dest/nfe:enderDest/nfe:nro")
    destinatario_bairro = texto_xml(root, ".//nfe:dest/nfe:enderDest/nfe:xBairro")
    destinatario_municipio = texto_xml(root, ".//nfe:dest/nfe:enderDest/nfe:xMun")
    destinatario_uf = texto_xml(root, ".//nfe:dest/nfe:enderDest/nfe:UF")
    destinatario_cep = texto_xml(root, ".//nfe:dest/nfe:enderDest/nfe:CEP")

    valor_total_produtos = numero(texto_xml(root, ".//nfe:total/nfe:ICMSTot/nfe:vProd"))
    valor_total_nota = numero(texto_xml(root, ".//nfe:total/nfe:ICMSTot/nfe:vNF"))

    transportadora_nome = texto_xml(root, ".//nfe:transp/nfe:transporta/nfe:xNome")
    transportadora_cnpj = limpar_cnpj(texto_xml(root, ".//nfe:transp/nfe:transporta/nfe:CNPJ"))
    frete_por_conta = normalizar_mod_frete(texto_xml(root, ".//nfe:transp/nfe:modFrete"))

    observacao_nota = texto_xml(root, ".//nfe:infAdic/nfe:infCpl")
    doc_origem_encontrado_xml = extrair_doc_origem_da_obs(observacao_nota)

    # compra/xPed é importantíssimo no teu cenário
    xped = texto_xml(root, ".//nfe:compra/nfe:xPed")
    if xped and not doc_origem_encontrado_xml:
        doc_origem_encontrado_xml = xped

    data_emissao = parse_data(data_emissao_raw)
    data_saida = parse_data(data_saida_raw)
    data_programada = parse_data(data_programada_raw)

    cabecalho = {
        "arquivo_xml": caminho_xml.name,
        "emitente_nome": emitente_nome,
        "emitente_cnpj": emitente_cnpj,
        "emitente_raiz_cnpj": raiz_cnpj(emitente_cnpj),
        "numero_nota": numero_nota,
        "natureza_operacao": natureza_operacao,
        "destinatario_nome": destinatario_nome,
        "destinatario_cnpj": destinatario_cnpj,
        "destinatario_raiz_cnpj": raiz_cnpj(destinatario_cnpj),
        "endereco_destinatario": destinatario_endereco,
        "numero_endereco_destinatario": destinatario_numero,
        "bairro_destinatario": destinatario_bairro,
        "municipio_destinatario": destinatario_municipio,
        "uf_destinatario": destinatario_uf,
        "cep_destinatario": destinatario_cep,
        "data_emissao": data_str(data_emissao),
        "data_saida_expedicao": data_str(data_saida),
        "data_programada_xml": data_str(data_programada),
        "valor_total_produtos": valor_total_produtos,
        "valor_total_nota": valor_total_nota,
        "transportadora_nome": transportadora_nome,
        "transportadora_cnpj": transportadora_cnpj,
        "frete_por_conta": frete_por_conta,
        "observacao_nota": observacao_nota,
        "doc_origem_encontrado_xml": doc_origem_encontrado_xml,
        "xped_xml": xped,
    }

    itens = []
    for det in root.findall(".//nfe:det", NS):
        prod = det.find("nfe:prod", NS)
        imposto = det.find("nfe:imposto", NS)
        if prod is None:
            continue

        itens.append({
            "arquivo_xml": caminho_xml.name,
            "numero_nota": numero_nota,
            "doc_origem_encontrado_xml": doc_origem_encontrado_xml,
            "codigo_item_xml": texto_xml(prod, "nfe:cProd"),
            "descricao_item_xml": texto_xml(prod, "nfe:xProd"),
            "unidade_item_xml": texto_xml(prod, "nfe:uCom"),
            "quantidade_item_xml": numero(texto_xml(prod, "nfe:qCom")),
            "valor_unitario_item_xml": numero(texto_xml(prod, "nfe:vUnCom")),
            "valor_total_item_xml": numero(texto_xml(prod, "nfe:vProd")),
            "lote_item_xml": texto_xml(prod, "nfe:rastro/nfe:nLote"),
            "validade_item_xml": texto_xml(prod, "nfe:rastro/nfe:dVal"),
            "ipi_aliquota_item_xml": texto_xml(imposto, "nfe:IPI/nfe:IPITrib/nfe:pIPI") if imposto is not None else "",
            "ipi_valor_item_xml": texto_xml(imposto, "nfe:IPI/nfe:IPITrib/nfe:vIPI") if imposto is not None else "",
        })

    return cabecalho, itens

# =========================================================
# LEITURA OC DOC/TXT
# =========================================================
def limpar_texto_extraido(texto: str) -> str:
    if not texto:
        return ""

    texto = texto.replace("\x00", "")
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")
    texto = "".join(
        ch for ch in texto
        if ch == "\n" or ch == "\t" or ord(ch) >= 32
    )
    return texto

def converter_doc_word_para_txt(caminho_doc: Path) -> str:
    try:
        import win32com.client
    except Exception as e:
        raise ValueError("pywin32 não está instalado. Rode: py -m pip install pywin32") from e

    caminho_doc_abs = str(caminho_doc.resolve())
    caminho_txt = PASTA_RESULTADO / f"_convertido_{caminho_doc.stem}.txt"
    caminho_txt_abs = str(caminho_txt.resolve())

    word = None
    doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        doc = word.Documents.Open(
            caminho_doc_abs,
            ReadOnly=True,
            ConfirmConversions=False,
            AddToRecentFiles=False
        )

        doc.SaveAs2(caminho_txt_abs, FileFormat=2)  # wdFormatText
        doc.Close(False)
        word.Quit()

        texto = caminho_txt.read_text(encoding="cp1252", errors="ignore")
        return limpar_texto_extraido(texto)

    except Exception as e:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass

        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass

        raise ValueError(f"Falha ao converter DOC pelo Word: {e}")

def ler_doc_texto(caminho_doc: Path) -> str:
    if arquivo_vazio(caminho_doc):
        raise ValueError("Arquivo DOC/TXT da OC vazio ou inexistente")

    ext = caminho_doc.suffix.lower()
    texto = ""

    if ext in [".doc", ".docx"]:
        texto = converter_doc_word_para_txt(caminho_doc)
    else:
        bruto = caminho_doc.read_bytes()
        for enc in ("utf-8", "utf-16", "utf-16le", "latin1", "cp1252"):
            try:
                candidato = bruto.decode(enc, errors="ignore")
                candidato = limpar_texto_extraido(candidato)
                if len(candidato) > len(texto):
                    texto = candidato
            except Exception:
                pass

    if not texto.strip():
        raise ValueError(f"Não foi possível extrair texto útil de {caminho_doc.name}")

    try:
        (PASTA_RESULTADO / f"debug_texto_lido_{caminho_doc.stem}.txt").write_text(texto, encoding="utf-8")
    except Exception:
        pass

    return texto

def extrair_cabecalho_oc_doc(caminho_doc: Path):
    texto = ler_doc_texto(caminho_doc)
    texto_flat = normalize_spaces(texto)

    numero_oc = buscar_regex(texto_flat, r"ORDEM\s+DE\s+COMPRA\s+(\d+)", default="")
    if not numero_oc:
        numero_oc = buscar_regex(texto_flat, r"Pedido\s*:\s*(\d+)", default="")

    doc_origem = buscar_regex(texto_flat, r"Doc\.?\s*Origem\s*:\s*(\d+)", default="")
    pedido = buscar_regex(texto_flat, r"Pedido\s*:\s*(\d+)", default="")
    data_pedido = buscar_regex(texto_flat, r"Data\s*:\s*(\d{2}/\d{2}/\d{4})", default="")
    prazo_entrega = buscar_regex(texto_flat, r"PRAZO\s+DE\s+ENTREGA\s*:\s*(\d{2}/\d{2}/\d{4})", default="")

    cliente_cnpj = limpar_cnpj(buscar_regex(
        texto_flat, r"C\.N\.P\.J\.?\s*:\s*([0-9\.\-/]+)", default=""
    ))

    cliente_nome = buscar_regex(
        texto_flat,
        r"Doc\.?\s*Origem\s*:\s*\d+\s+(.+?)\s+Pedido\s*:",
        default=""
    )

    endereco_principal = buscar_regex(
        texto_flat,
        r"Pedido\s*:\s*\d+\s+(.+?)\s+Data\s*:",
        default=""
    )

    endereco_entrega = buscar_regex(
        texto_flat,
        r"LOCAL\s+ENTREGA\s*:\s*(.+?)\s+LOCAL\s+COBRANÇA",
        default=""
    )

    fornecedor_nome = buscar_regex(
        texto_flat,
        r"Fornecedor\s*:\s*\d+\s*-\s*(.+?)\s+Endereço\s*:",
        default=""
    )

    fornecedor_cnpj = limpar_cnpj(buscar_regex(
        texto_flat, r"CNPJ\s*:\s*([0-9\.\-/]+)", default=""
    ))

    valor_total_pedido = numero(buscar_regex(
        texto_flat, r"Valor\s+do\s+Pedido\s*:\s*([0-9\.,]+)", default=""
    ))

    return {
        "arquivo_oc_doc": caminho_doc.name,
        "numero_oc": numero_oc,
        "pedido": pedido,
        "doc_origem": doc_origem,
        "data_pedido": data_pedido,
        "prazo_entrega": prazo_entrega,
        "cliente_nome": cliente_nome,
        "cliente_cnpj": cliente_cnpj,
        "cliente_raiz_cnpj": raiz_cnpj(cliente_cnpj),
        "endereco_principal": endereco_principal,
        "endereco_entrega": endereco_entrega,
        "fornecedor_nome": fornecedor_nome,
        "fornecedor_cnpj": fornecedor_cnpj,
        "fornecedor_raiz_cnpj": raiz_cnpj(fornecedor_cnpj),
        "valor_total_pedido": valor_total_pedido,
        "texto_extraido_doc_preview": texto_flat[:500],
    }

def extrair_itens_oc_doc(caminho_doc: Path) -> pd.DataFrame:
    texto = ler_doc_texto(caminho_doc)
    linhas = [linha.rstrip() for linha in texto.splitlines()]
    registros = []
    item_atual = None

    for linha in linhas:
        limpa = normalize_spaces(linha)
        if not limpa:
            continue

        partes = [p.strip() for p in linha.split("\t") if p.strip()]

        # formato tabulado esperado no DOC exportado
        if partes and re.fullmatch(r"\d{5,6}", partes[0]):
            if item_atual:
                registros.append(item_atual)

            codigo_item_oc = partes[0]
            descricao_item_oc = partes[1] if len(partes) > 1 else ""
            unidade_item_oc = partes[2] if len(partes) > 2 else ""
            embalagem_item_oc = numero(partes[3]) if len(partes) > 3 else 0.0
            quantidade_item_oc = numero(partes[4]) if len(partes) > 4 else 0.0

            valor_unitario_item_oc = 0.0
            desconto_item_oc = 0.0
            ipi_item_oc = 0.0
            valor_total_item_oc = 0.0

            if len(partes) > 5:
                moedas = re.findall(r"\d+[.,]\d{2}", partes[5])
                if len(moedas) >= 1:
                    valor_unitario_item_oc = numero(moedas[0])
                if len(moedas) >= 2:
                    desconto_item_oc = numero(moedas[1])
                if len(moedas) >= 3:
                    ipi_item_oc = numero(moedas[2])

            if len(partes) > 6:
                valor_total_item_oc = numero(partes[6])

            item_atual = {
                "arquivo_oc_doc": caminho_doc.name,
                "codigo_item_oc": str(codigo_item_oc).strip(),
                "descricao_item_oc": descricao_item_oc,
                "unidade_item_oc": unidade_item_oc,
                "embalagem_item_oc": embalagem_item_oc,
                "quantidade_item_oc": quantidade_item_oc,
                "valor_unitario_item_oc": valor_unitario_item_oc,
                "desconto_item_oc": desconto_item_oc,
                "ipi_item_oc": ipi_item_oc,
                "valor_total_item_oc": valor_total_item_oc,
                "metodo_extracao": "doc_tabulado",
            }
            continue

        if item_atual:
            texto_linha = limpa.upper()

            fim_item = any(x in texto_linha for x in [
                "VALOR DO PEDIDO",
                "CONDIÇÕES DE PAGAMENTO",
                "SETOR SOLICITANTE",
                "DGS BRASIL",
                "PÁGINA ",
                "----",
            ])

            cabecalho_tabela = any(x in texto_linha for x in [
                "REF.MAT",
                "UN.COMP",
                "QTD.PEDIDA",
                "VALOR UNIT",
                "CÓDIGO",
                "DESCRIÇÃO DO ITEM",
                "FABRICANTE",
            ])

            if not fim_item and not cabecalho_tabela:
                item_atual["descricao_item_oc"] = normalize_spaces(
                    item_atual.get("descricao_item_oc", "") + " " + limpa
                )

    if item_atual:
        registros.append(item_atual)

    if not registros:
        try:
            (PASTA_RESULTADO / f"debug_itens_{caminho_doc.stem}.txt").write_text(texto, encoding="utf-8")
        except Exception:
            pass

    return pd.DataFrame(registros)

# =========================================================
# REGRAS DE ITEM
# =========================================================
def limpar_descricao(desc):
    stopwords = {
        "DE", "DA", "DO", "DAS", "DOS", "COM", "PARA", "EM",
        "LTDA", "SIMPLES", "UNIDADE", "UN"
    }
    palavras = normalizar_texto(desc).split()
    return " ".join([p for p in palavras if p not in stopwords])

def comparar_descricao_inteligente(a, b):
    p1 = set(limpar_descricao(a).split())
    p2 = set(limpar_descricao(b).split())
    if not p1 or not p2:
        return False
    inter = p1 & p2
    # threshold simples e seguro
    return len(inter) >= 2

def extrair_fator(descricao):
    """
    Extrai fator de expressões tipo 2X8, 10X50.
    Não captura coisas como 10FRX45CM, pois ali não há número x número "isolado".
    """
    desc = str(descricao or "").upper()
    m = re.search(r"\b(\d{1,3})\s*[Xx]\s*(\d{1,3})\b", desc)
    if not m:
        return None
    a = int(m.group(1))
    b = int(m.group(2))
    return a * b

def consolidar_itens_xml(xml_itens: list) -> list:
    """
    Consolida linhas repetidas da NF por código + descrição normalizada.
    Ex.: SLIP3515S aparece duas vezes na NF do teu caso; aqui vira 1 só consolidado.
    """
    mapa = {}

    for item in xml_itens:
        codigo = str(item.get("codigo_item_xml", "") or "").strip()
        descricao = str(item.get("descricao_item_xml", "") or "")
        chave = (codigo, limpar_descricao(descricao))

        if chave not in mapa:
            mapa[chave] = dict(item)
        else:
            mapa[chave]["quantidade_item_xml"] = numero(mapa[chave].get("quantidade_item_xml", 0)) + numero(item.get("quantidade_item_xml", 0))
            mapa[chave]["valor_total_item_xml"] = numero(mapa[chave].get("valor_total_item_xml", 0)) + numero(item.get("valor_total_item_xml", 0))

    return list(mapa.values())

def encontrar_item_oc_para_nf(item_xml: dict, df_oc: pd.DataFrame):
    """
    Estratégia de match:
    1) código exato
    2) código da NF dentro da descrição da OC
    3) código da OC dentro da descrição da NF
    4) descrição inteligente
    """
    if df_oc is None or df_oc.empty:
        return None, "OC_SEM_ITENS"

    codigo_xml = str(item_xml.get("codigo_item_xml", "") or "").strip()
    desc_xml = str(item_xml.get("descricao_item_xml", "") or "")

    # 1) código exato
    match = df_oc[df_oc["codigo_item_oc"].astype(str).str.strip() == codigo_xml]
    if not match.empty:
        return match.iloc[0], "MATCH_CODIGO_EXATO"

    # 2) código XML dentro da descrição da OC
    candidatos = df_oc[
        df_oc["descricao_item_oc"].astype(str).str.upper().str.contains(re.escape(codigo_xml), na=False)
    ]
    if not candidatos.empty:
        return candidatos.iloc[0], "MATCH_CODIGO_EM_DESCRICAO_OC"

    # 3) descrição da NF contém código da OC
    for _, row in df_oc.iterrows():
        cod_oc = str(row.get("codigo_item_oc", "") or "").strip()
        if cod_oc and cod_oc in desc_xml:
            return row, "MATCH_CODIGO_OC_EM_DESCRICAO_NF"

    # 4) descrição inteligente
    for _, row in df_oc.iterrows():
        if comparar_descricao_inteligente(desc_xml, row.get("descricao_item_oc", "")):
            return row, "MATCH_DESCRICAO_INTELIGENTE"

    return None, "ITEM_NAO_ENCONTRADO"

def comparar_itens(xml_itens: list, df_oc_itens: pd.DataFrame):
    """
    Regra correta:
    - Todo item da NF precisa existir na OC
    - A OC pode ter mais itens que a NF
    - Recebimento parcial é permitido: qtd_nf <= qtd_oc
    """
    relatorio = []
    itens_ok = 0
    itens_erro = 0

    if df_oc_itens is None or df_oc_itens.empty:
        for item in xml_itens:
            relatorio.append({
                "codigo_item_xml": item.get("codigo_item_xml", ""),
                "descricao_item_xml": item.get("descricao_item_xml", ""),
                "status_item": "ERRO",
                "motivo_item": "OC_SEM_ITENS",
                "metodo_match": ""
            })
        return relatorio, 0, len(xml_itens)

    xml_itens = consolidar_itens_xml(xml_itens)

    for item_xml in xml_itens:
        codigo_xml = str(item_xml.get("codigo_item_xml", "") or "").strip()
        desc_xml = str(item_xml.get("descricao_item_xml", "") or "")
        qtd_xml_original = numero(item_xml.get("quantidade_item_xml", 0))
        vl_unit_xml = numero(item_xml.get("valor_unitario_item_xml", 0))
        vl_total_xml = numero(item_xml.get("valor_total_item_xml", 0))

        fator = extrair_fator(desc_xml)
        qtd_xml_ajustada = fator if fator else qtd_xml_original

        item_oc, metodo = encontrar_item_oc_para_nf(item_xml, df_oc_itens)
        motivos = []

        if item_oc is None:
            motivos.append("ITEM_NAO_ENCONTRADO")
            status = "ERRO"
        else:
            qtd_oc = numero(item_oc.get("quantidade_item_oc", 0))
            vl_unit_oc = numero(item_oc.get("valor_unitario_item_oc", 0))
            vl_total_oc = numero(item_oc.get("valor_total_item_oc", 0))
            embalagem_oc = numero(item_oc.get("embalagem_item_oc", 0))
            unidade_xml = normalizar_texto(item_xml.get("unidade_item_xml", ""))
            unidade_oc = normalizar_texto(item_oc.get("unidade_item_oc", ""))

            # Quantidade: NF pode ser parcial, então <= OC
            if qtd_xml_ajustada > qtd_oc + 0.01:
                # tenta regra simples com embalagem/fator da OC
                convertido = qtd_oc * embalagem_oc if embalagem_oc and embalagem_oc > 1 else qtd_oc
                if qtd_xml_ajustada > convertido + 0.01:
                    motivos.append("QTD_ACIMA_DA_OC")

            # Unidade: informativo, não mata sozinho
            if unidade_xml and unidade_oc and unidade_xml != unidade_oc:
                motivos.append("UNIDADE_DIVERGENTE")

            # Valor unitário: tolerância
            if vl_unit_oc > 0 and abs(vl_unit_xml - vl_unit_oc) > 0.05:
                motivos.append("VALOR_UNITARIO_DIVERGENTE")

            # Valor total do item na NF não pode estourar o valor do item correspondente na OC
            if vl_total_oc > 0 and vl_total_xml > vl_total_oc + 0.05:
                motivos.append("VALOR_TOTAL_ACIMA_DA_OC")

            status = "OK" if not any(m for m in motivos if m not in {"UNIDADE_DIVERGENTE"}) else "ERRO"

        if status == "OK":
            itens_ok += 1
        else:
            itens_erro += 1

        relatorio.append({
            "codigo_item_xml": codigo_xml,
            "descricao_item_xml": desc_xml,
            "quantidade_item_xml_original": qtd_xml_original,
            "quantidade_item_xml_ajustada": qtd_xml_ajustada,
            "fator_detectado": fator or "",
            "status_item": status,
            "motivo_item": "|".join(motivos),
            "metodo_match": metodo,
        })

    return relatorio, itens_ok, itens_erro

# =========================================================
# PAREAMENTO
# =========================================================
def listar_arquivos():
    xmls = sorted(PASTA_XML.glob("*.xml"))
    docs_oc = sorted(
        list(PASTA_OC.glob("*.doc")) +
        list(PASTA_OC.glob("*.docx")) +
        list(PASTA_OC.glob("*.txt")) +
        list(PASTA_OC.glob("*.csv"))
    )
    return xmls, docs_oc

def montar_pares():
    xmls, docs_oc = listar_arquivos()
    total = max(len(xmls), len(docs_oc)) if (xmls or docs_oc) else 0
    pares = []
    for i in range(total):
        pares.append({
            "xml": xmls[i] if i < len(xmls) else None,
            "oc_doc": docs_oc[i] if i < len(docs_oc) else None,
        })
    return pares

def mostrar_arquivos_encontrados():
    xmls, docs_oc = listar_arquivos()

    print("\nXML encontrados:")
    if xmls:
        for arq in xmls:
            print(f"- {arq.name} ({arq.stat().st_size} bytes)")
    else:
        print("- Nenhum XML encontrado")

    print("\nOCs DOC/TXT encontradas:")
    if docs_oc:
        for arq in docs_oc:
            print(f"- {arq.name} ({arq.stat().st_size} bytes)")
    else:
        print("- Nenhum DOC/TXT de OC encontrado")

# =========================================================
# CONFERÊNCIA DE CABEÇALHO
# =========================================================
def conferir_prazo(data_emissao_str: str, prazo_entrega_str: str):
    de = parse_data(data_emissao_str)
    pe = parse_data(prazo_entrega_str)

    if not de or not pe:
        return {"data_limite_7_dias": "", "status_prazo": "NAO_AVALIADO"}

    limite = de + timedelta(days=7)
    status = "EM_DIA" if pe <= limite else "FORA_DO_PRAZO"
    return {
        "data_limite_7_dias": data_str(limite),
        "status_prazo": status,
    }

def conferir_cabecalho(xml_head: dict, oc_head: dict, df_itens_oc: pd.DataFrame):
    motivos = []

    # Regra principal: Doc Origem na NF
    doc_origem_oc = str(oc_head.get("doc_origem", "") or "")
    doc_origem_xml = str(xml_head.get("doc_origem_encontrado_xml", "") or "")
    xped_xml = str(xml_head.get("xped_xml", "") or "")

    doc_ok = bool(doc_origem_oc) and (
        doc_origem_oc == doc_origem_xml
        or doc_origem_oc == xped_xml
        or doc_origem_oc in str(xml_head.get("observacao_nota", ""))
    )
    if not doc_ok:
        motivos.append("DOC_ORIGEM_NAO_REFERENCIADO")

    # Fornecedor: raiz do CNPJ e nome normalizado/alias
    forn_raiz_ok = raiz_cnpj(oc_head.get("fornecedor_cnpj", "")) == raiz_cnpj(xml_head.get("emitente_cnpj", ""))
    forn_nome_ok = comparar_nome_fornecedor(
        oc_head.get("fornecedor_nome", ""),
        xml_head.get("emitente_nome", "")
    )

    fornecedor_ok = forn_raiz_ok or forn_nome_ok
    if not fornecedor_ok:
        motivos.append("FORNECEDOR_DIVERGENTE")

    # Cliente / destinatário
    cliente_ok = comparar_texto_simples(
        oc_head.get("cliente_nome", ""),
        xml_head.get("destinatario_nome", "")
    )
    cliente_cnpj_ok = raiz_cnpj(oc_head.get("cliente_cnpj", "")) == raiz_cnpj(xml_head.get("destinatario_cnpj", ""))
    if not (cliente_ok or cliente_cnpj_ok):
        motivos.append("CLIENTE_DIVERGENTE")

    # Endereço (informativo, não mata cabeçalho sozinho)
    endereco_ok = comparar_texto_simples(
        oc_head.get("endereco_entrega", ""),
        xml_head.get("endereco_destinatario", "")
    )

    # Valor: NF parcial é permitida, então o total da NF deve ser <= total da OC
    total_oc = numero(oc_head.get("valor_total_pedido", 0.0))
    total_nf = numero(xml_head.get("valor_total_nota", 0.0))
    valor_total_ok = (total_oc <= 0) or (total_nf <= total_oc + 0.05)
    if not valor_total_ok:
        motivos.append("VALOR_TOTAL_ACIMA_DA_OC")

    prazo = conferir_prazo(xml_head.get("data_emissao", ""), oc_head.get("prazo_entrega", ""))

    status_geral = "APROVADA_CABECALHO" if not any(m in motivos for m in ["DOC_ORIGEM_NAO_REFERENCIADO", "FORNECEDOR_DIVERGENTE", "CLIENTE_DIVERGENTE", "VALOR_TOTAL_ACIMA_DA_OC"]) else "DIVERGENTE_CABECALHO"

    return {
        **xml_head,
        **oc_head,
        "doc_origem_referenciado_na_obs": "SIM" if doc_ok else "NAO",
        "fornecedor_ok": "SIM" if fornecedor_ok else "NAO",
        "cliente_ok": "SIM" if (cliente_ok or cliente_cnpj_ok) else "NAO",
        "endereco_ok": "SIM" if endereco_ok else "NAO",
        "valor_total_ok": "SIM" if valor_total_ok else "NAO",
        "data_limite_7_dias": prazo["data_limite_7_dias"],
        "status_prazo": prazo["status_prazo"],
        "motivo_cabecalho": "|".join(motivos),
        "status_geral_cabecalho": status_geral,
    }

# =========================================================
# PROCESSAMENTO PRINCIPAL
# =========================================================
def processar():
    garantir_pastas()
    mostrar_arquivos_encontrados()

    cab_xml_rows = []
    itens_xml_rows = []
    cab_oc_rows = []
    itens_oc_rows = []
    conferencias_rows = []
    conferencia_itens_rows = []
    programacao_rows = []
    logs = []

    pares = montar_pares()

    if not pares:
        print("\nNenhum arquivo encontrado para processar.")
        return

    print("\nIniciando processamento do Planejamento de Recebimento...\n")

    for idx, par in enumerate(pares, start=1):
        arq_xml = par["xml"]
        arq_oc_doc = par["oc_doc"]

        print(f"--- Par {idx} ---")
        print("XML   :", arq_xml.name if arq_xml else "Nenhum")
        print("OC DOC:", arq_oc_doc.name if arq_oc_doc else "Nenhum")

        log = {
            "arquivo_xml": arq_xml.name if arq_xml else "",
            "arquivo_oc_doc": arq_oc_doc.name if arq_oc_doc else "",
            "status_xml": "",
            "status_oc_doc_cabecalho": "",
            "status_oc_doc_itens": "",
            "status_conferencia_cabecalho": "",
            "status_conferencia_itens": "",
            "observacao": "",
        }

        xml_head = None
        oc_head = None
        itens_xml = []
        df_itens_oc = pd.DataFrame()

        # XML
        if arq_xml is not None:
            try:
                xml_head, itens_xml = extrair_xml(arq_xml)
                cab_xml_rows.append(xml_head)
                itens_xml_rows.extend(itens_xml)
                log["status_xml"] = "OK"
                print(f"XML processado | NF {xml_head.get('numero_nota', '')} | itens {len(itens_xml)}")
            except Exception as e:
                log["status_xml"] = "ERRO"
                log["observacao"] += f"Erro XML: {e} | "
                print(f"Erro XML: {e}")
        else:
            log["status_xml"] = "NAO_ENCONTRADO"
            log["observacao"] += "Sem XML pareado | "

        # OC
        if arq_oc_doc is not None:
            try:
                oc_head = extrair_cabecalho_oc_doc(arq_oc_doc)
                cab_oc_rows.append(oc_head)
                log["status_oc_doc_cabecalho"] = "OK"
                print(
                    f"Cabecalho OC processado | OC {oc_head.get('numero_oc', '')} | "
                    f"Doc. Origem {oc_head.get('doc_origem', '')} | "
                    f"Fornecedor {oc_head.get('fornecedor_nome', '')}"
                )
            except Exception as e:
                log["status_oc_doc_cabecalho"] = "ERRO"
                log["observacao"] += f"Erro cabecalho OC DOC: {e} | "
                print(f"Erro cabecalho OC DOC: {e}")

            try:
                df_itens_oc = extrair_itens_oc_doc(arq_oc_doc)
                if not df_itens_oc.empty:
                    itens_oc_rows.extend(df_itens_oc.to_dict(orient="records"))
                log["status_oc_doc_itens"] = "OK" if not df_itens_oc.empty else "SEM_ITENS_EXTRAIDOS"
                print(f"Itens OC processados | itens {len(df_itens_oc)}")
            except Exception as e:
                log["status_oc_doc_itens"] = "ERRO"
                log["observacao"] += f"Erro itens OC DOC: {e} | "
                print(f"Erro itens OC DOC: {e}")
        else:
            log["status_oc_doc_cabecalho"] = "NAO_ENCONTRADO"
            log["status_oc_doc_itens"] = "NAO_ENCONTRADO"
            log["observacao"] += "Sem DOC/TXT de OC pareado | "

        # Conferência
        if xml_head is not None and oc_head is not None:
            try:
                conf = conferir_cabecalho(xml_head, oc_head, df_itens_oc)
                conferencias_rows.append(conf)
                log["status_conferencia_cabecalho"] = "OK"

                rel_itens, itens_ok, itens_erro = comparar_itens(itens_xml, df_itens_oc)
                status_itens = "OK" if itens_erro == 0 else "ERRO"
                log["status_conferencia_itens"] = status_itens

                for row in rel_itens:
                    conferencia_itens_rows.append({
                        "arquivo_xml": arq_xml.name,
                        "arquivo_oc_doc": arq_oc_doc.name,
                        "numero_nota": xml_head.get("numero_nota", ""),
                        "numero_oc": oc_head.get("numero_oc", ""),
                        **row
                    })

                motivos = []
                if conf["status_geral_cabecalho"] != "APROVADA_CABECALHO":
                    if conf.get("motivo_cabecalho"):
                        motivos.append(conf["motivo_cabecalho"])
                    else:
                        motivos.append("ERRO_CABECALHO")

                if status_itens != "OK":
                    motivos.append("ERRO_ITENS")

                status_final = "OK" if not motivos else "ERRO"

                data_programada = xml_head.get("data_programada_xml", "") or oc_head.get("prazo_entrega", "")

                programacao_rows.append({
                    "id": idx,
                    "data_programada": data_programada,
                    "numero_nota": xml_head.get("numero_nota", ""),
                    "numero_oc": oc_head.get("numero_oc", ""),
                    "doc_origem": oc_head.get("doc_origem", ""),
                    "fornecedor": xml_head.get("emitente_nome", ""),
                    "qtd_itens": len(consolidar_itens_xml(itens_xml)),
                    "itens_ok": itens_ok,
                    "itens_erro": itens_erro,
                    "valor_total": xml_head.get("valor_total_nota", 0),
                    "status": status_final,
                    "status_itens": status_itens,
                    "status_cabecalho": conf.get("status_geral_cabecalho", ""),
                    "motivo": "|".join([m for m in motivos if m]),
                    "responsavel": "OK" if status_final == "OK" else "EMPRESTIMO",
                })

                print(
                    f"Conferencia | Cabeçalho: {conf['status_geral_cabecalho']} | "
                    f"Itens: {status_itens} | Final: {status_final}"
                )

            except Exception as e:
                log["status_conferencia_cabecalho"] = "ERRO"
                log["status_conferencia_itens"] = "ERRO"
                log["observacao"] += f"Erro conferencia: {e} | "
                print(f"Erro conferencia: {e}")
        else:
            log["status_conferencia_cabecalho"] = "NAO_REALIZADA"
            log["status_conferencia_itens"] = "NAO_REALIZADA"

        print()
        logs.append(log)

    # DataFrames
    df_cab_xml = pd.DataFrame(cab_xml_rows)
    df_itens_xml = pd.DataFrame(itens_xml_rows)
    df_cab_oc = pd.DataFrame(cab_oc_rows)
    df_itens_oc = pd.DataFrame(itens_oc_rows)
    df_conf = pd.DataFrame(conferencias_rows)
    df_conf_itens = pd.DataFrame(conferencia_itens_rows)
    df_prog = pd.DataFrame(programacao_rows)
    df_logs = pd.DataFrame(logs)

    # Excel
    with pd.ExcelWriter(ARQUIVO_RESULTADO, engine="openpyxl") as writer:
        (df_cab_xml if not df_cab_xml.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="xml_cabecalho", index=False)
        (df_itens_xml if not df_itens_xml.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="xml_itens", index=False)
        (df_cab_oc if not df_cab_oc.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="oc_doc_cabecalho", index=False)
        (df_itens_oc if not df_itens_oc.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="oc_doc_itens", index=False)
        (df_conf if not df_conf.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="conferencia_cabecalho", index=False)
        (df_conf_itens if not df_conf_itens.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="conferencia_itens", index=False)
        (df_prog if not df_prog.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="programacao", index=False)
        (df_logs if not df_logs.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="log_processamento", index=False)

    # CSVs
    df_cab_xml.to_csv(PASTA_RESULTADO / "xml_cabecalho.csv", index=False, sep=";", encoding="utf-8-sig")
    df_itens_xml.to_csv(PASTA_RESULTADO / "xml_itens.csv", index=False, sep=";", encoding="utf-8-sig")
    df_cab_oc.to_csv(PASTA_RESULTADO / "oc_doc_cabecalho.csv", index=False, sep=";", encoding="utf-8-sig")
    df_itens_oc.to_csv(PASTA_RESULTADO / "oc_doc_itens.csv", index=False, sep=";", encoding="utf-8-sig")
    df_conf.to_csv(PASTA_RESULTADO / "conferencia_cabecalho.csv", index=False, sep=";", encoding="utf-8-sig")
    df_conf_itens.to_csv(PASTA_RESULTADO / "conferencia_itens.csv", index=False, sep=";", encoding="utf-8-sig")
    df_prog.to_csv(PASTA_RESULTADO / "programacao_recebimento.csv", index=False, sep=";", encoding="utf-8-sig")
    df_logs.to_csv(PASTA_RESULTADO / "log_processamento.csv", index=False, sep=";", encoding="utf-8-sig")

    print("PRONTO - Sistema refatorado e arquivos gerados")
    print("Arquivos principais:")
    print("- resultado/conferencia_cabecalho.csv")
    print("- resultado/conferencia_itens.csv")
    print("- resultado/programacao_recebimento.csv")
    print("- resultado/log_processamento.csv")

if __name__ == "__main__":
    processar()