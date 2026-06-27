import re
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import xml.etree.ElementTree as ET

# =========================================================
# MVP CONFERÊNCIA NF-e x OC DOC/TXT
# Parte 1:
# - Extrair cabeçalho XML
# - Extrair itens XML
# - Extrair cabeçalho OC em DOC/TXT
# - Extrair itens OC em DOC/TXT
# - Conferir cabeçalho
# Parte 2 futura:
# - Conferência item a item
# - Fator / unidade / embalagem
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
            # O último separador indica o decimal.
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


def extrair_doc_origem_da_obs(observacao: str) -> str:
    """
    Extrai o Doc. Origem da observação/dados adicionais da NF-e.
    Regra do MVP: validar Doc. Origem, não número da OC/pedido.
    """
    if not observacao:
        return ""

    padroes = [
        r"Doc\.?\s*Origem\s*[:\-]?\s*([0-9]+)",
        r"Doc\.?\s*Orig\.?\s*[:\-]?\s*([0-9]+)",
        r"Pedido\s+interno\s*:\s*([0-9]+)",
        r"Pedido\s*[:\-]?\s*([0-9]+)",
    ]

    for p in padroes:
        m = re.search(p, observacao, flags=re.IGNORECASE)
        if m:
            return m.group(1)
    return ""


# =========================================================
# EXTRAÇÃO XML NF-e
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

    data_emissao = parse_data(data_emissao_raw)
    data_saida = parse_data(data_saida_raw)

    cabecalho = {
        "arquivo_xml": caminho_xml.name,
        "emitente_nome": emitente_nome,
        "emitente_cnpj": emitente_cnpj,
        "numero_nota": numero_nota,
        "natureza_operacao": natureza_operacao,
        "destinatario_nome": destinatario_nome,
        "destinatario_cnpj": destinatario_cnpj,
        "endereco_destinatario": destinatario_endereco,
        "numero_endereco_destinatario": destinatario_numero,
        "bairro_destinatario": destinatario_bairro,
        "municipio_destinatario": destinatario_municipio,
        "uf_destinatario": destinatario_uf,
        "cep_destinatario": destinatario_cep,
        "data_emissao": data_str(data_emissao),
        "data_saida_expedicao": data_str(data_saida),
        "valor_total_produtos": valor_total_produtos,
        "valor_total_nota": valor_total_nota,
        "transportadora_nome": transportadora_nome,
        "transportadora_cnpj": transportadora_cnpj,
        "frete_por_conta": frete_por_conta,
        "observacao_nota": observacao_nota,
        "doc_origem_encontrado_xml": doc_origem_encontrado_xml,
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
    """
    Limpa caracteres invisíveis comuns em exportações .doc/.txt.
    """
    if not texto:
        return ""

    texto = texto.replace("\x00", "")
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")

    # Remove controles estranhos, preservando tab e quebra de linha
    texto = "".join(
        ch for ch in texto
        if ch == "\n" or ch == "\t" or ord(ch) >= 32
    )

    return texto


def limpar_texto_extraido(texto: str) -> str:
    """
    Limpa caracteres invisíveis comuns em exportações .doc/.txt.
    Preserva tabulação e quebra de linha.
    """
    if not texto:
        return ""

    texto = texto.replace("\x00", "")
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")

    texto = "".join(
        ch for ch in texto
        if ch == "\n" or ch == "\t" or ord(ch) >= 32
    )

    return texto


def score_texto_oc(texto: str) -> int:
    """
    Dá nota para o texto lido.
    A melhor leitura é a que contém palavras-chave reais da OC.
    """
    t = texto.upper()

    palavras = [
        "ORDEM DE COMPRA",
        "DOC. ORIGEM",
        "PEDIDO:",
        "FORNECEDOR",
        "CNPJ",
        "PRAZO DE ENTREGA",
        "VALOR DO PEDIDO",
        "REF.MAT",
        "UN.COMP",
        "QTD.PEDIDA",
    ]

    score = 0

    for p in palavras:
        if p in t:
            score += 10

    # Também valoriza textos maiores, mas não deixa tamanho vencer conteúdo ruim.
    score += min(len(texto) // 500, 10)

    return score


def extrair_strings_de_binario(bruto: bytes) -> str:
    """
    Fallback para .doc antigo/binário.
    Extrai sequências imprimíveis, parecido com comando 'strings'.
    """
    partes = []
    atual = []

    for b in bruto:
        ch = chr(b)

        if ch == "\t" or ch == "\n" or ch == "\r" or (32 <= b <= 126) or (160 <= b <= 255):
            atual.append(ch)
        else:
            if len(atual) >= 4:
                partes.append("".join(atual))
            atual = []

    if len(atual) >= 4:
        partes.append("".join(atual))

    return "\n".join(partes)


def limpar_texto_extraido(texto: str) -> str:
    """
    Limpa caracteres invisíveis e normaliza quebras de linha.
    """
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
    """
    Usa o Microsoft Word instalado no Windows para converter .doc/.docx em texto.
    Isso resolve .doc binário antigo, que read_text() não consegue ler corretamente.
    """
    try:
        import win32com.client
    except Exception as e:
        raise ValueError(
            "pywin32 não está instalado. Rode: py -m pip install pywin32"
        ) from e

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

        # FileFormat=2 => wdFormatText
        doc.SaveAs2(caminho_txt_abs, FileFormat=2)
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
    """
    Lê OC em .doc/.docx/.txt/.csv.

    Prioridade:
    1. Se for .doc ou .docx, tenta converter via Microsoft Word.
    2. Se for texto puro, tenta encodings comuns.
    3. Salva debug do texto lido para validação.
    """
    if arquivo_vazio(caminho_doc):
        raise ValueError("Arquivo DOC/TXT da OC vazio ou inexistente")

    ext = caminho_doc.suffix.lower()

    texto = ""

    # Melhor caminho para Word real
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
        (PASTA_RESULTADO / f"debug_texto_lido_{caminho_doc.stem}.txt").write_text(
            texto,
            encoding="utf-8"
        )
    except Exception:
        pass

    return texto

def extrair_cabecalho_oc_doc(caminho_doc: Path):
    texto = ler_doc_texto(caminho_doc)
    texto_flat = normalize_spaces(texto)

    try:
        (PASTA_RESULTADO / f"debug_texto_{caminho_doc.stem}.txt").write_text(
            texto,
            encoding="utf-8"
        )
    except Exception:
        pass

    numero_oc = buscar_regex(
        texto_flat,
        r"ORDEM\s+DE\s+COMPRA\s+(\d+)",
        default=""
    )

    if not numero_oc:
        numero_oc = buscar_regex(
            texto_flat,
            r"Pedido\s*:\s*(\d+)",
            default=""
        )

    doc_origem = buscar_regex(
        texto_flat,
        r"Doc\.?\s*Origem\s*:\s*(\d+)",
        default=""
    )

    pedido = buscar_regex(
        texto_flat,
        r"Pedido\s*:\s*(\d+)",
        default=""
    )

    data_pedido = buscar_regex(
        texto_flat,
        r"Data\s*:\s*(\d{2}/\d{2}/\d{4})",
        default=""
    )

    prazo_entrega = buscar_regex(
        texto_flat,
        r"PRAZO\s+DE\s+ENTREGA\s*:\s*(\d{2}/\d{2}/\d{4})",
        default=""
    )

    cliente_cnpj = limpar_cnpj(
        buscar_regex(
            texto_flat,
            r"C\.N\.P\.J\.?\s*:\s*([0-9\.\-/]+)",
            default=""
        )
    )

    cliente_nome = buscar_regex(
        texto_flat,
        r"Doc\.?\s*Origem\s*:\s*\d+\s+(.+?)\s+Pedido\s*:",
        default=""
    )

    if not cliente_nome:
        cliente_nome = buscar_regex(
            texto_flat,
            r"(REDE\s+D['’]OR\s+SAO\s+LUIZ\s+S\.A\.\s*-\s*HOSPITAL\s+ESPERANCA\s+OLINDA)",
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

    if not fornecedor_nome:
        fornecedor_nome = buscar_regex(
            texto_flat,
            r"\d+\s*-\s*(SUZANO\s+PAPEL\s+E\s+CELULOSE\s+S\.A\.)",
            default=""
        )

    fornecedor_cnpj = limpar_cnpj(
        buscar_regex(
            texto_flat,
            r"CNPJ\s*:\s*([0-9\.\-/]+)",
            default=""
        )
    )

    valor_total_pedido = numero(
        buscar_regex(
            texto_flat,
            r"Valor\s+do\s+Pedido\s*:\s*([0-9\.,]+)",
            default=""
        )
    )

    return {
        "arquivo_oc_doc": caminho_doc.name,
        "numero_oc": numero_oc,
        "pedido": pedido,
        "doc_origem": doc_origem,
        "data_pedido": data_pedido,
        "prazo_entrega": prazo_entrega,
        "cliente_nome": cliente_nome,
        "cliente_cnpj": cliente_cnpj,
        "endereco_principal": endereco_principal,
        "endereco_entrega": endereco_entrega,
        "fornecedor_nome": fornecedor_nome,
        "fornecedor_cnpj": fornecedor_cnpj,
        "valor_total_pedido": valor_total_pedido,
        "texto_extraido_doc_preview": texto_flat[:500],
    }

def _split_tabs_linha(linha: str):
    return [p.strip() for p in linha.split("\t") if p.strip()]


def _extrair_moedas_compactadas(texto: str):
    """
    Ex.: '19.390.000.00' -> ['19.39', '0.00', '0.00']
    """
    return re.findall(r"\d+[\.,]\d{2}", str(texto or ""))


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

        # Formato tabulado esperado:
        # 035990 | PAPEL COPIA | UNIDADE | 1 | 300.00 | 19.390.000.00 | 5,817.00
        if partes and re.fullmatch(r"\d{5,6}", partes[0]):
            if item_atual:
                registros.append(item_atual)

            codigo = partes[0]
            descricao = partes[1] if len(partes) > 1 else ""
            unidade = partes[2] if len(partes) > 2 else ""
            embalagem = numero(partes[3]) if len(partes) > 3 else 0.0
            quantidade = numero(partes[4]) if len(partes) > 4 else 0.0

            valor_unitario = 0.0
            desconto = 0.0
            ipi = 0.0
            valor_total = 0.0

            if len(partes) > 5:
                moedas = re.findall(r"\d+[.,]\d{2}", partes[5])

                if len(moedas) >= 1:
                    valor_unitario = numero(moedas[0])
                if len(moedas) >= 2:
                    desconto = numero(moedas[1])
                if len(moedas) >= 3:
                    ipi = numero(moedas[2])

            if len(partes) > 6:
                valor_total = numero(partes[6])

            item_atual = {
                "arquivo_oc_doc": caminho_doc.name,
                "codigo_item_oc": codigo,
                "descricao_item_oc": descricao,
                "unidade_item_oc": unidade,
                "embalagem_item_oc": embalagem,
                "quantidade_item_oc": quantidade,
                "valor_unitario_item_oc": valor_unitario,
                "desconto_item_oc": desconto,
                "ipi_item_oc": ipi,
                "valor_total_item_oc": valor_total,
                "metodo_extracao": "doc_tabulado",
            }

            continue

        # Continuação da descrição do item
        if item_atual:
            texto_linha = limpa.upper()

            fim_item = any(x in texto_linha for x in [
                "VALOR DO PEDIDO",
                "CONDIÇÕES DE PAGAMENTO",
                "SETOR SOLICITANTE",
                "DGS BRASIL",
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

    # Fallback se tabulação não veio preservada
    if not registros:
        texto_flat = normalize_spaces(texto)

        padrao = re.compile(
            r"(?P<codigo>\d{5,6})\s+"
            r"(?P<descricao>[A-Z0-9ÇÃÕÁÉÍÓÚÂÊÔÀÜ/\- ]+?)\s+"
            r"(?P<unidade>UNIDADE|CAIXA|PACOTE|UN|CX|PC)\s+"
            r"(?P<embalagem>\d+)\s+"
            r"(?P<quantidade>\d+[.,]\d{2})\s+"
            r"(?P<compactado>\d+[.,]\d{2}(?:\d+[.,]\d{2}){0,2})\s+"
            r"(?P<valor_total>\d{1,3}(?:,\d{3})*\.\d{2}|\d+[.,]\d{2})",
            flags=re.I
        )

        for m in padrao.finditer(texto_flat):
            moedas = re.findall(r"\d+[.,]\d{2}", m.group("compactado"))

            valor_unitario = numero(moedas[0]) if len(moedas) >= 1 else 0.0
            desconto = numero(moedas[1]) if len(moedas) >= 2 else 0.0
            ipi = numero(moedas[2]) if len(moedas) >= 3 else 0.0

            registros.append({
                "arquivo_oc_doc": caminho_doc.name,
                "codigo_item_oc": m.group("codigo"),
                "descricao_item_oc": normalize_spaces(m.group("descricao")),
                "unidade_item_oc": m.group("unidade"),
                "embalagem_item_oc": numero(m.group("embalagem")),
                "quantidade_item_oc": numero(m.group("quantidade")),
                "valor_unitario_item_oc": valor_unitario,
                "desconto_item_oc": desconto,
                "ipi_item_oc": ipi,
                "valor_total_item_oc": numero(m.group("valor_total")),
                "metodo_extracao": "doc_regex_fallback",
            })

    if not registros:
        try:
            (PASTA_RESULTADO / f"debug_itens_{caminho_doc.stem}.txt").write_text(
                texto,
                encoding="utf-8"
            )
        except Exception:
            pass

    return pd.DataFrame(registros)

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

    print("\n📂 XML encontrados:")
    if xmls:
        for arq in xmls:
            print(f"- {arq.name} ({arq.stat().st_size} bytes)")
    else:
        print("- Nenhum XML encontrado")

    print("\n📂 OCs DOC/TXT encontradas:")
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
    total_pedido_oc = numero(oc_head.get("valor_total_pedido", 0.0))

    if (not total_pedido_oc) and df_itens_oc is not None and not df_itens_oc.empty:
        total_pedido_oc = float(df_itens_oc["valor_total_item_oc"].fillna(0).sum())

    obs = xml_head.get("observacao_nota", "")
    doc_origem_oc = str(oc_head.get("doc_origem", "") or "")
    doc_origem_xml = str(xml_head.get("doc_origem_encontrado_xml", "") or "")

    doc_ok = False

    if obs and doc_origem_oc:
        doc_ok = doc_origem_oc in obs or doc_origem_oc.lstrip("0") in obs

    if not doc_ok and doc_origem_xml and doc_origem_oc:
        doc_ok = (
            doc_origem_xml == doc_origem_oc
            or doc_origem_xml.lstrip("0") == doc_origem_oc.lstrip("0")
        )

    fornecedor_ok = comparar_texto_simples(
        oc_head.get("fornecedor_nome", ""),
        xml_head.get("emitente_nome", "")
    )

    cnpj_fornecedor_ok = (
        limpar_cnpj(oc_head.get("fornecedor_cnpj", ""))
        == limpar_cnpj(xml_head.get("emitente_cnpj", ""))
        and bool(oc_head.get("fornecedor_cnpj", ""))
    )

    cliente_ok = comparar_texto_simples(
        oc_head.get("cliente_nome", ""),
        xml_head.get("destinatario_nome", "")
    )

    cnpj_cliente_ok = (
        limpar_cnpj(oc_head.get("cliente_cnpj", ""))
        == limpar_cnpj(xml_head.get("destinatario_cnpj", ""))
        and bool(oc_head.get("cliente_cnpj", ""))
    )

    endereco_ok = comparar_texto_simples(
        oc_head.get("endereco_entrega", ""),
        xml_head.get("endereco_destinatario", "")
    )

    valor_total_nota = numero(xml_head.get("valor_total_nota", 0.0))
    valor_total_ok = abs(numero(total_pedido_oc) - valor_total_nota) <= 0.01 if total_pedido_oc else False

    prazo = conferir_prazo(
        xml_head.get("data_emissao", ""),
        oc_head.get("prazo_entrega", "")
    )

    checks = {
        "doc_origem_referenciado_na_obs": "SIM" if doc_ok else "NAO",
        "fornecedor_ok": "SIM" if fornecedor_ok else "NAO",
        "cnpj_fornecedor_ok": "SIM" if cnpj_fornecedor_ok else "NAO",
        "cliente_ok": "SIM" if cliente_ok else "NAO",
        "cnpj_cliente_ok": "SIM" if cnpj_cliente_ok else "NAO",
        "endereco_ok": "SIM" if endereco_ok else "NAO",
        "valor_total_ok": "SIM" if valor_total_ok else "NAO",
        "data_limite_7_dias": prazo["data_limite_7_dias"],
        "status_prazo": prazo["status_prazo"],
    }

    essenciais = [
        checks["doc_origem_referenciado_na_obs"],
        checks["fornecedor_ok"],
        checks["cnpj_fornecedor_ok"],
        checks["valor_total_ok"],
    ]

    if all(v == "SIM" for v in essenciais) and checks["status_prazo"] in {"EM_DIA", "NAO_AVALIADO"}:
        status_geral = "APROVADA_CABECALHO"
    elif checks["doc_origem_referenciado_na_obs"] == "NAO":
        status_geral = "REJEITADA_DOC_ORIGEM_NAO_REFERENCIADO"
    else:
        status_geral = "DIVERGENTE_CABECALHO"

    return {
        **xml_head,
        **oc_head,
        "valor_total_pedido_calculado_ou_cabecalho": total_pedido_oc,
        **checks,
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
    logs = []

    pares = montar_pares()

    if not pares:
        print("\n⚠ Nenhum arquivo encontrado para processar.")
        return

    print("\n🚀 Iniciando processamento do MVP DOC/TXT...\n")

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
            "observacao": "",
        }

        xml_head = None
        oc_head = None
        df_itens_oc = pd.DataFrame()

        if arq_xml is not None:
            try:
                xml_head, itens_xml = extrair_xml(arq_xml)
                cab_xml_rows.append(xml_head)
                itens_xml_rows.extend(itens_xml)
                log["status_xml"] = "OK"
                print(f"✅ XML processado | NF {xml_head.get('numero_nota', '')} | itens {len(itens_xml)}")
            except Exception as e:
                log["status_xml"] = "ERRO"
                log["observacao"] += f"Erro XML: {e} | "
                print(f"❌ Erro XML: {e}")
        else:
            log["status_xml"] = "NAO_ENCONTRADO"
            log["observacao"] += "Sem XML pareado | "
            print("⚠ Sem XML pareado")

        if arq_oc_doc is not None:
            try:
                oc_head = extrair_cabecalho_oc_doc(arq_oc_doc)
                cab_oc_rows.append(oc_head)
                log["status_oc_doc_cabecalho"] = "OK"
                print(
                    f"✅ Cabeçalho OC DOC processado | OC {oc_head.get('numero_oc', '')} | "
                    f"Doc. Origem {oc_head.get('doc_origem', '')} | "
                    f"fornecedor {oc_head.get('fornecedor_nome', '')}"
                )
            except Exception as e:
                log["status_oc_doc_cabecalho"] = "ERRO"
                log["observacao"] += f"Erro cabeçalho OC DOC: {e} | "
                print(f"❌ Erro cabeçalho OC DOC: {e}")

            try:
                df_itens_oc = extrair_itens_oc_doc(arq_oc_doc)
                if not df_itens_oc.empty:
                    itens_oc_rows.extend(df_itens_oc.to_dict(orient="records"))
                log["status_oc_doc_itens"] = "OK" if not df_itens_oc.empty else "SEM_ITENS_EXTRAIDOS"
                print(f"✅ Itens OC DOC processados | itens {len(df_itens_oc)}")
            except Exception as e:
                log["status_oc_doc_itens"] = "ERRO"
                log["observacao"] += f"Erro itens OC DOC: {e} | "
                print(f"❌ Erro itens OC DOC: {e}")
        else:
            log["status_oc_doc_cabecalho"] = "NAO_ENCONTRADO"
            log["status_oc_doc_itens"] = "NAO_ENCONTRADO"
            log["observacao"] += "Sem DOC/TXT de OC pareado | "
            print("⚠ Sem DOC/TXT de OC pareado")

        if xml_head is not None and oc_head is not None:
            try:
                conf = conferir_cabecalho(xml_head, oc_head, df_itens_oc)
                conferencias_rows.append(conf)
                log["status_conferencia_cabecalho"] = "OK"
                print(f"✅ Conferência cabeçalho | status: {conf['status_geral_cabecalho']}")
            except Exception as e:
                log["status_conferencia_cabecalho"] = "ERRO"
                log["observacao"] += f"Erro conferência cabeçalho: {e} | "
                print(f"❌ Erro conferência cabeçalho: {e}")
        else:
            log["status_conferencia_cabecalho"] = "NAO_REALIZADA"
            print("⚠ Conferência não realizada por falta de XML ou OC DOC")

        print()
        logs.append(log)

    df_cab_xml = pd.DataFrame(cab_xml_rows)
    df_itens_xml = pd.DataFrame(itens_xml_rows)
    df_cab_oc = pd.DataFrame(cab_oc_rows)
    df_itens_oc = pd.DataFrame(itens_oc_rows)
    df_conf = pd.DataFrame(conferencias_rows)
    df_logs = pd.DataFrame(logs)

    with pd.ExcelWriter(ARQUIVO_RESULTADO, engine="openpyxl") as writer:
        (df_cab_xml if not df_cab_xml.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="xml_cabecalho", index=False)
        (df_itens_xml if not df_itens_xml.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="xml_itens", index=False)
        (df_cab_oc if not df_cab_oc.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="oc_doc_cabecalho", index=False)
        (df_itens_oc if not df_itens_oc.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="oc_doc_itens", index=False)
        (df_conf if not df_conf.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="conferencia_cabecalho", index=False)
        (df_logs if not df_logs.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="log_processamento", index=False)

    try:
        xls = pd.ExcelFile(ARQUIVO_RESULTADO, engine="openpyxl")
        print(f"✅ Excel gerado e validado: {ARQUIVO_RESULTADO}")
        print("Abas:", xls.sheet_names)
    except Exception as e:
        print(f"❌ Falha ao validar XLSX gerado: {e}")

    df_cab_xml.to_csv(PASTA_RESULTADO / "xml_cabecalho.csv", index=False, sep=";", encoding="utf-8-sig")
    df_itens_xml.to_csv(PASTA_RESULTADO / "xml_itens.csv", index=False, sep=";", encoding="utf-8-sig")
    df_cab_oc.to_csv(PASTA_RESULTADO / "oc_doc_cabecalho.csv", index=False, sep=";", encoding="utf-8-sig")
    df_itens_oc.to_csv(PASTA_RESULTADO / "oc_doc_itens.csv", index=False, sep=";", encoding="utf-8-sig")
    df_conf.to_csv(PASTA_RESULTADO / "conferencia_cabecalho.csv", index=False, sep=";", encoding="utf-8-sig")
    df_logs.to_csv(PASTA_RESULTADO / "log_processamento.csv", index=False, sep=";", encoding="utf-8-sig")

    print("\n✅ CSVs gerados para validação manual.")
    print("- resultado/xml_cabecalho.csv")
    print("- resultado/xml_itens.csv")
    print("- resultado/oc_doc_cabecalho.csv")
    print("- resultado/oc_doc_itens.csv")
    print("- resultado/conferencia_cabecalho.csv")
    print("- resultado/log_processamento.csv")

    print("\n📌 Resumo do MVP DOC/TXT")
    print(f"- Cabeçalhos XML: {len(df_cab_xml)}")
    print(f"- Itens XML: {len(df_itens_xml)}")
    print(f"- Cabeçalhos OC DOC: {len(df_cab_oc)}")
    print(f"- Itens OC DOC: {len(df_itens_oc)}")
    print(f"- Conferências de cabeçalho: {len(df_conf)}")
    print(f"- Logs: {len(df_logs)}")


if __name__ == "__main__":
    processar()
