"""
Motor de Planejamento de Recebimento
=====================================
Compara NF-e (XML) com Ordem de Compra (DOC/TXT/CSV) e classifica cada
nota em OK / EMPRESTIMO / RECUSADA, seguindo a lógica de reconciliação
descrita na especificação do sistema.

Principais melhorias desta revisão (foco em comparação de itens):
- Normalização de texto corrigida (acentos deixavam de ser removidos
  corretamente e quebravam palavras, ex.: "ATENÇÃO" -> "ATENO").
- Casamento de itens XML x OC reescrito para buscar, em cada etapa de
  prioridade, o MELHOR candidato entre TODOS os itens da OC (antes o
  código verificava as 5 regras por linha da OC, na ordem da planilha,
  o que podia casar um item errado só porque veio primeiro na lista).
- Match por descrição agora usa um score contínuo (similaridade de
  tokens relevantes + similaridade de sequência), com stopwords de
  itens (UN, CX, PARA, DE...) e limiar mínimo, reduzindo falso-positivo
  e falso-negativo.
- Fator de embalagem (2X8, 10X50...) deixou de SOBRESCREVER a
  quantidade fiscal do XML (isso podia distorcer a quantidade real);
  agora é apenas informativo e cruzado com o valor do item.
- Validação de consistência de cálculo do item (qtd x valor unitário
  ~= valor total) e de unidade divergente, previstas na especificação
  e que não existiam na versão anterior.
- Correção de bug de classificação: quando a NF não tem OC vinculada,
  a nota agora é corretamente lançada como RECUSADA na programação de
  recebimento (antes esse caso não gerava nenhuma linha de saída).
"""

import re
import sys
import unicodedata
from difflib import SequenceMatcher
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

TOL_QTD = 0.01
TOL_VALOR = 0.05

# Limiar mínimo de similaridade para aceitar um match por descrição
# inteligente (0 a 1). Abaixo disso o item é considerado não encontrado.
LIMIAR_SIMILARIDADE_DESC = 0.45
MIN_TOKENS_COMUNS = 2

# Palavras que aparecem com frequência em descrições de item e que,
# sozinhas, não indicam que dois itens são o mesmo produto.
STOPWORDS_DESC = {
    "DE", "DA", "DO", "DOS", "DAS", "PARA", "COM", "SEM", "E", "OU",
    "UN", "UND", "UNIDADE", "UNID", "EMB", "EMBALAGEM", "CX", "CAIXA",
    "PCT", "PC", "PCS", "KIT", "REF", "MOD", "MODELO", "TIPO", "COR",
    "N", "NO", "NA", "A", "O",
}

ALIASES_UNIDADE = {
    "UN": "UN", "UND": "UN", "UNID": "UN", "UNIDADE": "UN",
    "PC": "PC", "PCT": "PCT", "PCS": "PC",
    "CX": "CX", "CAIXA": "CX",
    "KG": "KG", "KGS": "KG",
    "G": "G", "GR": "G", "GRAMA": "G",
    "L": "L", "LT": "L", "LITRO": "L",
    "ML": "ML",
    "M": "M", "MT": "M", "METRO": "M",
    "FR": "FR", "FRASCO": "FR",
    "AMP": "AMP", "AMPOLA": "AMP",
    "CP": "CP", "CPR": "CP", "COMPRIMIDO": "CP",
    "CJ": "CJ", "CONJUNTO": "CJ",
    "PAR": "PAR",
    "RL": "RL", "ROLO": "RL",
}

# =========================================================
# UTILITÁRIOS DE TEXTO
# =========================================================
def garantir_pastas():
    PASTA_XML.mkdir(parents=True, exist_ok=True)
    PASTA_OC.mkdir(parents=True, exist_ok=True)
    PASTA_RESULTADO.mkdir(parents=True, exist_ok=True)


def arquivo_vazio(caminho: Path) -> bool:
    return (not caminho.exists()) or caminho.stat().st_size == 0


def normalize_spaces(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto or "")).strip()


def remover_acentos(texto: str) -> str:
    """Remove acentuação preservando a letra base (Á -> A, Ç -> C...).

    A versão anterior removia o caractere acentuado inteiro via regex,
    o que fragmentava palavras (ex.: 'ATENÇÃO' virava 'ATENO') e
    prejudicava comparações de nome de cliente/fornecedor e descrição
    de item.
    """
    texto = str(texto or "")
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalizar_texto(texto: str) -> str:
    """Normalização genérica (nomes, cliente, fornecedor, endereço)."""
    texto = remover_acentos(normalize_spaces(texto)).upper()
    texto = re.sub(r"[^A-Z0-9 ]", "", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_desc(desc: str) -> str:
    """Normalização de descrição de item, preservando separação de
    palavras (hífen vira espaço, não é apenas removido)."""
    texto = remover_acentos(normalize_spaces(desc)).upper()
    texto = texto.replace("-", " ")
    texto = re.sub(r"[^A-Z0-9 ]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_codigo(codigo: str) -> str:
    """Normalização de código de item: remove hífen, espaço e símbolo,
    mantendo apenas letras e números colados (MA05-MP -> MA05MP)."""
    texto = remover_acentos(str(codigo or "")).upper()
    texto = re.sub(r"[^A-Z0-9]", "", texto)
    return texto


def normalizar_unidade(unidade: str) -> str:
    chave = normalizar_codigo(unidade)
    return ALIASES_UNIDADE.get(chave, chave)


def limpar_cnpj(texto: str) -> str:
    return re.sub(r"\D", "", str(texto or ""))


def raiz_cnpj(cnpj: str) -> str:
    cnpj_limpo = limpar_cnpj(cnpj)
    return cnpj_limpo[:8] if len(cnpj_limpo) >= 8 else cnpj_limpo


def so_numeros(texto: str) -> str:
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
    aliases = [("MDR", "MEDERI")]
    for x, y in aliases:
        na = na.replace(x, y)
        nb = nb.replace(x, y)
    return na == nb or na in nb or nb in na


def normalizar_mod_frete(codigo):
    mapa = {
        "0": "Por conta do emitente",
        "1": "Por conta do destinatario/remetente",
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
# XML NF-e
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


def extrair_xml(caminho_xml: Path):
    if arquivo_vazio(caminho_xml):
        raise ValueError("Arquivo XML vazio ou inexistente")
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
    except ET.ParseError as e:
        raise ValueError(f"Arquivo XML invalido: {e}")

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
    texto = "".join(ch for ch in texto if ch == "\n" or ch == "\t" or ord(ch) >= 32)
    return texto


def converter_doc_word_para_txt(caminho_doc: Path) -> str:
    try:
        import win32com.client
    except Exception as e:
        raise ValueError("pywin32 nao esta instalado. Rode: py -m pip install pywin32") from e

    caminho_doc_abs = str(caminho_doc.resolve())
    caminho_txt = PASTA_RESULTADO / f"_convertido_{caminho_doc.stem}.txt"
    caminho_txt_abs = str(caminho_txt.resolve())

    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(caminho_doc_abs, ReadOnly=True, ConfirmConversions=False, AddToRecentFiles=False)
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
        raise ValueError(f"Nao foi possivel extrair texto util de {caminho_doc.name}")
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
    cliente_cnpj = limpar_cnpj(buscar_regex(texto_flat, r"C\.N\.P\.J\.?\s*:\s*([0-9\.\-/]+)", default=""))
    cliente_nome = buscar_regex(texto_flat, r"Doc\.?\s*Origem\s*:\s*\d+\s+(.+?)\s+Pedido\s*:", default="")
    endereco_principal = buscar_regex(texto_flat, r"Pedido\s*:\s*\d+\s+(.+?)\s+Data\s*:", default="")
    endereco_entrega = buscar_regex(texto_flat, r"LOCAL\s+ENTREGA\s*:\s*(.+?)\s+LOCAL\s+COBRANÇA", default="")
    fornecedor_nome = buscar_regex(texto_flat, r"Fornecedor\s*:\s*\d+\s*-\s*(.+?)\s+Endereço\s*:", default="")
    fornecedor_cnpj = limpar_cnpj(buscar_regex(texto_flat, r"CNPJ\s*:\s*([0-9\.\-/]+)", default=""))
    valor_total_pedido = numero(buscar_regex(texto_flat, r"Valor\s+do\s+Pedido\s*:\s*([0-9\.,]+)", default=""))
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
            fim_item = any(x in texto_linha for x in ["VALOR DO PEDIDO", "CONDIÇÕES DE PAGAMENTO", "SETOR SOLICITANTE", "DGS BRASIL", "PÁGINA ", "----"])
            cabecalho_tabela = any(x in texto_linha for x in ["REF.MAT", "UN.COMP", "QTD.PEDIDA", "VALOR UNIT", "CÓDIGO", "DESCRIÇÃO DO ITEM", "FABRICANTE"])
            if not fim_item and not cabecalho_tabela:
                item_atual["descricao_item_oc"] = normalize_spaces(item_atual.get("descricao_item_oc", "") + " " + limpa)
    if item_atual:
        registros.append(item_atual)
    if not registros:
        try:
            (PASTA_RESULTADO / f"debug_itens_{caminho_doc.stem}.txt").write_text(texto, encoding="utf-8")
        except Exception:
            pass
    return pd.DataFrame(registros)

# =========================================================
# COMPARAÇÃO INTELIGENTE DE ITENS (XML x OC)
# =========================================================
def extrair_fator(descricao):
    """Detecta padrões de embalagem tipo 2X8, 10X50 na descrição.

    Agora é usado apenas como informação auxiliar/auditoria — NÃO
    sobrescreve mais a quantidade fiscal do XML, que é a fonte de
    verdade (a versão anterior fazia `qtd_xml = fator`, o que podia
    substituir uma quantidade correta por um número extraído de forma
    heurística da descrição).
    """
    desc = str(descricao or "").upper()
    m = re.search(r"\b(\d{1,3})\s*[Xx]\s*(\d{1,3})\b", desc)
    if not m:
        return None
    return int(m.group(1)) * int(m.group(2))


def calcular_similaridade_descricao(desc_xml_norm: str, desc_oc_norm: str) -> float:
    """Score contínuo (0-1) de quão parecidas são duas descrições já
    normalizadas. Combina:
    - Jaccard de tokens relevantes (ignorando stopwords de item);
    - razão de similaridade de sequência (difflib), que capta
      proximidade mesmo quando o vocabulário difere um pouco.
    Contenção direta (uma descrição contida na outra) tem peso extra,
    pois é um sinal muito forte de correspondência.
    """
    if not desc_xml_norm or not desc_oc_norm:
        return 0.0

    tokens_xml = {t for t in desc_xml_norm.split() if t not in STOPWORDS_DESC and len(t) > 1}
    tokens_oc = {t for t in desc_oc_norm.split() if t not in STOPWORDS_DESC and len(t) > 1}

    if not tokens_xml or not tokens_oc:
        jaccard = 0.0
    else:
        intersecao = tokens_xml & tokens_oc
        uniao = tokens_xml | tokens_oc
        jaccard = len(intersecao) / len(uniao) if uniao else 0.0

    seq_ratio = SequenceMatcher(None, desc_xml_norm, desc_oc_norm).ratio()
    score = (jaccard * 0.6) + (seq_ratio * 0.4)

    if desc_xml_norm in desc_oc_norm or desc_oc_norm in desc_xml_norm:
        score = max(score, 0.85)

    return round(score, 4)


def consolidar_itens_xml(xml_itens: list) -> list:
    """Agrupa linhas repetidas do XML pelo mesmo item (código
    normalizado + descrição normalizada), somando quantidade e valor.
    """
    mapa = {}
    ordem = []
    for item in xml_itens:
        codigo = str(item.get("codigo_item_xml", "") or "").strip()
        descricao = str(item.get("descricao_item_xml", "") or "")
        chave = (normalizar_codigo(codigo), normalizar_desc(descricao))
        if chave not in mapa:
            mapa[chave] = dict(item)
            ordem.append(chave)
        else:
            mapa[chave]["quantidade_item_xml"] = numero(mapa[chave].get("quantidade_item_xml", 0)) + numero(item.get("quantidade_item_xml", 0))
            mapa[chave]["valor_total_item_xml"] = numero(mapa[chave].get("valor_total_item_xml", 0)) + numero(item.get("valor_total_item_xml", 0))
    return [mapa[chave] for chave in ordem]


def indexar_itens_oc(df_oc: pd.DataFrame) -> list:
    """Pré-processa os itens da OC uma única vez (código/descrição
    normalizados), evitando reprocessar texto a cada item do XML."""
    indice = []
    if df_oc is None or df_oc.empty:
        return indice
    for pos, row in enumerate(df_oc.to_dict(orient="records")):
        codigo = str(row.get("codigo_item_oc", "")).strip()
        descricao = str(row.get("descricao_item_oc", ""))
        qtd = numero(row.get("quantidade_item_oc", 0))
        indice.append({
            "pos": pos,
            "codigo": codigo,
            "codigo_norm": normalizar_codigo(codigo),
            "descricao": descricao,
            
            "descricao_norm": normalizar_desc(descricao),
            "descricao_codigo_norm": normalizar_codigo(descricao),

            "unidade": str(row.get("unidade_item_oc", "")),
            "valor_unitario": numero(row.get("valor_unitario_item_oc", 0)),
            "valor_total": numero(row.get("valor_total_item_oc", 0)),
            "quantidade_total": qtd,
            "quantidade_disponivel": qtd,
        })
    return indice


def localizar_item_oc(item_xml: dict, indice_oc: list):
    """Busca o item da OC correspondente a um item do XML.

    Diferente da versão anterior (que percorria as linhas da OC e, em
    cada linha, testava as 5 regras em sequência — o que fazia o
    primeiro item da planilha "ganhar" mesmo quando não era o melhor
    match), aqui cada critério de prioridade é testado contra TODOS os
    itens da OC antes de passar para o próximo critério. Isso garante
    que "código exato" sempre vença "match por descrição", não importa
    a ordem das linhas na OC.

    Itens da OC com saldo já consumido por outra linha do XML (relevante
    quando duas descrições diferentes disputam a mesma linha da OC no
    match por descrição) são preteridos, mas continuam elegíveis se não
    houver alternativa — ficando marcados via score/õmétodo para auditoria.
    """
    codigo_xml = str(item_xml.get("codigo_item_xml", "")).strip()
    codigo_xml_norm = normalizar_codigo(codigo_xml)
    desc_xml = str(item_xml.get("descricao_item_xml", ""))
    desc_xml_norm = normalizar_desc(desc_xml)

    disponiveis = [c for c in indice_oc if c["quantidade_disponivel"] > 1e-6]
    pool = disponiveis if disponiveis else indice_oc

    # 1) código exato
    for c in pool:
        if codigo_xml and codigo_xml == c["codigo"]:
            return c, "COD_EXATO", 1.0

    # 2) código do XML dentro da descrição da OC
    for c in pool:
        if codigo_xml and codigo_xml in c["descricao"]:
            return c, "COD_EM_DESC", 0.95

    # 3) código relaxado (sem hífen/espaço/símbolo)
    for c in pool:
        if codigo_xml_norm and codigo_xml_norm == c["codigo_norm"]:
            return c, "COD_RELAXADO", 0.9
    for c in pool:
        
        if codigo_xml_norm and codigo_xml_norm in c["descricao_codigo_norm"]:
            return c, "COD_RELAXADO_EM_DESC", 0.85


    # 4) código da OC dentro da descrição da NF (fallback adicional)
    for c in pool:
        if c["codigo"] and c["codigo"] in desc_xml:
            return c, "COD_OC_EM_XML", 0.8

    # 5) descrição inteligente — escolhe o MELHOR candidato, não o primeiro
    melhor, melhor_score = None, 0.0
    for c in pool:
        score = calcular_similaridade_descricao(desc_xml_norm, c["descricao_norm"])
        tokens_comuns = len(set(desc_xml_norm.split()) & set(c["descricao_norm"].split()))
        contido = desc_xml_norm in c["descricao_norm"] or c["descricao_norm"] in desc_xml_norm
        if score >= LIMIAR_SIMILARIDADE_DESC and (tokens_comuns >= MIN_TOKENS_COMUNS or contido):
            if score > melhor_score:
                melhor, melhor_score = c, score
    if melhor is not None:
        return melhor, "DESC_MATCH", melhor_score

    # Último recurso: código exato bate, mas só em item já esgotado por
    # outra linha do XML — ainda é o item certo, só registramos o alerta.
    if disponiveis != indice_oc:
        for c in indice_oc:
            if codigo_xml and codigo_xml == c["codigo"]:
                return c, "COD_EXATO_ITEM_OC_ESGOTADO", 1.0

    return None, "NAO_ENCONTRADO", 0.0


def avaliar_itens(xml_itens: list, df_oc: pd.DataFrame) -> dict:
    erro_critico = False
    itens_ok = 0
    itens_erro = 0
    motivos = []
    detalhes = []

    itens_consolidados = consolidar_itens_xml(xml_itens)

    if df_oc is None or df_oc.empty:
        return {
            "erro_critico": True,
            "itens_ok": 0,
            "itens_erro": len(itens_consolidados),
            "motivos": ["OC_SEM_ITENS"],
            "detalhes": [],
        }

    indice_oc = indexar_itens_oc(df_oc)

    for item in itens_consolidados:
        codigo_xml = str(item.get("codigo_item_xml", "")).strip()
        desc_xml = str(item.get("descricao_item_xml", ""))
        qtd_xml = numero(item.get("quantidade_item_xml", 0))
        vl_unit_xml = numero(item.get("valor_unitario_item_xml", 0))
        vl_total_xml = numero(item.get("valor_total_item_xml", 0))
        unidade_xml = str(item.get("unidade_item_xml", ""))
        fator = extrair_fator(desc_xml)

        match, metodo, score = localizar_item_oc(item, indice_oc)
        item_motivos = []

        if match is None:
            erro_critico = True
            itens_erro += 1
            item_motivos.append("ITEM_NAO_ENCONTRADO")
            motivos.append(f"{codigo_xml}|ITEM_NAO_ENCONTRADO")
            detalhes.append({
                "codigo_item_xml": codigo_xml,
                "descricao_item_xml": desc_xml,
                "status_item": "ERRO",
                "motivo_item": "ITEM_NAO_ENCONTRADO",
                "metodo_match": metodo,
                "score_match": 0.0,
                "quantidade_item_xml": qtd_xml,
                "quantidade_item_oc": "",
                "valor_unitario_item_xml": vl_unit_xml,
                "valor_unitario_item_oc": "",
            })
            continue

        qtd_oc = numero(match.get("quantidade_total", 0))
        vl_unit_oc = numero(match.get("valor_unitario", 0))
        vl_total_oc = numero(match.get("valor_total", 0))
        unidade_oc = str(match.get("unidade", ""))

        # baixa informativa de saldo (não altera a regra de qtd_xml x qtd_oc,
        # que segue comparando com a quantidade total do item na OC)
        match["quantidade_disponivel"] = max(0.0, match["quantidade_disponivel"] - qtd_xml)

        # --- Quantidade ---
        if qtd_xml > qtd_oc + TOL_QTD:
            erro_critico = True
            itens_erro += 1
            item_motivos.append("QTD_ACIMA")
            motivos.append(f"{codigo_xml}|QTD_ACIMA")
        else:
            itens_ok += 1

        # --- Valor unitário ---
        if abs(vl_unit_xml - vl_unit_oc) > TOL_VALOR:
            item_motivos.append("VALOR_DIVERGENTE")
            motivos.append(f"{codigo_xml}|VALOR_DIVERGENTE")

        # --- Consistência de cálculo do item (qtd * vl_unit ~= vl_total) ---
        vl_total_calculado = qtd_xml * vl_unit_xml
        if vl_total_xml and abs(vl_total_xml - vl_total_calculado) > max(TOL_VALOR, vl_total_calculado * 0.01):
            item_motivos.append("CALCULO_ITEM_INCONSISTENTE")
            motivos.append(f"{codigo_xml}|CALCULO_ITEM_INCONSISTENTE")

        # --- Valor faturado do item não deve superar indevidamente o item da OC ---
        if vl_total_oc and vl_total_xml > vl_total_oc + TOL_VALOR:
            item_motivos.append("VALOR_ITEM_ACIMA_OC")
            motivos.append(f"{codigo_xml}|VALOR_ITEM_ACIMA_OC")

        # --- Unidade ---
        u_xml_norm = normalizar_unidade(unidade_xml)
        u_oc_norm = normalizar_unidade(unidade_oc)
        if u_xml_norm and u_oc_norm and u_xml_norm != u_oc_norm:
            item_motivos.append("UNIDADE_DIVERGENTE")
            motivos.append(f"{codigo_xml}|UNIDADE_DIVERGENTE")

        # --- Fator de embalagem (informativo) ---
        if fator:
            item_motivos.append(f"FATOR_EMBALAGEM_DETECTADO:{fator}")

        if metodo == "COD_EXATO_ITEM_OC_ESGOTADO":
            item_motivos.append("OC_ITEM_JA_UTILIZADO_POR_OUTRA_LINHA")

        status_item = "ERRO" if ("QTD_ACIMA" in item_motivos or "ITEM_NAO_ENCONTRADO" in item_motivos) else "OK"

        detalhes.append({
            "codigo_item_xml": codigo_xml,
            "descricao_item_xml": desc_xml,
            "status_item": status_item,
            "motivo_item": "|".join(item_motivos),
            "metodo_match": metodo,
            "score_match": score,
            "quantidade_item_xml": qtd_xml,
            "quantidade_item_oc": qtd_oc,
            "valor_unitario_item_xml": vl_unit_xml,
            "valor_unitario_item_oc": vl_unit_oc,
        })

    return {
        "erro_critico": erro_critico,
        "itens_ok": itens_ok,
        "itens_erro": itens_erro,
        "motivos": motivos,
        "detalhes": detalhes,
    }

# =========================================================
# PAREAMENTO
# =========================================================
def listar_arquivos():
    xmls = sorted(PASTA_XML.glob("*.xml"))
    docs_oc = sorted(list(PASTA_OC.glob("*.doc")) + list(PASTA_OC.glob("*.docx")) + list(PASTA_OC.glob("*.txt")) + list(PASTA_OC.glob("*.csv")))
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
    return {"data_limite_7_dias": data_str(limite), "status_prazo": status}


def avaliar_cabecalho(xml_head, oc_head):
    motivos = []
    erro_critico = False

    doc_oc = so_numeros(oc_head.get("doc_origem", ""))
    xped = so_numeros(xml_head.get("xped_xml", ""))
    obs_nf = so_numeros(xml_head.get("observacao_nota", ""))

    doc_ok = (doc_oc == xped) or (doc_oc and doc_oc in obs_nf)
    if not doc_ok:
        erro_critico = True
        motivos.append("DOC_ORIGEM")

    raiz_nf = raiz_cnpj(xml_head.get("emitente_cnpj", ""))
    raiz_oc = raiz_cnpj(oc_head.get("fornecedor_cnpj", ""))
    if raiz_nf != raiz_oc:
        erro_critico = True
        motivos.append("CNPJ_RAIZ")
    elif limpar_cnpj(xml_head.get("emitente_cnpj", "")) != limpar_cnpj(oc_head.get("fornecedor_cnpj", "")):
        motivos.append("FILIAL")

    
    cliente_ok_texto = comparar_texto_simples(
    oc_head.get("cliente_nome", ""),
    xml_head.get("destinatario_nome", "")
    )

    cliente_ok_cnpj = raiz_cnpj(oc_head.get("cliente_cnpj", "")) == raiz_cnpj(xml_head.get("destinatario_cnpj", ""))

    cliente_ok = cliente_ok_texto or cliente_ok_cnpj

    if not cliente_ok:
        erro_critico = True
        motivos.append("CLIENTE")


    valor_oc = numero(oc_head.get("valor_total_pedido", 0))
    valor_nf = numero(xml_head.get("valor_total_nota", 0))
    if abs(valor_nf - valor_oc) > TOL_VALOR:
        if valor_nf > valor_oc:
            erro_critico = True
            motivos.append("VALOR_ACIMA")
        else:
            motivos.append("NF_PARCIAL")

    prazo = conferir_prazo(xml_head.get("data_emissao", ""), oc_head.get("prazo_entrega", ""))

    return {
        "erro_critico": erro_critico,
        "motivos": motivos,
        "status_prazo": prazo["status_prazo"],
        "data_limite_7_dias": prazo["data_limite_7_dias"],
        "doc_origem_ok": doc_ok,
        "cliente_ok": cliente_ok,
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

        if xml_head is not None and oc_head is not None:
            try:
                cab = avaliar_cabecalho(xml_head, oc_head)
                conferencias_rows.append({
                    "numero_nota": xml_head.get("numero_nota", ""),
                    "numero_oc": oc_head.get("numero_oc", ""),
                    "status_cabecalho": "OK" if not cab["erro_critico"] else "ERRO",
                    "motivo_cabecalho": "|".join(cab.get("motivos", [])),
                    "status_prazo": cab.get("status_prazo", ""),
                    "data_limite_7_dias": cab.get("data_limite_7_dias", ""),
                })

                itens = avaliar_itens(itens_xml, df_itens_oc)
                itens_ok = itens["itens_ok"]
                itens_erro = itens["itens_erro"]
                status_itens = "ERRO" if itens["erro_critico"] else "OK"
                log["status_conferencia_itens"] = status_itens

                for det in itens.get("detalhes", []):
                    conferencia_itens_rows.append({
                        "arquivo_xml": arq_xml.name,
                        "arquivo_oc_doc": arq_oc_doc.name,
                        "numero_nota": xml_head.get("numero_nota", ""),
                        "numero_oc": oc_head.get("numero_oc", ""),
                        **det,
                    })

                if cab["erro_critico"] or itens["erro_critico"]:
                    status_final = "EMPRESTIMO"
                else:
                    status_final = "OK"

                motivos = []
                motivos.extend(cab.get("motivos", []))
                motivos.extend(itens.get("motivos", []))
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
                    "status_cabecalho": "OK" if not cab["erro_critico"] else "ERRO",
                    "motivo": "|".join(motivos),
                    "responsavel": "OK" if status_final == "OK" else "EMPRESTIMO",
                })

                print(
                    f"Conferencia | Cabecalho: {'OK' if not cab['erro_critico'] else 'ERRO'} | "
                    f"Itens: {status_itens} | Final: {status_final}"
                )
            except Exception as e:
                log["status_conferencia_cabecalho"] = "ERRO"
                log["status_conferencia_itens"] = "ERRO"
                log["observacao"] += f"Erro conferencia: {e} | "
                print(f"Erro conferencia: {e}")

        elif xml_head is not None and oc_head is None:
            # Regra de negócio: sem OC vinculada, a NF é RECUSADA.
            # A versão anterior deixava esse caso sem nenhuma linha de
            # saída em programacao_recebimento — corrigido aqui.
            log["status_conferencia_cabecalho"] = "RECUSADA_SEM_OC"
            log["status_conferencia_itens"] = "NAO_AVALIADO"
            programacao_rows.append({
                "id": idx,
                "data_programada": xml_head.get("data_programada_xml", ""),
                "numero_nota": xml_head.get("numero_nota", ""),
                "numero_oc": "",
                "doc_origem": xml_head.get("doc_origem_encontrado_xml", ""),
                "fornecedor": xml_head.get("emitente_nome", ""),
                "qtd_itens": len(consolidar_itens_xml(itens_xml)),
                "itens_ok": 0,
                "itens_erro": 0,
                "valor_total": xml_head.get("valor_total_nota", 0),
                "status": "RECUSADA",
                "status_itens": "NAO_AVALIADO",
                "status_cabecalho": "NAO_AVALIADO",
                "motivo": "OC_NAO_VINCULADA",
                "responsavel": "RECUSADA",
            })
            print("Classificacao: RECUSADA (OC nao vinculada / nao encontrada)")
        else:
            log["status_conferencia_cabecalho"] = "NAO_REALIZADA"
            log["status_conferencia_itens"] = "NAO_REALIZADA"

        print()
        logs.append(log)

    df_cab_xml = pd.DataFrame(cab_xml_rows)
    df_itens_xml = pd.DataFrame(itens_xml_rows)
    df_cab_oc = pd.DataFrame(cab_oc_rows)
    df_itens_oc = pd.DataFrame(itens_oc_rows)
    df_conf = pd.DataFrame(conferencias_rows)
    df_conf_itens = pd.DataFrame(conferencia_itens_rows)
    df_prog = pd.DataFrame(programacao_rows)
    df_logs = pd.DataFrame(logs)

    with pd.ExcelWriter(ARQUIVO_RESULTADO, engine="openpyxl") as writer:
        (df_cab_xml if not df_cab_xml.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="xml_cabecalho", index=False)
        (df_itens_xml if not df_itens_xml.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="xml_itens", index=False)
        (df_cab_oc if not df_cab_oc.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="oc_doc_cabecalho", index=False)
        (df_itens_oc if not df_itens_oc.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="oc_doc_itens", index=False)
        (df_conf if not df_conf.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="conferencia_cabecalho", index=False)
        (df_conf_itens if not df_conf_itens.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="conferencia_itens", index=False)
        (df_prog if not df_prog.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="programacao", index=False)
        (df_logs if not df_logs.empty else pd.DataFrame(columns=["sem_dados"])).to_excel(writer, sheet_name="log_processamento", index=False)

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
