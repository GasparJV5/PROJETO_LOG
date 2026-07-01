import sys
import subprocess
from pathlib import Path
from datetime import date
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
PASTA_XML = BASE_DIR / "entrada" / "xml"
PASTA_OC = BASE_DIR / "entrada" / "oc"
PASTA_RESULTADO = BASE_DIR / "resultado"
ARQUIVO_PROGRAMACAO = PASTA_RESULTADO / "programacao_recebimento.csv"

# Status oficiais gerados pelo motor (main.py). O painel antigo filtrava
# por "ERRO", que nunca é gerado pelo motor (ele gera OK/EMPRESTIMO/
# RECUSADA) — o filtro nunca funcionava. Corrigido aqui.
STATUS_VALIDOS = ["OK", "EMPRESTIMO", "RECUSADA"]


def garantir_pastas():
    PASTA_XML.mkdir(parents=True, exist_ok=True)
    PASTA_OC.mkdir(parents=True, exist_ok=True)
    PASTA_RESULTADO.mkdir(parents=True, exist_ok=True)


def salvar_upload(uploaded_file, destino: Path):
    destino.parent.mkdir(parents=True, exist_ok=True)
    with open(destino, "wb") as f:
        f.write(uploaded_file.getbuffer())


def executar_motor():
    comando = [sys.executable, str(BASE_DIR / "main.py")]
    return subprocess.run(
        comando,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(BASE_DIR)
    )


def carregar_programacao(data_programada=None):
    if not ARQUIVO_PROGRAMACAO.exists():
        return pd.DataFrame()

    df = pd.read_csv(ARQUIVO_PROGRAMACAO, sep=";")

    defaults = {
        "id": "",
        "data_programada": "",
        "numero_nota": "",
        "numero_oc": "",
        "doc_origem": "",
        "fornecedor": "",
        "qtd_itens": 0,
        "valor_total": 0,
        "status": "",
        "status_itens": "",
        "status_cabecalho": "",
        "responsavel": "",
        "motivo": "",
        "itens_ok": 0,
        "itens_erro": 0,
    }

    for col, default in defaults.items():
        if col not in df.columns:
            df[col] = default

    df["status"] = df["status"].astype(str).str.upper().str.strip()

    if data_programada is not None:
        data_str = data_programada.strftime("%d/%m/%Y") if hasattr(data_programada, "strftime") else str(data_programada)
        vazios = df["data_programada"].astype(str).str.strip().isin(["", "nan", "None"])
        df.loc[vazios, "data_programada"] = data_str
        df.to_csv(ARQUIVO_PROGRAMACAO, sep=";", index=False)

    return df


def carregar_detalhe_itens():
    caminho = PASTA_RESULTADO / "conferencia_itens.csv"
    if not caminho.exists():
        return pd.DataFrame()
    return pd.read_csv(caminho, sep=";")


def formatar_moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def render_card(row, df_itens_detalhe):
    with st.container(border=True):
        st.markdown(f"**NF:** {row.get('numero_nota', '')}  |  **Fornecedor:** {row.get('fornecedor', '')}")
        st.markdown(f"**OC:** {row.get('numero_oc', '') or '—'}  |  **Doc. Origem:** {row.get('doc_origem', '') or '—'}")
        st.markdown(f"**Itens:** {row.get('qtd_itens', 0)} (OK: {row.get('itens_ok', 0)} / Erro: {row.get('itens_erro', 0)})")
        st.markdown(f"**Valor Total:** {formatar_moeda(row.get('valor_total', 0))}")
        if str(row.get("motivo", "")).strip() not in ["", "nan"]:
            st.markdown(f"**Motivo:** {row.get('motivo', '')}")
        st.caption(f"Cabeçalho: {row.get('status_cabecalho', '')} | Itens: {row.get('status_itens', '')} | Responsável: {row.get('responsavel', '')}")

        if not df_itens_detalhe.empty:
            detalhe_nf = df_itens_detalhe[df_itens_detalhe["numero_nota"].astype(str) == str(row.get("numero_nota", ""))]
            if not detalhe_nf.empty:
                with st.expander("Ver comparação item a item"):
                    colunas = [c for c in [
                        "codigo_item_xml", "descricao_item_xml", "status_item", "motivo_item",
                        "metodo_match", "score_match", "quantidade_item_xml", "quantidade_item_oc",
                        "valor_unitario_item_xml", "valor_unitario_item_oc",
                    ] if c in detalhe_nf.columns]
                    st.dataframe(detalhe_nf[colunas], use_container_width=True, hide_index=True)


def mostrar_cards(df: pd.DataFrame, df_itens_detalhe: pd.DataFrame):
    if df.empty:
        st.info("Nenhum resultado disponível ainda.")
        return

    st.subheader("Painel de Recebimento")

    status_filtro = st.selectbox("Filtrar status", ["TODOS"] + STATUS_VALIDOS, index=0)
    if status_filtro != "TODOS":
        df = df[df["status"] == status_filtro]

    if df.empty:
        st.warning("Nenhum card para o filtro selecionado.")
        return

    agrupado = df.groupby("data_programada", dropna=False)

    for data_prog, grupo in agrupado:
        data_prog = data_prog if str(data_prog).strip() not in ["", "nan", "None"] else "Sem Data"
        st.markdown(f"## {data_prog}")

        col_ok, col_emprestimo, col_recusada = st.columns(3)

        with col_ok:
            st.markdown("### ✅ OK")
            grupo_ok = grupo[grupo["status"] == "OK"]
            if grupo_ok.empty:
                st.write("Nenhum card OK")
            else:
                for _, row in grupo_ok.iterrows():
                    render_card(row, df_itens_detalhe)

        with col_emprestimo:
            st.markdown("### ⚠️ Empréstimo (divergência)")
            grupo_emp = grupo[grupo["status"] == "EMPRESTIMO"]
            if grupo_emp.empty:
                st.write("Nenhum card em empréstimo")
            else:
                for _, row in grupo_emp.iterrows():
                    render_card(row, df_itens_detalhe)

        with col_recusada:
            st.markdown("### ❌ Recusada")
            grupo_rec = grupo[grupo["status"] == "RECUSADA"]
            if grupo_rec.empty:
                st.write("Nenhuma nota recusada")
            else:
                for _, row in grupo_rec.iterrows():
                    render_card(row, df_itens_detalhe)


def main():
    st.set_page_config(page_title="Planejamento de Recebimento", page_icon="📦", layout="wide")
    garantir_pastas()

    st.title("Planejamento de Recebimento")
    st.caption("Upload de XML + OC, processamento do motor e visualização em cards por dia.")

    with st.container(border=True):
        st.subheader("1) Incluir arquivos para comparação")
        col_a, col_b = st.columns(2)

        with col_a:
            data_programada = st.date_input("Data programada", value=date.today(), format="DD/MM/YYYY")
            xml_file = st.file_uploader("Selecione o XML da NF-e", type=["xml"], accept_multiple_files=False)

        with col_b:
            oc_file = st.file_uploader(
                "Selecione a OC (DOC, DOCX, TXT ou CSV)",
                type=["doc", "docx", "txt", "csv"],
                accept_multiple_files=False,
            )

        processar = st.button("Processar comparação", type="primary", use_container_width=True)

    if processar:
        if xml_file is None or oc_file is None:
            st.error("Selecione o XML e a OC antes de processar.")
            st.stop()

        for arq in PASTA_XML.glob("*"):
            if arq.is_file():
                arq.unlink(missing_ok=True)

        for arq in PASTA_OC.glob("*"):
            if arq.is_file():
                arq.unlink(missing_ok=True)

        salvar_upload(xml_file, PASTA_XML / xml_file.name)
        salvar_upload(oc_file, PASTA_OC / oc_file.name)

        with st.spinner("Processando XML + OC..."):
            resultado = executar_motor()

        if resultado.returncode != 0:
            st.error("Erro ao executar o motor de comparação.")
            st.code(resultado.stderr or resultado.stdout or "Sem detalhes.")
            st.stop()

        st.success("Comparação executada com sucesso.")

        if resultado.stdout:
            with st.expander("Ver log do processamento"):
                st.code(resultado.stdout)

        df = carregar_programacao(data_programada)
        df_itens_detalhe = carregar_detalhe_itens()
        mostrar_cards(df, df_itens_detalhe)

    else:
        st.subheader("2) Cards já processados")
        df = carregar_programacao()
        df_itens_detalhe = carregar_detalhe_itens()
        mostrar_cards(df, df_itens_detalhe)


if __name__ == "__main__":
    main()
