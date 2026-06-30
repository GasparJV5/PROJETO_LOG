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

    if data_programada is not None:
        data_str = data_programada.strftime("%d/%m/%Y") if hasattr(data_programada, "strftime") else str(data_programada)
        vazios = df["data_programada"].astype(str).str.strip().isin(["", "nan", "None"])
        df.loc[vazios, "data_programada"] = data_str
        df.to_csv(ARQUIVO_PROGRAMACAO, sep=";", index=False)

    return df


def formatar_moeda(valor):
    try:
        valor = float(valor)
    except:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def mostrar_cards(df: pd.DataFrame):
    if df.empty:
        st.info("Nenhum resultado disponível ainda.")
        return

    st.subheader("Painel de Recebimento")

    status_filtro = st.selectbox("Filtrar status", ["TODOS", "OK", "ERRO"], index=0)
    if status_filtro != "TODOS":
        df = df[df["status"].astype(str).str.upper() == status_filtro]

    if df.empty:
        st.warning("Nenhum card para o filtro selecionado.")
        return

    agrupado = df.groupby("data_programada", dropna=False)

    for data_prog, grupo in agrupado:
        data_prog = data_prog if str(data_prog).strip() not in ["", "nan", "None"] else "Sem Data"
        st.markdown(f"## {data_prog}")

        col1, col2 = st.columns(2)
        ok = grupo[grupo["status"].astype(str).str.upper() == "OK"]
        erro = grupo[grupo["status"].astype(str).str.upper() == "ERRO"]

        with col1:
            st.markdown("### OK")
            if ok.empty:
                st.write("Nenhum card OK")
            else:
                for _, row in ok.iterrows():
                    with st.container(border=True):
                        st.markdown(f"**NF:** {row.get('numero_nota', '')}")
                        st.markdown(f"**Fornecedor:** {row.get('fornecedor', '')}")
                        st.markdown(f"**OC:** {row.get('numero_oc', '')}")
                        st.markdown(f"**Doc. Origem:** {row.get('doc_origem', '')}")
                        st.markdown(f"**Qtd Itens:** {row.get('qtd_itens', 0)}")
                        st.markdown(f"**Valor Total:** {formatar_moeda(row.get('valor_total', 0))}")
                        st.caption(f"Responsável: {row.get('responsavel', '')}")

        with col2:
            st.markdown("### ERRO")
            if erro.empty:
                st.write("Nenhum card com erro")
            else:
                for _, row in erro.iterrows():
                    with st.expander(f"NF {row.get('numero_nota', '')} | {row.get('fornecedor', '')}"):
                        st.markdown(f"**OC:** {row.get('numero_oc', '')}")
                        st.markdown(f"**Doc. Origem:** {row.get('doc_origem', '')}")
                        st.markdown(f"**Qtd Itens:** {row.get('qtd_itens', 0)}")
                        st.markdown(f"**Valor Total:** {formatar_moeda(row.get('valor_total', 0))}")
                        st.markdown(f"**Status Cabeçalho:** {row.get('status_cabecalho', '')}")
                        st.markdown(f"**Status Itens:** {row.get('status_itens', '')}")
                        st.markdown(f"**Motivo:** {row.get('motivo', '')}")
                        st.caption(f"Responsável: {row.get('responsavel', '')}")


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
        mostrar_cards(df)

    else:
        st.subheader("2) Cards já processados")
        df = carregar_programacao()
        mostrar_cards(df)


if __name__ == "__main__":
    main()