# -*- coding: utf-8 -*-
"""
app.py — Painel RH TJMA · Auxílio-Bolsa (GDG)
Adm. Jesus Martins Oliveira Junior — COGEX-MA/TJMA
MAT-RHBOLSAS-STREAMLIT-001

Login: RH / RH@123
Fonte de dados (ordem de prioridade):
  1) GET em link público .CSV (Google Sheets publicado como CSV)
  2) Upload manual de arquivo .XLSX/.XLS completo
  3) CSV sanitizado local (dados_sanitizados.csv, 100% dos registros reais)
Regra dura: 100% dos registros — sem amostragem, sem invenção de linha/valor.
"""
import io
import datetime
import requests
import pandas as pd
import streamlit as st
import plotly.express as px

# ==========================================================================
# CONFIG
# ==========================================================================
APP_VERSION = "v1.0.0"
APP_TITLE = "Painel RH TJMA · Auxílio-Bolsa"
LOGIN_USER = "RH"
LOGIN_PASS = "RH@123"

# Link público CSV (Google Sheets publicado). Ajustar aqui quando o link
# público oficial estiver disponível. Formato exigido: export?format=csv
CSV_PUBLICO_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1VR3D-L_D0AGYa_y8-Rya5wNEV6NysmuVkpOv5w0cQHA/export?format=csv"
)
CSV_LOCAL_FALLBACK = "dados_sanitizados.csv"

TAG_LABELS = {
    "ATIVO-EM-CURSO": ("Ativo em curso", "🟢"),
    "PENDENTE-DOCUMENTACAO": ("Pendente documentação", "📄"),
    "PENDENTE-COMPROVACAO": ("Pendente comprovação", "🟠"),
    "CONCLUIDO-DIPLOMADO": ("Concluído/diplomado", "🎓"),
    "SEM-DADOS-SUFICIENTES": ("Sem dados suficientes", "⚪"),
    "OCORRENCIA-PROCESSUAL": ("Ocorrência processual", "⚠️"),
    "IDENTIFICADO-SEM-STATUS-REGISTRADO": ("Sem status registrado", "➖"),
}

st.set_page_config(page_title=APP_TITLE, page_icon="🎓", layout="wide")

# ==========================================================================
# ESTILO — verde / branco / azul + tema escuro (azul/branco), contraste OK
# ==========================================================================
def aplicar_estilo(modo_escuro: bool):
    if modo_escuro:
        bg = "#0B1220"
        bg_card = "#131C2E"
        texto = "#F5F7FA"
        azul = "#3B82F6"
        verde = "#22C55E"
        borda = "#1E2A44"
    else:
        bg = "#FFFFFF"
        bg_card = "#F3F8F4"
        texto = "#0B1220"
        azul = "#1D4ED8"
        verde = "#15803D"
        borda = "#D6E4DC"

    st.markdown(f"""
    <style>
        .stApp {{ background-color: {bg}; color: {texto}; }}
        [data-testid="stMetric"] {{
            background-color: {bg_card};
            border: 1px solid {borda};
            border-radius: 10px;
            padding: 12px 14px;
        }}
        [data-testid="stMetricValue"] {{ color: {azul}; }}
        [data-testid="stMetricLabel"] {{ color: {texto}; }}
        h1, h2, h3 {{ color: {azul}; }}
        .rodape-app {{
            color: {texto}; opacity: 0.55; font-size: 0.75rem;
            text-align: right; margin-top: 2rem; border-top: 1px solid {borda};
            padding-top: 8px;
        }}
        .badge-verde {{ color: {verde}; font-weight: 600; }}
        section[data-testid="stSidebar"] {{ background-color: {bg_card}; }}
    </style>
    """, unsafe_allow_html=True)
    return azul, verde


# ==========================================================================
# LOGIN
# ==========================================================================
def tela_login():
    st.markdown(f"## 🎓 {APP_TITLE}")
    st.caption(f"Controle de deploy: {APP_VERSION}")
    with st.form("login_form"):
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        if usuario == LOGIN_USER and senha == LOGIN_PASS:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")


# ==========================================================================
# INGESTÃO DE DADOS (GET público -> upload .xlsx -> CSV local sanitizado)
# ==========================================================================
@st.cache_data(ttl=300, show_spinner=False)
def carregar_de_csv_publico(url: str) -> pd.DataFrame | None:
    try:
        resp = requests.get(url, timeout=8)
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        if df.shape[1] <= 1:
            return None
        return df
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def carregar_de_csv_local(path: str) -> pd.DataFrame | None:
    try:
        return pd.read_csv(path)
    except Exception:
        return None


def carregar_de_xlsx(arquivo) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(arquivo, engine="openpyxl")
        aba_preferida = None
        for nome in xls.sheet_names:
            if "GERAL" in nome.upper():
                aba_preferida = nome
                break
        aba = aba_preferida or xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=aba)
        df = df.dropna(axis=1, how="all")
        df = df.dropna(axis=0, how="all")
        return df
    except Exception as e:
        st.sidebar.error(f"Falha ao ler o arquivo enviado: {e}")
        return None


def obter_dados():
    st.sidebar.markdown("### 📥 Fonte de dados")
    origem_forcada = st.sidebar.file_uploader(
        "Upload manual (.xlsx / .xls) — alternativa ao GET público",
        type=["xlsx", "xls"]
    )

    if origem_forcada is not None:
        df = carregar_de_xlsx(origem_forcada)
        if df is not None:
            st.sidebar.success(f"Carregado do upload manual · {len(df)} registros")
            return df, "Upload manual (.xlsx)"

    df = carregar_de_csv_publico(CSV_PUBLICO_URL)
    if df is not None:
        st.sidebar.success(f"Carregado via GET público (CSV) · {len(df)} registros")
        return df, "GET público (CSV)"

    df = carregar_de_csv_local(CSV_LOCAL_FALLBACK)
    if df is not None:
        st.sidebar.warning(f"GET público falhou — usando CSV sanitizado local · {len(df)} registros")
        return df, "CSV sanitizado local (fallback)"

    st.sidebar.error("Nenhuma fonte de dados disponível (GET, upload e CSV local falharam).")
    return pd.DataFrame(), "Nenhuma"


# ==========================================================================
# KPIs E FILTROS
# ==========================================================================
def coluna_existente(df, nome):
    return nome if nome in df.columns else None


def montar_kpis(df: pd.DataFrame):
    total = len(df)
    col_tag = coluna_existente(df, "TAG_TIPOLOGIA")
    contagem = df[col_tag].value_counts(dropna=True) if col_tag else pd.Series(dtype=int)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de registros", total)
    c2.metric("🟢 Ativos em curso", int(contagem.get("ATIVO-EM-CURSO", 0)))
    c3.metric("🎓 Concluídos/diplomados", int(contagem.get("CONCLUIDO-DIPLOMADO", 0)))
    c4.metric("⚠️ Ocorrência processual", int(contagem.get("OCORRENCIA-PROCESSUAL", 0)))
    c5.metric("🟠 Pendentes (doc. + comprov.)",
              int(contagem.get("PENDENTE-DOCUMENTACAO", 0) + contagem.get("PENDENTE-COMPROVACAO", 0)))


def montar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.markdown("### 🔎 Filtros")
    busca = st.sidebar.text_input("Busca livre (nome, matrícula, processo...)")

    df_filtrado = df.copy()

    for col, rotulo in [
        ("COMARCA", "Comarca"),
        ("TIPO DE BOLSA", "Tipo de bolsa"),
        ("STATUS", "Status"),
        ("TAG_TIPOLOGIA", "Situação (tag)"),
    ]:
        if col in df.columns:
            opcoes = sorted([str(v) for v in df[col].dropna().unique().tolist()])
            selecionadas = st.sidebar.multiselect(rotulo, opcoes)
            if selecionadas:
                df_filtrado = df_filtrado[df_filtrado[col].astype(str).isin(selecionadas)]

    if busca:
        mask = pd.Series(False, index=df_filtrado.index)
        for col in df_filtrado.columns:
            mask = mask | df_filtrado[col].astype(str).str.contains(busca, case=False, na=False)
        df_filtrado = df_filtrado[mask]

    if st.sidebar.button("🧹 Limpar filtros"):
        st.rerun()

    return df_filtrado


def montar_graficos(df: pd.DataFrame, azul, verde):
    col1, col2 = st.columns(2)

    if "TAG_TIPOLOGIA" in df.columns:
        cont = df["TAG_TIPOLOGIA"].value_counts(dropna=True).reset_index()
        cont.columns = ["Situação", "Registros"]
        cont["Situação"] = cont["Situação"].map(lambda t: TAG_LABELS.get(t, (t, ""))[0])
        fig1 = px.pie(cont, names="Situação", values="Registros",
                       title="Distribuição por situação (TAG_TIPOLOGIA)",
                       color_discrete_sequence=[azul, verde, "#F59E0B", "#EF4444", "#94A3B8", "#8B5CF6", "#0EA5E9"])
        col1.plotly_chart(fig1, use_container_width=True)

    if "TIPO DE BOLSA" in df.columns:
        cont2 = df["TIPO DE BOLSA"].value_counts(dropna=True).reset_index()
        cont2.columns = ["Tipo de bolsa", "Registros"]
        fig2 = px.bar(cont2, x="Tipo de bolsa", y="Registros",
                       title="Registros por tipo de bolsa",
                       color_discrete_sequence=[verde])
        col2.plotly_chart(fig2, use_container_width=True)

    if "COMARCA" in df.columns:
        cont3 = df["COMARCA"].value_counts(dropna=True).reset_index().head(15)
        cont3.columns = ["Comarca", "Registros"]
        fig3 = px.bar(cont3, x="Registros", y="Comarca", orientation="h",
                       title="Top 15 comarcas por nº de registros",
                       color_discrete_sequence=[azul])
        fig3.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig3, use_container_width=True)


def montar_tabela_e_export(df: pd.DataFrame):
    st.markdown("### 📋 Registros")
    st.caption(f"{len(df)} registro(s) após filtro")
    st.dataframe(df, use_container_width=True, height=420)

    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    nome_arquivo = f"RH_TJMA_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    st.download_button(
        "⬇️ Exportar CSV (RH TJMA)",
        data=csv_bytes,
        file_name=nome_arquivo,
        mime="text/csv",
    )


def rodape():
    st.markdown(
        f'<div class="rodape-app">ADM, Jesus e A, 107805 · Painel RH TJMA · {APP_VERSION}</div>',
        unsafe_allow_html=True,
    )


# ==========================================================================
# MAIN
# ==========================================================================
def main():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    modo_escuro = st.sidebar.toggle("🌙 Tema escuro (azul/branco)", value=False)
    azul, verde = aplicar_estilo(modo_escuro)

    if not st.session_state["autenticado"]:
        tela_login()
        rodape()
        return

    st.sidebar.markdown(f"**Sessão:** {LOGIN_USER}")
    if st.sidebar.button("Sair"):
        st.session_state["autenticado"] = False
        st.rerun()

    st.title(f"🎓 {APP_TITLE}")

    df, origem = obter_dados()
    st.caption(f"Fonte ativa: **{origem}**")

    if df.empty:
        rodape()
        return

    df_filtrado = montar_filtros(df)
    montar_kpis(df_filtrado)
    montar_graficos(df_filtrado, azul, verde)
    montar_tabela_e_export(df_filtrado)
    rodape()


if __name__ == "__main__":
    main()
