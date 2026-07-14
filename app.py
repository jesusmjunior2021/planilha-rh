# -*- coding: utf-8 -*-
"""
app.py — Painel RH TJMA ·
Adm. JMOJ-V1.107805
MAT-RHBOLSAS-STREAMLIT-002

Login: RH / RH@123

Fonte de dados (ordem de prioridade) — SEMPRE a planilha real completa,
todas as abas, todas as colunas, estrutura original preservada:
  1) GET do workbook público completo (.xlsx, todas as abas) via export do
     Google Sheets — NÃO usa CSV de aba única, porque CSV exporta só 1 aba.
  2) Upload manual de arquivo .xlsx/.xls completo (todas as abas).
  3) Cópia local da planilha real (fallback offline), mesma estrutura.
Regra dura: 100% dos registros e 100% das abas — sem amostragem, sem
recorte para 1 aba só, sem invenção de linha/valor/coluna.
"""
import io
import datetime
import requests
import pandas as pd
import streamlit as st

# ==========================================================================
# CONFIG
# ==========================================================================
APP_VERSION = "v1.1.0"
APP_TITLE = "Painel RH TJMA · Auxílio-Bolsa"
LOGIN_USER = "RH"
LOGIN_PASS = "RH@123"

GOOGLE_SHEET_ID = "1-udWUaMYkU8dhZtiHarPbDrRRTzOsDbH_HDTcwJM_c8"
# export?format=xlsx traz o WORKBOOK INTEIRO (todas as abas), diferente do
# export?format=csv que só traz 1 aba (gid). É por isso que a fonte primária
# do GET precisa ser xlsx, não csv, quando a exigência é "todas as abas".
XLSX_PUBLICO_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=xlsx"

XLSX_LOCAL_FALLBACK = "planilha_bolsas_COM_FILTROS.xlsx"

TAG_LABELS = {
    "ATIVO-EM-CURSO": ("Ativo em curso", "🟢"),
    "PENDENTE-DOCUMENTACAO": ("Pendente documentação", "📄"),
    "PENDENTE-COMPROVACAO": ("Pendente comprovação", "🟠"),
    "CONCLUIDO-DIPLOMADO": ("Concluído/diplomado", "🎓"),
    "SEM-DADOS-SUFICIENTES": ("Sem dados suficientes", "⚪"),
    "OCORRENCIA-PROCESSUAL": ("Ocorrência processual", "⚠️"),
    "IDENTIFICADO-SEM-STATUS-REGISTRADO": ("Sem status registrado", "➖"),
}

# Colunas candidatas a filtro/gráfico, na ordem de prioridade — só aparecem
# na UI se realmente existirem na aba selecionada (nada é forçado).
COLUNAS_PRIORITARIAS = ["TAG_TIPOLOGIA", "STATUS", "COMARCA", "TIPO DE BOLSA", "CARGO"]

st.set_page_config(page_title=APP_TITLE, page_icon="🎓", layout="wide")

# ==========================================================================
# ESTILO — verde / branco / azul + tema escuro (azul/branco), contraste OK
# ==========================================================================
def aplicar_estilo(modo_escuro: bool):
    if modo_escuro:
        bg, bg_card, texto, azul, verde, borda = (
            "#0B1220", "#131C2E", "#F5F7FA", "#3B82F6", "#22C55E", "#1E2A44"
        )
    else:
        bg, bg_card, texto, azul, verde, borda = (
            "#FFFFFF", "#F3F8F4", "#0B1220", "#1D4ED8", "#15803D", "#D6E4DC"
        )

    st.markdown(f"""
    <style>
        .stApp {{ background-color: {bg}; color: {texto}; }}
        [data-testid="stMetric"] {{
            background-color: {bg_card}; border: 1px solid {borda};
            border-radius: 10px; padding: 12px 14px;
        }}
        [data-testid="stMetricValue"] {{ color: {azul}; }}
        [data-testid="stMetricLabel"] {{ color: {texto}; }}
        h1, h2, h3 {{ color: {azul}; }}
        .rodape-app {{
            color: {texto}; opacity: 0.55; font-size: 0.75rem;
            text-align: right; margin-top: 2rem; border-top: 1px solid {borda};
            padding-top: 8px;
        }}
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
# LEITURA FIEL DA PLANILHA REAL — TODAS AS ABAS, TODAS AS COLUNAS
# (mesma lógica das etapas de sanitização já validadas: detecta banner
# mesclado, cabeçalho real, última linha real; remove só o 100% vazio)
# ==========================================================================
def _detect_header_row(ws):
    row1_filled = sum(1 for c in ws[1] if c.value not in (None, ""))
    row2_filled = sum(1 for c in ws[2] if c.value not in (None, "")) if ws.max_row >= 2 else 0
    return 2 if (row1_filled <= 2 and row2_filled > row1_filled) else 1


def _last_data_row(ws, header_row):
    last_row = header_row
    for row in ws.iter_rows(min_row=header_row):
        if any(c.value not in (None, "") for c in row):
            last_row = row[0].row
    return last_row


def ler_workbook_completo(source) -> dict:
    """Lê TODAS as abas do workbook, preservando cabeçalho/estrutura reais de
    cada aba (independentes entre si). Retorna {nome_aba: DataFrame}."""
    wb = openpyxl.load_workbook(source, data_only=True)
    abas = {}
    for nome in wb.sheetnames:
        ws = wb[nome]
        header_row = _detect_header_row(ws)
        end_row = _last_data_row(ws, header_row)
        if end_row <= header_row:
            continue

        headers_raw = [c.value for c in ws[header_row]]
        col_idx_validas = [i + 1 for i, h in enumerate(headers_raw) if h not in (None, "")]
        if not col_idx_validas:
            continue
        headers = [headers_raw[i - 1] for i in col_idx_validas]

        rows = []
        for r in range(header_row + 1, end_row + 1):
            vals = [ws.cell(row=r, column=c).value for c in col_idx_validas]
            if all(v in (None, "") for v in vals):
                continue
            rows.append(vals)

        # deduplicar nomes de coluna repetidos na mesma aba (mantém a ordem
        # e o dado real; só evita erro de DataFrame com colunas duplicadas)
        vistos = {}
        headers_final = []
        for h in headers:
            vistos[h] = vistos.get(h, 0) + 1
            headers_final.append(h if vistos[h] == 1 else f"{h} ({vistos[h]})")

        abas[nome] = pd.DataFrame(rows, columns=headers_final)
    return abas


@st.cache_data(ttl=300, show_spinner=False)
def carregar_de_xlsx_publico(url: str):
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
        abas = ler_workbook_completo(io.BytesIO(resp.content))
        return abas if abas else None
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def carregar_de_xlsx_local(path: str):
    try:
        abas = ler_workbook_completo(path)
        return abas if abas else None
    except Exception:
        return None


def carregar_de_upload(arquivo):
    try:
        abas = ler_workbook_completo(arquivo)
        return abas if abas else None
    except Exception as e:
        st.sidebar.error(f"Falha ao ler o arquivo enviado: {e}")
        return None


def obter_dados():
    st.sidebar.markdown("### 📥 Fonte de dados")
    upload = st.sidebar.file_uploader(
        "Upload manual (.xlsx / .xls completo) — alternativa ao GET público",
        type=["xlsx", "xls"]
    )

    if upload is not None:
        abas = carregar_de_upload(upload)
        if abas:
            st.sidebar.success(f"Carregado do upload manual · {len(abas)} aba(s)")
            return abas, "Upload manual (.xlsx completo)"

    abas = carregar_de_xlsx_publico(XLSX_PUBLICO_URL)
    if abas:
        st.sidebar.success(f"Carregado via GET público (workbook completo) · {len(abas)} aba(s)")
        return abas, "GET público — workbook completo (.xlsx)"

    abas = carregar_de_xlsx_local(XLSX_LOCAL_FALLBACK)
    if abas:
        st.sidebar.warning(f"GET público falhou — usando cópia local da planilha real · {len(abas)} aba(s)")
        return abas, "Cópia local da planilha real (fallback)"

    st.sidebar.error("Nenhuma fonte de dados disponível (GET, upload e cópia local falharam).")
    return {}, "Nenhuma"


# ==========================================================================
# PAINEL GERAL (visão de todas as abas — como o router da planilha original)
# ==========================================================================
def montar_painel_geral(abas: dict):
    st.markdown("### 📁 Painel — todas as abas da planilha real")
    linhas = [{"Aba": nome, "Registros": len(df), "Colunas": len(df.columns)}
              for nome, df in abas.items()]
    st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)


# ==========================================================================
# KPIs, FILTROS E GRÁFICOS — por aba selecionada, 100% dinâmico
# ==========================================================================
def montar_kpis(df: pd.DataFrame):
    total = len(df)
    tem_tag = "TAG_TIPOLOGIA" in df.columns

    if tem_tag:
        contagem = df["TAG_TIPOLOGIA"].value_counts(dropna=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total de registros", total)
        c2.metric("🟢 Ativos em curso", int(contagem.get("ATIVO-EM-CURSO", 0)))
        c3.metric("🎓 Concluídos/diplomados", int(contagem.get("CONCLUIDO-DIPLOMADO", 0)))
        c4.metric("⚠️ Ocorrência processual", int(contagem.get("OCORRENCIA-PROCESSUAL", 0)))
        c5.metric("🟠 Pendentes (doc. + comprov.)",
                  int(contagem.get("PENDENTE-DOCUMENTACAO", 0) + contagem.get("PENDENTE-COMPROVACAO", 0)))
    else:
        st.metric("Total de registros nesta aba", total)


def montar_filtros(df: pd.DataFrame, chave: str) -> pd.DataFrame:
    st.sidebar.markdown("### 🔎 Filtros (aba atual)")
    busca = st.sidebar.text_input("Busca livre (todas as colunas)", key=f"busca_{chave}")

    df_filtrado = df.copy()
    for col in COLUNAS_PRIORITARIAS:
        if col in df.columns:
            opcoes = sorted([str(v) for v in df[col].dropna().unique().tolist()])
            if 0 < len(opcoes) <= 300:
                selecionadas = st.sidebar.multiselect(col.title(), opcoes, key=f"filtro_{col}_{chave}")
                if selecionadas:
                    df_filtrado = df_filtrado[df_filtrado[col].astype(str).isin(selecionadas)]

    if busca:
        mask = pd.Series(False, index=df_filtrado.index)
        for col in df_filtrado.columns:
            mask = mask | df_filtrado[col].astype(str).str.contains(busca, case=False, na=False)
        df_filtrado = df_filtrado[mask]

    return df_filtrado


def montar_graficos(df: pd.DataFrame, azul, verde):
    col1, col2 = st.columns(2)
    tem_grafico = False

    if "TAG_TIPOLOGIA" in df.columns:
        cont = df["TAG_TIPOLOGIA"].value_counts(dropna=True).reset_index()
        cont.columns = ["Situação", "Registros"]
        cont["Situação"] = cont["Situação"].map(lambda t: TAG_LABELS.get(t, (t, ""))[0])
        fig1 = px.pie(cont, names="Situação", values="Registros",
                      title="Distribuição por situação (TAG_TIPOLOGIA)",
                      color_discrete_sequence=[azul, verde, "#F59E0B", "#EF4444", "#94A3B8", "#8B5CF6", "#0EA5E9"])
        col1.plotly_chart(fig1, use_container_width=True)
        tem_grafico = True

    if "TIPO DE BOLSA" in df.columns:
        cont2 = df["TIPO DE BOLSA"].value_counts(dropna=True).reset_index()
        cont2.columns = ["Tipo de bolsa", "Registros"]
        fig2 = px.bar(cont2, x="Tipo de bolsa", y="Registros",
                      title="Registros por tipo de bolsa", color_discrete_sequence=[verde])
        col2.plotly_chart(fig2, use_container_width=True)
        tem_grafico = True

    if "COMARCA" in df.columns:
        cont3 = df["COMARCA"].value_counts(dropna=True).reset_index().head(15)
        cont3.columns = ["Comarca", "Registros"]
        fig3 = px.bar(cont3, x="Registros", y="Comarca", orientation="h",
                      title="Top 15 comarcas por nº de registros", color_discrete_sequence=[azul])
        fig3.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig3, use_container_width=True)
        tem_grafico = True

    if not tem_grafico:
        st.caption("Esta aba não possui colunas-padrão (TAG_TIPOLOGIA / TIPO DE BOLSA / COMARCA) para gráfico automático.")


def montar_tabela_e_export(df: pd.DataFrame, nome_aba: str):
    st.markdown(f"### 📋 Registros — {nome_aba}")
    st.caption(f"{len(df)} registro(s) após filtro")
    st.dataframe(df, use_container_width=True, height=420)

    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    slug = nome_aba.strip().replace(" ", "_").replace("/", "-")
    nome_arquivo = f"RH_TJMA_export_{slug}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    st.download_button("⬇️ Exportar CSV (RH TJMA)", data=csv_bytes, file_name=nome_arquivo, mime="text/csv")


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

    abas, origem = obter_dados()
    st.caption(f"Fonte ativa: **{origem}**")

    if not abas:
        rodape()
        return

    montar_painel_geral(abas)

    nomes = list(abas.keys())
    padrao = next((n for n in nomes if "GERAL" in n.upper()), nomes[0])
    aba_selecionada = st.selectbox("Selecionar aba", nomes, index=nomes.index(padrao))

    df = abas[aba_selecionada]
    df_filtrado = montar_filtros(df, chave=aba_selecionada)
    montar_kpis(df_filtrado)
    montar_graficos(df_filtrado, azul, verde)
    montar_tabela_e_export(df_filtrado, aba_selecionada)
    rodape()


if __name__ == "__main__":
    main()
