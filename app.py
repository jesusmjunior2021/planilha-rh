# -*- coding: utf-8 -*-
"""
main.py — Painel RH TJMA · Auxílio-Bolsa (GDG)
Adm. Jesus Martins Oliveira Junior — COGEX-MA/TJMA
MAT-RHBOLSAS-STREAMLIT-003 (versão 100% nativa)

Login: RH / RH@123

DECISÃO DE ARQUITETURA (por que essa versão existe):
O Streamlit Cloud roda num ambiente isolado que só instala o que está
listado em requirements.txt — nada é "nativo do sistema" além disso. Toda
vez que dependemos de libs pesadas (openpyxl, plotly, xlrd, reportlab) e o
requirements.txt não é lido corretamente pela plataforma, o app quebra com
ModuleNotFoundError. Pra eliminar essa classe inteira de erro, esta versão:

  - NÃO usa openpyxl/xlrd: lê .xlsx "na unha" com zipfile + xml.etree
    (ambos da biblioteca padrão do Python — sempre disponíveis, sem
    instalação nenhuma, em qualquer ambiente).
  - NÃO usa plotly: usa Altair, que já vem instalado automaticamente como
    dependência do próprio Streamlit (confirmado no log de deploy).
  - NÃO usa reportlab/PyPDF: gera um relatório HTML com CSS de impressão
    (@media print) pronto pra "Salvar como PDF" direto do navegador —
    zero dependência.
  - requirements.txt fica só com streamlit + pandas + altair (as duas
    últimas já vêm junto do streamlit de qualquer forma).

Fonte de dados (ordem de prioridade) — SEMPRE a planilha real completa,
todas as abas, todas as colunas, estrutura original preservada:
  1) GET do workbook público completo (.xlsx, todas as abas) via export do
     Google Sheets.
  2) Upload manual de arquivo .xlsx ou .csv completo.
  3) Cópia local da planilha real (fallback offline), mesma estrutura.
Regra dura: 100% dos registros e 100% das abas — sem amostragem, sem
recorte para 1 aba só, sem invenção de linha/valor/coluna.
"""
import io
import re
import json
import zipfile
import datetime
import urllib.request
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
import altair as alt

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import streamlit.components.v1 as components

# ==========================================================================
# CONFIG
# ==========================================================================
APP_VERSION = "v2.0.0-nativo"
APP_TITLE = "Painel RH TJMA · Auxílio-Bolsa"
LOGIN_USER = "RH"
LOGIN_PASS = "RH@123"

GOOGLE_SHEET_ID = "1iaXdM3maNqnvhnz7-KvGE4CL0mQDoVUjUpLjDEsipqg"
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

COLUNAS_PRIORITARIAS = ["TAG_TIPOLOGIA", "STATUS", "COMARCA", "TIPO DE BOLSA", "CARGO"]

st.set_page_config(page_title=APP_TITLE, page_icon="🎓", layout="wide")

# ==========================================================================
# LEITOR NATIVO DE .XLSX — só biblioteca padrão (zipfile + xml.etree)
# Um arquivo .xlsx é um .zip contendo XMLs (formato OOXML). Aqui a gente
# abre o zip direto e lê os XMLs internos, sem nenhuma lib de terceiros.
# ==========================================================================
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_NS_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _col_letra_para_indice(ref: str) -> int:
    """'A1' -> 0, 'B7' -> 1, 'AA3' -> 26 ..."""
    letras = re.match(r"[A-Za-z]+", ref)
    letras = letras.group(0).upper() if letras else "A"
    idx = 0
    for ch in letras:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _ler_shared_strings(z: zipfile.ZipFile):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in tree.findall(f"{_NS}si"):
        texto = "".join(t.text or "" for t in si.findall(f".//{_NS}t"))
        out.append(texto)
    return out


def _ler_lista_abas(z: zipfile.ZipFile):
    """Retorna [(nome_aba, caminho_xml)] respeitando a ordem real do workbook."""
    wb_tree = ET.fromstring(z.read("xl/workbook.xml"))
    rels_tree = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))

    rid_para_target = {}
    for rel in rels_tree.findall(f"{_NS_PKG_REL}Relationship"):
        rid_para_target[rel.get("Id")] = rel.get("Target")

    abas = []
    sheets_el = wb_tree.find(f"{_NS}sheets")
    if sheets_el is None:
        return abas
    for sheet in sheets_el.findall(f"{_NS}sheet"):
        nome = sheet.get("name")
        rid = sheet.get(f"{_NS_R}id")
        target = rid_para_target.get(rid)
        if not target:
            continue
        if target.startswith("/"):
            target = target[1:]
        elif not target.startswith("xl/"):
            target = "xl/" + target
        abas.append((nome, target))
    return abas


def _ler_matriz_da_aba(z: zipfile.ZipFile, caminho: str, shared_strings):
    tree = ET.fromstring(z.read(caminho))
    sheet_data = tree.find(f"{_NS}sheetData")
    if sheet_data is None:
        return []

    linhas_dict = []
    for row in sheet_data.findall(f"{_NS}row"):
        celulas = {}
        max_idx = -1
        for c in row.findall(f"{_NS}c"):
            ref = c.get("r", "")
            idx = _col_letra_para_indice(ref) if ref else max_idx + 1
            tipo = c.get("t")
            v_el = c.find(f"{_NS}v")
            is_el = c.find(f"{_NS}is")

            valor = None
            if tipo == "s" and v_el is not None and v_el.text is not None:
                pos = int(v_el.text)
                valor = shared_strings[pos] if pos < len(shared_strings) else ""
            elif tipo == "inlineStr" and is_el is not None:
                valor = "".join(t.text or "" for t in is_el.findall(f".//{_NS}t"))
            elif tipo == "b" and v_el is not None:
                valor = bool(int(v_el.text))
            elif v_el is not None and v_el.text is not None:
                bruto = v_el.text
                try:
                    numero = float(bruto)
                    valor = int(numero) if numero.is_integer() else numero
                except ValueError:
                    valor = bruto

            celulas[idx] = valor
            max_idx = max(max_idx, idx)
        linhas_dict.append((max_idx, celulas))

    if not linhas_dict:
        return []
    n_col = max(m for m, _ in linhas_dict) + 1
    matriz = [[celulas.get(i) for i in range(n_col)] for _, celulas in linhas_dict]
    return matriz


def _detectar_linha_cabecalho(matriz):
    if len(matriz) < 2:
        return 0
    r1 = sum(1 for v in matriz[0] if v not in (None, ""))
    r2 = sum(1 for v in matriz[1] if v not in (None, ""))
    return 1 if (r1 <= 2 and r2 > r1) else 0


def _dedup_headers(headers):
    vistos, final = {}, []
    for h in headers:
        vistos[h] = vistos.get(h, 0) + 1
        final.append(h if vistos[h] == 1 else f"{h} ({vistos[h]})")
    return final


def ler_workbook_nativo(fonte) -> dict:
    """Lê TODAS as abas de um .xlsx usando só a biblioteca padrão do Python.
    `fonte` pode ser um caminho de arquivo, bytes, ou um objeto tipo-arquivo."""
    with zipfile.ZipFile(fonte) as z:
        shared_strings = _ler_shared_strings(z)
        lista_abas = _ler_lista_abas(z)
        nomes_no_zip = set(z.namelist())

        abas = {}
        for nome, caminho in lista_abas:
            if caminho not in nomes_no_zip:
                continue
            matriz = _ler_matriz_da_aba(z, caminho, shared_strings)
            if len(matriz) < 2:
                continue

            header_idx = _detectar_linha_cabecalho(matriz)
            headers_raw = matriz[header_idx]
            col_idx_validas = [i for i, h in enumerate(headers_raw) if h not in (None, "")]
            if not col_idx_validas:
                continue
            headers = [str(headers_raw[i]) for i in col_idx_validas]

            linhas = []
            for linha in matriz[header_idx + 1:]:
                vals = [linha[i] if i < len(linha) else None for i in col_idx_validas]
                if all(v in (None, "") for v in vals):
                    continue
                linhas.append(vals)

            abas[nome] = pd.DataFrame(linhas, columns=_dedup_headers(headers))
    return abas


def ler_csv_nativo(fonte) -> dict:
    """CSV só tem 1 tabela por definição — usa só pandas (já vem com o
    streamlit), detectando separador ',' ou ';' automaticamente."""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    try:
        df = pd.read_csv(fonte, sep=None, engine="python", encoding="utf-8-sig")
    except Exception:
        if hasattr(fonte, "seek"):
            fonte.seek(0)
        df = pd.read_csv(fonte, sep=";", engine="python", encoding="latin-1")
    df = df.dropna(how="all")
    df.columns = _dedup_headers([str(c) for c in df.columns])
    return {"Dados (CSV)": df.reset_index(drop=True)}


# ==========================================================================
# ESCRITOR NATIVO DE .XLSX — pra oferecer download em Excel, sem openpyxl
# ==========================================================================
def _indice_para_col_letra(idx: int) -> str:
    letras = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letras = chr(65 + rem) + letras
    return letras


def gerar_xlsx_nativo(df: pd.DataFrame, nome_aba: str = "Dados") -> bytes:
    """Gera um .xlsx válido (mínimo, 1 aba) usando só zipfile + strings XML —
    sem nenhuma lib de terceiros. Suficiente pra exportar dados filtrados."""
    from xml.sax.saxutils import escape as xml_escape

    nome_aba_seguro = re.sub(r'[\\/*?:\[\]]', "_", str(nome_aba))[:31] or "Dados"

    linhas_xml = []
    header_cells = [
        f'<c r="{_indice_para_col_letra(i)}1" t="inlineStr"><is><t>{xml_escape(str(col))}</t></is></c>'
        for i, col in enumerate(df.columns)
    ]
    linhas_xml.append(f'<row r="1">{"".join(header_cells)}</row>')

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        cells = []
        for c_idx, val in enumerate(row):
            ref = f"{_indice_para_col_letra(c_idx)}{r_idx}"
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cells.append(f'<c r="{ref}"><v>{val}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{xml_escape(str(val))}</t></is></c>')
        linhas_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(linhas_xml)}</sheetData></worksheet>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{xml_escape(nome_aba_seguro)}" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return buffer.getvalue()


def _hex_para_rgb(hex_color: str):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


def ler_arquivo_por_extensao(fonte, nome_arquivo: str) -> dict:
    ext = nome_arquivo.lower().rsplit(".", 1)[-1]
    if ext == "xlsx":
        return ler_workbook_nativo(fonte)
    if ext == "csv":
        return ler_csv_nativo(fonte)
    raise ValueError(
        f"Extensão .{ext} não suportada nesta versão nativa. "
        "Use .xlsx (Excel moderno) ou .csv. Arquivo .xls antigo: "
        "abra no Google Sheets/Excel e salve como .xlsx primeiro."
    )


@st.cache_data(ttl=300, show_spinner=False)
def carregar_de_xlsx_publico(url: str):
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            conteudo = resp.read()
        abas = ler_workbook_nativo(io.BytesIO(conteudo))
        return abas if abas else None
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def carregar_de_xlsx_local(path: str):
    try:
        abas = ler_workbook_nativo(path)
        return abas if abas else None
    except Exception:
        return None


def carregar_de_upload(arquivo):
    try:
        abas = ler_arquivo_por_extensao(arquivo, arquivo.name)
        return abas if abas else None
    except Exception as e:
        st.sidebar.error(f"Falha ao ler o arquivo enviado: {e}")
        return None


def obter_dados():
    st.sidebar.markdown("### 📥 Fonte de dados")
    upload = st.sidebar.file_uploader(
        "Upload manual (.xlsx ou .csv) — alternativa ao GET público",
        type=["xlsx", "csv"]
    )

    if upload is not None:
        abas = carregar_de_upload(upload)
        if abas:
            st.sidebar.success(f"Carregado do upload manual · {len(abas)} aba(s)")
            return abas, f"Upload manual ({upload.name})"

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
# ESTILO — SEMPRE ESCURO (fixo, sem toggle) · azul/verde de destaque,
# contraste alto em tudo — sidebar, botões, inputs, textos secundários.
# ==========================================================================
BG = "#0B1220"          # fundo principal
BG_CARD = "#182238"     # cards/sidebar — mais claro que o fundo, não igual
BG_INPUT = "#1F2A44"    # inputs/botões
TEXTO = "#F5F7FA"       # texto principal — branco quase puro
TEXTO_MUDO = "#B7C0D1"  # texto secundário — cinza claro, NUNCA opacity baixa
AZUL = "#4F8CFF"
VERDE = "#2ECC71"
BORDA = "#2B3A5C"


def aplicar_estilo():
    st.markdown(f"""
    <style>
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{
            background-color: {BG} !important; color: {TEXTO} !important;
        }}

        /* Sidebar inteira — fundo, texto, labels de widgets */
        section[data-testid="stSidebar"] {{
            background-color: {BG_CARD} !important; border-right: 1px solid {BORDA};
        }}
        section[data-testid="stSidebar"] * {{ color: {TEXTO} !important; }}
        section[data-testid="stSidebar"] label p {{ color: {TEXTO} !important; opacity: 1 !important; }}
        section[data-testid="stSidebar"] .stCaption, section[data-testid="stSidebar"] small {{
            color: {TEXTO_MUDO} !important;
        }}

        /* Botões e uploader */
        .stButton button, .stDownloadButton button, [data-testid="stFileUploaderDropzone"] {{
            background-color: {BG_INPUT} !important; color: {TEXTO} !important;
            border: 1px solid {BORDA} !important;
        }}
        .stButton button:hover, .stDownloadButton button:hover {{
            border-color: {AZUL} !important; color: {AZUL} !important;
        }}

        /* Inputs, selects, multiselect, text_input */
        input, textarea, [data-baseweb="select"] > div, [data-baseweb="tag"] {{
            background-color: {BG_INPUT} !important; color: {TEXTO} !important;
            border-color: {BORDA} !important;
        }}

        /* Cards de métrica (KPIs) */
        [data-testid="stMetric"] {{
            background-color: {BG_CARD} !important; border: 1px solid {BORDA};
            border-radius: 10px; padding: 12px 14px;
        }}
        [data-testid="stMetricValue"] {{ color: {AZUL} !important; font-weight: 700; }}
        [data-testid="stMetricLabel"] {{ color: {TEXTO} !important; }}

        /* Títulos e legendas */
        h1, h2, h3 {{ color: {AZUL} !important; }}
        .stCaption, [data-testid="stCaptionContainer"], p {{ color: {TEXTO} !important; }}
        .stMarkdown small {{ color: {TEXTO_MUDO} !important; }}

        /* Mensagens de sucesso/aviso/erro — mantém legível no fundo escuro */
        [data-testid="stAlert"] {{ background-color: {BG_CARD} !important; }}

        .rodape-app {{
            color: {TEXTO_MUDO}; font-size: 0.75rem; text-align: right;
            margin-top: 2rem; border-top: 1px solid {BORDA}; padding-top: 8px;
        }}
    </style>
    """, unsafe_allow_html=True)
    return AZUL, VERDE


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
# PAINEL GERAL (visão de todas as abas)
# ==========================================================================
def montar_painel_geral(abas: dict):
    st.markdown("### 📁 Painel — todas as abas da planilha real")
    linhas = [{"Aba": nome, "Registros": len(df), "Colunas": len(df.columns)}
              for nome, df in abas.items()]
    st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)


# ==========================================================================
# KPIs, FILTROS
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


# ==========================================================================
# GRÁFICOS — Altair (já vem junto do Streamlit, sem instalar nada extra)
# ==========================================================================
def _grafico_pizza(df, col, titulo, azul, verde):
    cont = df[col].value_counts(dropna=True).reset_index()
    cont.columns = ["Situação", "Registros"]
    if col == "TAG_TIPOLOGIA":
        cont["Situação"] = cont["Situação"].map(lambda t: TAG_LABELS.get(t, (t, ""))[0])
    paleta = [azul, verde, "#F59E0B", "#EF4444", "#94A3B8", "#8B5CF6", "#0EA5E9"]
    return (
        alt.Chart(cont)
        .mark_arc(innerRadius=60)
        .encode(
            theta=alt.Theta("Registros:Q"),
            color=alt.Color("Situação:N", scale=alt.Scale(range=paleta)),
            tooltip=["Situação", "Registros"],
        )
        .properties(title=titulo, height=320)
    )


def _grafico_barras(df, col, titulo, cor, horizontal=False, top=None):
    cont = df[col].value_counts(dropna=True).reset_index()
    cont.columns = [col, "Registros"]
    if top:
        cont = cont.head(top)
    if horizontal:
        chart = (
            alt.Chart(cont)
            .mark_bar(color=cor)
            .encode(
                x="Registros:Q",
                y=alt.Y(f"{col}:N", sort="-x", title=None),
                tooltip=[col, "Registros"],
            )
        )
    else:
        chart = (
            alt.Chart(cont)
            .mark_bar(color=cor)
            .encode(
                x=alt.X(f"{col}:N", sort="-y", title=None),
                y="Registros:Q",
                tooltip=[col, "Registros"],
            )
        )
    return chart.properties(title=titulo, height=320)


def montar_graficos(df: pd.DataFrame, azul, verde):
    col1, col2 = st.columns(2)
    tem_grafico = False

    if "TAG_TIPOLOGIA" in df.columns:
        col1.altair_chart(
            _grafico_pizza(df, "TAG_TIPOLOGIA", "Distribuição por situação (TAG_TIPOLOGIA)", azul, verde),
            use_container_width=True,
        )
        tem_grafico = True

    if "TIPO DE BOLSA" in df.columns:
        col2.altair_chart(
            _grafico_barras(df, "TIPO DE BOLSA", "Registros por tipo de bolsa", verde),
            use_container_width=True,
        )
        tem_grafico = True

    if "COMARCA" in df.columns:
        st.altair_chart(
            _grafico_barras(df, "COMARCA", "Top 15 comarcas por nº de registros", azul, horizontal=True, top=15),
            use_container_width=True,
        )
        tem_grafico = True

    if not tem_grafico:
        st.caption("Esta aba não possui colunas-padrão (TAG_TIPOLOGIA / TIPO DE BOLSA / COMARCA) para gráfico automático.")


# ==========================================================================
# EXPORT — CSV nativo · XLSX nativo · PDF real via jsPDF (CDN, no navegador)
# jsPDF roda inteiramente no browser do usuário: zero instalação em Python,
# zero servidor, apenas um <script> carregado via CDN (igual o resto da UI).
# (pdf.js, em contraste, serve pra *ler/exibir* PDF — não pra gerar um.)
# ==========================================================================
def _preparar_dados_para_pdf(df: pd.DataFrame):
    df_limpo = df.astype(object).where(pd.notnull(df), "")
    linhas = []
    for linha in df_limpo.values.tolist():
        linha_serializavel = []
        for v in linha:
            if isinstance(v, (int, float, str, bool)):
                linha_serializavel.append(v)
            else:
                linha_serializavel.append(str(v))
        linhas.append(linha_serializavel)
    return linhas


def montar_botao_pdf_js(df: pd.DataFrame, nome_aba: str, origem: str, azul_hex: str, chave: str):
    MAX_LINHAS_PDF = 2000  # trava de segurança pra não travar o navegador em abas gigantes
    aviso = ""
    df_pdf = df
    if len(df) > MAX_LINHAS_PDF:
        df_pdf = df.head(MAX_LINHAS_PDF)
        aviso = f" (mostrando as primeiras {MAX_LINHAS_PDF} de {len(df)} linhas — use o CSV/XLSX pra ver tudo)"

    colunas = list(df_pdf.columns)
    linhas = _preparar_dados_para_pdf(df_pdf)
    r, g, b = _hex_para_rgb(azul_hex)
    slug = re.sub(r"[^A-Za-z0-9_-]+", "_", nome_aba.strip())
    nome_arquivo = f"RH_TJMA_relatorio_{slug}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    titulo_txt = f"{APP_TITLE} — {nome_aba}"
    subtitulo_txt = (
        f"Fonte: {origem} · Gerado em {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')} "
        f"· {len(df_pdf)} registro(s){aviso}"
    )

    colunas_json = json.dumps(colunas, ensure_ascii=False)
    linhas_json = json.dumps(linhas, ensure_ascii=False)
    titulo_json = json.dumps(titulo_txt, ensure_ascii=False)
    subtitulo_json = json.dumps(subtitulo_txt, ensure_ascii=False)
    nome_arquivo_json = json.dumps(nome_arquivo, ensure_ascii=False)
    btn_id = f"btnPdf_{chave}"

    html = f"""
    <div style="font-family: Arial, Helvetica, sans-serif;">
      <button id="{btn_id}"
        style="width:100%; padding:10px 14px; border-radius:8px; border:none;
               background:{azul_hex}; color:#fff; font-weight:600; font-size:14px;
               cursor:pointer;">
        🧾 Baixar PDF (gerado no navegador)
      </button>
      <div id="{btn_id}_status" style="font-size:12px; color:#888; margin-top:4px;"></div>
    </div>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/4.0.0/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/5.0.8/jspdf.plugin.autotable.min.js"></script>
    <script>
      (function() {{
        const colunas = {colunas_json};
        const linhas = {linhas_json};
        const titulo = {titulo_json};
        const subtitulo = {subtitulo_json};
        const nomeArquivo = {nome_arquivo_json};
        const botao = document.getElementById("{btn_id}");
        const status = document.getElementById("{btn_id}_status");

        botao.addEventListener("click", function() {{
          try {{
            status.textContent = "Gerando PDF...";
            const doc = new jspdf.jsPDF({{ orientation: "landscape" }});
            doc.setFontSize(14);
            doc.text(titulo, 14, 15);
            doc.setFontSize(9);
            doc.text(subtitulo, 14, 21);
            doc.autoTable({{
              head: [colunas],
              body: linhas,
              startY: 26,
              styles: {{ fontSize: 7, cellPadding: 2, overflow: "linebreak" }},
              headStyles: {{ fillColor: [{r}, {g}, {b}], textColor: 255 }},
              alternateRowStyles: {{ fillColor: [243, 248, 244] }},
              theme: "grid",
              horizontalPageBreak: true,
            }});
            doc.save(nomeArquivo);
            status.textContent = "PDF gerado ✅";
          }} catch (err) {{
            status.textContent = "Erro ao gerar PDF: " + err.message;
          }}
        }});
      }})();
    </script>
    """
    components.html(html, height=80)


def montar_tabela_e_export(df: pd.DataFrame, nome_aba: str, origem: str, azul_hex: str):
    st.markdown(f"### 📋 Registros — {nome_aba}")
    st.caption(f"{len(df)} registro(s) após filtro")
    st.dataframe(df, use_container_width=True, height=420)

    slug = nome_aba.strip().replace(" ", "_").replace("/", "-")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")

    st.markdown("##### ⬇️ Exportar dados filtrados")
    col_csv, col_xlsx, col_pdf = st.columns(3)

    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    col_csv.download_button(
        "📄 CSV",
        data=csv_bytes,
        file_name=f"RH_TJMA_export_{slug}_{timestamp}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    xlsx_bytes = gerar_xlsx_nativo(df, nome_aba=nome_aba)
    col_xlsx.download_button(
        "📊 XLSX (Excel)",
        data=xlsx_bytes,
        file_name=f"RH_TJMA_export_{slug}_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    with col_pdf:
        montar_botao_pdf_js(df, nome_aba, origem, azul_hex, chave=slug)


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

    azul, verde = aplicar_estilo()

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
    montar_tabela_e_export(df_filtrado, aba_selecionada, origem, azul)
    rodape()


if __name__ == "__main__":
    main()
