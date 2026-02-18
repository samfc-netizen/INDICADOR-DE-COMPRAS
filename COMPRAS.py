import pandas as pd
import streamlit as st
import plotly.express as px
import re

st.set_page_config(page_title="Indicador de Compras", layout="wide")

EXCEL_PATH = "Indicador de compras.xlsx"

MESES_PT = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}
MESES_LABELS = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]


# -----------------------------
# Utilidades
# -----------------------------
def strip_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def colnorm(name: str) -> str:
    s = str(name).strip().upper()
    s = s.replace("\t", " ")
    s = " ".join(s.split())
    return s

def find_col(df: pd.DataFrame, target_norm: str):
    mapping = {colnorm(c): c for c in df.columns}
    return mapping.get(target_norm)

def to_float(series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)

def to_datetime_safe(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def month_number_from_text(s):
    s = "" if pd.isna(s) else str(s).strip().upper()
    return MESES_PT.get(s, pd.NA)

def parse_mes_to_num(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip().upper()

    try:
        n = int(float(s.replace(",", ".")))
        if 1 <= n <= 12:
            return n
    except Exception:
        pass

    if s in MESES_PT:
        return MESES_PT[s]

    for nome, num in MESES_PT.items():
        if nome in s:
            return num

    m = re.search(r"\b(\d{1,2})\b", s)
    if m:
        try:
            n = int(m.group(1))
            if 1 <= n <= 12:
                return n
        except Exception:
            pass

    return pd.NA

def supplier_key(s: str) -> str:
    s = "" if pd.isna(s) else str(s).strip().upper()
    s = re.sub(r"^\s*\d+\s*-\s*", "", s)
    s = "".join(ch if ch.isalnum() or ch.isspace() else " " for ch in s)
    s = " ".join(s.split())
    return s

def brl(v) -> str:
    try:
        v = float(v)
    except Exception:
        v = 0.0
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

def pct_str(v: float) -> str:
    try:
        s = f"{v*100:,.2f}"
    except Exception:
        s = "0,00"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s}%"

def style_dif(val):
    try:
        v = float(val)
    except Exception:
        v = 0.0
    if v > 0:
        return "color:#0a7a2f;font-weight:900;"
    if v < 0:
        return "color:#b00020;font-weight:900;"
    return "color:#333;font-weight:800;"

def nota_key(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits.lstrip("0")

def find_sheet_name(excel: pd.ExcelFile, wanted: str):
    wn = colnorm(wanted)
    for sh in excel.sheet_names:
        if colnorm(sh) == wn:
            return sh
    for sh in excel.sheet_names:
        if wn in colnorm(sh):
            return sh
    return None

def most_frequent_nonempty(series: pd.Series) -> str:
    s = series.dropna().astype(str).map(lambda x: x.strip()).replace("", pd.NA).dropna()
    if s.empty:
        return ""
    return str(s.value_counts().index[0])


# -----------------------------
# Carregamento
# -----------------------------
@st.cache_data(show_spinner=False)
def load_data(path: str):
    xls = pd.ExcelFile(path)

    sh_cmv = find_sheet_name(xls, "CMV E ESTOQUE")
    sh_ent = find_sheet_name(xls, "NOTAS ENTRADAS")
    sh_citel = find_sheet_name(xls, "NOTAS CITEL")
    sh_sellout = find_sheet_name(xls, "SELLOUT")

    if sh_cmv is None:
        raise ValueError(f"Não encontrei a aba 'CMV E ESTOQUE'. Abas: {xls.sheet_names}")
    if sh_ent is None:
        raise ValueError(f"Não encontrei a aba 'NOTAS ENTRADAS'. Abas: {xls.sheet_names}")
    if sh_citel is None:
        raise ValueError(f"Não encontrei a aba 'NOTAS CITEL'. Abas: {xls.sheet_names}")

    df_cmv = strip_cols(pd.read_excel(xls, sh_cmv))
    df_ent = strip_cols(pd.read_excel(xls, sh_ent))
    df_citel = strip_cols(pd.read_excel(xls, sh_citel))
    df_sellout = strip_cols(pd.read_excel(xls, sh_sellout)) if sh_sellout else None

    # ---- CMV E ESTOQUE ----
    col_cmv_forn = find_col(df_cmv, "FORNECEDOR")
    col_cmv_cmv = find_col(df_cmv, "CMV")
    col_cmv_mes = find_col(df_cmv, "MÊS") or find_col(df_cmv, "MES")
    col_cmv_linha = find_col(df_cmv, "LINHA")
    col_cmv_marca = find_col(df_cmv, "MARCA")
    col_cmv_vlr_est = find_col(df_cmv, "VLR ESTOQUE") or find_col(df_cmv, "VLR_ESTOQUE") or find_col(df_cmv, "VALOR ESTOQUE")

    if col_cmv_forn is None and df_cmv.shape[1] >= 2:
        col_cmv_forn = df_cmv.columns[1]

    if col_cmv_forn is None:
        raise ValueError(f"CMV E ESTOQUE: não encontrei FORNECEDOR. Colunas: {list(df_cmv.columns)}")
    if col_cmv_cmv is None:
        raise ValueError(f"CMV E ESTOQUE: não encontrei CMV. Colunas: {list(df_cmv.columns)}")
    if col_cmv_linha is None:
        raise ValueError(f"CMV E ESTOQUE: não encontrei LINHA. Colunas: {list(df_cmv.columns)}")
    if col_cmv_vlr_est is None:
        raise ValueError(f"CMV E ESTOQUE: não encontrei VLR ESTOQUE. Colunas: {list(df_cmv.columns)}")

    df_cmv = df_cmv.copy()
    df_cmv["FORNECEDOR_CMV"] = df_cmv[col_cmv_forn].astype(str).fillna("").str.strip()
    df_cmv["FORN_KEY"] = df_cmv["FORNECEDOR_CMV"].map(supplier_key)
    df_cmv["CMV_VALOR"] = to_float(df_cmv[col_cmv_cmv])
    df_cmv["ESTOQUE_VALOR"] = to_float(df_cmv[col_cmv_vlr_est])
    df_cmv["LINHA"] = df_cmv[col_cmv_linha].astype(str).fillna("").str.strip()
    df_cmv["MARCA"] = df_cmv[col_cmv_marca].astype(str).fillna("").str.strip() if col_cmv_marca else ""

    if col_cmv_mes is not None:
        df_cmv["MES_NUM"] = df_cmv[col_cmv_mes].map(month_number_from_text)
    else:
        df_cmv["MES_NUM"] = pd.NA

    # ---- NOTAS CITEL ----
    col_citel_forn = find_col(df_citel, "FORNECEDOR")
    col_citel_val = find_col(df_citel, "VL_NOTA_FISCAL")
    col_citel_dt = find_col(df_citel, "DT_EMISSAO")
    col_citel_doc = find_col(df_citel, "NR_DOCUMENTO")

    if col_citel_forn is None:
        raise ValueError(f"NOTAS CITEL: não encontrei FORNECEDOR. Colunas: {list(df_citel.columns)}")
    if col_citel_val is None:
        raise ValueError(f"NOTAS CITEL: não encontrei VL_NOTA_FISCAL. Colunas: {list(df_citel.columns)}")
    if col_citel_dt is None:
        raise ValueError(f"NOTAS CITEL: não encontrei DT_EMISSAO. Colunas: {list(df_citel.columns)}")
    if col_citel_doc is None:
        raise ValueError(f"NOTAS CITEL: não encontrei NR_DOCUMENTO. Colunas: {list(df_citel.columns)}")

    df_citel = df_citel.copy()
    df_citel["FORNECEDOR_CITEL"] = df_citel[col_citel_forn].astype(str).fillna("").str.strip()
    df_citel["FORN_KEY"] = df_citel["FORNECEDOR_CITEL"].map(supplier_key)
    df_citel["COMPRA_VALOR"] = to_float(df_citel[col_citel_val])

    df_citel["DATA_DT"] = to_datetime_safe(df_citel[col_citel_dt])
    df_citel["ANO"] = df_citel["DATA_DT"].dt.year
    df_citel["MES_NUM"] = df_citel["DATA_DT"].dt.month

    df_citel["NR_DOCUMENTO"] = df_citel[col_citel_doc]
    df_citel["NOTA_KEY"] = df_citel["NR_DOCUMENTO"].map(nota_key)

    # ---- NOTAS ENTRADAS ----
    col_ent_data = find_col(df_ent, "DATA")
    col_ent_vr = (
        find_col(df_ent, "VR. CONTÁBIL")
        or find_col(df_ent, "VR CONTÁBIL")
        or find_col(df_ent, "VR CONTABIL")
        or find_col(df_ent, "VR. CONTABIL")
    )
    col_ent_nf = (
        find_col(df_ent, "NR NOTA FISCAL")
        or find_col(df_ent, "NR_NOTA_FISCAL")
        or find_col(df_ent, "NR NOTA")
        or find_col(df_ent, "NOTA FISCAL")
    )
    col_ent_fornecedor = (
        find_col(df_ent, "DESCRIÇÃO")
        or find_col(df_ent, "DESCRICAO")
        or find_col(df_ent, "FORNECEDOR")
    )
    col_ent_marca = find_col(df_ent, "MARCA")
    col_ent_linha = find_col(df_ent, "LINHA")
    col_ent_grupo = find_col(df_ent, "GRUPO") or find_col(df_ent, "GRUPO/SEGMENTO") or find_col(df_ent, "SEGMENTO")

    if col_ent_vr is None:
        raise ValueError(f"NOTAS ENTRADAS: não encontrei VR. CONTÁBIL. Colunas: {list(df_ent.columns)}")
    if col_ent_nf is None:
        raise ValueError(f"NOTAS ENTRADAS: não encontrei 'NR NOTA FISCAL' (ou variações). Colunas: {list(df_ent.columns)}")
    if col_ent_fornecedor is None:
        raise ValueError(f"NOTAS ENTRADAS: não encontrei DESCRIÇÃO/FORNECEDOR. Colunas: {list(df_ent.columns)}")
    if col_ent_linha is None:
        raise ValueError(f"NOTAS ENTRADAS: não encontrei LINHA. Colunas: {list(df_ent.columns)}")

    df_ent = df_ent.copy()
    df_ent["VR_CONTABIL"] = to_float(df_ent[col_ent_vr])
    df_ent["NR_NOTA_FISCAL"] = df_ent[col_ent_nf]
    df_ent["NOTA_KEY"] = df_ent["NR_NOTA_FISCAL"].map(nota_key)
    df_ent["FORNECEDOR_ENT"] = df_ent[col_ent_fornecedor].astype(str).fillna("").str.strip()
    df_ent["FORN_KEY"] = df_ent["FORNECEDOR_ENT"].map(supplier_key)
    df_ent["MARCA"] = df_ent[col_ent_marca].astype(str).fillna("").str.strip() if col_ent_marca else ""
    df_ent["LINHA"] = df_ent[col_ent_linha].astype(str).fillna("").str.strip()
    df_ent["GRUPO"] = df_ent[col_ent_grupo].astype(str).fillna("").str.strip() if col_ent_grupo else ""

    if col_ent_data is not None:
        df_ent["DATA_DT"] = to_datetime_safe(df_ent[col_ent_data])
        df_ent["ANO"] = df_ent["DATA_DT"].dt.year
        df_ent["MES_NUM"] = df_ent["DATA_DT"].dt.month
    else:
        df_ent["DATA_DT"] = pd.NaT
        df_ent["ANO"] = pd.NA
        df_ent["MES_NUM"] = pd.NA

    # ---- SELLOUT ----
    if df_sellout is not None:
        col_so_fat = find_col(df_sellout, "FATURAMENTO")
        col_so_forn = (
            find_col(df_sellout, "FORNECEDOR")
            or find_col(df_sellout, "NM FORNECEDOR")
            or find_col(df_sellout, "NM_FORNECEDOR")
            or find_col(df_sellout, "EMITENTE")
            or find_col(df_sellout, "CLIENTE")
        )
        col_so_marca = find_col(df_sellout, "MARCA")
        col_so_linha = find_col(df_sellout, "LINHA")

        col_so_mes = find_col(df_sellout, "MÊS") or find_col(df_sellout, "MES")
        col_so_ano = find_col(df_sellout, "ANO")
        col_so_data = find_col(df_sellout, "DATA") or find_col(df_sellout, "DT") or find_col(df_sellout, "DT_VENDA")

        col_so_cod = find_col(df_sellout, "CÓDIGO") or find_col(df_sellout, "CODIGO")
        col_so_desc_prod = (
            find_col(df_sellout, "DESCRIÇÃO DO PRODUTO")
            or find_col(df_sellout, "DESCRICAO DO PRODUTO")
            or find_col(df_sellout, "DESCRIÇÃO PRODUTO")
            or find_col(df_sellout, "DESCRICAO PRODUTO")
        )
        col_so_qtd = (
            find_col(df_sellout, "QTD. FATUR")
            or find_col(df_sellout, "QTD FATUR")
            or find_col(df_sellout, "QTDE FATUR")
            or find_col(df_sellout, "QUANTIDADE")
        )

        if col_so_fat is None:
            raise ValueError(f"SELLOUT: não encontrei FATURAMENTO. Colunas: {list(df_sellout.columns)}")
        if col_so_forn is None:
            raise ValueError(f"SELLOUT: não encontrei FORNECEDOR. Colunas: {list(df_sellout.columns)}")
        if col_so_mes is None and col_so_data is None:
            raise ValueError(
                f"SELLOUT: não encontrei coluna MÊS/MES nem DATA. "
                f"Confira se o nome está exatamente 'MÊS'. Colunas: {list(df_sellout.columns)}"
            )

        df_sellout = df_sellout.copy()
        df_sellout["FATURAMENTO"] = to_float(df_sellout[col_so_fat])
        df_sellout["FORNECEDOR_SELLOUT"] = df_sellout[col_so_forn].astype(str).fillna("").str.strip()
        df_sellout["FORN_KEY"] = df_sellout["FORNECEDOR_SELLOUT"].map(supplier_key)
        df_sellout["MARCA"] = df_sellout[col_so_marca].astype(str).fillna("").str.strip() if col_so_marca else ""
        df_sellout["LINHA"] = df_sellout[col_so_linha].astype(str).fillna("").str.strip() if col_so_linha else ""

        df_sellout["CODIGO"] = df_sellout[col_so_cod].astype(str).fillna("").str.strip() if col_so_cod else ""
        df_sellout["DESCRICAO_PRODUTO"] = df_sellout[col_so_desc_prod].astype(str).fillna("").str.strip() if col_so_desc_prod else ""
        df_sellout["QTD_FATUR"] = to_float(df_sellout[col_so_qtd]) if col_so_qtd else 0.0

        df_sellout["DATA_DT"] = pd.NaT
        df_sellout["MES_NUM"] = pd.NA
        df_sellout["ANO"] = pd.NA

        if col_so_data is not None:
            df_sellout["DATA_DT"] = to_datetime_safe(df_sellout[col_so_data])
            if df_sellout["DATA_DT"].notna().any():
                df_sellout["MES_NUM"] = df_sellout["DATA_DT"].dt.month
                df_sellout["ANO"] = df_sellout["DATA_DT"].dt.year

        if col_so_mes is not None:
            df_sellout["MES_NUM"] = df_sellout[col_so_mes].map(parse_mes_to_num)

        if col_so_ano is not None:
            df_sellout["ANO"] = pd.to_numeric(df_sellout[col_so_ano], errors="coerce")

    return df_cmv, df_citel, df_ent, df_sellout


try:
    df_cmv, df_citel, df_ent, df_sellout = load_data(EXCEL_PATH)
except Exception as e:
    st.error(f"Erro ao ler '{EXCEL_PATH}': {e}")
    st.stop()


# -----------------------------
# Sidebar: Página + filtros
# -----------------------------
st.sidebar.title("Navegação")
page = st.sidebar.selectbox("Página", ["COMPRAS", "SELLOUT"])

st.sidebar.divider()
st.sidebar.subheader("Filtros")

anos_citel = sorted(df_citel["ANO"].dropna().astype(int).unique().tolist())
anos_ent = sorted(df_ent["ANO"].dropna().astype(int).unique().tolist())
anos_sellout = []
if df_sellout is not None and "ANO" in df_sellout.columns and df_sellout["ANO"].notna().any():
    anos_sellout = sorted(df_sellout["ANO"].dropna().astype(int).unique().tolist())

anos = sorted(set(anos_citel + anos_ent + anos_sellout))
sel_anos = st.sidebar.multiselect("Ano", options=anos, default=anos)

sel_meses = st.sidebar.multiselect("Mês", options=MESES_LABELS, default=MESES_LABELS)
sel_meses_num = [MESES_PT[m] for m in sel_meses if m in MESES_PT]

def apply_month_year_filter(df, apply_year=True, apply_month=True):
    if df is None:
        return None
    out = df.copy()
    if apply_year and sel_anos and "ANO" in out.columns and out["ANO"].notna().any():
        out = out[out["ANO"].isin(sel_anos)]
    if apply_month and sel_meses_num and "MES_NUM" in out.columns:
        out = out[out["MES_NUM"].isin(sel_meses_num)]
    return out

df_citel_f = apply_month_year_filter(df_citel, apply_year=True, apply_month=True)
df_ent_f = apply_month_year_filter(df_ent, apply_year=True, apply_month=True)
df_cmv_f = apply_month_year_filter(df_cmv, apply_year=False, apply_month=True)
df_sellout_f = apply_month_year_filter(df_sellout, apply_year=True, apply_month=True) if df_sellout is not None else None


# -----------------------------
# PAGE: COMPRAS
# -----------------------------
def render_compras_page():
    st.title("INDICADORES DE COMPRAS")

    # Cards topo
    total_compras_citel = float(df_citel_f["COMPRA_VALOR"].sum())
    total_vendas_cmv = float(df_cmv_f["CMV_VALOR"].sum())
    dif_topo = total_vendas_cmv - total_compras_citel

    if total_compras_citel != 0:
        dif_pct = dif_topo / total_compras_citel
    elif total_vendas_cmv != 0:
        dif_pct = dif_topo / total_vendas_cmv
    else:
        dif_pct = 0.0

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("### TOTAL COMPRAS (CITEL)")
        st.markdown(f"<div style='font-size:28px;font-weight:900'>{brl(total_compras_citel)}</div>", unsafe_allow_html=True)
    with c2:
        st.markdown("### TOTAL VENDAS CMV")
        st.markdown(f"<div style='font-size:28px;font-weight:900'>{brl(total_vendas_cmv)}</div>", unsafe_allow_html=True)
    with c3:
        color = "#0a7a2f" if dif_topo >= 0 else "#b00020"
        st.markdown("### DIFERENÇA (CMV - COMPRAS)")
        st.markdown(
            f"<div style='font-size:28px;font-weight:1000;color:{color}'>{brl(dif_topo)} "
            f"<span style='font-size:18px;font-weight:900'>({pct_str(dif_pct)})</span></div>",
            unsafe_allow_html=True
        )

    st.divider()

    # Tabela por fornecedor (CITEL x CMV)
    st.subheader("Tabela por Fornecedor — Compras (CITEL) x CMV (Autcom)")

    nome_canon = (
        df_cmv_f.groupby(["FORN_KEY", "FORNECEDOR_CMV"], as_index=False)
        .size()
        .sort_values(["FORN_KEY", "size"], ascending=[True, False])
        .drop_duplicates("FORN_KEY")[["FORN_KEY", "FORNECEDOR_CMV"]]
        .rename(columns={"FORNECEDOR_CMV": "FORNECEDOR"})
    )

    vendas = df_cmv_f.groupby("FORN_KEY", as_index=False).agg(**{"VENDAS CMV": ("CMV_VALOR", "sum")})
    compras = df_citel_f.groupby("FORN_KEY", as_index=False).agg(**{"COMPRAS FORNECEDOR": ("COMPRA_VALOR", "sum")})

    tab = nome_canon.merge(vendas, on="FORN_KEY", how="left").merge(compras, on="FORN_KEY", how="left")
    tab["VENDAS CMV"] = tab["VENDAS CMV"].fillna(0.0)
    tab["COMPRAS FORNECEDOR"] = tab["COMPRAS FORNECEDOR"].fillna(0.0)
    tab["DIF (CMV - COMPRAS)"] = tab["VENDAS CMV"] - tab["COMPRAS FORNECEDOR"]
    tab = tab[~((tab["VENDAS CMV"] == 0) & (tab["COMPRAS FORNECEDOR"] == 0))].copy()
    tab = tab[["FORNECEDOR", "COMPRAS FORNECEDOR", "VENDAS CMV", "DIF (CMV - COMPRAS)"]].sort_values("COMPRAS FORNECEDOR", ascending=False)

    st.dataframe(
        tab.style
          .format({"COMPRAS FORNECEDOR": brl, "VENDAS CMV": brl, "DIF (CMV - COMPRAS)": brl})
          .applymap(style_dif, subset=["DIF (CMV - COMPRAS)"]),
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    # Conciliação CITEL x ENTRADAS
    st.subheader("Conciliação de Compras: CITEL x ENTRADAS")

    total_compras_entradas = float(df_ent_f["VR_CONTABIL"].sum())
    dif_citel_vs_ent = total_compras_citel - total_compras_entradas
    color2 = "#0a7a2f" if dif_citel_vs_ent >= 0 else "#b00020"

    d1, d2, d3 = st.columns(3)
    with d1:
        st.markdown("### TOTAL COMPRAS (CITEL)")
        st.markdown(f"<div style='font-size:26px;font-weight:900'>{brl(total_compras_citel)}</div>", unsafe_allow_html=True)
    with d2:
        st.markdown("### TOTAL COMPRAS (ENTRADAS)")
        st.markdown(f"<div style='font-size:26px;font-weight:900'>{brl(total_compras_entradas)}</div>", unsafe_allow_html=True)
    with d3:
        st.markdown("### DIFERENÇA (CITEL - ENTRADAS)")
        st.markdown(f"<div style='font-size:26px;font-weight:1000;color:{color2}'>{brl(dif_citel_vs_ent)}</div>", unsafe_allow_html=True)

    st.divider()

    # Nuvem de notas (CITEL NR_DOCUMENTO vs ENTRADAS NR NOTA FISCAL)
    st.subheader("Notas no CITEL que não constam em ENTRADAS (por Número da Nota)")

    set_citel = set(df_citel_f["NOTA_KEY"].dropna().astype(str).tolist())
    set_ent = set(df_ent_f["NOTA_KEY"].dropna().astype(str).tolist())
    missing = sorted([k for k in set_citel if k and (k not in set_ent)])

    st.caption(
        f"Comparação direta: CITEL (NR_DOCUMENTO) vs ENTRADAS (NR NOTA FISCAL), normalizados (só dígitos, sem zeros). "
        f"Encontradas **{len(missing)}** notas no CITEL que não aparecem em ENTRADAS no recorte selecionado."
    )

    if len(missing) == 0:
        st.success("Nenhuma nota pendente: todas as notas do CITEL aparecem em ENTRADAS (no recorte selecionado).")
    else:
        max_show = 250
        show = missing[:max_show]
        extra = len(missing) - len(show)

        tags_html = """
        <style>
          .tagwrap { line-height: 2.2; }
          .tag {
            display:inline-block;
            padding: 4px 10px;
            margin: 4px 6px 0 0;
            border-radius: 16px;
            background: #f2f2f2;
            font-size: 13px;
            font-weight: 800;
          }
        </style>
        <div class="tagwrap">
        """
        for d in show:
            tags_html += f"<span class='tag'>{d}</span>"
        tags_html += "</div>"
        if extra > 0:
            tags_html += f"<div style='margin-top:10px;font-weight:800'>+{extra} notas (não exibidas)</div>"

        st.markdown(tags_html, unsafe_allow_html=True)

        with st.expander("Ver lista completa (tabela)"):
            det = df_citel_f[df_citel_f["NOTA_KEY"].isin(missing)].copy()
            det["DATA_EMISSAO"] = det["DATA_DT"].dt.strftime("%d/%m/%Y")
            det_view = det[["NOTA_KEY", "NR_DOCUMENTO", "FORNECEDOR_CITEL", "DATA_EMISSAO", "COMPRA_VALOR"]].rename(
                columns={
                    "NOTA_KEY": "CHAVE_NOTA",
                    "NR_DOCUMENTO": "NR_DOCUMENTO (CITEL)",
                    "FORNECEDOR_CITEL": "FORNECEDOR",
                    "COMPRA_VALOR": "VALOR (VL_NOTA_FISCAL)"
                }
            )
            det_view = det_view.sort_values(["FORNECEDOR", "DATA_EMISSAO", "NR_DOCUMENTO (CITEL)"])
            st.dataframe(
                det_view.style.format({"VALOR (VL_NOTA_FISCAL)": brl}),
                use_container_width=True,
                hide_index=True
            )

    st.divider()

    # ==========================================================
    # ✅ DRILL — Fornecedor → Marca → Linha → Grupo | ENTRADAS x CMV/Estoque
    # + Treemap + Totais do recorte
    # ==========================================================
    st.subheader("Drill — Fornecedor → Marca → Linha → Grupo | Compras (ENTRADAS) x CMV/Estoque (CMV E ESTOQUE)")

    if df_ent_f.empty:
        st.info("Sem dados em NOTAS ENTRADAS no recorte selecionado.")
        return

    # Lista de fornecedores (ENTRADAS)
    forn_list = (
        df_ent_f.groupby(["FORN_KEY", "FORNECEDOR_ENT"], as_index=False)
        .agg(TOTAL=("VR_CONTABIL", "sum"))
        .sort_values("TOTAL", ascending=False)
    )

    options_forn = ["(Todos)"] + forn_list["FORNECEDOR_ENT"].tolist()
    sel_forn = st.selectbox(
        "Selecione o Fornecedor (ENTRADAS → DESCRIÇÃO) ou (Todos)",
        options=options_forn,
        index=0,
        key="drill_forn_select"
    )

    if sel_forn == "(Todos)":
        ent_base = df_ent_f.copy()
        cmv_base = df_cmv_f.copy()
        sel_forn_key = None
    else:
        sel_forn_key = forn_list.loc[forn_list["FORNECEDOR_ENT"] == sel_forn, "FORN_KEY"].iloc[0]
        ent_base = df_ent_f[df_ent_f["FORN_KEY"] == sel_forn_key].copy()
        cmv_base = df_cmv_f[df_cmv_f["FORN_KEY"] == sel_forn_key].copy()

    # Filtro de marcas (somente ENTRADAS)
    marcas = sorted([m for m in ent_base["MARCA"].dropna().astype(str).unique().tolist() if m.strip() != ""])
    sel_marcas = st.multiselect("Filtrar Marcas (ENTRADAS)", options=marcas, default=marcas, key="drill_marcas_multiselect")
    if sel_marcas:
        ent_base = ent_base[ent_base["MARCA"].isin(sel_marcas)].copy()

    if ent_base.empty:
        st.warning("Esse recorte ficou sem dados em ENTRADAS (verifique o fornecedor/marcas).")
        return

    has_grupo = ent_base["GRUPO"].astype(str).str.strip().ne("").any()
    group_cols = ["MARCA", "LINHA"] + (["GRUPO"] if has_grupo else [])

    # Compras por bloco (MARCA/LINHA/GRUPO)
    ent_agg = ent_base.groupby(group_cols, as_index=False).agg(COMPRAS=("VR_CONTABIL", "sum"))

    # Compras totais por LINHA (para rateio)
    comp_por_linha = ent_agg.groupby("LINHA", as_index=False).agg(COMPRAS_LINHA=("COMPRAS", "sum"))

    # CMV/Estoque por LINHA
    cmv_agg_linha = (
        cmv_base.groupby("LINHA", as_index=False)
        .agg(
            VENDAS_CMV_LINHA=("CMV_VALOR", "sum"),
            ESTOQUE_LINHA=("ESTOQUE_VALOR", "sum"),
        )
    )

    dr = ent_agg.merge(comp_por_linha, on="LINHA", how="left").merge(cmv_agg_linha, on="LINHA", how="left")
    dr["VENDAS_CMV_LINHA"] = dr["VENDAS_CMV_LINHA"].fillna(0.0)
    dr["ESTOQUE_LINHA"] = dr["ESTOQUE_LINHA"].fillna(0.0)
    dr["COMPRAS_LINHA"] = dr["COMPRAS_LINHA"].fillna(0.0)

    # Rateio de CMV/Estoque para MARCA/GRUPO dentro da LINHA proporcional às compras
    dr["VENDAS_CMV"] = 0.0
    dr["VLR_ESTOQUE"] = 0.0
    mask = dr["COMPRAS_LINHA"] > 0
    dr.loc[mask, "VENDAS_CMV"] = dr.loc[mask, "VENDAS_CMV_LINHA"] * (dr.loc[mask, "COMPRAS"] / dr.loc[mask, "COMPRAS_LINHA"])
    dr.loc[mask, "VLR_ESTOQUE"] = dr.loc[mask, "ESTOQUE_LINHA"] * (dr.loc[mask, "COMPRAS"] / dr.loc[mask, "COMPRAS_LINHA"])

    dr["DIF (CMV - COMPRAS)"] = dr["VENDAS_CMV"] - dr["COMPRAS"]

    total_comp = float(dr["COMPRAS"].sum())
    total_vend = float(dr["VENDAS_CMV"].sum())
    total_est = float(dr["VLR_ESTOQUE"].sum())

    dr["PART_COMPRA_%"] = (dr["COMPRAS"] / total_comp) if total_comp != 0 else 0.0
    dr["PART_VENDA_%"] = (dr["VENDAS_CMV"] / total_vend) if total_vend != 0 else 0.0
    dr["PART_ESTOQUE_%"] = (dr["VLR_ESTOQUE"] / total_est) if total_est != 0 else 0.0

    cols_show = group_cols + ["COMPRAS", "VENDAS_CMV", "VLR_ESTOQUE", "DIF (CMV - COMPRAS)", "PART_COMPRA_%", "PART_VENDA_%", "PART_ESTOQUE_%"]
    dr_show = dr[cols_show].sort_values("COMPRAS", ascending=False)

    st.dataframe(
        dr_show.style
          .format({
              "COMPRAS": brl,
              "VENDAS_CMV": brl,
              "VLR_ESTOQUE": brl,
              "DIF (CMV - COMPRAS)": brl,
              "PART_COMPRA_%": lambda x: pct_str(float(x)),
              "PART_VENDA_%": lambda x: pct_str(float(x)),
              "PART_ESTOQUE_%": lambda x: pct_str(float(x)),
          })
          .applymap(style_dif, subset=["DIF (CMV - COMPRAS)"]),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("##### Participação por Linha (Mapa / Treemap)")

    path_cols = ["MARCA"]
    if has_grupo:
        path_cols.append("GRUPO")
    path_cols.append("LINHA")

    g1, g2, g3 = st.columns(3)
    with g1:
        fig_comp = px.treemap(dr, path=path_cols, values="COMPRAS", title="Compras (ENTRADAS)")
        fig_comp.update_layout(margin=dict(t=50, l=10, r=10, b=10))
        st.plotly_chart(fig_comp, use_container_width=True)
    with g2:
        fig_vend = px.treemap(dr, path=path_cols, values="VENDAS_CMV", title="Vendas (CMV)")
        fig_vend.update_layout(margin=dict(t=50, l=10, r=10, b=10))
        st.plotly_chart(fig_vend, use_container_width=True)
    with g3:
        fig_est = px.treemap(dr, path=path_cols, values="VLR_ESTOQUE", title="Valor de Estoque (CMV E ESTOQUE)")
        fig_est.update_layout(margin=dict(t=50, l=10, r=10, b=10))
        st.plotly_chart(fig_est, use_container_width=True)

    # Totais do recorte do Drill
    t1, t2, t3 = st.columns(3)
    with t1:
        st.markdown("#### TOTAL COMPRAS (recorte Drill)")
        st.markdown(f"<div style='font-size:26px;font-weight:900'>{brl(total_comp)}</div>", unsafe_allow_html=True)
    with t2:
        st.markdown("#### TOTAL VENDAS CMV (recorte Drill)")
        st.markdown(f"<div style='font-size:26px;font-weight:900'>{brl(total_vend)}</div>", unsafe_allow_html=True)
    with t3:
        st.markdown("#### TOTAL VALOR ESTOQUE (recorte Drill)")
        st.markdown(f"<div style='font-size:26px;font-weight:900'>{brl(total_est)}</div>", unsafe_allow_html=True)

    st.divider()

    # ==========================================================
    # ✅ NOVO: Estoque por Fornecedor (CMV E ESTOQUE) + Drill por LINHA
    # (SEM alterar funcionalidades existentes)
    # ==========================================================
    st.subheader("VALOR DE ESTOQUE")

    # Nome canônico do fornecedor (CMV)
    nome_canon_cmv = (
        df_cmv_f.groupby(["FORN_KEY", "FORNECEDOR_CMV"], as_index=False)
        .size()
        .sort_values(["FORN_KEY", "size"], ascending=[True, False])
        .drop_duplicates("FORN_KEY")[["FORN_KEY", "FORNECEDOR_CMV"]]
        .rename(columns={"FORNECEDOR_CMV": "FORNECEDOR"})
    )

    est_forn = df_cmv_f.groupby("FORN_KEY", as_index=False).agg(VLR_ESTOQUE=("ESTOQUE_VALOR", "sum"))
    est_forn = nome_canon_cmv.merge(est_forn, on="FORN_KEY", how="left")
    est_forn["VLR_ESTOQUE"] = est_forn["VLR_ESTOQUE"].fillna(0.0)

    total_est_geral = float(est_forn["VLR_ESTOQUE"].sum())
    est_forn["PART_ESTOQUE_%"] = (est_forn["VLR_ESTOQUE"] / total_est_geral) if total_est_geral != 0 else 0.0
    est_forn = est_forn.sort_values("VLR_ESTOQUE", ascending=False)

    st.dataframe(
        est_forn[["FORNECEDOR", "VLR_ESTOQUE", "PART_ESTOQUE_%"]]
        .style.format({
            "VLR_ESTOQUE": brl,
            "PART_ESTOQUE_%": lambda x: pct_str(float(x)),
        }),
        use_container_width=True,
        hide_index=True
    )

    # Drill: selecionar fornecedor e mostrar linhas e participação dentro do fornecedor
    options_est = ["(Selecione)"] + est_forn["FORNECEDOR"].astype(str).tolist()
    sel_est_forn = st.selectbox("Selecionar Fornecedor (para detalhar Linhas)", options=options_est, index=0, key="estoque_forn_select")

    if sel_est_forn != "(Selecione)":
        sel_key = est_forn.loc[est_forn["FORNECEDOR"] == sel_est_forn, "FORN_KEY"].iloc[0]
        base = df_cmv_f[df_cmv_f["FORN_KEY"] == sel_key].copy()

        est_total_f = float(base["ESTOQUE_VALOR"].sum())
        by_linha = (
            base.groupby("LINHA", as_index=False)
            .agg(VLR_ESTOQUE=("ESTOQUE_VALOR", "sum"))
            .sort_values("VLR_ESTOQUE", ascending=False)
        )
        by_linha["PART_NO_FORNECEDOR_%"] = (by_linha["VLR_ESTOQUE"] / est_total_f) if est_total_f != 0 else 0.0

        st.markdown("#### Linhas do Fornecedor — Valor de Estoque + Participação no Fornecedor (%)")
        st.dataframe(
            by_linha.style.format({
                "VLR_ESTOQUE": brl,
                "PART_NO_FORNECEDOR_%": lambda x: pct_str(float(x)),
            }),
            use_container_width=True,
            hide_index=True
        )


# -----------------------------
# PAGE: SELLOUT
# -----------------------------
def render_sellout_page():
    st.title("Indicadores de Sellout ")

    if df_sellout_f is None:
        st.warning("Aba **SELLOUT** não encontrada neste Excel.")
        return

    if df_sellout_f.empty:
        st.info("Sem dados em SELLOUT no recorte selecionado.")
        return

    # Resumo: Sellout x CMV por fornecedor
    so_forn = df_sellout_f.groupby("FORN_KEY", as_index=False).agg(FATURAMENTO_SELLOUT=("FATURAMENTO", "sum"))
    cmv_forn_sum = df_cmv_f.groupby("FORN_KEY", as_index=False).agg(CMV=("CMV_VALOR", "sum"))

    nome_cmv = (
        df_cmv_f.groupby(["FORN_KEY", "FORNECEDOR_CMV"], as_index=False)
        .size()
        .sort_values(["FORN_KEY", "size"], ascending=[True, False])
        .drop_duplicates("FORN_KEY")[["FORN_KEY", "FORNECEDOR_CMV"]]
        .rename(columns={"FORNECEDOR_CMV": "FORNECEDOR"})
    )
    nome_so = (
        df_sellout_f.groupby(["FORN_KEY", "FORNECEDOR_SELLOUT"], as_index=False)
        .size()
        .sort_values(["FORN_KEY", "size"], ascending=[True, False])
        .drop_duplicates("FORN_KEY")[["FORN_KEY", "FORNECEDOR_SELLOUT"]]
        .rename(columns={"FORNECEDOR_SELLOUT": "FORNECEDOR"})
    )

    sell_tab = so_forn.merge(cmv_forn_sum, on="FORN_KEY", how="left")
    sell_tab["CMV"] = sell_tab["CMV"].fillna(0.0)
    sell_tab = sell_tab.merge(nome_cmv, on="FORN_KEY", how="left")
    sell_tab = sell_tab.merge(nome_so, on="FORN_KEY", how="left", suffixes=("", "_SO"))
    sell_tab["FORNECEDOR"] = sell_tab["FORNECEDOR"].fillna(sell_tab["FORNECEDOR_SO"]).fillna("")

    sell_tab["MARKUP"] = sell_tab.apply(lambda r: (r["FATURAMENTO_SELLOUT"] / r["CMV"]) if r["CMV"] != 0 else 0.0, axis=1)
    total_sellout = float(sell_tab["FATURAMENTO_SELLOUT"].sum())
    sell_tab["PART_FORNECEDOR_%"] = (sell_tab["FATURAMENTO_SELLOUT"] / total_sellout) if total_sellout != 0 else 0.0

    sell_tab = sell_tab[["FORNECEDOR", "FATURAMENTO_SELLOUT", "CMV", "MARKUP", "PART_FORNECEDOR_%"]].sort_values("FATURAMENTO_SELLOUT", ascending=False)

    st.subheader("Resumo — Fornecedor | Faturamento | CMV | Markup | Participação")
    st.dataframe(
        sell_tab.style.format({
            "FATURAMENTO_SELLOUT": brl,
            "CMV": brl,
            "MARKUP": lambda x: f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
            "PART_FORNECEDOR_%": lambda x: pct_str(float(x)),
        }),
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    # Drill por LINHA (fornecedores + produtos)
    st.subheader("Drill por LINHA — Fornecedores e Produtos (participações %)")

    linhas_all = sorted([x for x in df_sellout_f["LINHA"].dropna().astype(str).unique().tolist() if x.strip() != ""])
    if not linhas_all:
        st.info("Sem LINHA preenchida no SELLOUT para o recorte selecionado.")
        return

    sel_line_global = st.selectbox("Selecione a LINHA (SELLOUT)", options=linhas_all, index=0, key="sellout_line_global_select")

    so_line_all = df_sellout_f[df_sellout_f["LINHA"].astype(str) == str(sel_line_global)].copy()
    total_line = float(so_line_all["FATURAMENTO"].sum())

    by_forn = (
        so_line_all.groupby(["FORN_KEY", "FORNECEDOR_SELLOUT"], as_index=False)
        .agg(FATURAMENTO=("FATURAMENTO", "sum"))
        .sort_values("FATURAMENTO", ascending=False)
    )
    by_forn["% FORNECEDOR / LINHA"] = (by_forn["FATURAMENTO"] / total_line) if total_line != 0 else 0.0

    st.markdown("#### Fornecedores da Linha")
    st.dataframe(
        by_forn[["FORNECEDOR_SELLOUT", "FATURAMENTO", "% FORNECEDOR / LINHA"]]
            .rename(columns={"FORNECEDOR_SELLOUT": "FORNECEDOR"})
            .style.format({
                "FATURAMENTO": brl,
                "% FORNECEDOR / LINHA": lambda x: pct_str(float(x)),
            }),
        use_container_width=True,
        hide_index=True
    )

    desc_canon_all = (
        so_line_all.groupby("CODIGO", as_index=False)["DESCRICAO_PRODUTO"]
        .apply(most_frequent_nonempty)
        .rename(columns={"DESCRICAO_PRODUTO": "DESCRIÇÃO DO PRODUTO"})
    )

    prod_all = (
        so_line_all.groupby("CODIGO", as_index=False)
        .agg(FATURAMENTO=("FATURAMENTO", "sum"), QTD_FATUR=("QTD_FATUR", "sum"))
        .sort_values("FATURAMENTO", ascending=False)
    )
    prod_all = prod_all.merge(desc_canon_all, on="CODIGO", how="left")
    prod_all["DESCRIÇÃO DO PRODUTO"] = prod_all["DESCRIÇÃO DO PRODUTO"].fillna("")
    prod_all["% PRODUTO / LINHA"] = (prod_all["FATURAMENTO"] / total_line) if total_line != 0 else 0.0

    st.markdown("#### Produtos da Linha (participação %)")
    st.dataframe(
        prod_all[["CODIGO", "DESCRIÇÃO DO PRODUTO", "FATURAMENTO", "QTD_FATUR", "% PRODUTO / LINHA"]]
            .style.format({
                "FATURAMENTO": brl,
                "QTD_FATUR": lambda x: f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "% PRODUTO / LINHA": lambda x: pct_str(float(x)),
            }),
        use_container_width=True,
        hide_index=True
    )


# -----------------------------
# Render
# -----------------------------
if page == "COMPRAS":
    render_compras_page()
else:
    render_sellout_page()
