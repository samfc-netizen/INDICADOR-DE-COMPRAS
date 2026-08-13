import pandas as pd
import streamlit as st
import plotly.express as px
import re
import unicodedata
import os

st.set_page_config(page_title="Indicador de Compras", layout="wide")
APP_VERSION = "2026-08-13.12 — Filtro mensal checklist + página Orçamento"
FILTER_STATE_VERSION = "orcamento-checklist-v4"


GIRO_NOTAS_PATH = "GIRO E NOTAS.xlsx"
CAD_FORNECEDORES_PATH = "CADASTRO DE FORNECEDORES.csv"
CAD_PRODUTOS_PATH = "CADASTRO PRODUTOS GERAL.csv"
SELLOUT_PATH = "sellout.csv"
NOTAS_ENTRADA_PATH = "NOTAS DE ENTRADA.csv"

# Itens administrativos/serviços que não devem compor os indicadores comerciais.
EXCLUDED_PRODUCT_KEYS = {"22940"}  # PRESTAÇÃO DE SERVIÇOS

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

def parse_brl_value(value) -> float:
    """Converte texto monetário brasileiro (ex.: R$ 1.770.000,00) em float."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    txt = str(value).strip()
    if not txt:
        return 0.0
    txt = txt.replace("R$", "").replace(" ", "")
    if "," in txt:
        txt = txt.replace(".", "").replace(",", ".")
    else:
        txt = re.sub(r"[^0-9.\-]", "", txt)
    try:
        return float(txt)
    except Exception:
        return 0.0


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
    """Normaliza número da NF sem incorporar o sufixo do documento de entrada.

    Exemplos:
      000000091642-NF-009 -> 91642
      91642.0             -> 91642
      000000091642        -> 91642
    """
    if pd.isna(x):
        return ""

    s = str(x).strip()
    if not s:
        return ""

    # Em NOTAS DE ENTRADA, tudo após -NF- identifica espécie/empresa,
    # não faz parte do número fiscal. Aceita espaços e variações de caixa.
    s = re.split(r"\s*-\s*NF\s*-", s, maxsplit=1, flags=re.IGNORECASE)[0]

    # Excel pode entregar o documento numérico como 91642.0.
    s = re.sub(r"[.,]0+$", "", s)

    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    return digits.lstrip("0") or "0"

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
def read_report_csv(path: str, required_terms=(), source_token=None):
    """Lê CSV exportado pelo Citel, localizando automaticamente a linha do cabeçalho."""
    encodings = ("utf-8-sig", "latin1", "cp1252")
    last_error = None
    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding, errors="strict") as arq:
                linhas = arq.readlines()
            header_idx = None
            terms = [colnorm(x) for x in required_terms]
            for i, linha in enumerate(linhas[:250]):
                norm = colnorm(linha.replace(";", " "))
                if terms and all(t in norm for t in terms):
                    header_idx = i
                    break
            if header_idx is None:
                # fallback: primeira linha com vários separadores
                for i, linha in enumerate(linhas[:250]):
                    if linha.count(";") >= 3:
                        header_idx = i
                        break
            if header_idx is None:
                raise ValueError("Cabeçalho tabular não localizado.")
            return strip_cols(pd.read_csv(
                path, sep=";", skiprows=header_idx, encoding=encoding,
                decimal=",", thousands=".", dtype=str, engine="python"
            ))
        except Exception as exc:
            last_error = exc
    raise ValueError(f"Não foi possível ler {path}: {last_error}")


def only_digits(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\D", "", str(value))


def cnpj_key(value) -> str:
    """Normaliza CNPJ para 14 dígitos quando a informação está íntegra."""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if not s:
        return ""

    # Formato textual comum: 45.985.371/0001-08
    digits = re.sub(r"\D", "", s)
    if len(digits) <= 14 and ("E" not in s.upper()):
        return digits.zfill(14) if digits else ""

    # Número simples eventualmente lido como 45985371000108.0
    try:
        normalized = s.replace(".", "").replace(",", ".") if "," in s else s
        number = float(normalized)
        if number >= 0:
            return str(int(round(number))).zfill(14)[-14:]
    except Exception:
        pass
    return ""


def supplier_name_key(value) -> str:
    """Chave robusta para comparar razões sociais, ignorando acentos e sufixos jurídicos."""
    if pd.isna(value):
        return ""
    s = str(value).strip().upper()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    s = re.sub(r"^\s*\d+\s*[-–:]\s*", "", s)
    s = re.sub(r"\b(LTDA|LIMITADA|S/?A|SA|EIRELI|MEI|ME|EPP)\b\.?", " ", s)
    s = re.sub(r"\b(INDUSTRIA|IND|COMERCIO|COM|IMPORTACAO|EXPORTACAO)\b", " ", s)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return " ".join(s.split())


def product_key(value) -> str:
    """Normaliza códigos como 020004.84658 para 84658 e 84658.0 para 84658."""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if "." in s and re.fullmatch(r"\d+\.\d+", s):
        left, right = s.split(".", 1)
        # Código Citel composto: prefixo.Produto
        if len(right) >= 4 and not set(right) <= {"0"}:
            s = right
        elif set(right) <= {"0"}:
            s = left
    digits = only_digits(s)
    return digits.lstrip("0") or "0"


def parse_number_br(series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    s = series.fillna("").astype(str).str.strip()
    both = s.str.contains(",", regex=False) & s.str.contains(".", regex=False)
    s.loc[both] = s.loc[both].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    comma = s.str.contains(",", regex=False) & ~s.str.contains(".", regex=False)
    s.loc[comma] = s.loc[comma].str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def canonical_supplier(name, brand="") -> str:
    """Consolida razão social e aplica as divisões comerciais solicitadas."""
    raw = "" if pd.isna(name) else str(name).upper().strip()
    brand_n = "" if pd.isna(brand) else str(brand).upper().strip()
    raw = re.sub(r"^\s*\d+\s*[-–:]\s*", "", raw)
    raw = re.sub(r"\b(LTDA|S\s*/?\s*A|SA|EIRELI|ME|EPP)\b\.?", " ", raw)
    raw = re.sub(r"[^A-Z0-9À-Ü]+", " ", raw)
    raw = " ".join(raw.split())

    # Equivalências confirmadas entre razões sociais, abreviações, alterações
    # societárias e artefatos de HTML (&amp; convertido em AMP).
    alias_key = unicodedata.normalize("NFKD", raw).encode("ASCII", "ignore").decode("ASCII")

    # Consolidação explícita ATLAS. O sufixo societário (S/A, S A, SA) já foi
    # removido acima; por isso tratamos as formas normalizadas diretamente.
    # Isso garante que compras, CMV, sellout, histórico e orçamento usem a
    # mesma chave de fornecedor.
    if alias_key in {"ATLAS", "ATLAS S A", "PINCEIS ATLAS", "PINCEIS ATLAS S A"}:
        return "PINCEIS ATLAS S A"

    supplier_aliases = {
        "A AMP S TECHNOLOGIES IND E COM": "A S TECHNOLOGIES INDUSTRIA E COMERCIO",
        "A S TECHNOLOGIES INDUSTRIA E COMERCIO": "A S TECHNOLOGIES INDUSTRIA E COMERCIO",
        "A S TECHNOLOGIES INDUSTRIA E COMERCIO S A": "A S TECHNOLOGIES INDUSTRIA E COMERCIO",
        "SIKA": "SIKA",
        "SIKA S A": "SIKA",
        "ARCOM": "ARCOM",
        "ARCOM S A": "ARCOM",
        "MARLUVAS EQUIPAMENTOS DE SEGURANCA": "MARLUVAS EQUIPAMENTOS DE SEGURANCA",
        "LIMA PERGHER INDUSTRIAL E COMERCIO": "LIMA PERGHER INDUSTRIA E COMERCIO",
        "LIMA PERGHERN INDUSTRIA E COMERCIO": "LIMA PERGHER INDUSTRIA E COMERCIO",
        "J PROLAB IND E COM DE PRODUTOS PARA LABORATORIO": "J PROLAB INDUSTRIA E COMERCIO DE PRODUTOS PARA LABORATORIO",
        "J PROLAB IND E COMERCIO PRODUTOS PARA LABORATORIO": "J PROLAB INDUSTRIA E COMERCIO DE PRODUTOS PARA LABORATORIO",
        "COM ALV LIMPEZA E DESCARTAVEIS": "COM ALVORADA LIMPEZA E DESCARTAVEIS",
        "COM ALVORADA LIMPEZA E DESCARTAVEIS": "COM ALVORADA LIMPEZA E DESCARTAVEIS",
        "FABRICIO AMP RAFAEL TINTAS": "FABRICIO RAFAEL TINTAS",
        "FABRICIO RAFAEL TINTAS": "FABRICIO RAFAEL TINTAS",
        "BLACK AMP DECKER DO BRASIL": "BLACK DECKER DO BRASIL",
        "BLACK DECKER DO BRASIL": "BLACK DECKER DO BRASIL",
        "BAND EQUIP E MAQ REPUXADORAS": "BAND EQUIPAMENTOS E MAQUINAS REPUXADORAS",
        "BAND EQUIPAMENTOS E MAQUINAS REPUXADORAS": "BAND EQUIPAMENTOS E MAQUINAS REPUXADORAS",
        "CARVIC COM PROD E EQ AUTOMOTIVOS": "CARVIC COMERCIO DE PRODUTOS E EQUIPAMENTOS AUTOMOTIVOS",
        "CARVIC COM VER DE PROD E EQ P REP AUTO": "CARVIC COMERCIO DE PRODUTOS E EQUIPAMENTOS AUTOMOTIVOS",
        "EQUIPAMENTOS P PINTURA MAJAM": "EQUIPAMENTOS PARA PINTURA MAJAM",
        "EQUIPAMENTOS PARA PINTURA MAJAM EM RECUPERACAO JUDICI": "EQUIPAMENTOS PARA PINTURA MAJAM",
        "IND DE PLASTICOS HERC": "INDUSTRIA DE PLASTICOS HERC",
        "INDUSTRIA DE PLASTICOS HERC": "INDUSTRIA DE PLASTICOS HERC",
        "INDA STRIA DE PLA STICOS HERC": "INDUSTRIA DE PLASTICOS HERC",
        "MATRIZ CEILANDIA ETITEC SOLUCOES INTELIGENTES": "ETITEC SOLUCOES INTELIGENTES",
        "ETITEC SOLUCOES INTELIGENTES": "ETITEC SOLUCOES INTELIGENTES",
        "NOVA CASA DIST MAT CONST S A DF FL 02": "NOVA CASA DISTRIBUIDORA DE MATERIAIS PARA CONSTRUCAO",
        "NOVA CASA DIST MAT CONST DF FL 02": "NOVA CASA DISTRIBUIDORA DE MATERIAIS PARA CONSTRUCAO",
        "NOVA CASA DIST MAT P CONST": "NOVA CASA DISTRIBUIDORA DE MATERIAIS PARA CONSTRUCAO",
        "OTTO BAUMGART INDUSTRIA E COMERCIO S A": "OTTO BAUMGART INDUSTRIA E COMERCIO",
        "OTTO BAUMGART IND E COM": "OTTO BAUMGART INDUSTRIA E COMERCIO",
        "TERRA UTIL COMERCIO DE FERRAMENTAS E UTILIDADES": "TERRA UTILIDADES DE FERRAMENTAS E MAQUINAS",
        "TERRA UTIL DE MAQUINAS FERRAMENTAS E UTILID": "TERRA UTILIDADES DE FERRAMENTAS E MAQUINAS",
        "PPG IND DO BRASIL TINTAS E VERNIZES": "PPG INDUSTRIAL DO BRASIL TINTAS E VERNIZES",
        "PPG INDUSTRIAL DO BRASIL TINTAS E VERNIZES": "PPG INDUSTRIAL DO BRASIL TINTAS E VERNIZES",
        "SHERWIN WILLIAMS DO BRASIL IND E COM": "SHERWIN WILLIAMS DO BRASIL INDUSTRIA E COMERCIO",
        "SHERWIN WILLIAMS DO BRASIL INDUSTRIA E COMERCIO": "SHERWIN WILLIAMS DO BRASIL INDUSTRIA E COMERCIO",
        "G A OESTE INDUSTRIA DE TINTAS": "GOIAS G A OESTE INDUSTRIA DE TINTAS",
        "GOIAS G AMP A OESTE INDUSTRIA DE TINTAS": "GOIAS G A OESTE INDUSTRIA DE TINTAS",
        "GOIAS INDUSTRIA DE TINTAS": "GOIAS G A OESTE INDUSTRIA DE TINTAS",
        "INFINITY INDUSTRIAL": "INFINITY INDUSTRIAL",
        "INFINITY INDUSTRIAL ATACADO": "INFINITY INDUSTRIAL",
        "AUTOAMERICA": "AUTOAMERICA IMP EXP DE PROD AUTO",
        "AUTOAMERICA IMP EXP DE PROD AUTO": "AUTOAMERICA IMP EXP DE PROD AUTO",
        # EVC INDUSTRIAL e VONIXX são o mesmo fornecedor. Consolidação usada
        # em Compras, CMV, Sellout, Histórico e Orçamento.
        "EVC INDUSTRIAL": "VONIXX INDUSTRIA E COMERCIO DE POLIDORES",
        "EVC INDUSTRIAL LTDA": "VONIXX INDUSTRIA E COMERCIO DE POLIDORES",
        "VONIXX": "VONIXX INDUSTRIA E COMERCIO DE POLIDORES",
        "VONIXX INDUSTRIA E COMERCIO DE POLIDORES": "VONIXX INDUSTRIA E COMERCIO DE POLIDORES",
        # PINCEIS ATLAS S A e ATLAS S A representam o mesmo fornecedor.
        "PINCEIS ATLAS": "PINCEIS ATLAS S A",
        "PINCEIS ATLAS S A": "PINCEIS ATLAS S A",
        "ATLAS": "PINCEIS ATLAS S A",
        "ATLAS S A": "PINCEIS ATLAS S A",
    }
    raw = supplier_aliases.get(alias_key, raw)

    # Consolida todas as variações de razão social/cadastro da Roberlo em um
    # único fornecedor. Ex.: ROBERLO DO BRASIL, ROBERLO BRASIL, ROBERLO LTDA.
    if "ROBERLO" in alias_key or "ROBERLO" in raw:
        return "ROBERLO"

    if any(marca in brand_n for marca in {"SIKKENS", "WANDA", "TECH FLEET", "TECHFLEET"}):
        return "AKZO NOBEL AUTO"
    if any(marca in brand_n for marca in {"CORAL", "MACTRA", "HAMMERITE", "CETOL", "SPARLACK"}):
        return "AKZO NOBEL DECOR"
    if "SAINT" in raw and "GOBAIN" in raw:
        if "TEK BOND" in brand_n or "TEKBOND" in brand_n:
            return "SAINT GOBAIN TEK BOND"
        if "NORTON" in brand_n:
            return "SAINT GOBAIN NORTON"
        return "SAINT GOBAIN"
    if re.search(r"(^| )3M( |$)", raw):
        return "3M DO BRASIL"
    if "AKZO" in raw and "NOBEL" in raw:
        return "AKZO NOBEL"
    return raw


def supplier_division(name, brand="") -> str:
    """Classificação comercial única, usada por todas as fontes do dashboard."""
    supplier = canonical_supplier(name, brand)
    brand_n = "" if pd.isna(brand) else str(brand).upper().strip()

    if supplier in {"AKZO NOBEL AUTO", "AKZO NOBEL DECOR"}:
        return supplier
    if supplier == "AKZO NOBEL":
        if any(x in brand_n for x in {"SIKKENS", "WANDA", "TECH FLEET", "TECHFLEET"}):
            return "AKZO NOBEL AUTO"
        if any(x in brand_n for x in {"CORAL", "MACTRA", "HAMMERITE", "CETOL", "SPARLACK", "ALABASTINE"}):
            return "AKZO NOBEL DECOR"
        # Itens auxiliares sem marca (balança/espectrofotômetro) pertencem ao
        # fluxo automotivo; nunca mantemos o agrupador genérico no dashboard.
        return "AKZO NOBEL AUTO"

    if supplier in {"SAINT GOBAIN NORTON", "SAINT GOBAIN TEK BOND"}:
        return supplier
    if supplier == "SAINT GOBAIN":
        if "NORTON" in brand_n:
            return "SAINT GOBAIN NORTON"
        if "TEK BOND" in brand_n or "TEKBOND" in brand_n:
            return "SAINT GOBAIN TEK BOND"
        # Norton é a divisão padrão quando o cadastro antigo não informa marca.
        return "SAINT GOBAIN NORTON"

    # A razão social geral da Sherwin também aparece no cadastro dos produtos
    # automotivos. A descrição/marca distingue Lazzuril/Lazzudur e demais itens
    # automotivos, preservando o fornecedor geral como grupo separado.
    if supplier == "SHERWIN WILLIAMS DO BRASIL INDUSTRIA E COMERCIO":
        auto_terms = {
            "LAZZURIL", "LAZZUDUR", "SHERTRUCK", "SPECTRAPRIME",
            "PRIMER", "VERNIZ", "POLIESTER", "CATALISADOR", "ENDURECEDOR",
            "THINNER", "MASSA", "MASCARAMENTO", "WASH", "PEGA POEIRA",
        }
        if any(term in brand_n for term in auto_terms):
            return "SHERWIN WILLIAMS DO BRASIL IND E COM DIV AUTOMOTIVA"

    return supplier


def supplier_family(name) -> str:
    """Família usada para conciliar a NF do CITEL com as divisões das entradas."""
    supplier = canonical_supplier(name)
    if supplier.startswith("AKZO NOBEL"):
        return "AKZO NOBEL"
    if supplier.startswith("SAINT GOBAIN"):
        return "SAINT GOBAIN"
    return supplier


def build_supplier_key(name, brand="") -> str:
    return supplier_key(canonical_supplier(name, brand))


def empty_citel():
    return pd.DataFrame(columns=[
        "FORNECEDOR_CITEL", "FORN_KEY", "COMPRA_VALOR", "DATA_DT", "ANO",
        "MES_NUM", "NR_DOCUMENTO", "NOTA_KEY"
    ])


@st.cache_data(show_spinner=False)
def load_data(giro_path: str, cad_forn_path: str, cad_prod_path: str, sellout_path: str, entradas_path: str, cache_token=None):
    # ---------------- CADASTROS ----------------
    cad_forn = read_report_csv(cad_forn_path, ("CÓDIGO", "NOME DO FORNECEDOR", "C.P.F./C.N.P.J."), cache_token)
    c_cnpj = find_col(cad_forn, "C.P.F./C.N.P.J.") or find_col(cad_forn, "CPF/CNPJ")
    c_nome_f = find_col(cad_forn, "NOME DO FORNECEDOR")
    if c_cnpj is None or c_nome_f is None:
        raise ValueError(f"CADASTRO DE FORNECEDORES: colunas obrigatórias ausentes. Colunas: {list(cad_forn.columns)}")
    cad_forn["CNPJ_KEY"] = cad_forn[c_cnpj].map(cnpj_key)
    cad_forn["FORNECEDOR_CAD"] = cad_forn[c_nome_f].fillna("").astype(str).str.strip()
    cad_forn["NOME_KEY"] = cad_forn["FORNECEDOR_CAD"].map(supplier_name_key)

    cad_cnpj_valid = cad_forn[cad_forn["CNPJ_KEY"] != ""].drop_duplicates("CNPJ_KEY", keep="last")
    cnpj_to_supplier = dict(zip(cad_cnpj_valid["CNPJ_KEY"], cad_cnpj_valid["FORNECEDOR_CAD"]))

    # Mapa de nomes só usa chaves que apontam para um único fornecedor.
    cad_nome_valid = cad_forn[cad_forn["NOME_KEY"] != ""].copy()
    nome_counts = cad_nome_valid.groupby("NOME_KEY")["FORNECEDOR_CAD"].nunique()
    unique_name_keys = set(nome_counts[nome_counts == 1].index)
    cad_nome_valid = cad_nome_valid[cad_nome_valid["NOME_KEY"].isin(unique_name_keys)].drop_duplicates("NOME_KEY", keep="last")
    nome_to_supplier = dict(zip(cad_nome_valid["NOME_KEY"], cad_nome_valid["FORNECEDOR_CAD"]))

    cad_prod = read_report_csv(cad_prod_path, ("Cód.Item", "Desc. Fornecedor", "Desc. Linha/Grupo"), cache_token)
    c_prod_cod = find_col(cad_prod, "CÓD.ITEM") or find_col(cad_prod, "COD.ITEM") or find_col(cad_prod, "CÓDIGO")
    c_prod_forn = find_col(cad_prod, "DESC. FORNECEDOR")
    c_prod_linha = find_col(cad_prod, "DESC. LINHA/GRUPO")
    c_prod_marca = find_col(cad_prod, "DESC. MARCA")
    if not all([c_prod_cod, c_prod_forn, c_prod_linha]):
        raise ValueError(f"CADASTRO PRODUTOS GERAL: colunas obrigatórias ausentes. Colunas: {list(cad_prod.columns)}")
    cad_prod["COD_KEY"] = cad_prod[c_prod_cod].map(product_key)
    cad_prod["FORNECEDOR_PROD"] = cad_prod[c_prod_forn].fillna("").astype(str).str.strip()
    cad_prod["LINHA_PROD"] = cad_prod[c_prod_linha].fillna("").astype(str).str.strip()
    cad_prod["MARCA_PROD"] = cad_prod[c_prod_marca].fillna("").astype(str).str.strip() if c_prod_marca else ""
    cad_prod = cad_prod[cad_prod["COD_KEY"] != ""].drop_duplicates("COD_KEY", keep="last")
    prod_lookup = cad_prod[["COD_KEY", "FORNECEDOR_PROD", "LINHA_PROD", "MARCA_PROD"]]

    # ---------------- GIRO / CMV / ESTOQUE + NOTAS CITEL ----------------
    xls = pd.ExcelFile(giro_path)
    giro_sheet = None
    notas_sheet = None
    for sh in xls.sheet_names:
        probe = strip_cols(pd.read_excel(xls, sheet_name=sh, nrows=3))
        norms = {colnorm(c) for c in probe.columns}
        if {"CÓDIGO", "CMV", "VLR ESTOQUE"}.issubset(norms):
            giro_sheet = sh
        if "NR_CNPJ_EMITENTE" in norms:
            notas_sheet = sh
    if giro_sheet is None:
        raise ValueError(f"GIRO E NOTAS: não encontrei aba com CÓDIGO, CMV e VLR ESTOQUE. Abas: {xls.sheet_names}")

    giro = strip_cols(pd.read_excel(xls, sheet_name=giro_sheet))
    c_g_cod = find_col(giro, "CÓDIGO") or find_col(giro, "CODIGO")
    c_g_desc = find_col(giro, "DESCRIÇÃO DO ITEM") or find_col(giro, "DESCRICAO DO ITEM")
    c_g_marca = find_col(giro, "MARCA")
    c_g_cmv = find_col(giro, "CMV")
    c_g_est = find_col(giro, "VLR ESTOQUE") or find_col(giro, "VLR_ESTOQUE")
    c_g_mes = find_col(giro, "MÊS") or find_col(giro, "MES")
    if not all([c_g_cod, c_g_cmv, c_g_est]):
        raise ValueError(f"GIRO: colunas obrigatórias ausentes. Colunas: {list(giro.columns)}")
    giro["COD_KEY"] = giro[c_g_cod].map(product_key)
    giro = giro[~giro["COD_KEY"].isin(EXCLUDED_PRODUCT_KEYS)].copy()
    giro = giro.merge(prod_lookup, on="COD_KEY", how="left")
    giro["MARCA"] = giro[c_g_marca].fillna("").astype(str).str.strip() if c_g_marca else giro["MARCA_PROD"]
    giro["DESCRICAO_ITEM"] = giro[c_g_desc].fillna("").astype(str).str.strip() if c_g_desc else ""
    giro["LINHA"] = giro["LINHA_PROD"].fillna("").astype(str).str.strip()
    giro_hint = giro["MARCA"].fillna("").astype(str) + " " + giro["DESCRICAO_ITEM"].fillna("").astype(str)
    giro["FORNECEDOR_CMV"] = [supplier_division(n, h) for n, h in zip(giro["FORNECEDOR_PROD"], giro_hint)]
    giro["FORNECEDOR_CMV"] = giro["FORNECEDOR_CMV"].replace("", "FORNECEDOR NÃO IDENTIFICADO").fillna("FORNECEDOR NÃO IDENTIFICADO")
    giro["FORN_KEY"] = giro["FORNECEDOR_CMV"].map(supplier_key)
    giro["CMV_VALOR"] = parse_number_br(giro[c_g_cmv])
    giro["ESTOQUE_VALOR"] = parse_number_br(giro[c_g_est])
    giro["MES_NUM"] = giro[c_g_mes].map(parse_mes_to_num) if c_g_mes else pd.NA
    df_cmv = giro

    df_citel = empty_citel()
    if notas_sheet is not None:
        notas = strip_cols(pd.read_excel(xls, sheet_name=notas_sheet))
        c_n_cnpj = find_col(notas, "NR_CNPJ_EMITENTE")
        c_n_nome = find_col(notas, "NM_EMITENTE") or find_col(notas, "NOME EMITENTE") or find_col(notas, "FORNECEDOR")
        c_n_val = find_col(notas, "VL_NOTA_FISCAL") or find_col(notas, "VALOR NOTA") or find_col(notas, "VR. CONTÁBIL")
        c_n_dt = find_col(notas, "DT_EMISSAO") or find_col(notas, "DATA")
        c_n_doc = find_col(notas, "NR_DOCUMENTO") or find_col(notas, "DOCUMENTO") or find_col(notas, "NR NOTA FISCAL")
        if all([c_n_cnpj, c_n_val, c_n_dt, c_n_doc]):
            notas["CNPJ_ORIGINAL"] = notas[c_n_cnpj]
            notas["CNPJ_KEY"] = notas[c_n_cnpj].map(cnpj_key)
            notas["NOME_EMITENTE_ORIGINAL"] = notas[c_n_nome].fillna("").astype(str).str.strip() if c_n_nome else ""
            notas["NOME_KEY"] = notas["NOME_EMITENTE_ORIGINAL"].map(supplier_name_key)

            por_cnpj = notas["CNPJ_KEY"].map(cnpj_to_supplier)
            por_nome_cadastro = notas["NOME_KEY"].map(nome_to_supplier)
            nome_da_nota = notas["NOME_EMITENTE_ORIGINAL"].replace("", pd.NA)

            # Ordem: CNPJ íntegro -> nome localizado no cadastro -> razão social da própria NF.
            fornecedor_base = por_cnpj.fillna(por_nome_cadastro).fillna(nome_da_nota)
            notas["METODO_IDENTIFICACAO"] = "NÃO IDENTIFICADO"
            notas.loc[por_cnpj.notna(), "METODO_IDENTIFICACAO"] = "CNPJ"
            notas.loc[por_cnpj.isna() & por_nome_cadastro.notna(), "METODO_IDENTIFICACAO"] = "NOME/CADASTRO"
            notas.loc[por_cnpj.isna() & por_nome_cadastro.isna() & nome_da_nota.notna(), "METODO_IDENTIFICACAO"] = "NOME DA NF"

            notas["FORNECEDOR_CITEL"] = fornecedor_base.map(canonical_supplier).replace("", "FORNECEDOR NÃO IDENTIFICADO").fillna("FORNECEDOR NÃO IDENTIFICADO")
            notas["FORN_KEY"] = notas["FORNECEDOR_CITEL"].map(supplier_key)
            notas["COMPRA_VALOR"] = parse_number_br(notas[c_n_val])
            notas["DATA_DT"] = to_datetime_safe(notas[c_n_dt])
            notas["ANO"] = notas["DATA_DT"].dt.year
            notas["MES_NUM"] = notas["DATA_DT"].dt.month
            notas["NR_DOCUMENTO"] = notas[c_n_doc]
            notas["NOTA_KEY"] = notas["NR_DOCUMENTO"].map(nota_key)
            df_citel = notas

    # ---------------- NOTAS DE ENTRADA ANALÍTICAS ----------------
    ent = read_report_csv(entradas_path, ("DOCUMENTO", "FORNECEDOR", "VR. CONTÁBIL"), cache_token)
    c_e_doc = find_col(ent, "DOCUMENTO")
    c_e_forn = find_col(ent, "FORNECEDOR")
    c_e_data = find_col(ent, "DATA")
    c_e_cod = find_col(ent, "CÓDIGO") or find_col(ent, "CODIGO")
    c_e_desc = find_col(ent, "DESCRIÇÃO DO ITEM") or find_col(ent, "DESCRICAO DO ITEM")
    c_e_marca = find_col(ent, "MARCA")
    c_e_val = find_col(ent, "VR. CONTÁBIL") or find_col(ent, "VR CONTABIL")
    if not all([c_e_doc, c_e_forn, c_e_data, c_e_val]):
        raise ValueError(f"NOTAS DE ENTRADA: colunas obrigatórias ausentes. Colunas: {list(ent.columns)}")
    ent["COD_KEY"] = ent[c_e_cod].map(product_key) if c_e_cod else ""
    ent = ent[~ent["COD_KEY"].isin(EXCLUDED_PRODUCT_KEYS)].copy()
    ent = ent.merge(prod_lookup[["COD_KEY", "LINHA_PROD"]], on="COD_KEY", how="left")
    ent["MARCA"] = ent[c_e_marca].fillna("").astype(str).str.strip() if c_e_marca else ""
    ent["DESCRICAO_ITEM"] = ent[c_e_desc].fillna("").astype(str).str.strip() if c_e_desc else ""
    ent_hint = ent["MARCA"].fillna("").astype(str) + " " + ent["DESCRICAO_ITEM"].fillna("").astype(str)
    ent["FORNECEDOR_ENT"] = [supplier_division(n, h) for n, h in zip(ent[c_e_forn], ent_hint)]
    ent["FORN_KEY"] = ent["FORNECEDOR_ENT"].map(supplier_key)
    ent["VR_CONTABIL"] = parse_number_br(ent[c_e_val])
    ent["NR_NOTA_FISCAL"] = ent[c_e_doc]
    ent["NOTA_KEY"] = ent["NR_NOTA_FISCAL"].map(nota_key)
    ent["LINHA"] = ent["LINHA_PROD"].fillna("").astype(str).str.strip()
    ent["GRUPO"] = ""
    ent["DATA_DT"] = to_datetime_safe(ent[c_e_data])
    ent["ANO"] = ent["DATA_DT"].dt.year
    ent["MES_NUM"] = ent["DATA_DT"].dt.month
    df_ent = ent

    # As notas do CITEL não possuem item/marca. Para aplicar as mesmas divisões
    # comerciais, distribui o valor de cada NF conforme as linhas analíticas da
    # mesma nota em ENTRADAS. O rateio preserva exatamente o total original.
    if not df_citel.empty:
        df_citel["FAMILIA_FORNECEDOR"] = df_citel["FORNECEDOR_CITEL"].map(supplier_family)
        df_ent["FAMILIA_FORNECEDOR"] = df_ent["FORNECEDOR_ENT"].map(supplier_family)

        divisoes = {
            "AKZO NOBEL AUTO", "AKZO NOBEL DECOR",
            "SAINT GOBAIN NORTON", "SAINT GOBAIN TEK BOND",
        }
        rateio = (
            df_ent[df_ent["FORNECEDOR_ENT"].isin(divisoes) & (df_ent["NOTA_KEY"] != "")]
            .groupby(["NOTA_KEY", "FAMILIA_FORNECEDOR", "FORNECEDOR_ENT"], as_index=False)
            .agg(VALOR_DIVISAO=("VR_CONTABIL", "sum"))
        )
        if not rateio.empty:
            totais_rateio = (
                rateio.groupby(["NOTA_KEY", "FAMILIA_FORNECEDOR"])["VALOR_DIVISAO"]
                .transform("sum")
            )
            rateio["PESO_DIVISAO"] = rateio["VALOR_DIVISAO"] / totais_rateio.where(totais_rateio != 0, 1.0)
            rateio = rateio.rename(columns={"FORNECEDOR_ENT": "DIVISAO_CITEL"})

            citel_rateado = df_citel.merge(
                rateio[["NOTA_KEY", "FAMILIA_FORNECEDOR", "DIVISAO_CITEL", "PESO_DIVISAO"]],
                on=["NOTA_KEY", "FAMILIA_FORNECEDOR"], how="left"
            )
            citel_rateado["COMPRA_VALOR"] = (
                citel_rateado["COMPRA_VALOR"] * citel_rateado["PESO_DIVISAO"].fillna(1.0)
            )
            citel_rateado["FORNECEDOR_CITEL"] = (
                citel_rateado["DIVISAO_CITEL"].fillna(citel_rateado["FORNECEDOR_CITEL"])
            )
            citel_rateado["FORN_KEY"] = citel_rateado["FORNECEDOR_CITEL"].map(supplier_key)
            df_citel = citel_rateado.drop(columns=["DIVISAO_CITEL", "PESO_DIVISAO"])

        # Notas sem correspondência direta em ENTRADAS são distribuídas pela
        # composição real da respectiva família. Isso impede grupos genéricos
        # sem inventar valor e mantém o total do CITEL inalterado.
        genericos = {"AKZO NOBEL", "SAINT GOBAIN"}
        mask_generico = df_citel["FORNECEDOR_CITEL"].isin(genericos)
        if mask_generico.any():
            pesos_familia = (
                df_ent[df_ent["FORNECEDOR_ENT"].isin(divisoes)]
                .groupby(["FAMILIA_FORNECEDOR", "FORNECEDOR_ENT"], as_index=False)
                .agg(VALOR_DIVISAO=("VR_CONTABIL", "sum"))
            )
            if not pesos_familia.empty:
                totais_familia = pesos_familia.groupby("FAMILIA_FORNECEDOR")["VALOR_DIVISAO"].transform("sum")
                pesos_familia["PESO_FAMILIA"] = pesos_familia["VALOR_DIVISAO"] / totais_familia.where(totais_familia != 0, 1.0)
                pesos_familia = pesos_familia.rename(columns={"FORNECEDOR_ENT": "DIVISAO_FALLBACK"})

                citel_ok = df_citel[~mask_generico].copy()
                citel_fallback = df_citel[mask_generico].merge(
                    pesos_familia[["FAMILIA_FORNECEDOR", "DIVISAO_FALLBACK", "PESO_FAMILIA"]],
                    on="FAMILIA_FORNECEDOR", how="left"
                )
                tem_fallback = citel_fallback["DIVISAO_FALLBACK"].notna()
                citel_fallback.loc[tem_fallback, "COMPRA_VALOR"] *= citel_fallback.loc[tem_fallback, "PESO_FAMILIA"]
                citel_fallback.loc[tem_fallback, "FORNECEDOR_CITEL"] = citel_fallback.loc[tem_fallback, "DIVISAO_FALLBACK"]
                citel_fallback["FORN_KEY"] = citel_fallback["FORNECEDOR_CITEL"].map(supplier_key)
                df_citel = pd.concat([
                    citel_ok,
                    citel_fallback.drop(columns=["DIVISAO_FALLBACK", "PESO_FAMILIA"])
                ], ignore_index=True)

    # ---------------- SELLOUT ----------------
    so = read_report_csv(sellout_path, ("FORNECEDOR", "CÓDIGO", "FATURAMENTO"), cache_token)
    c_s_cod = find_col(so, "CÓDIGO") or find_col(so, "CODIGO")
    c_s_forn = find_col(so, "FORNECEDOR")
    c_s_desc = find_col(so, "DESCRIÇÃO DO PRODUTO") or find_col(so, "DESCRICAO DO PRODUTO")
    c_s_fat = find_col(so, "FATURAMENTO")
    c_s_qtd = find_col(so, "QTD. FATUR") or find_col(so, "QTD FATUR")
    c_s_mes = find_col(so, "MÊS") or find_col(so, "MES")
    c_s_ano = find_col(so, "ANO")
    if not all([c_s_cod, c_s_fat]):
        raise ValueError(f"SELLOUT: colunas obrigatórias ausentes. Colunas: {list(so.columns)}")
    so["COD_KEY"] = so[c_s_cod].map(product_key)
    so = so[~so["COD_KEY"].isin(EXCLUDED_PRODUCT_KEYS)].copy()
    so = so.merge(prod_lookup, on="COD_KEY", how="left")
    so["MARCA"] = so["MARCA_PROD"].fillna("").astype(str).str.strip()
    so["DESCRICAO_PRODUTO"] = so[c_s_desc].fillna("").astype(str).str.strip() if c_s_desc else ""
    sellout_hint = so["MARCA"] + " " + so["DESCRICAO_PRODUTO"]
    fornecedor_prod = so["FORNECEDOR_PROD"].fillna("").astype(str).str.strip()
    fornecedor_relatorio = so[c_s_forn].fillna("").astype(str).str.strip() if c_s_forn else pd.Series("", index=so.index)
    fornecedor_base_so = fornecedor_prod.where(fornecedor_prod != "", fornecedor_relatorio)
    so["FORNECEDOR_SELLOUT"] = [supplier_division(n, h) for n, h in zip(fornecedor_base_so, sellout_hint)]
    # Segunda chance: alguns cadastros têm fornecedor inválido, embora o relatório
    # de Sellout possua o agrupamento correto na própria linha.
    vazios_so = so["FORNECEDOR_SELLOUT"].fillna("").astype(str).str.strip() == ""
    if vazios_so.any():
        so.loc[vazios_so, "FORNECEDOR_SELLOUT"] = [
            supplier_division(n, m)
            for n, m in zip(fornecedor_relatorio.loc[vazios_so], sellout_hint.loc[vazios_so])
        ]
    so["FORNECEDOR_SELLOUT"] = so["FORNECEDOR_SELLOUT"].replace("", "FORNECEDOR NÃO IDENTIFICADO").fillna("FORNECEDOR NÃO IDENTIFICADO")
    so["FORN_KEY"] = so["FORNECEDOR_SELLOUT"].map(supplier_key)
    so["LINHA"] = so["LINHA_PROD"].fillna("").astype(str).str.strip()
    so["FATURAMENTO"] = parse_number_br(so[c_s_fat])
    so["QTD_FATUR"] = parse_number_br(so[c_s_qtd]) if c_s_qtd else 0.0
    so["CODIGO"] = so["COD_KEY"]

    # Competência do Sellout. Quando existe a coluna MÊS, ela prevalece sobre
    # o período geral do cabeçalho e permite os filtros/históricos mensais reais.
    text = open(sellout_path, "r", encoding="latin1", errors="ignore").read()[:8000]
    mi = re.search(r"Data Inicial \(Entrada\)\s*:\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    mf = re.search(r"Data Final \(Entrada\)\s*:\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    data_ref = pd.to_datetime(mf.group(1), dayfirst=True) if mf else (pd.to_datetime(mi.group(1), dayfirst=True) if mi else pd.NaT)

    if c_s_mes:
        so["MES_NUM"] = so[c_s_mes].map(parse_mes_to_num).astype("Int64")
    else:
        so["MES_NUM"] = data_ref.month if pd.notna(data_ref) else pd.NA

    if c_s_ano:
        so["ANO"] = pd.to_numeric(so[c_s_ano], errors="coerce").astype("Int64")
    elif c_s_mes:
        # Ao existir detalhamento mensal, o ano deve acompanhar as demais bases
        # do carregamento. O cabeçalho pode guardar uma data antiga do relatório
        # original e não deve eliminar todo o Sellout no filtro do ano atual.
        anos_referencia = pd.concat([
            df_citel["ANO"] if "ANO" in df_citel.columns else pd.Series(dtype=float),
            df_ent["ANO"] if "ANO" in df_ent.columns else pd.Series(dtype=float),
        ]).dropna()
        ano_referencia = int(anos_referencia.max()) if not anos_referencia.empty else (data_ref.year if pd.notna(data_ref) else pd.NA)
        so["ANO"] = ano_referencia
    elif pd.notna(data_ref):
        so["ANO"] = data_ref.year
    else:
        so["ANO"] = pd.NA

    so["DATA_DT"] = pd.to_datetime(
        {"year": pd.to_numeric(so["ANO"], errors="coerce"),
         "month": pd.to_numeric(so["MES_NUM"], errors="coerce"),
         "day": 1},
        errors="coerce",
    )
    df_sellout = so

    return df_cmv, df_citel, df_ent, df_sellout


try:
    # O token por data/tamanho invalida o cache quando qualquer arquivo é
    # substituído, mesmo que permaneça com exatamente o mesmo nome/caminho.
    arquivos_origem = [
        GIRO_NOTAS_PATH, CAD_FORNECEDORES_PATH, CAD_PRODUTOS_PATH,
        SELLOUT_PATH, NOTAS_ENTRADA_PATH,
    ]
    cache_token = tuple(
        (os.path.getmtime(path), os.path.getsize(path)) for path in arquivos_origem
    )
    df_cmv, df_citel, df_ent, df_sellout = load_data(
        GIRO_NOTAS_PATH, CAD_FORNECEDORES_PATH, CAD_PRODUTOS_PATH,
        SELLOUT_PATH, NOTAS_ENTRADA_PATH, cache_token
    )
except Exception as e:
    st.error(f"Erro ao carregar as bases: {e}")
    st.stop()


# -----------------------------
# Exportação PDF simples (sem dependências adicionais)
# -----------------------------
def _pdf_escape_text(value) -> bytes:
    text = str(value if value is not None else "")
    text = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    return text.encode("cp1252", errors="replace")


def build_budget_pdf(orcamento_df: pd.DataFrame, cmv_base: float, periodo_txt: str) -> bytes:
    """Gera PDF multipágina com fontes padrão do PDF, sem reportlab/fpdf."""
    rows = []
    for _, row in orcamento_df.iterrows():
        rows.append((
            str(row.get("FORNECEDOR", "")),
            float(row.get("CMV PERÍODO_NUM", 0.0)),
            float(row.get("COMPRAS PERÍODO_NUM", 0.0)),
            float(row.get("DIFERENÇA_NUM", 0.0)),
            float(row.get("PARTICIPAÇÃO", 0.0)),
            float(row.get("ORÇAMENTO FINAL_NUM", 0.0)),
        ))

    total = sum(r[5] for r in rows)
    page_w, page_h = 842, 595
    left, top = 28, 555
    line_h = 17
    rows_per_page = 25
    chunks = [rows[i:i + rows_per_page] for i in range(0, len(rows), rows_per_page)] or [[]]

    objects = [b"<< /Type /Catalog /Pages 2 0 R >>"]
    page_ids, content_ids = [], []
    next_id = 4
    for _ in chunks:
        page_ids.append(next_id); next_id += 1
        content_ids.append(next_id); next_id += 1
    kids = b" ".join(f"{pid} 0 R".encode() for pid in page_ids)
    objects.append(b"<< /Type /Pages /Kids [" + kids + b"] /Count " + str(len(page_ids)).encode() + b" >>")
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")

    for page_idx, chunk in enumerate(chunks):
        cmds = []
        def txt(x, y, size, text):
            cmds.append(b"BT /F1 " + str(size).encode() + b" Tf " + f"{x} {y} Td ".encode() + b"(" + _pdf_escape_text(text) + b") Tj ET")

        txt(left, top, 17, "ORÇAMENTO DE COMPRAS")
        txt(left, top - 24, 10, f"Período de participação: {periodo_txt}")
        txt(left, top - 40, 10, f"CMV base mensal: {brl(cmv_base)}")
        txt(left, top - 56, 10, f"Orçamento total final: {brl(total)}")

        y = top - 84
        txt(left, y, 9, "FORNECEDOR")
        txt(355, y, 9, "CMV")
        txt(455, y, 9, "COMPRAS")
        txt(555, y, 9, "DIF.")
        txt(650, y, 9, "%")
        txt(700, y, 9, "ORÇAMENTO")
        y -= 8
        cmds.append(f"{left} {y} m 817 {y} l S".encode())
        y -= 15

        start_row = page_idx * rows_per_page + 1
        for n, (fornecedor, cmv, compras, dif, part, valor) in enumerate(chunk, start=start_row):
            nome = fornecedor if len(fornecedor) <= 41 else fornecedor[:38] + "..."
            txt(left, y, 8, f"{n}. {nome}")
            txt(355, y, 8, brl(cmv))
            txt(455, y, 8, brl(compras))
            txt(555, y, 8, brl(dif))
            txt(650, y, 8, pct_str(part))
            txt(700, y, 8, brl(valor))
            y -= line_h

        txt(left, 22, 8, f"Página {page_idx + 1} de {len(chunks)}")
        stream = b"\n".join(cmds)
        page_obj = f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {page_w} {page_h}] /Resources << /Font << /F1 3 0 R >> >> /Contents {content_ids[page_idx]} 0 R >>".encode()
        content_obj = b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream"
        objects.extend([page_obj, content_obj])

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for i, obj in enumerate(objects, start=1):
        offsets.append(len(out))
        out.extend(f"{i} 0 obj\n".encode()); out.extend(obj); out.extend(b"\nendobj\n")
    xref = len(out)
    out.extend(f"xref\n0 {len(objects)+1}\n".encode())
    out.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        out.extend(f"{off:010d} 00000 n \n".encode())
    out.extend(f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode())
    return bytes(out)


# -----------------------------
# Sidebar: Página + filtros
# -----------------------------
st.sidebar.title("Navegação")
page = st.sidebar.selectbox("Página", ["COMPRAS", "SELLOUT", "HISTÓRICO POR FORNECEDOR", "ORÇAMENTO"])

st.sidebar.divider()
st.sidebar.subheader("Filtros")

anos_citel = sorted(df_citel["ANO"].dropna().astype(int).unique().tolist())
anos_ent = sorted(df_ent["ANO"].dropna().astype(int).unique().tolist())
anos_sellout = []
if df_sellout is not None and "ANO" in df_sellout.columns and df_sellout["ANO"].notna().any():
    anos_sellout = sorted(df_sellout["ANO"].dropna().astype(int).unique().tolist())

anos = sorted(set(anos_citel + anos_ent + anos_sellout))

# A introdução da competência mensal mudou o significado do filtro do Sellout.
# Limpa uma única vez seleções antigas persistidas na sessão do Streamlit.
if st.session_state.get("filter_state_version") != FILTER_STATE_VERSION:
    st.session_state["filter_state_version"] = FILTER_STATE_VERSION
    st.session_state["filtro_anos_aplicado"] = anos
    st.session_state["filtro_meses_aplicado"] = MESES_LABELS
    st.session_state.pop("filtro_anos_input", None)
    st.session_state.pop("filtro_meses_input", None)

# Também remove opções que deixaram de existir após a troca dos arquivos.
anos_aplicados = st.session_state.get("filtro_anos_aplicado", anos)
st.session_state["filtro_anos_aplicado"] = [x for x in anos_aplicados if x in anos]
meses_aplicados = st.session_state.get("filtro_meses_aplicado", MESES_LABELS)
st.session_state["filtro_meses_aplicado"] = [x for x in meses_aplicados if x in MESES_LABELS]
# Os filtros globais ficam dentro de um formulário para evitar que cada clique
# provoque a reconstrução completa de todas as tabelas e gráficos.
with st.sidebar.form("form_filtros_globais", clear_on_submit=False):
    sel_anos_input = st.multiselect(
        "Ano",
        options=anos,
        default=st.session_state.get("filtro_anos_aplicado", anos),
        key="filtro_anos_input",
    )
    st.markdown("**Meses**")
    meses_atualmente_aplicados = st.session_state.get("filtro_meses_aplicado", MESES_LABELS)
    mes_cols = st.columns(2)
    meses_marcados = {}
    for idx, mes in enumerate(MESES_LABELS):
        meses_marcados[mes] = mes_cols[idx % 2].checkbox(
            mes.title(),
            value=(mes in meses_atualmente_aplicados),
            key=f"filtro_mes_check_{MESES_PT[mes]}",
        )
    sel_meses_input = [mes for mes in MESES_LABELS if meses_marcados.get(mes, False)]
    aplicar_filtros = st.form_submit_button("Aplicar filtros", use_container_width=True)

if aplicar_filtros or "filtro_anos_aplicado" not in st.session_state:
    st.session_state["filtro_anos_aplicado"] = sel_anos_input
    st.session_state["filtro_meses_aplicado"] = sel_meses_input

sel_anos = st.session_state.get("filtro_anos_aplicado", anos)
sel_meses = st.session_state.get("filtro_meses_aplicado", MESES_LABELS)
sel_meses_num = [MESES_PT[m] for m in sel_meses if m in MESES_PT]

def apply_month_year_filter(df, apply_year=True, apply_month=True):
    if df is None:
        return None
    out = df
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
    st.caption(f"Versão do aplicativo: {APP_VERSION}")
    st.title("INDICADORES DE COMPRAS")

    if df_citel.empty:
        st.warning("A planilha GIRO E NOTAS atual não possui uma aba com NR_CNPJ_EMITENTE. Os indicadores de Compras CITEL ficarão zerados até essa aba ser incluída; CMV, estoque, entradas e sellout continuam disponíveis.")

    # Resumo geral do período selecionado
    total_compras_citel = float(df_citel_f["COMPRA_VALOR"].sum())
    total_vendas_cmv = float(df_cmv_f["CMV_VALOR"].sum())
    dif_topo = total_vendas_cmv - total_compras_citel

    if total_compras_citel != 0:
        dif_pct = dif_topo / total_compras_citel
    elif total_vendas_cmv != 0:
        dif_pct = dif_topo / total_vendas_cmv
    else:
        dif_pct = 0.0

    # Quantidade real de competências selecionadas (ano x mês), usada nas médias.
    periodos_selecionados = max(len(sel_anos) * len(sel_meses_num), 1)
    media_compras = total_compras_citel / periodos_selecionados
    media_cmv = total_vendas_cmv / periodos_selecionados
    media_diferenca = dif_topo / periodos_selecionados

    st.subheader("Resumo do Período Selecionado")

    resumo_periodo = pd.DataFrame({
        "INDICADOR": ["Total Compras", "Total CMV", "Diferença (CMV - Compras)"],
        "VALOR": [total_compras_citel, total_vendas_cmv, dif_topo],
        "% SOBRE COMPRAS": [1.0 if total_compras_citel else 0.0,
                              (total_vendas_cmv / total_compras_citel) if total_compras_citel else 0.0,
                              dif_pct],
    })

    st.dataframe(
        resumo_periodo.style
        .format({
            "VALOR": brl,
            "% SOBRE COMPRAS": lambda x: pct_str(float(x)),
        })
        .map(style_dif, subset=["VALOR"]),
        use_container_width=True,
        hide_index=True,
    )

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Média de Compras por Período", brl(media_compras))
    with m2:
        st.metric("Média de CMV por Período", brl(media_cmv))
    with m3:
        st.metric("Média da Diferença", brl(media_diferenca))
    with m4:
        st.metric("Períodos Considerados", periodos_selecionados)

    st.caption("A média considera cada combinação de ano e mês selecionada nos filtros.")

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
      .map(style_dif, subset=["DIF (CMV - COMPRAS)"]),
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
        f"Comparação direta: GIRO E NOTAS/NOTAS (NR_DOCUMENTO) vs NOTAS DE ENTRADA (DOCUMENTO), removendo o sufixo -NF-009 e zeros à esquerda. "
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
    # Aplicação explícita evita um rerun pesado a cada marca marcada/desmarcada.
    marcas_state_key = f"drill_marcas_aplicadas::{sel_forn_key or 'TODOS'}"
    marcas_default = st.session_state.get(marcas_state_key, marcas)
    marcas_default = [m for m in marcas_default if m in marcas]

    with st.form("form_filtro_marcas", clear_on_submit=False):
        sel_marcas_input = st.multiselect(
            "Filtrar Marcas (ENTRADAS)",
            options=marcas,
            default=marcas_default,
            key=f"drill_marcas_input::{sel_forn_key or 'TODOS'}",
        )
        aplicar_marcas = st.form_submit_button("Aplicar marcas")

    if aplicar_marcas or marcas_state_key not in st.session_state:
        st.session_state[marcas_state_key] = sel_marcas_input

    sel_marcas = st.session_state.get(marcas_state_key, marcas)
    if sel_marcas:
        ent_base = ent_base[ent_base["MARCA"].isin(sel_marcas)]

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
          .map(style_dif, subset=["DIF (CMV - COMPRAS)"]),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("##### Participação por Linha (Mapa / Treemap)")
    exibir_mapas = st.toggle(
        "Exibir mapas detalhados",
        value=False,
        help="Os três mapas são pesados. Ative somente quando precisar analisá-los.",
        key="exibir_treemaps_compras",
    )

    if exibir_mapas:
        path_cols = ["MARCA"]
        if has_grupo:
            path_cols.append("GRUPO")
        path_cols.append("LINHA")

        # Remove linhas sem valor antes de montar os gráficos, reduzindo memória e JSON.
        dr_plot = dr.loc[(dr["COMPRAS"] != 0) | (dr["VENDAS_CMV"] != 0) | (dr["VLR_ESTOQUE"] != 0)].copy()
        g1, g2, g3 = st.columns(3)
        with g1:
            fig_comp = px.treemap(dr_plot, path=path_cols, values="COMPRAS", title="Compras (ENTRADAS)")
            fig_comp.update_layout(margin=dict(t=50, l=10, r=10, b=10))
            st.plotly_chart(fig_comp, use_container_width=True, key="treemap_compras")
        with g2:
            fig_vend = px.treemap(dr_plot, path=path_cols, values="VENDAS_CMV", title="Vendas (CMV)")
            fig_vend.update_layout(margin=dict(t=50, l=10, r=10, b=10))
            st.plotly_chart(fig_vend, use_container_width=True, key="treemap_vendas")
        with g3:
            fig_est = px.treemap(dr_plot, path=path_cols, values="VLR_ESTOQUE", title="Valor de Estoque (CMV E ESTOQUE)")
            fig_est.update_layout(margin=dict(t=50, l=10, r=10, b=10))
            st.plotly_chart(fig_est, use_container_width=True, key="treemap_estoque")
    else:
        st.caption("Mapas desativados para manter os filtros leves e estáveis.")

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
        anos_detectados = sorted(df_sellout["ANO"].dropna().astype(int).unique().tolist()) if "ANO" in df_sellout.columns else []
        meses_detectados_num = sorted(df_sellout["MES_NUM"].dropna().astype(int).unique().tolist()) if "MES_NUM" in df_sellout.columns else []
        meses_detectados = [MESES_LABELS[m - 1] for m in meses_detectados_num if 1 <= m <= 12]
        st.warning(
            "Diagnóstico do Sellout — "
            f"linhas carregadas: {len(df_sellout):,}; "
            f"anos reconhecidos: {anos_detectados or 'nenhum'}; "
            f"meses reconhecidos: {', '.join(meses_detectados) or 'nenhum'}; "
            f"filtro de ano aplicado: {sel_anos or 'todos'}; "
            f"filtro de mês aplicado: {sel_meses or 'todos'}."
        )
        if st.button("Restaurar filtros e exibir Sellout", key="reset_filtros_sellout"):
            st.session_state["filtro_anos_aplicado"] = anos
            st.session_state["filtro_meses_aplicado"] = MESES_LABELS
            st.session_state.pop("filtro_anos_input", None)
            st.session_state.pop("filtro_meses_input", None)
            st.rerun()
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

    # Novo KPI: Fornecedor -> Linhas -> Produtos
    st.subheader("Drill por Fornecedor — Linhas e Produtos")

    nome_so_drill = (
        df_sellout_f.groupby(["FORN_KEY", "FORNECEDOR_SELLOUT"], as_index=False)
        .size()
        .sort_values(["FORN_KEY", "size"], ascending=[True, False])
        .drop_duplicates("FORN_KEY")[["FORN_KEY", "FORNECEDOR_SELLOUT"]]
        .rename(columns={"FORNECEDOR_SELLOUT": "FORNECEDOR"})
    )

    forn_drill_tab = (
        df_sellout_f.groupby("FORN_KEY", as_index=False)
        .agg(FATURAMENTO=("FATURAMENTO", "sum"))
        .merge(nome_so_drill, on="FORN_KEY", how="left")
        .sort_values("FATURAMENTO", ascending=False)
    )

    if forn_drill_tab.empty:
        st.info("Sem fornecedores no SELLOUT para o recorte selecionado.")
    else:
        forn_options = forn_drill_tab["FORNECEDOR"].fillna("").astype(str).tolist()
        sel_forn_drill = st.selectbox(
            "Selecione o Fornecedor (SELLOUT)",
            options=forn_options,
            index=0,
            key="sellout_fornecedor_drill_select"
        )

        sel_forn_key = forn_drill_tab.loc[forn_drill_tab["FORNECEDOR"] == sel_forn_drill, "FORN_KEY"].iloc[0]
        so_forn_base = df_sellout_f[df_sellout_f["FORN_KEY"] == sel_forn_key].copy()

        linhas_forn = (
            so_forn_base[so_forn_base["LINHA"].astype(str).str.strip() != ""]
            .groupby("LINHA", as_index=False)
            .agg(FATURAMENTO=("FATURAMENTO", "sum"))
            .sort_values("FATURAMENTO", ascending=False)
        )

        total_forn = float(linhas_forn["FATURAMENTO"].sum())
        linhas_forn["% LINHA / FORNECEDOR"] = (linhas_forn["FATURAMENTO"] / total_forn) if total_forn != 0 else 0.0

        st.markdown("#### Linhas do Fornecedor")
        st.dataframe(
            linhas_forn.style.format({
                "FATURAMENTO": brl,
                "% LINHA / FORNECEDOR": lambda x: pct_str(float(x)),
            }),
            use_container_width=True,
            hide_index=True
        )

        if linhas_forn.empty:
            st.info("Esse fornecedor não possui linha preenchida no recorte selecionado.")
        else:
            linha_options = linhas_forn["LINHA"].astype(str).tolist()
            sel_linha_forn = st.selectbox(
                "Selecione a Linha do Fornecedor",
                options=linha_options,
                index=0,
                key="sellout_linha_fornecedor_select"
            )

            so_line_forn = so_forn_base[so_forn_base["LINHA"].astype(str) == str(sel_linha_forn)].copy()
            total_line_forn = float(so_line_forn["FATURAMENTO"].sum())

            desc_canon_forn = (
                so_line_forn.groupby("CODIGO", as_index=False)["DESCRICAO_PRODUTO"]
                .apply(most_frequent_nonempty)
                .rename(columns={"DESCRICAO_PRODUTO": "DESCRIÇÃO DO PRODUTO"})
            )

            prod_line_forn = (
                so_line_forn.groupby("CODIGO", as_index=False)
                .agg(FATURAMENTO=("FATURAMENTO", "sum"))
                .sort_values("FATURAMENTO", ascending=False)
            )
            prod_line_forn = prod_line_forn.merge(desc_canon_forn, on="CODIGO", how="left")
            prod_line_forn["DESCRIÇÃO DO PRODUTO"] = prod_line_forn["DESCRIÇÃO DO PRODUTO"].fillna("")
            prod_line_forn["% PRODUTO / LINHA"] = (prod_line_forn["FATURAMENTO"] / total_line_forn) if total_line_forn != 0 else 0.0

            st.markdown("#### Produtos da Linha Selecionada")
            st.dataframe(
                prod_line_forn[["CODIGO", "DESCRIÇÃO DO PRODUTO", "FATURAMENTO", "% PRODUTO / LINHA"]]
                .style.format({
                    "FATURAMENTO": brl,
                    "% PRODUTO / LINHA": lambda x: pct_str(float(x)),
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
# PAGE: HISTÓRICO POR FORNECEDOR
# -----------------------------
def render_historico_fornecedor_page():
    st.title("HISTÓRICO MENSAL POR FORNECEDOR")
    st.caption("Acompanhe, mês a mês, as Compras, o CMV e o Sellout do fornecedor selecionado.")

    # Cadastro canônico de fornecedores reunindo todas as fontes disponíveis.
    cadastros = []
    if not df_citel.empty:
        cadastros.append(
            df_citel[["FORN_KEY", "FORNECEDOR_CITEL"]]
            .rename(columns={"FORNECEDOR_CITEL": "FORNECEDOR"})
        )
    if not df_cmv.empty:
        cadastros.append(
            df_cmv[["FORN_KEY", "FORNECEDOR_CMV"]]
            .rename(columns={"FORNECEDOR_CMV": "FORNECEDOR"})
        )
    if df_sellout is not None and not df_sellout.empty:
        cadastros.append(
            df_sellout[["FORN_KEY", "FORNECEDOR_SELLOUT"]]
            .rename(columns={"FORNECEDOR_SELLOUT": "FORNECEDOR"})
        )

    if not cadastros:
        st.info("Não há fornecedores disponíveis nas bases carregadas.")
        return

    fornecedores_base = pd.concat(cadastros, ignore_index=True)
    fornecedores_base["FORNECEDOR"] = fornecedores_base["FORNECEDOR"].fillna("").astype(str).str.strip()
    fornecedores_base = fornecedores_base[
        (fornecedores_base["FORN_KEY"].astype(str).str.strip() != "")
        & (fornecedores_base["FORNECEDOR"] != "")
    ]

    fornecedores = (
        fornecedores_base.groupby("FORN_KEY", as_index=False)
        .agg(FORNECEDOR=("FORNECEDOR", most_frequent_nonempty))
        .sort_values("FORNECEDOR")
    )

    if fornecedores.empty:
        st.info("Não foi possível identificar fornecedores válidos nas bases.")
        return

    anos_hist = sorted(set(anos_citel + anos_sellout))
    if not anos_hist:
        anos_hist = anos if anos else []

    f1, f2 = st.columns([2, 1])
    with f1:
        fornecedor_nome = st.selectbox(
            "Fornecedor",
            options=fornecedores["FORNECEDOR"].tolist(),
            key="historico_fornecedor_select",
        )
    with f2:
        ano_padrao = anos_hist[-1] if anos_hist else None
        ano_hist = st.selectbox(
            "Ano",
            options=anos_hist,
            index=(len(anos_hist) - 1) if anos_hist else 0,
            key="historico_fornecedor_ano",
        ) if anos_hist else None

    fornecedor_key = fornecedores.loc[
        fornecedores["FORNECEDOR"] == fornecedor_nome, "FORN_KEY"
    ].iloc[0]

    # Compras: fonte CITEL, com ano e mês da data de emissão.
    compras_base = df_citel[df_citel["FORN_KEY"] == fornecedor_key].copy()
    if ano_hist is not None and "ANO" in compras_base.columns:
        compras_base = compras_base[compras_base["ANO"] == ano_hist]
    compras_mes = compras_base.groupby("MES_NUM")["COMPRA_VALOR"].sum()

    # CMV: a planilha CMV E ESTOQUE possui competência mensal, mas não possui ano.
    cmv_base_hist = df_cmv[df_cmv["FORN_KEY"] == fornecedor_key].copy()
    cmv_mes = cmv_base_hist.groupby("MES_NUM")["CMV_VALOR"].sum()

    # Sellout: utiliza ano e mês disponíveis na aba SELLOUT.
    if df_sellout is not None:
        sellout_base = df_sellout[df_sellout["FORN_KEY"] == fornecedor_key].copy()
        if ano_hist is not None and "ANO" in sellout_base.columns and sellout_base["ANO"].notna().any():
            sellout_base = sellout_base[sellout_base["ANO"] == ano_hist]
        sellout_mes = sellout_base.groupby("MES_NUM")["FATURAMENTO"].sum()
    else:
        sellout_mes = pd.Series(dtype=float)

    historico = pd.DataFrame({"MES_NUM": range(1, 13)})
    historico["MÊS"] = historico["MES_NUM"].map(
        {i + 1: nome for i, nome in enumerate(MESES_LABELS)}
    )
    historico["COMPRAS"] = historico["MES_NUM"].map(compras_mes).fillna(0.0)
    historico["CMV"] = historico["MES_NUM"].map(cmv_mes).fillna(0.0)
    historico["SELLOUT"] = historico["MES_NUM"].map(sellout_mes).fillna(0.0)
    historico["DIF. CMV - COMPRAS"] = historico["CMV"] - historico["COMPRAS"]

    meses_com_movimento = int(
        ((historico[["COMPRAS", "CMV", "SELLOUT"]].abs().sum(axis=1)) > 0).sum()
    )
    divisor_media = meses_com_movimento if meses_com_movimento > 0 else 12

    total_compras = float(historico["COMPRAS"].sum())
    total_cmv = float(historico["CMV"].sum())
    total_sellout = float(historico["SELLOUT"].sum())

    st.subheader(f"Resumo — {fornecedor_nome}" + (f" | {ano_hist}" if ano_hist else ""))
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Total Compras", brl(total_compras), help="Soma das compras do fornecedor no ano selecionado.")
    with k2:
        st.metric("Total CMV", brl(total_cmv), help="Soma do CMV mensal disponível na base CMV E ESTOQUE.")
    with k3:
        st.metric("Total Sellout", brl(total_sellout), help="Soma do faturamento Sellout do fornecedor no ano selecionado.")
    with k4:
        st.metric("Meses com Movimento", meses_com_movimento)

    a1, a2, a3 = st.columns(3)
    with a1:
        st.metric("Média Mensal de Compras", brl(total_compras / divisor_media))
    with a2:
        st.metric("Média Mensal de CMV", brl(total_cmv / divisor_media))
    with a3:
        st.metric("Média Mensal de Sellout", brl(total_sellout / divisor_media))

    st.subheader("Histórico mensal")
    st.dataframe(
        historico[["MÊS", "COMPRAS", "CMV", "SELLOUT", "DIF. CMV - COMPRAS"]]
        .style
        .format({
            "COMPRAS": brl,
            "CMV": brl,
            "SELLOUT": brl,
            "DIF. CMV - COMPRAS": brl,
        })
        .map(style_dif, subset=["DIF. CMV - COMPRAS"]),
        use_container_width=True,
        hide_index=True,
    )

    grafico = historico.melt(
        id_vars=["MES_NUM", "MÊS"],
        value_vars=["COMPRAS", "CMV", "SELLOUT"],
        var_name="INDICADOR",
        value_name="VALOR",
    )
    fig = px.line(
        grafico,
        x="MÊS",
        y="VALOR",
        color="INDICADOR",
        markers=True,
        category_orders={"MÊS": MESES_LABELS},
        title="Evolução mensal — Compras x CMV x Sellout",
    )
    fig.update_layout(
        xaxis_title="Mês",
        yaxis_title="Valor (R$)",
        legend_title_text="Indicador",
        hovermode="x unified",
        margin=dict(t=55, l=10, r=10, b=10),
    )
    fig.update_yaxes(tickprefix="R$ ", separatethousands=True)
    st.plotly_chart(fig, use_container_width=True, key="grafico_historico_fornecedor")

    st.caption(
        "Observação: a aba CMV E ESTOQUE possui o mês da competência, mas não possui uma coluna de ano. "
        "Por isso, o CMV exibido representa os valores mensais disponíveis nessa aba; Compras e Sellout respeitam o ano selecionado."
    )


# -----------------------------
# PAGE: ORÇAMENTO
# -----------------------------
def render_orcamento_page():
    st.caption(f"Versão do aplicativo: {APP_VERSION}")
    st.title("ORÇAMENTO")
    st.caption(
        "O CMV base é calculado pela média dos 3 meses mais recentes disponíveis na base de CMV. "
        "A participação de cada fornecedor é calculada sobre o CMV do período selecionado nos filtros globais."
    )

    if df_cmv.empty:
        st.warning("Não há dados de CMV disponíveis para montar o orçamento.")
        return

    cmv_mes = (
        df_cmv.dropna(subset=["MES_NUM"])
        .groupby("MES_NUM", as_index=False)["CMV_VALOR"].sum()
        .sort_values("MES_NUM")
    )
    cmv_mes = cmv_mes[cmv_mes["CMV_VALOR"].abs() > 0].copy()
    ultimos_meses = cmv_mes.tail(3)
    media_ult_trimestre = float(ultimos_meses["CMV_VALOR"].mean()) if not ultimos_meses.empty else 0.0
    nomes_ultimos = [MESES_LABELS[int(m)-1].title() for m in ultimos_meses["MES_NUM"].tolist()]

    if "orcamento_cmv_base_texto" not in st.session_state:
        st.session_state["orcamento_cmv_base_texto"] = brl(media_ult_trimestre)

    def _usar_media_automatica():
        st.session_state["orcamento_cmv_base_texto"] = brl(media_ult_trimestre)
        st.session_state.pop("orcamento_editor", None)
        st.session_state.pop("orcamento_assinatura", None)

    c1, c2, c3 = st.columns([1.2, 1.2, 1])
    with c1:
        cmv_base_texto = st.text_input(
            "CMV base mensal para orçamento (editável)",
            key="orcamento_cmv_base_texto",
            help="Use o padrão brasileiro, por exemplo: R$ 1.770.000,00",
        )
        cmv_base = parse_brl_value(cmv_base_texto)
        st.caption(f"Valor considerado: **{brl(cmv_base)}**")
    with c2:
        st.metric("Média automática — últimos 3 meses", brl(media_ult_trimestre))
        st.caption(" + ".join(nomes_ultimos) if nomes_ultimos else "Sem meses válidos")
    with c3:
        st.button("Usar média automática", use_container_width=True, on_click=_usar_media_automatica)

    participacao = (
        df_cmv_f.groupby(["FORN_KEY", "FORNECEDOR_CMV"], as_index=False)
        .agg(**{"CMV PERÍODO": ("CMV_VALOR", "sum")})
        .rename(columns={"FORNECEDOR_CMV": "FORNECEDOR"})
    )
    compras_periodo = (
        df_citel_f.groupby("FORN_KEY", as_index=False)
        .agg(**{"COMPRAS PERÍODO": ("COMPRA_VALOR", "sum")})
    ) if not df_citel_f.empty else pd.DataFrame(columns=["FORN_KEY", "COMPRAS PERÍODO"])

    participacao = participacao.merge(compras_periodo, on="FORN_KEY", how="left")
    participacao["COMPRAS PERÍODO"] = participacao["COMPRAS PERÍODO"].fillna(0.0)
    participacao["DIFERENÇA (CMV - COMPRAS)"] = participacao["CMV PERÍODO"] - participacao["COMPRAS PERÍODO"]
    participacao = participacao[participacao["CMV PERÍODO"] > 0].copy()

    total_part = float(participacao["CMV PERÍODO"].sum())
    if total_part <= 0:
        st.warning("O período selecionado não possui CMV positivo para calcular a participação dos fornecedores.")
        return

    participacao["PARTICIPAÇÃO"] = participacao["CMV PERÍODO"] / total_part
    participacao["PARTICIPAÇÃO %"] = participacao["PARTICIPAÇÃO"].map(pct_str)
    participacao["ORÇAMENTO CALCULADO"] = cmv_base * participacao["PARTICIPAÇÃO"]
    participacao["ORÇAMENTO FINAL"] = participacao["ORÇAMENTO CALCULADO"]
    participacao = participacao.sort_values("PARTICIPAÇÃO", ascending=False).reset_index(drop=True)

    assinatura = (tuple(sel_anos), tuple(sel_meses_num), round(float(cmv_base), 2), tuple(participacao["FORN_KEY"].tolist()))
    if st.session_state.get("orcamento_assinatura") != assinatura:
        st.session_state["orcamento_assinatura"] = assinatura
        st.session_state.pop("orcamento_editor", None)

    st.subheader("Distribuição do orçamento por fornecedor")
    st.caption(
        "Todos os valores monetários são exibidos no padrão brasileiro. "
        "Edite somente **ORÇAMENTO FINAL**, por exemplo: R$ 595.000,00."
    )

    editor_df = pd.DataFrame({
        "FORNECEDOR": participacao["FORNECEDOR"],
        "CMV PERÍODO": participacao["CMV PERÍODO"].map(brl),
        "COMPRAS PERÍODO": participacao["COMPRAS PERÍODO"].map(brl),
        "DIFERENÇA (CMV - COMPRAS)": participacao["DIFERENÇA (CMV - COMPRAS)"].map(brl),
        "PARTICIPAÇÃO": participacao["PARTICIPAÇÃO %"],
        "ORÇAMENTO CALCULADO": participacao["ORÇAMENTO CALCULADO"].map(brl),
        "ORÇAMENTO FINAL": participacao["ORÇAMENTO FINAL"].map(brl),
    })

    editado = st.data_editor(
        editor_df, use_container_width=True, hide_index=True,
        disabled=["FORNECEDOR", "CMV PERÍODO", "COMPRAS PERÍODO", "DIFERENÇA (CMV - COMPRAS)", "PARTICIPAÇÃO", "ORÇAMENTO CALCULADO"],
        column_config={
            "FORNECEDOR": st.column_config.TextColumn("Fornecedor", width="large"),
            "CMV PERÍODO": st.column_config.TextColumn("CMV no período"),
            "COMPRAS PERÍODO": st.column_config.TextColumn("Compras no período"),
            "DIFERENÇA (CMV - COMPRAS)": st.column_config.TextColumn("Dif. CMV - Compras"),
            "PARTICIPAÇÃO": st.column_config.TextColumn("Participação"),
            "ORÇAMENTO CALCULADO": st.column_config.TextColumn("Orçamento calculado"),
            "ORÇAMENTO FINAL": st.column_config.TextColumn("Orçamento final (editável)", help="Digite no padrão R$ 595.000,00"),
        }, key="orcamento_editor",
    )

    editado["ORÇAMENTO FINAL_NUM"] = editado["ORÇAMENTO FINAL"].map(parse_brl_value)
    total_final = float(editado["ORÇAMENTO FINAL_NUM"].sum())
    diferenca_base = total_final - float(cmv_base)
    total_cmv_periodo = float(participacao["CMV PERÍODO"].sum())
    total_compras_periodo = float(participacao["COMPRAS PERÍODO"].sum())
    total_dif_periodo = total_cmv_periodo - total_compras_periodo

    k1, k2, k3, k4 = st.columns(4)
    with k1: st.metric("CMV base", brl(cmv_base))
    with k2: st.metric("Orçamento final", brl(total_final))
    with k3: st.metric("CMV - Compras do período", brl(total_dif_periodo))
    with k4: st.metric("Diferença após ajustes", brl(diferenca_base))

    periodo_txt = ", ".join([m.title() for m in sel_meses]) if sel_meses else "Todos os meses disponíveis"
    st.caption(f"Participação calculada com base no CMV de: {periodo_txt}.")

    pdf_df = pd.DataFrame({
        "FORNECEDOR": participacao["FORNECEDOR"].values,
        "CMV PERÍODO_NUM": participacao["CMV PERÍODO"].values,
        "COMPRAS PERÍODO_NUM": participacao["COMPRAS PERÍODO"].values,
        "DIFERENÇA_NUM": participacao["DIFERENÇA (CMV - COMPRAS)"].values,
        "PARTICIPAÇÃO": participacao["PARTICIPAÇÃO"].values,
        "ORÇAMENTO FINAL_NUM": editado["ORÇAMENTO FINAL_NUM"].values,
    })
    pdf_bytes = build_budget_pdf(pdf_df, cmv_base, periodo_txt)
    st.download_button("Baixar orçamento final em PDF", data=pdf_bytes, file_name="orcamento_compras.pdf", mime="application/pdf", use_container_width=True)


# -----------------------------
# Render
# -----------------------------
if page == "COMPRAS":
    render_compras_page()
elif page == "SELLOUT":
    render_sellout_page()
elif page == "HISTÓRICO POR FORNECEDOR":
    render_historico_fornecedor_page()
else:
    render_orcamento_page()
