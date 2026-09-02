import pandas as pd
import streamlit as st
import plotly.express as px
import re
import unicodedata
import os
from datetime import datetime
from zoneinfo import ZoneInfo

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None

st.set_page_config(page_title="Indicador de Compras", layout="wide")
APP_VERSION = "2026-09-02.22 — Memória de fornecedores no Google Sheets"
FILTER_STATE_VERSION = "orcamento-checklist-v4"


GIRO_NOTAS_PATH = "GIRO E NOTAS.xlsx"
CAD_FORNECEDORES_PATH = "CADASTRO DE FORNECEDORES.csv"
CAD_PRODUTOS_PATH = "CADASTRO PRODUTOS GERAL.csv"
SELLOUT_PATH = "sellout.csv"
NOTAS_ENTRADA_PATH = "NOTAS DE ENTRADA.csv"
SUPPLIER_MEMORY_PATH = "MEMORIA_FORNECEDORES.csv"  # fallback local, usado apenas sem Secrets
SUPPLIER_MEMORY_SPREADSHEET_ID = "1jx--9QLyCTeH1KC6DNhoaJ4dTd0RhzabotK5YUSgq8A"
SUPPLIER_MEMORY_WORKSHEET = "MEMORIA_FORNECEDORES"

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
# Memória de fornecedor por produto
# Google Sheets em produção + fallback CSV para desenvolvimento local
# -----------------------------
SUPPLIER_MEMORY_COLUMNS = [
    "COD_KEY", "FORNECEDOR_VALIDADO", "FORNECEDORES_JA_VISTOS", "ATUALIZADO_EM"
]

def _supplier_memory_uses_sheets() -> bool:
    """Retorna True quando as credenciais do Google estiverem configuradas no Streamlit Secrets."""
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False

@st.cache_resource(show_spinner=False)
def _supplier_memory_worksheet():
    """Abre a aba persistente da memória de fornecedores no Google Sheets."""
    if gspread is None or Credentials is None:
        raise RuntimeError(
            "Dependências do Google Sheets ausentes. Adicione gspread e google-auth ao requirements.txt."
        )

    try:
        service_account_info = dict(st.secrets["gcp_service_account"])
    except Exception as exc:
        raise RuntimeError(
            "Credenciais gcp_service_account não encontradas nos Secrets do Streamlit."
        ) from exc

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_info(
        service_account_info, scopes=scopes
    )
    client = gspread.authorize(credentials)

    spreadsheet = client.open_by_key(SUPPLIER_MEMORY_SPREADSHEET_ID)
    try:
        worksheet = spreadsheet.worksheet(SUPPLIER_MEMORY_WORKSHEET)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=SUPPLIER_MEMORY_WORKSHEET, rows=1000, cols=len(SUPPLIER_MEMORY_COLUMNS)
        )

    # Garante cabeçalho mesmo em uma aba recém-criada/vazia.
    header = worksheet.row_values(1)
    if header != SUPPLIER_MEMORY_COLUMNS:
        worksheet.update(
            range_name=f"A1:D1",
            values=[SUPPLIER_MEMORY_COLUMNS],
        )
    return worksheet

@st.cache_data(ttl=60, show_spinner=False)
def _load_supplier_memory_from_sheets(cache_buster: int = 0) -> pd.DataFrame:
    """Carrega a memória do Sheets; cache_buster permite invalidação imediata após gravação."""
    worksheet = _supplier_memory_worksheet()
    rows = worksheet.get_all_records(expected_headers=SUPPLIER_MEMORY_COLUMNS)
    if not rows:
        return pd.DataFrame(columns=SUPPLIER_MEMORY_COLUMNS)
    mem = pd.DataFrame(rows).fillna("")
    for col in SUPPLIER_MEMORY_COLUMNS:
        if col not in mem.columns:
            mem[col] = ""
    return mem[SUPPLIER_MEMORY_COLUMNS]

def _clear_supplier_memory_cache():
    """Invalida somente o cache leve da memória, sem limpar as bases pesadas do dashboard."""
    _load_supplier_memory_from_sheets.clear()

def load_supplier_memory(path: str = SUPPLIER_MEMORY_PATH) -> pd.DataFrame:
    """Lê as decisões manuais. Em produção usa Google Sheets; sem Secrets usa CSV local."""
    if _supplier_memory_uses_sheets():
        mem = _load_supplier_memory_from_sheets().copy()
    else:
        if not os.path.exists(path):
            return pd.DataFrame(columns=SUPPLIER_MEMORY_COLUMNS)
        try:
            mem = pd.read_csv(path, sep=";", dtype=str, encoding="utf-8-sig").fillna("")
        except Exception:
            return pd.DataFrame(columns=SUPPLIER_MEMORY_COLUMNS)

    for col in SUPPLIER_MEMORY_COLUMNS:
        if col not in mem.columns:
            mem[col] = ""
    mem["COD_KEY"] = mem["COD_KEY"].map(product_key)
    mem = mem[mem["COD_KEY"] != ""].drop_duplicates("COD_KEY", keep="last")
    return mem[SUPPLIER_MEMORY_COLUMNS]

def supplier_memory_map(path: str = SUPPLIER_MEMORY_PATH) -> dict:
    mem = load_supplier_memory(path)
    if mem.empty:
        return {}
    return dict(zip(mem["COD_KEY"], mem["FORNECEDOR_VALIDADO"]))

def save_supplier_decision(cod_key: str, fornecedor_validado: str, fornecedores_vistos, path: str = SUPPLIER_MEMORY_PATH):
    """Persiste uma decisão. No Sheets atualiza apenas a linha do produto, sem regravar a tabela inteira."""
    cod_key = product_key(cod_key)
    fornecedor_validado = str(fornecedor_validado or "").strip()
    vistos = sorted({str(x).strip() for x in fornecedores_vistos if str(x).strip()})
    if not cod_key or not fornecedor_validado:
        raise ValueError("Código e fornecedor validado são obrigatórios.")

    atualizado_em = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%Y-%m-%d %H:%M:%S")
    values = [
        cod_key,
        fornecedor_validado,
        " || ".join(vistos),
        atualizado_em,
    ]

    if _supplier_memory_uses_sheets():
        worksheet = _supplier_memory_worksheet()

        # Localiza COD_KEY na coluna A. Se existe, atualiza somente a linha;
        # se não existe, acrescenta uma nova associação.
        try:
            cell = worksheet.find(cod_key, in_column=1)
        except Exception:
            cell = None

        if cell and cell.row > 1:
            worksheet.update(
                range_name=f"A{cell.row}:D{cell.row}",
                values=[values],
            )
        else:
            worksheet.append_row(values, value_input_option="RAW")

        _clear_supplier_memory_cache()
        return

    # Fallback local para desenvolvimento sem Secrets.
    mem = load_supplier_memory(path)
    nova = dict(zip(SUPPLIER_MEMORY_COLUMNS, values))
    mem = mem[mem["COD_KEY"] != cod_key].copy()
    mem = pd.concat([mem, pd.DataFrame([nova])], ignore_index=True)
    mem.to_csv(path, sep=";", index=False, encoding="utf-8-sig")

def remove_supplier_decision(cod_key: str, path: str = SUPPLIER_MEMORY_PATH):
    """Remove uma decisão da memória e faz o produto voltar para análise."""
    cod_key = product_key(cod_key)
    if not cod_key:
        return

    if _supplier_memory_uses_sheets():
        worksheet = _supplier_memory_worksheet()
        try:
            cell = worksheet.find(cod_key, in_column=1)
        except Exception:
            cell = None

        if cell and cell.row > 1:
            worksheet.delete_rows(cell.row)
        _clear_supplier_memory_cache()
        return

    # Fallback local para desenvolvimento sem Secrets.
    mem = load_supplier_memory(path)
    if mem.empty:
        return
    mem = mem[mem["COD_KEY"] != cod_key].copy()
    mem.to_csv(path, sep=";", index=False, encoding="utf-8-sig")


def build_product_description_map(*dfs) -> dict:
    """Monta uma descrição amigável por código usando as bases já carregadas."""
    desc_map = {}
    candidates = ("DESCRICAO_ITEM", "DESCRICAO_PRODUTO")
    for df in dfs:
        if df is None or df.empty or "COD_KEY" not in df.columns:
            continue
        for desc_col in candidates:
            if desc_col not in df.columns:
                continue
            base = df[["COD_KEY", desc_col]].copy()
            base["COD_KEY"] = base["COD_KEY"].map(product_key)
            base[desc_col] = base[desc_col].fillna("").astype(str).str.strip()
            base = base[(base["COD_KEY"] != "") & (base[desc_col] != "")]
            for _, row in base.iterrows():
                cod = row["COD_KEY"]
                if cod not in desc_map:
                    desc_map[cod] = row[desc_col]
    return desc_map


def _valid_supplier_name(value) -> bool:
    txt = str(value or "").strip()
    return bool(txt) and txt != "FORNECEDOR NÃO IDENTIFICADO"

def build_supplier_review_candidates(df_cmv, df_ent, df_sellout, memory_path: str = SUPPLIER_MEMORY_PATH) -> pd.DataFrame:
    """
    Localiza produtos cujo fornecedor mudou/diverge entre as bases.
    Uma divergência já analisada não reaparece; somente fornecedor novo reabre a revisão.
    """
    rows = []

    def add_source(df, supplier_col, source, desc_col=None):
        if df is None or df.empty or "COD_KEY" not in df.columns or supplier_col not in df.columns:
            return
        cols = ["COD_KEY", supplier_col] + ([desc_col] if desc_col and desc_col in df.columns else [])
        base = df[cols].copy()
        base["COD_KEY"] = base["COD_KEY"].map(product_key)
        base = base[base["COD_KEY"] != ""]
        for _, r in base.iterrows():
            forn = str(r.get(supplier_col, "") or "").strip()
            if not _valid_supplier_name(forn):
                continue
            rows.append({
                "COD_KEY": r["COD_KEY"],
                "FORNECEDOR": forn,
                "FONTE": source,
                "DESCRICAO": str(r.get(desc_col, "") or "").strip() if desc_col else "",
            })

    add_source(df_cmv, "FORNECEDOR_CMV_ORIGINAL", "CADASTRO/GIRO", "DESCRICAO_ITEM")
    add_source(df_ent, "FORNECEDOR_ENT_ORIGINAL", "ENTRADAS", "DESCRICAO_ITEM")
    add_source(df_sellout, "FORNECEDOR_SELLOUT_ORIGINAL", "SELLOUT", "DESCRICAO_PRODUTO")

    ev = pd.DataFrame(rows)
    if ev.empty:
        return pd.DataFrame(columns=["COD_KEY", "DESCRICAO", "FORNECEDOR_ATUAL", "SUGESTAO", "CANDIDATOS", "NOVOS", "FONTES"])

    mem = load_supplier_memory(memory_path)
    mem_idx = mem.set_index("COD_KEY").to_dict("index") if not mem.empty else {}
    out = []
    for cod, grp in ev.groupby("COD_KEY"):
        counts = grp["FORNECEDOR"].value_counts()
        candidatos = counts.index.tolist()
        if len(candidatos) <= 1:
            continue

        m = mem_idx.get(cod)
        if m:
            vistos = {x.strip() for x in str(m.get("FORNECEDORES_JA_VISTOS", "")).split("||") if x.strip()}
            novos = [x for x in candidatos if x not in vistos]
            if not novos:
                continue
            atual = str(m.get("FORNECEDOR_VALIDADO", "")).strip()
            sugestao = atual or candidatos[0]
        else:
            novos = candidatos
            # Prioriza evidência operacional (Entradas/Sellout) sobre o cadastro atual.
            op = grp[grp["FONTE"].isin(["ENTRADAS", "SELLOUT"])]
            sugestao = str(op["FORNECEDOR"].value_counts().index[0]) if not op.empty else candidatos[0]
            atual = sugestao

        desc = most_frequent_nonempty(grp["DESCRICAO"])
        fontes = "; ".join(
            f"{f}: {', '.join(g['FORNECEDOR'].value_counts().index.tolist())}"
            for f, g in grp.groupby("FONTE")
        )
        out.append({
            "COD_KEY": cod, "DESCRICAO": desc, "FORNECEDOR_ATUAL": atual,
            "SUGESTAO": sugestao, "CANDIDATOS": candidatos, "NOVOS": novos, "FONTES": fontes
        })
    return pd.DataFrame(out)


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
def load_data(giro_path: str, cad_forn_path: str, cad_prod_path: str, sellout_path: str, entradas_path: str, memory_path: str = SUPPLIER_MEMORY_PATH, cache_token=None):
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
    cad_prod["FORNECEDOR_PROD_ORIGINAL"] = cad_prod["FORNECEDOR_PROD"]
    cad_prod["LINHA_PROD"] = cad_prod[c_prod_linha].fillna("").astype(str).str.strip()
    cad_prod["MARCA_PROD"] = cad_prod[c_prod_marca].fillna("").astype(str).str.strip() if c_prod_marca else ""
    cad_prod = cad_prod[cad_prod["COD_KEY"] != ""].drop_duplicates("COD_KEY", keep="last")

    # Uma decisão manual passa a prevalecer sobre alterações futuras do cadastro.
    memory_map = supplier_memory_map(memory_path)
    if memory_map:
        mem_supplier = cad_prod["COD_KEY"].map(memory_map).fillna("")
        cad_prod["FORNECEDOR_PROD"] = mem_supplier.where(mem_supplier != "", cad_prod["FORNECEDOR_PROD"])
    prod_lookup = cad_prod[["COD_KEY", "FORNECEDOR_PROD", "FORNECEDOR_PROD_ORIGINAL", "LINHA_PROD", "MARCA_PROD"]]

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
    giro["FORNECEDOR_CMV_ORIGINAL"] = [supplier_division(n, h) for n, h in zip(giro["FORNECEDOR_PROD_ORIGINAL"], giro_hint)]
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

            # Ajuste específico da aba COMPRAS: os dois agrupamentos Sherwin Williams
            # chegam invertidos em relação ao resultado desejado no dashboard.
            # Fazemos a troca somente na base de compras (CITEL), sem alterar CMV,
            # estoque, sellout ou demais indicadores.
            sherwin_swap = {
                "SHERWIN WILLIAMS DO BRASIL INDUSTRIA E COMERCIO":
                    "SHERWIN WILLIAMS DO BRASIL IND E COM DIV AUTOMOTIVA",
                "SHERWIN WILLIAMS DO BRASIL IND E COM DIV AUTOMOTIVA":
                    "SHERWIN WILLIAMS DO BRASIL INDUSTRIA E COMERCIO",
            }
            notas["FORNECEDOR_CITEL"] = notas["FORNECEDOR_CITEL"].replace(sherwin_swap)

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
    ent["FORNECEDOR_ENT_ORIGINAL"] = [supplier_division(n, h) for n, h in zip(ent[c_e_forn], ent_hint)]
    ent["FORNECEDOR_ENT"] = ent["FORNECEDOR_ENT_ORIGINAL"]
    if memory_map:
        mem_ent = ent["COD_KEY"].map(memory_map).fillna("")
        ent.loc[mem_ent != "", "FORNECEDOR_ENT"] = mem_ent[mem_ent != ""]
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
    # Mantém uma versão sem a memória para detectar mudanças futuras.
    fornecedor_original_so = so["FORNECEDOR_PROD_ORIGINAL"].fillna("").astype(str).str.strip()
    fornecedor_original_so = fornecedor_original_so.where(fornecedor_original_so != "", fornecedor_relatorio)
    so["FORNECEDOR_SELLOUT_ORIGINAL"] = [supplier_division(n, h) for n, h in zip(fornecedor_original_so, sellout_hint)]
    so["FORNECEDOR_SELLOUT"] = [supplier_division(n, h) for n, h in zip(fornecedor_base_so, sellout_hint)]
    if memory_map:
        mem_so = so["COD_KEY"].map(memory_map).fillna("")
        so.loc[mem_so != "", "FORNECEDOR_SELLOUT"] = mem_so[mem_so != ""]
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
    # A memória agora pode vir do Google Sheets. O fingerprint invalida o cache
    # das bases quando uma associação é alterada, sem depender do mtime de CSV local.
    _mem_for_token = load_supplier_memory(SUPPLIER_MEMORY_PATH)
    _memory_fingerprint = tuple(
        _mem_for_token.fillna("").astype(str).itertuples(index=False, name=None)
    )
    cache_token = tuple(
        (os.path.getmtime(path), os.path.getsize(path)) for path in arquivos_origem
    ) + (_memory_fingerprint,)
    df_cmv, df_citel, df_ent, df_sellout = load_data(
        GIRO_NOTAS_PATH, CAD_FORNECEDORES_PATH, CAD_PRODUTOS_PATH,
        SELLOUT_PATH, NOTAS_ENTRADA_PATH, SUPPLIER_MEMORY_PATH, cache_token
    )
except Exception as e:
    st.error(f"Erro ao carregar as bases: {e}")
    st.stop()

# Divergências de fornecedor são calculadas com as versões ORIGINAIS das bases.
# Assim, uma decisão gravada corrige o painel sem esconder a chegada de um fornecedor novo.


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
# Interface BI — Dauto + Única
# -----------------------------
BRAND_IMAGE_B64 = 'iVBORw0KGgoAAAANSUhEUgAAALkAAAB8CAYAAAA1iBFbAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAACJNSURBVHhe7Z15tCx1feA/v1q6q/tub2MVRIMIBg+IqOwIPEBQBsRtTDTiKBrHiQvumJkobiRzkvGYM2OUBIQIamZIYjDxGImOgjrqAR3IoBJZ30N4vPfu0re7q2v/zR/V1V1dXVVd3X3vfffd259z+nTV9/er6q5Pf/tXv9qFlFIyYcIGRkkGJkzYaEySfMKGZ5LkEzY8kySfsOGZJPmEDc8kySdseMSB2IW4kh8phEiGJqSwmZ2vepKv8uz7ONh+gNVgrZ2zzr2vSpKvwixHZj3LXynWk2/WofMVS/IVms2qsd7Ej8t69x3nQLsfK8nHmDSX+HxXQ9BqzHOtWGnnafNbLT+rNd9BDJ3kQ1bvY9zp44wrbdzp15JxvY07fZJx3I0z7SgMleRDVO0wyjTjMIrAUaZZK0bxN8o04zKKw1GmGYVCSV6gSg/D1o/Im24UIcNOM2z91STPRRqrVX9UJ8NMN0zdURiY5AOKOxSpV6TOKBSRVKQOQ9RbTYp6KlJvcJ285c2ftoirInUihqk7DJlJnhHuI69eetngBYlPN3jB+z8jb5q8sjhF660k6b76yauXXpa9LOn1Q/Id9E+XVz+vLE7ResOQmuQpoT7y6vSW9X5pKSWO62M7Hp7v43o+vh/g+ZIgCAgCSdCeXhECRREoioKmClRVQddUNFWlXNIo6WqKlGJ/kLyyiCJ1Voo8nxF5dQY5t10Px/HxfB/PC0LvQdu5lARBd/oD7Z0h6hWhJ8nzJEZk1cmSLKXE9Xws28VyPGzHw/cDjLLeEaZpKrqmoioKQgii5ZMynN4Pwh/F8/zOH8SyXVRVoVzSMEoaRllH15LyB4vPiscpUmdUsnzGyaqT79zDcnzstveizsPpN5b3TpJniYzIKk8TLaXE9wNatotpOVi2R7mkMV0tUzF0SroWm2Z0HNejZbk0TBvb8TDKGlWjRKWso6rhjxcyvvRB5aOQ5TQiqzzLuecHtCyHlu3Ssj2MVXDOQehdyJBkvIe08m6stwVxXJ+madG0XDRVZXbaYLpaRlHG/7J5BIGkYdosNyw832fK0JmqGimr1vB7p8lLi8UZVD4MaU7jpJUPcm7aHqqirJlzDhLvIoh3xhIMLbpl0zAdqpUSc9MGRlnv1FlLLNul1rAwWw7T1RJTlXJCen4LkxaLyCsrSprXiLSyLOeu59MwbRqmTbVSPqDOKeSdkZM9r2wQmUleRLZsb7DUTYt606ZS1tk2N4Wuqz3THShc12eh1qRlu8xMlZmpGiiK6Ev2NIFpsYi8skGkeY1IK0tz7geSxjp1zip6p0B5GqlJnpSdJhrAbDksNy2khB3bpqkcwFYkj5btsn+hgRAwO2VQrZSgR1i69OR4nLyyLJJe4yTL8pzXTYsgWN/OWSXvFChP0pPkSdG9sa7sQEqWGy3qTZsts1W2zlZ7plmvLC6bLC2bzEyVmZ2uoIjBrUtyPE5eWZI0txHJsjTnUkpqDYvlhsXWuYPHORne6fhL954Vi8grS9JJ8qTobqy3JXE9j6W6hecFHLp9hnJp5bba1wLb8dg7X0fTFLbMGOha+P3jyZ4mMC1GTjxOmlsy4mnOHden1rTwXP+gdM5A78MnelY8DREEQd/elbSWBMCyPRaXm5RLOodun4lNcfCxd76O7bhsnZ3CKK9eoifdRiTj2c5dFpfNDeGcNfQeR/i+32M7S3bTcliqW8xUy2ybO3hWlXks1EzqpsWWmQpTxnj9xbR4MpEjkvFM5y2bxbrF7JSxYZyzBt6T9CR5XHb8h2iYFot1i22zVeZmKp34RqBWb7GwbLJ1xmC6anTiobzRW5ZkIseJl2U5b5o2iw2LrTOVDeecAt67w/F4OnllxG9JkSW70W5Nts9NbUjZczMVts9NsVi3aLTsTjx0ELrIS9g4ReoVSvCWzUK9tSEblYhB3rvD8Xg6eWX033elV3bLdqm1W/DZ6e6/baMxO22wbbZKrW7Rsp1OfBTh5JSnzyPp3OmsNTeyc1bBexYKsZnGZ+J4PrV6i5kpY8O2JnHmZirMTBnU6haO53ficeFxsoQHQZAMQWb9pHOPWrsPvhmcs4Les+IAStrMAhlQq7fQdW1DbfAMYttcFV3XqNVb+LKbrGmNQDfeP54V7x1Pcd6wN51zVsB7RFZcSa4WJJLlho3rBRy2AXZZDcth22dwvYB6w0LGznEJKS48IlkeT/C481rDwnX9TemcVfAeJ9ZdCWWblku9aW9a2bSF15s2puV0hMelpglOxpLjyVif88bmds6I3pOk1Wl3V0LZfhDQbNrMzVQOyqNqK0W5pDE3U6XRdPCDICG8+AZRerzbKknCK3MaTYsts9VN7ZwV9J6kZ+9Kw3Twpdx0fcI0ts1V22f7dbf6SQjvjfXT/4P0r3bDHzT8vAkx783ubkUyvGeRdBx2V2ifm9y02bFluqfCZuaQrdM0mjau5w/sJybFdlrrjP3sofOAetPikK39zq+//nosy+pMv5FfSQ7ZOk3DdEbynoYSzaTetDFKOhVj/Z66udZUDB2jpHda80H9xHzhvd0UgHrTwiiXJs4TFPEeJy0ejym0zxAzWy7btkxWmUm2banSNB1sx+uJhxL7ux/dsu6wbF8YHMdxPcyWM3GeQdJ7sn+e9JlGVEehfSK+UdZW9GLXjUJJ1zDKGmart1UhJjEuPH4wKO2HiKZvmg5GeWUvMN5IJL2nkWxMslA8P8C03IPqJPy1ZutsFdNy8fwwgZOtSj5hqyOl7EwXXlU/cT6I8bx3UVqWi6IoB/QC2PWOUdZRFIWW5XZiyZYjSuTkcBoty0VMnA8kzXuSQa25lBKlZbvMbfATgVaCuWmDlu12E7lQq9K/sSmlxJo4L8xo3ntRLNtjulpOxickmK6WsWyvs+ocBdm+AdDEeXFWwrtilDRUdfKkw0GoqoJR0rBst7NxGW9VZGwPSjQe0duKh3cTmzgvRtx7Vms+qMuiRLcJmDCYaqWEbffuSswjTbjteJ1bM0woxrDekyjV9jV2EwZTNUrd/bYpCUzOLsRo2HY8Js6Ho2qUcNzwXPMs72muI5TNflLQMJRLGq4f4HqxRCZ+eDraEOrf4KR9awnPCzb9iVjDEnmPEp2ULksek47hkJRLWt/Rz6K4rk+pfRuGCcNRLmk9ST4Mwye5lKO9NghGScdtX6aVXC1GJOPRuOv5VEqTfeOjkOY9vpaMx5OkPmkiD9loYt/7c4L5Pe2ZSgQCiFYbEqTorEFEqYJ69NGUTj6Jnju9t5GOC66DmKoOXO0caKSU1Ootmi2bHVunespEu4sihNJ20f8j7F8yma6UC12/ef3113PNNddQLvfvagzbjXAPjhDhLRniauPl0KtdUQRShrdcjuLJDFCUcJ7RfKLPWC2KzDvyvn1Ltb284TRh7qXfQqRTZ5gkl1Jif/NbiKOfgX7i89p5LRDtvI7nefguwfOw/uV7lI49Fu2E43qMS8/F/PPPE1Q0pt/xH9sJsn6RUmJaDvOLTQ7dPtUrNJbkUgZhksSSXErJ3oUmO7ZOF9rwzEtyy/Z46NeLuK6PoioceeQ0W7cYqKqC5wU0TYfdu5fxvAAZwLOePcfTe5rMzRkcccQ0luXxwAP7ePaz59i3z6Rpep0/pkBw7LFbMAydet3m6b1Ntm4x2L69iq4rnXueR3+AlbgHepEkT3ofJsmHyioBuLseR33WMSiKhvB85BNPIk0TRdVQNA1F1cA0kU8+GcbKBtqzj8Hdu6evyfB27ca5458wLr+yPff1j6YqeH7yqvJs4uW+H6CtwP7xWs3mVa+9nZ0Xf4WdF93GGWfezAO/2Ifn+ezaXeOSS7/G+RfcyoUXf4WdF9/Gffft5feuuoNbvnw/Ukoe31Xjkld8jXvufYq3v+ObXPSyr3D+ztu48KKvcMnLv8o99+7h+3c9zikvupHzLriVF774Jv7hjgex7ai7EG5fLC5aBEG40R0E4StaS0Rrkug1LsN6jzO0cSkURHsy799+Tf3q38f+7vdiFSTOt7/L8rveh4ztThPt7ySlhMBH+gHOA7/GuPYDqIcdCjLAD3xkECA9j2B+kcCyexZGBgEEPkhJEHjd+UuJ9HwC30N6bvjyva5dKZGOQzC/iHSdsayritLzEKlhCAKJqgytvI+SrvKed72YG294BXd8/bUcd9w2fu+qO3hqT5OPX3c3y8s2f/O1V/HXN1/OB99/Otu3VcI9O277IJaU2LaHrqt84fOX8qUbL2Nutsx/euep/OMd/54jj5jm/R/4DiefdCjf+PrreNMbn8+733snv/zV/vbKOeDOOx/lBz/czeKSxc9/voe77tpFo+Fg2x7751tYto/r+jhO+D4u43gf2ninvZUS6djI3zyJbDZ7KzUa+E/u6QlFXy9YXmbxzPNp3XILU5fupHLxRShCwfmb26lffAXBnj0sXPVWFs44m9ollxM8+WRnHubf3s7+s87De+RRll52Oa1P/zHS9bB//BOW3vQWai86i8UXnMHiKaexcN7LkLYNUuLPL7D0+jeydPo51N76DoJWqzPPYVEUgZ9yb5Vk/zs5TjvJV2L1blkef/7f72F+3uSFpxzONe99CY89VmP/PpMf/Z8nkMC3/vlhbr7lfj71mR+y5+kGMvZUPQh/EE0VPPe523neCdsplVSOOmqWk046lEceXaJluXz2zy7i1FMP55prTqPVcnnooUVk+6kit//dL7lw57MIgoClmk2prPLha7/Lzbfcz+c+91P+4gv38vFP3MWtt/0rP/zR7vjXH4m4956GL7ErMa2FHzrJZezVbSl7K8iwsx8LdpEyQNbrSNuJtprCr2rb0GjiPfAL5M/uY/rP/hRlZor6Jz8N0cI5NtTq4PvIRp3ANAmaTcwPXotWLmFc+0Gm/vg6pq//JNN/dC20z9W2vnUn8jdPUf3Tz8DPf45z/7925lmUSF60QRaPxYfDg0G9srunARTrfw4iCCSu62M7Pqbp8sMfPcFUVadUVjEMlfn5Fvfcu4eHHl7E84LwcYWq4P779zK/0KJlugQy7L+qqkBRFYQIE0lVFQ4/fArPDbjn3qdoNl1+9ct5FEWwbasBAsymy3HHbaNS0VGEwjOPnkVVBae+8Aief+IODj9siscfrzE3Z1Ct6Dz44HxyEYYm7r0o0W8wdJJH3Y4iP1VUl/Bnj8VF7x+jjRSA4yA8l/JF56OccybBA7/Ce+TR/i6GBBEENG66BakIyh/7Q4zXXEn5FZdSesWllC+8AKEoBLaNc9NNlM4/j/LO8xAnnoh1w01IN/v0zaL0txpFrIyPogoUIfj8F+7lVa+5nZtuuo/LLnsOz3jGDGeecRSHHFLlL/7HJXzkQ2dQLqtUKzovu+i3+PFPfsOVr76dd73n26hqdyNSEQJdV1DVcC/N8cdv5+WXHsu1H/3fvPJVt/P77/wmJ/72IZx08mEIwhb1lVccjxCCubkyz3rWFk594RG8+aqTOP30o3jb207hT66/gHe/68VceeXxXP3WU5KLsKYMneRFfkeZaNyjWHc4WRoSbS+HpYLq1W9G1Ou4d9/V1/IKwj66f++9qMcei7ZjB0JRQSjdl5TYP/gBcu9+1DNfgjCqlM4/H/nTn+A9/EjP/Ioi27vU+hN8MCL1jzE81YrGW95yMv/usudy1plH87E/OofPffZi5mbLXP3WF+A4Pu9577ep1x3edvUpHH74FB/5yJlc97FzOefsoznn3GfyB+88lSOPnAYEs7Ml3vzmk3n+iYcghKBa0fnEdS/lfdecxmkvOZI3X3USt//PV7FlLtzTc9ihUxz/3O0IAZqm9L0MQ8MwNCrtd00bPs2SpHkv6nL4T5exPBdhQorYVi9IROCj9LTdJP4dMvW/En7ndpkEMbsV/Z1vx/yrWwgct93Ut+sCAoGiqgjSu0fS9fD+5fso27aiv+RFoKqoZ52GROB+53t9f5wipG08ZslO/iCKIkbeeIozN2dw7YfP4vpPn8enPnkeV7/1BVSrOooSdhm+etsr2T9v8oEPfYeHHl5g1+5lnn66yZWvPJ6Pf+xcPnXdS/nEdS/l6GfOEQQBW7cavP+a0zj11CPwgwDTdPG8gIsufDave+3zOO+lx/CNf3qIn/0sPDYSdXPWkjTvRRl6qk6eCYHYMgcI3J/8FP/xXfhP7sF7+BGsb/wjYrb/Ngt0Uj1A2E6YZFKGG0XLDVAUpBCdv4BQFYyLLkLxfMy//7twj0mnnQepCLRzz8Z78EG8p/eAH+55iV7B3v14P/oxyuWXom7fAUKgn3AC6vN+G/eeewiSG8wF8IOwjzsKWRutoxBuzvQeBIKwZX3ByYdx57d+l5tvuozHHqvxmtf9LefvvJUzzr6ZM8++mXPP/zLn77yVSy79Ghdd8lUuuPA2zt95K+ee92XOPOsWTj/rFs485xbO23krL7v0q1xx5f/iP/+X7/HUnkbvh60h43gf6mAQUlL/wl9SfcPvoM7OID2X5Xd/AP/uH4DrholHgJzbQvUj76dyxRUgBM4vfom/bz+Vc89GOg5Lr34N7FvGuOZdqNUKXm2Z1qf/K/p/eD3lU06l/gfXsP3+exGVMtJ1Mf/oOpzv3oVy7un437yTuW/+A0uvfwPGeS+l8rE/pHbFqwk8n8pVb0KbmwMBUtORAqwPfJTpO+9APeKIzsGm1ne+i/nBjzD3lb9GO/74ngNUWUSaTMth/2KTQ7eF12f2HhBSOi2dpP9sxH0LTXZsmxn7YFARwo3TAMtyefKpBj/72R6e+E2dVstlcdFiuW5jNj0CKSnpKjMzJaamdGZmSpTL4UXEO3ZUeM6xWznyyBmmp0tMz5Taz+VMftp4FNkYj3vvORgkxMADQkMnee2Lf8nU7/4O2mx43z5pO3j334esN8IkFwJ27EA74TgUvQQS3Ad+ib9/P8a5Z4MQuLt30/pvn8X9zt0QSBRVQb3iMqbe9268+++ncc2H2PajuxHlEsgA+5GHsd74drx9e1Bfezlzn/wUSy+/gtI5Z1H96IfxHnsc5+Zbsf/520izhUASTE+jHX0Uym8dw/RnPoHQ9I6MoNWi9oar0M88g6n3v7eQ5EhTrd6i0bLZPhcemh8myRdqLaarxW7LPG6Sx/H98DHw0QGb+IEbCFeLihAIpf3eXkOEe1tE2IKKyN7KU8R/rd6iYVr9h/VXI8nNv/862vOfj37cc3rifUQf5Hq0vn832hGHUjrxRJASe9du9Gccjvebp1AO2U6w5ym0o55J4HkopRLeE0+gP/MYpGUjyiWc/fNoM1Wc+SWMww4jaDZRqhX8Wg3t0EMJLBtsm8C2EFUDSmXkch2lWsX3XGi0UKoVqFYRrotfW0bMTCGXamhHH4kQgx/02m2NGwSBz9xMeI3mMElea9ioisoh29K7cnFWMsnXO0WSfN9CAz/w2TJjDJ3kw/XJhcDYeQHuA79g6YYvUbvhRmpfvJHlG26kdsNNLN/wpe7wF2+i9sW/onHbV1GmDPQTTgAhkEGA88D/w9+7H+exR5EBeHv3gRQ49/1fpOMSPLUXJNgP/htIif/wowihEuzeBQr48/N4LQt//3xYvn8fXm0B5/FdSNNC0XV8s4n0fYJ6E+/XD+E/9TQCCJoN3F8/RLBvAevHP07/g+bQsh10bfCfIg1dU3ueqDChOON4H64lp91qB0HPIftchEAqCkJR2ntNoo+Lf2z7oFB7g7G9VdUe71bJR3YrRf/k6LP65hn77IJb7JGmh3bt47Dt0+jt3WLDtOSu67N3weQ5xxzSKcti0pL3Evc+dEs+9AUAQoCqInS92EvTUKIEj6YXond/dvSlhIjO84yNt1+d6bJe7emiaeOf1TdPpfsaAtvx0FWlk+DDousqmqaMfNHFZmVc74ppTVafRTGt4VaZaS1USVeYOB+OYb0nUZpm732gJ2TTNG1Kpf7WJC2Zk0R1SprKxPlwNE2bcqk/yYt4B1Asx8cf48Ytm4UgkFiO37kIuajgJKWSiu14E+cF8f0Ay/EptZN8FO+KUVJpTFqWgTRMG6Ok9lz0MIpwTVXa996eOC9CmvdhUSpljVrDSsYnJKg1LCqJK+37k1yiKEpizxHtWJdySZk4L0ia9zT6f4tuTKkYJYIgwLLHP/V0oxLdGq6SczheSVyfmhyPEEJQMXTkxPlAsrz37CLs2bfcu/swQlEEVMoai8tmsmxCm8Vlk0pZa+/J7Jc4LEKAUdZYmDjPZaW8KxJJpazRslwcd7L/NonjerQsl0pZ65wH39OSDJCfVbdS1rDtifMsxvUeR6H96Iqpis780qRlSTK/ZFI1xn/UTPJHKekaVUNnYeI8lZXyDqBE8itlDdtxc+/qv9loWS6241JtP50tvgEZeYsnb1ge3nOFvv5il65zHct2Js4TJL0PasHDWO/Gfs80LcsKT1+RUDcdLMfjmUds65lgs7LrqQWMksZMtdQ+vaZfthDd61WFEO0bC4XnsITB8E0mHrESOW+YDq2J8x7SvCcbFVHgnJWITtMkhGCqoqMIwUJtsgpdqJmobSdprQcZrcowCCGoGOGG1cR5yGp4V4haICSKEFQNjaXl5qY+ich2PJaWTSpG+KeXOa1EeqxbP95lSc5DEl63WDV0Fmub2zkp3klxliSM5Z9I2+mTE7UsZZ2pSok9+5d7Km4m9uxfZqpSolLW+iR3kjeWxN060cGgbv3kcHwe0XilrDNd3dzOSfGet1club2TVieip7sSMV0toauCp+frndhm4en5OroqmKl2b7GcJ3BYkvOKfsiZagldUzalc1bZe6e7AnRWy4oQTFdLOI7L/CbqK87XTBzHZbpaQrRb3LTWJCK+4RMvDuv2d1myfrTos6Yr2qZzzpjek12VtPp93ZVo5rqmMl3RaTRa1Oqj3zvwYKFWb9FotJiulDrnLmf1xZPnokSk1Y3TSfpM5xrTFZ36JnHOkN7TnJITj0jtrkTDRruvuFBrsryBTyhablgs1JpMV0sY7ZOB0nzEyWtNugzXmtN2PrMJnLOq3nvp6a6I2KoiilXK4f7K+aXGhmxdavUW80sNZqrhBg9tD4Nak+QPELlLDmeRrJt0Plsts1BrbkjnjOE9SVrdOEIIhOO6Mn6goue9/UeRMrxgoG7aTFcNtm/pfZTIwcr8UpOGaTNTLWGU1K6k2IGfZDKGxXGZvX3H6ELmaDikWz/qmsQ9Zzlv2R5102FmauM4ZwW8h+P9ffasJO/rrqTOXAiMkspstUzTtNi7AfYA7J2v0zQtZhOik8sdf88iXp62C7Hnh0js+iLHeaWsMTdV2jDOGeA9Ii2WRto0aQjHdSXJFjzWwoj4MymlxPN9GqaL58NhO2YOumdS2o7H0/vraCpMV3U0tVd0tLpMiu68p7QmcYTofwx2WLf/2Z5FnbteQNNy8Tx5UDpnFbxH8YjkeDwmXNeTadLjw0npgZQ0Ww7NlsvWuSm2zob3BVzvLC6bLNaaTFXCA17RLdFIiI7G80XTl+Rx0fFEDwnL8hI9z7mUkkbLpWE6bNty8DhnlbzHXceH43Smz0ry5HtSOu2zxZpWeCj6kO0zVMrr8xmVLdtlX3t1P2VoVBJnt6WJTn3PaE0iH3HZyVg4Xqw1j96znJu2h5Tr2zlr4D0iOZ6MCdf1JGNIDwJJ03IxLZeKUWL73BS63n/7gAOB6/rM15q0LIeqoTNlhPfwZgVEx8tkbGMzTjwe+VqJRPeDANPyaNkeRllfV85ZQ+/J4Tg9dTzP79wpblTpst1vtBwP0/KoVstsmTYwDlArY9kuSw0L07SpGhpGSevcXoyxRUOyNRGJfnhWvDs8fqJLKXH9AMv2MC2XatU4oM45QN7TSMaF5/mSuLwRpUfDrhdeoNtyfDRVZW6mwnS1PPIN1IsSBJKGaVOrt/B8n0pJxSjrPZJZBdHxsriLtFh3vLdsJZy3bBfbDVAVZc2cs068D4x5nifjLQsjSE+WEbsZj+162O2b8sxMGVQMfUUuaSJ2HWC9aWE7HuWSSlnXMEpq6uox/j6M6G4sX3TcR1a8Oz6a82QsQkqJH0gs28PxfCzHDy88WGHnjOg9Gl4N73HS4u0kZ2jp0XBeCxPh+RLbcXG8ANcL8H2JUdYwyjolXUXTVHRNRVWiO5ZG84h+vADX8/E8H8f1sWwXy/ZQVYGuKZQ0hXJJR4s9xyYpN3qPS06W9byniI6XJ4cjZE7/vH94NZ372G4QPjB2COfhfA4+7+TFfb/bJy8qPRnLa2HSxl0vwHF9/EDi+QFSSjxfEgRBeGdoKRHtLy1EeIBFU8NdS5qqoCptyYmNreRCJgWmtSJp9cYVnVzeiDR/4zhPK0sbl+3+u+sG+IEMnzwRBAQyfAKFbK95IexI5Xkv6Wrf3WWTLpKe0rwn67AC3rMQvt/bJx8kPT6cjA0SnxWLkyzP+/IRyTpJKYNakZ7YGKKJPi9jGZPxtXCeHE+SV563nKSUJx2ttfcsOklOAenxOmlio/ek+HjZoFgR0hYoGYskx8vShPXECohOG4+Ix7OWLRkf5DzNc3x4rZyTsdzJ2Gp5T35OnLwySCQ5BaT31umXHR9OdmOS5eOSXLhoPCk5PpwWo6DotPGItHjWcibjB5NzUpZ1vXlPIoIg6HuiSlx673jxFiY5niV/XPIEJ8fThvsls2Ki85Y1WXYgnaeVZy1TxGp5T86HlHlHZMWTiKC9xZG2oGGsX9SgFmbQeHwB4/FBJBeqiODkeM9whuje8WR5L1nxiLxlS5Z1x9ePc1KWcbW8J6cnZZ5x8sridJKcjAWPi0+Wjyo+Iis+iKyFS8aLSO6WFfvhkgwqZ8ByppWttfOkDxKfE5G1rMn4anvPK0ujJ8kjkiLi0pPlSRnZ0+bHRiFtYZOxvvGMVqR3PFmezqDyOHnLnFa2Xp2TsdzJWN/4AfIOGUlOipQ86RQQnxVbCdIWOhkbRnJWLM6g8jQGLX9aeZ73lXKeVqfI8qXVScbSW28YpXtCgfI0MpOcjIXPk06KeFLqxMkrSzJoAZPlScEkJPeOd0mLxRlUnseg5U0rP5DOI/KWOa0s6X5c73llg8hN8og0KYPEkyGfjLqjkLXgScEUlExOnAFlwzLIQVr5enBOjof16r3zROZBErLKk+J7Y12y5K8EaXLpkzO6ZAqUj0KapzhZ5evBeUSa+/Xmveex42mi4uSV95aFXyyvPmP8CGli4yRbj95YL1nxOEXqjMogR+TUSXPeH+9lVOcM5Z2xk5uCdYrQk+QRKaEe8sqzxNNXtnL0y1iZ5GaIeuNQ1EtWvQPhnFQ3g70zoCyiSJ2ipCY5BeUMqtNf3v/F++tkk73gvfPIrhcyqJyCdVaaIi4G1ekvT1+O/nr5pPs4OLxnJjlDihhUN7+86IJlz6OInCJ1GKLeapDvqUuRetl1Rlm+9HkNcjWoPM4wdYchN8kjClTpMEzdiGGmGVbEMPWHqbvaDONkmLqMUJ8R3AxTf5i6o1AoySOGqAoj1F8phpU2bP21ZFiHw9ZfSUbxOMo0wzJUkjOmxHGmzWNUUaNOdyAY1d2o0xVhHH/jTDssQyd5xIiTpVJ0XislZqXms9YU9VSUovNbSV8rOa+iiN17FostaQ5FZR1oDoTg1eJgcc4qej/qsC3JUCojt+RZrPDsxma1BK8n1ptz1pn3FU/yOKs464GsJ8lrycR5P6ua5ElW46PWq9j1wsT5Gif5hPVJVgocbMmcxSTJJ2x40p/VN2HCBmKS5BM2PP8f6eEysKru4MUAAAAASUVORK5CYII='

def inject_bi_css():
    st.markdown("""
    <style>
    :root { --navy:#0c2a5b; --ink:#17243a; --muted:#6f7c90; --line:#e6ebf2; --soft:#f6f8fb; }
    .stApp { background:#f7f9fc; }
    [data-testid="stHeader"] { background:rgba(247,249,252,.92); }
    .block-container { padding-top:1.4rem; padding-bottom:2rem; max-width:1600px; }
    [data-testid="stSidebar"] { background:#ffffff; border-right:1px solid #e4e9f1; }
    [data-testid="stSidebar"] .block-container { padding-top:1rem; }
    [data-testid="stSidebar"] hr { margin:.8rem 0 1rem; border-color:#edf0f5; }
    [data-testid="stSidebar"] label { color:#33425a; font-size:.82rem; }
    [data-testid="stSidebar"] div[role="radiogroup"] label { background:#fff; border:1px solid transparent; padding:.48rem .55rem; border-radius:9px; margin:.08rem 0; }
    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) { background:#0c2a5b; color:white; }
    [data-testid="stMetric"] { background:#fff; border:1px solid #e4e9f1; border-radius:14px; padding:14px 16px; box-shadow:0 1px 2px rgba(20,43,77,.04); }
    [data-testid="stMetricLabel"] { color:#6b778b; font-size:.78rem; font-weight:700; text-transform:uppercase; letter-spacing:.02em; }
    [data-testid="stMetricValue"] { color:#142a4d; font-size:1.55rem; font-weight:800; }
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border:1px solid #e4e9f1; border-radius:12px; overflow:hidden; background:white; }
    .bi-header { display:flex; justify-content:space-between; gap:16px; align-items:flex-start; margin:0 0 18px; }
    .bi-title { color:#10284f; font-size:2rem; font-weight:850; margin:0; line-height:1.05; }
    .bi-subtitle { color:#738096; font-size:.95rem; margin-top:7px; }
    .bi-period { background:#fff; border:1px solid #e4e9f1; border-radius:12px; padding:10px 14px; min-width:220px; color:#233a60; }
    .bi-period small { color:#7b8798; display:block; font-weight:700; }
    .bi-period strong { font-size:.95rem; }
    .bi-section { background:#fff; border:1px solid #e4e9f1; border-radius:14px; padding:14px 18px 8px; margin:16px 0 10px; }
    .bi-section-title { font-weight:800; color:#172c51; font-size:1.02rem; margin-bottom:3px; }
    .bi-section-sub { color:#7a8798; font-size:.82rem; margin-bottom:4px; }
    .brand-wrap { text-align:center; padding:4px 0 10px; }
    .brand-wrap img { width:190px; max-width:100%; border-radius:8px; }
    .brand-caption { color:#8a94a5; font-size:.7rem; margin-top:3px; }
    .filter-title { font-size:.72rem; font-weight:800; letter-spacing:.07em; color:#78859a; margin-top:.3rem; text-transform:uppercase; }
    .stButton>button, .stDownloadButton>button, .stFormSubmitButton>button { border-radius:9px; font-weight:700; }
    .stFormSubmitButton>button { background:#0c2a5b; color:white; border-color:#0c2a5b; }
    h1,h2,h3 { color:#14294b; }
    </style>
    """, unsafe_allow_html=True)

def periodo_resumo():
    anos_txt = ", ".join(str(x) for x in sel_anos) if sel_anos else "Todos os anos"
    if not sel_meses_num:
        meses_txt = "Todos os meses"
    elif len(sel_meses_num) == 12:
        meses_txt = "Jan–Dez"
    else:
        abbr = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
        meses_txt = ", ".join(abbr[m-1] for m in sel_meses_num)
    return f"{meses_txt} / {anos_txt}"

def bi_header(title, subtitle):
    st.markdown(f"<div class='bi-header'><div><div class='bi-title'>{title}</div><div class='bi-subtitle'>{subtitle}</div></div><div class='bi-period'><small>PERÍODO SELECIONADO</small><strong>{periodo_resumo()}</strong></div></div>", unsafe_allow_html=True)

def section_header(title, subtitle=""):
    st.markdown(f"<div class='bi-section'><div class='bi-section-title'>{title}</div><div class='bi-section-sub'>{subtitle}</div></div>", unsafe_allow_html=True)

inject_bi_css()

# -----------------------------
# Sidebar: Página + filtros
# -----------------------------
st.sidebar.markdown(f"<div class='brand-wrap'><img src='data:image/png;base64,{BRAND_IMAGE_B64}'><div class='brand-caption'>PAINEL DE GESTÃO</div></div>", unsafe_allow_html=True)
st.sidebar.markdown("<div class='filter-title'>Navegação</div>", unsafe_allow_html=True)
page = st.sidebar.radio(
    "Página", ["COMPRAS", "SELLOUT", "HISTÓRICO POR FORNECEDOR", "REVISÃO DE FORNECEDORES", "ORÇAMENTO"],
    format_func=lambda x: {"COMPRAS":"🛒  Compras", "SELLOUT":"📈  Sellout", "HISTÓRICO POR FORNECEDOR":"◷  Histórico por Fornecedor", "REVISÃO DE FORNECEDORES":"✓  Revisão de Fornecedores", "ORÇAMENTO":"◎  Orçamento"}[x],
    label_visibility="collapsed",
)

st.sidebar.divider()
st.sidebar.markdown("<div class='filter-title'>Filtros do período</div>", unsafe_allow_html=True)

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
    mes_cols = st.columns(3)
    meses_marcados = {}
    for idx, mes in enumerate(MESES_LABELS):
        meses_marcados[mes] = mes_cols[idx % 3].checkbox(
            ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"][idx],
            value=(mes in meses_atualmente_aplicados),
            key=f"filtro_mes_check_{MESES_PT[mes]}",
        )
    sel_meses_input = [mes for mes in MESES_LABELS if meses_marcados.get(mes, False)]
    aplicar_filtros = st.form_submit_button("Aplicar filtros", use_container_width=True)

st.sidebar.caption(f"Versão {APP_VERSION}")

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
    bi_header("Indicadores de Compras", "Análise consolidada de Compras, CMV e Entradas")

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
    # Ex.: Jan a Mar de um ano = 3 períodos; Jan a Mar de dois anos = 6 períodos.
    periodos_selecionados = max(len(sel_anos) * len(sel_meses_num), 1)
    media_compras = total_compras_citel / periodos_selecionados
    media_cmv = total_vendas_cmv / periodos_selecionados
    media_diferenca = dif_topo / periodos_selecionados

    cmv_sobre_compras = (total_vendas_cmv / total_compras_citel) if total_compras_citel else 0.0
    section_header("Visão executiva do período", "Principais indicadores para acompanhamento do nível de compras")
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1: st.metric("Total Compras (CITEL)", brl(total_compras_citel))
    with m2: st.metric("Total CMV", brl(total_vendas_cmv))
    with m3: st.metric("Diferença CMV - Compras", brl(dif_topo), delta=pct_str(dif_pct))
    with m4: st.metric("CMV x Compras", pct_str(cmv_sobre_compras))
    with m5: st.metric("Períodos Considerados", periodos_selecionados)

    st.caption("Médias mensais calculadas de acordo com o período selecionado")
    a1, a2, a3 = st.columns(3)
    with a1:
        st.metric("Média Mensal de Compras", brl(media_compras))
    with a2:
        st.metric("Média Mensal de CMV", brl(media_cmv))
    with a3:
        st.metric("Média Mensal da Diferença", brl(media_diferenca))

    # Tabela por fornecedor (CITEL x CMV)
    section_header("Compras x CMV por Fornecedor", "Participação, compras realizadas, CMV e diferença por fornecedor")

    nome_canon = (
        df_cmv_f.groupby(["FORN_KEY", "FORNECEDOR_CMV"], as_index=False)
        .size()
        .sort_values(["FORN_KEY", "size"], ascending=[True, False])
        .drop_duplicates("FORN_KEY")[["FORN_KEY", "FORNECEDOR_CMV"]]
        .rename(columns={"FORNECEDOR_CMV": "FORNECEDOR"})
    )

    vendas = df_cmv_f.groupby("FORN_KEY", as_index=False).agg(**{"CMV": ("CMV_VALOR", "sum")})
    compras = df_citel_f.groupby("FORN_KEY", as_index=False).agg(**{"COMPRAS FORNECEDOR": ("COMPRA_VALOR", "sum")})

    tab = nome_canon.merge(vendas, on="FORN_KEY", how="left").merge(compras, on="FORN_KEY", how="left")
    tab["CMV"] = tab["CMV"].fillna(0.0)
    tab["COMPRAS FORNECEDOR"] = tab["COMPRAS FORNECEDOR"].fillna(0.0)
    tab["DIF (CMV - COMPRAS)"] = tab["CMV"] - tab["COMPRAS FORNECEDOR"]
    total_cmv_tab = float(tab["CMV"].sum())
    total_comp_tab = float(tab["COMPRAS FORNECEDOR"].sum())
    tab["% CMV"] = tab["CMV"] / total_cmv_tab if total_cmv_tab else 0.0
    tab["% COMPRAS"] = tab["COMPRAS FORNECEDOR"] / total_comp_tab if total_comp_tab else 0.0
    tab = tab[~((tab["CMV"] == 0) & (tab["COMPRAS FORNECEDOR"] == 0))].copy()
    tab = tab[["FORNECEDOR", "COMPRAS FORNECEDOR", "CMV", "DIF (CMV - COMPRAS)", "% CMV", "% COMPRAS"]].sort_values("COMPRAS FORNECEDOR", ascending=False)

    st.dataframe(
    tab.style
      .format({"COMPRAS FORNECEDOR": brl, "CMV": brl, "DIF (CMV - COMPRAS)": brl, "% CMV": lambda x: pct_str(float(x)), "% COMPRAS": lambda x: pct_str(float(x))})
      .map(style_dif, subset=["DIF (CMV - COMPRAS)"]),
    use_container_width=True,
    hide_index=True
    )

    st.divider()

    # Conciliação CITEL x ENTRADAS
    section_header("Conciliação de Compras: CITEL x ENTRADAS", "Validação do total fiscal contra as entradas analíticas")

    total_compras_entradas = float(df_ent_f["VR_CONTABIL"].sum())
    dif_citel_vs_ent = total_compras_citel - total_compras_entradas
    color2 = "#0a7a2f" if dif_citel_vs_ent >= 0 else "#b00020"

    dif_ent_pct = (dif_citel_vs_ent / total_compras_citel) if total_compras_citel else 0.0
    d1, d2, d3 = st.columns(3)
    with d1: st.metric("Total Compras (CITEL)", brl(total_compras_citel))
    with d2: st.metric("Total Compras (ENTRADAS)", brl(total_compras_entradas))
    with d3: st.metric("Diferença CITEL - ENTRADAS", brl(dif_citel_vs_ent), delta=pct_str(dif_ent_pct))

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
    bi_header("Indicadores de Sellout", "Desempenho de faturamento, participação e giro por fornecedor")

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
    bi_header("Orçamento de Compras", "Distribuição orientada pelo CMV e participação dos fornecedores")
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

    section_header("Distribuição do orçamento por fornecedor", "Edite somente o orçamento final; os demais campos são calculados automaticamente")
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
# PAGE: REVISÃO DE FORNECEDORES
# -----------------------------
@st.fragment
def render_supplier_review_page():
    # A revisão roda isoladamente. Assim, salvar uma decisão não recarrega
    # GIRO, NOTAS, SELLOUT e demais bases pesadas do dashboard.
    df_supplier_review = build_supplier_review_candidates(
        df_cmv, df_ent, df_sellout, SUPPLIER_MEMORY_PATH
    )
    bi_header("Revisão de Fornecedores", "Memória por produto: confirme apenas quando surgir uma divergência nova")

    st.info(
        "Quando um produto aparece com fornecedores diferentes, você decide qual deve prevalecer. "
        "Com o Google configurado, a decisão fica salva permanentemente na aba MEMORIA_FORNECEDORES "
        "e a mesma divergência não será perguntada novamente."
    )

    if df_supplier_review.empty:
        st.success("Nenhuma divergência nova de fornecedor para revisar.")
    else:
        st.warning(f"Há {len(df_supplier_review)} produto(s) com fornecedor novo ou divergente.")
        for idx, row in df_supplier_review.reset_index(drop=True).iterrows():
            cod = str(row["COD_KEY"])
            desc = str(row.get("DESCRICAO", "") or "")
            candidatos = list(row.get("CANDIDATOS", []))
            sugestao = str(row.get("SUGESTAO", "") or "")
            novos = list(row.get("NOVOS", []))
            titulo = f"{cod} — {desc if desc else 'Produto sem descrição'}"
            with st.expander(titulo, expanded=(idx == 0)):
                c1, c2 = st.columns([1, 1])
                with c1:
                    st.markdown(f"**Fornecedor validado/sugerido:** {sugestao}")
                    st.markdown(f"**Novo(s) fornecedor(es) detectado(s):** {', '.join(novos)}")
                with c2:
                    st.caption("Evidências encontradas nas bases")
                    st.write(row.get("FONTES", ""))

                options = candidatos.copy()
                if sugestao and sugestao not in options:
                    options.insert(0, sugestao)
                default_idx = options.index(sugestao) if sugestao in options else 0
                escolhido = st.selectbox(
                    "Qual fornecedor deve ficar vinculado a este produto?",
                    options=options, index=default_idx, key=f"supplier_choice_{cod}_{idx}"
                )
                st.caption(
                    "Se você mantiver o fornecedor atual, o novo fornecedor ficará registrado como já analisado "
                    "e não voltará a gerar pergunta. Se surgir uma terceira opção no futuro, a revisão reaparece."
                )
                if st.button("Salvar decisão", key=f"supplier_save_{cod}_{idx}", type="primary"):
                    save_supplier_decision(cod, escolhido, candidatos, SUPPLIER_MEMORY_PATH)
                    st.toast(f"Decisão salva: {cod} → {escolhido}", icon="✅")
                    # Reexecuta apenas esta tela de revisão, não o dashboard inteiro.
                    # A decisão some da fila imediatamente e o próximo item aparece.
                    st.rerun(scope="fragment")

    st.divider()
    st.subheader("Associações já validadas")
    st.caption(
        "Aqui você pode corrigir uma decisão anterior sem editar a planilha manualmente. "
        "Alterar substitui a associação salva; remover faz o produto voltar para análise."
    )

    mem = load_supplier_memory(SUPPLIER_MEMORY_PATH)
    if mem.empty:
        st.caption("Ainda não há decisões gravadas.")
    else:
        desc_map = build_product_description_map(df_cmv, df_ent, df_sellout)

        c_busca, c_total = st.columns([3, 1])
        with c_busca:
            busca_mem = st.text_input(
                "Buscar associação",
                placeholder="Digite código, produto ou fornecedor...",
                key="supplier_memory_search"
            ).strip().upper()
        with c_total:
            st.metric("Associações salvas", len(mem))

        mem_view = mem.copy()
        mem_view["DESCRICAO"] = mem_view["COD_KEY"].map(desc_map).fillna("")
        if busca_mem:
            mask = (
                mem_view["COD_KEY"].astype(str).str.upper().str.contains(busca_mem, regex=False)
                | mem_view["DESCRICAO"].astype(str).str.upper().str.contains(busca_mem, regex=False)
                | mem_view["FORNECEDOR_VALIDADO"].astype(str).str.upper().str.contains(busca_mem, regex=False)
                | mem_view["FORNECEDORES_JA_VISTOS"].astype(str).str.upper().str.contains(busca_mem, regex=False)
            )
            mem_view = mem_view[mask].copy()

        if mem_view.empty:
            st.info("Nenhuma associação encontrada para essa busca.")
        else:
            for idx, row in mem_view.reset_index(drop=True).iterrows():
                cod = str(row["COD_KEY"])
                desc = str(row.get("DESCRICAO", "") or "").strip()
                atual = str(row.get("FORNECEDOR_VALIDADO", "") or "").strip()
                vistos = [
                    x.strip() for x in str(row.get("FORNECEDORES_JA_VISTOS", "") or "").split("||")
                    if x.strip()
                ]
                if atual and atual not in vistos:
                    vistos.insert(0, atual)

                updated = str(row.get("ATUALIZADO_EM", "") or "").strip()
                titulo = f"{cod} — {desc if desc else 'Produto sem descrição'}"

                with st.expander(titulo, expanded=False):
                    st.markdown(f"**Fornecedor atual:** :blue[{atual}]")
                    if vistos:
                        st.caption("Fornecedores já encontrados: " + " • ".join(vistos))
                    if updated:
                        st.caption(f"Última alteração: {updated}")

                    novo_fornecedor = st.selectbox(
                        "Fornecedor correto para este produto",
                        options=vistos,
                        index=vistos.index(atual) if atual in vistos else 0,
                        key=f"memory_edit_choice_{cod}_{idx}"
                    )

                    col_save, col_delete = st.columns([1, 1])
                    with col_save:
                        if st.button(
                            "Salvar alteração",
                            key=f"memory_edit_save_{cod}_{idx}",
                            type="primary",
                            use_container_width=True
                        ):
                            save_supplier_decision(
                                cod, novo_fornecedor, vistos, SUPPLIER_MEMORY_PATH
                            )
                            st.toast(
                                f"Associação atualizada: {cod} → {novo_fornecedor}",
                                icon="✅"
                            )
                            st.rerun(scope="fragment")

                    with col_delete:
                        confirmar = st.checkbox(
                            "Confirmar remoção",
                            key=f"memory_delete_confirm_{cod}_{idx}"
                        )
                        if st.button(
                            "Remover associação",
                            key=f"memory_delete_{cod}_{idx}",
                            disabled=not confirmar,
                            use_container_width=True
                        ):
                            remove_supplier_decision(cod, SUPPLIER_MEMORY_PATH)
                            st.toast(
                                f"Associação removida para o produto {cod}.",
                                icon="🗑️"
                            )
                            st.rerun(scope="fragment")

# -----------------------------
# Render
# -----------------------------
if page == "COMPRAS":
    render_compras_page()
elif page == "SELLOUT":
    render_sellout_page()
elif page == "HISTÓRICO POR FORNECEDOR":
    render_historico_fornecedor_page()
elif page == "REVISÃO DE FORNECEDORES":
    render_supplier_review_page()
else:
    render_orcamento_page()
