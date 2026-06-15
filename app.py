import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st


APP_TITLE = "Gestión de Publicaciones Pendientes - Aurora"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

MAESTRO_FILE = DATA_DIR / "maestro_sku_ean.xlsx"
PUBLICACIONES_FILE = DATA_DIR / "publicaciones_mercado_libre.xlsx"
PACKS_FILE = DATA_DIR / "packs.xlsx"

# URL de Apps Script ya configurada.
# En Streamlit Cloud puedes sobrescribirla desde Secrets si cambia el despliegue.
DEFAULT_APP_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw8k8UkeHtHdAcAFUKvBtHfELH7byRdM0hXao5-OjqeCbI1KL3JxaQfFebgq7_4fzoy/exec"
DEFAULT_APP_SCRIPT_TOKEN = "aurora_publicaciones_2026"

ESTADOS = [
    "SIN STOCK",
    "LLEGÓ STOCK",
    "PENDIENTE VERIFICAR FÍSICO",
    "STOCK FÍSICO CONFIRMADO",
    "EN PROCESO DE PUBLICACIÓN",
    "PUBLICADO",
    "NO ENCONTRADO FÍSICO",
    "NO PUBLICABLE",
    "FALTA FOTO",
    "FALTA INFORMACIÓN",
    "CUBIERTO POR PACK",
    "CUBIERTO POR UNIDAD",
    "PRODUCTO NUEVO CON STOCK",
    "PRODUCTO NUEVO SIN STOCK",
    "PRODUCTO NUEVO PUBLICADO",
]

ESTADOS_MANUALES_PROTEGIDOS = {
    "STOCK FÍSICO CONFIRMADO",
    "EN PROCESO DE PUBLICACIÓN",
    "PUBLICADO",
    "NO ENCONTRADO FÍSICO",
    "NO PUBLICABLE",
    "FALTA FOTO",
    "FALTA INFORMACIÓN",
    "PRODUCTO NUEVO PUBLICADO",
}

MOTIVOS = [
    "",
    "No encontrado físicamente",
    "Producto dañado",
    "Producto incompleto",
    "No apto para Mercado Libre",
    "Sin foto",
    "Sin información técnica",
    "SKU duplicado",
    "EAN incorrecto",
    "Pack mal definido",
    "Solo venta presencial",
    "No rentable",
    "Descontinuado",
    "Producto nuevo detectado en inventario",
    "Otro",
]


# ============================================================
# Utilidades
# ============================================================

def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def norm_col(value: str) -> str:
    value = "" if value is None else str(value).strip().lower()
    replacements = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "ñ": "n", ".": " ", "_": " ", "-": " ", "/": " "
    }
    for a, b in replacements.items():
        value = value.replace(a, b)
    return " ".join(value.split())


def clean_sku(value) -> str:
    if pd.isna(value):
        return ""
    value = str(value).strip()
    if value.endswith(".0"):
        value = value[:-2]
    return value.strip()


def to_number(value) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value).strip()
    if not text:
        return 0.0

    text = text.replace("$", "").replace(" ", "")

    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return 0.0


def read_excel_any(path_or_file) -> pd.DataFrame:
    return pd.read_excel(path_or_file, dtype=str)


def find_column(df: pd.DataFrame, candidates: List[str], required: bool = True) -> Optional[str]:
    normalized = {norm_col(c): c for c in df.columns}
    candidate_norms = [norm_col(c) for c in candidates]

    for cand in candidate_norms:
        if cand in normalized:
            return normalized[cand]

    for cand in candidate_norms:
        for norm_name, original in normalized.items():
            if cand and cand in norm_name:
                return original

    if required:
        raise ValueError(
            f"No se encontró columna requerida. Busqué: {candidates}. "
            f"Columnas disponibles: {list(df.columns)}"
        )
    return None


def safe_file_exists(path: Path) -> bool:
    return path.exists() and path.is_file()


def export_excel(df: pd.DataFrame) -> bytes:
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="cola_marketing")
    return output.getvalue()


# ============================================================
# API Apps Script
# ============================================================

def get_api_config() -> Tuple[str, str]:
    """
    Primero intenta leer Streamlit Secrets.
    Si no existen, usa la URL y token configurados por defecto.
    """
    try:
        url = st.secrets.get("APP_SCRIPT_URL", DEFAULT_APP_SCRIPT_URL)
        token = st.secrets.get("APP_SCRIPT_TOKEN", DEFAULT_APP_SCRIPT_TOKEN)
    except Exception:
        url = DEFAULT_APP_SCRIPT_URL
        token = DEFAULT_APP_SCRIPT_TOKEN

    return url, token


def api_call(action: str, payload: Optional[dict] = None, timeout: int = 60) -> dict:
    url, token = get_api_config()
    if not url or not token:
        raise RuntimeError("Faltan APP_SCRIPT_URL o APP_SCRIPT_TOKEN en los Secrets de Streamlit.")

    body = {
        "token": token,
        "action": action,
        "payload": payload or {},
    }

    response = requests.post(url, json=body, timeout=timeout)
    response.raise_for_status()
    data = response.json()

    if not data.get("ok"):
        raise RuntimeError(data.get("error", "Error desconocido desde Apps Script"))

    return data


def api_get_data() -> Tuple[pd.DataFrame, pd.DataFrame]:
    data = api_call("get_data", {}, timeout=60)
    estado_rows = data.get("estado_actual", [])
    inventario_rows = data.get("inventario_actual", [])

    estado_df = pd.DataFrame(estado_rows)
    inv_df = pd.DataFrame(inventario_rows)

    if estado_df.empty:
        estado_df = pd.DataFrame(columns=[
            "sku", "estado", "origen", "descripcion", "familia", "ean",
            "stock_anterior", "stock_actual", "responsable", "motivo",
            "observacion", "link_publicacion", "accion", "estado_anterior",
            "fecha_actualizacion"
        ])

    if inv_df.empty:
        inv_df = pd.DataFrame(columns=[
            "sku", "descripcion", "familia", "stock_actual",
            "costo_promedio", "saldo_valor", "fecha_carga"
        ])

    return estado_df, inv_df


def api_replace_inventory(inv_df: pd.DataFrame, usuario: str, productos_nuevos: int, llegaron_stock: int) -> None:
    rows = []
    for _, r in inv_df.iterrows():
        rows.append({
            "sku": clean_sku(r.get("SKU", "")),
            "descripcion": str(r.get("Articulo", "") or ""),
            "familia": str(r.get("Familia", "") or ""),
            "stock_actual": float(r.get("Stock", 0) or 0),
            "costo_promedio": float(r.get("CostoPromedio", 0) or 0),
            "saldo_valor": float(r.get("SaldoValor", 0) or 0),
        })

    payload = {
        "usuario": usuario,
        "fecha_carga": now_iso(),
        "productos_nuevos": int(productos_nuevos),
        "llegaron_stock": int(llegaron_stock),
        "items": rows,
    }

    api_call("replace_inventory", payload, timeout=180)


def api_upsert_product(payload: dict) -> None:
    api_call("upsert_product", payload, timeout=60)


def api_bulk_upsert_products(items: List[dict]) -> None:
    if not items:
        return
    api_call("bulk_upsert_products", {"items": items}, timeout=120)


# ============================================================
# Lectura archivos base desde GitHub/repo
# ============================================================

@st.cache_data(show_spinner=False)
def load_maestro() -> pd.DataFrame:
    if not safe_file_exists(MAESTRO_FILE):
        raise FileNotFoundError(f"No existe {MAESTRO_FILE}")

    df = read_excel_any(MAESTRO_FILE)

    sku_col = find_column(df, ["SKU"])
    desc_col = find_column(df, ["Descripción", "Descripcion", "Articulo", "Artículo", "Nombre"], required=False)
    ean_col = find_column(df, ["codigo de barras", "código de barras", "ean", "codigo universal"], required=False)
    fam_col = find_column(df, ["Familia", "Categoría", "Categoria"], required=False)

    out = pd.DataFrame()
    out["SKU"] = df[sku_col].map(clean_sku)
    out["DescripcionMaestro"] = df[desc_col].fillna("").astype(str) if desc_col else ""
    out["EAN"] = df[ean_col].fillna("").astype(str) if ean_col else ""
    out["FamiliaMaestro"] = df[fam_col].fillna("").astype(str) if fam_col else ""
    out = out[out["SKU"] != ""].drop_duplicates(subset=["SKU"], keep="first")
    return out


@st.cache_data(show_spinner=False)
def load_publicaciones() -> pd.DataFrame:
    if not safe_file_exists(PUBLICACIONES_FILE):
        raise FileNotFoundError(f"No existe {PUBLICACIONES_FILE}")

    df = read_excel_any(PUBLICACIONES_FILE)

    sku_col = find_column(df, ["SKU"])
    title_col = find_column(df, ["Titulo", "Título", "Publicación", "Publicacion", "Nombre", "Descripción", "Descripcion"], required=False)
    link_col = find_column(df, ["Link", "Permalink", "URL", "Enlace"], required=False)
    estado_col = find_column(df, ["Estado", "Status"], required=False)

    out = pd.DataFrame()
    out["SKU"] = df[sku_col].map(clean_sku)
    out["TituloML"] = df[title_col].fillna("").astype(str) if title_col else ""
    out["LinkML"] = df[link_col].fillna("").astype(str) if link_col else ""
    out["EstadoML"] = df[estado_col].fillna("").astype(str) if estado_col else ""
    out = out[out["SKU"] != ""].drop_duplicates(subset=["SKU"], keep="first")
    return out


@st.cache_data(show_spinner=False)
def load_packs() -> pd.DataFrame:
    if not safe_file_exists(PACKS_FILE):
        return pd.DataFrame(columns=["SKU_PACK", "SKU_UNIDAD", "CANTIDAD"])

    df = read_excel_any(PACKS_FILE)
    columns_norm = {norm_col(c): c for c in df.columns}

    pack_col = None
    unit_col = None
    qty_col = None

    for norm_name, original in columns_norm.items():
        if "sku" in norm_name and ("pack" in norm_name or "combo" in norm_name):
            pack_col = original
        if "sku" in norm_name and ("unidad" in norm_name or "componente" in norm_name or "producto" in norm_name):
            unit_col = original
        if "cantidad" in norm_name or norm_name in {"cant", "qty", "unidades"}:
            qty_col = original

    sku_like = [original for norm_name, original in columns_norm.items() if "sku" in norm_name]
    if pack_col is None and len(sku_like) >= 1:
        pack_col = sku_like[0]
    if unit_col is None and len(sku_like) >= 2:
        unit_col = sku_like[1]

    if not pack_col or not unit_col:
        return pd.DataFrame(columns=["SKU_PACK", "SKU_UNIDAD", "CANTIDAD"])

    out = pd.DataFrame()
    out["SKU_PACK"] = df[pack_col].map(clean_sku)
    out["SKU_UNIDAD"] = df[unit_col].map(clean_sku)
    out["CANTIDAD"] = df[qty_col].map(to_number) if qty_col else 1
    out["CANTIDAD"] = out["CANTIDAD"].replace(0, 1)
    out = out[(out["SKU_PACK"] != "") & (out["SKU_UNIDAD"] != "")]
    return out.drop_duplicates()


def load_inventory_from_upload(uploaded_file) -> pd.DataFrame:
    df = read_excel_any(uploaded_file)

    sku_col = find_column(df, ["SKU"])
    stock_col = find_column(df, ["Q. Saldo Consolidado", "Saldo Consolidado", "Stock", "Cantidad", "Existencia"])
    art_col = find_column(df, ["Artículo", "Articulo", "Descripción", "Descripcion", "Nombre"], required=False)
    fam_col = find_column(df, ["Familia", "Categoría", "Categoria"], required=False)
    costo_col = find_column(df, ["Costo promedio", "Costo Promedio", "Costo"], required=False)
    saldo_valor_col = find_column(df, ["$ Saldo", "Saldo valor", "Valor saldo"], required=False)

    out = pd.DataFrame()
    out["SKU"] = df[sku_col].map(clean_sku)
    out["Articulo"] = df[art_col].fillna("").astype(str) if art_col else ""
    out["Familia"] = df[fam_col].fillna("").astype(str) if fam_col else ""
    out["Stock"] = df[stock_col].map(to_number)
    out["CostoPromedio"] = df[costo_col].map(to_number) if costo_col else 0
    out["SaldoValor"] = df[saldo_valor_col].map(to_number) if saldo_valor_col else 0

    out = out[out["SKU"] != ""]
    out = out.groupby("SKU", as_index=False).agg({
        "Articulo": "first",
        "Familia": "first",
        "Stock": "sum",
        "CostoPromedio": "max",
        "SaldoValor": "sum",
    })

    return out


def normalize_inventory_from_api(inv_api_df: pd.DataFrame) -> pd.DataFrame:
    if inv_api_df.empty:
        return pd.DataFrame(columns=["SKU", "Articulo", "Familia", "Stock", "CostoPromedio", "SaldoValor"])

    df = inv_api_df.copy()

    out = pd.DataFrame()
    out["SKU"] = df.get("sku", "").map(clean_sku)
    out["Articulo"] = df.get("descripcion", "").fillna("").astype(str)
    out["Familia"] = df.get("familia", "").fillna("").astype(str)
    out["Stock"] = df.get("stock_actual", 0).map(to_number)
    out["CostoPromedio"] = df.get("costo_promedio", 0).map(to_number)
    out["SaldoValor"] = df.get("saldo_valor", 0).map(to_number)

    out = out[out["SKU"] != ""]
    return out


# ============================================================
# Motor de gestión
# ============================================================

def relation_maps(packs: pd.DataFrame, publicaciones: pd.DataFrame):
    published_file = set(publicaciones["SKU"].astype(str))
    pack_to_units = {}
    unit_to_packs = {}

    for _, row in packs.iterrows():
        pack = clean_sku(row["SKU_PACK"])
        unit = clean_sku(row["SKU_UNIDAD"])

        if not pack or not unit:
            continue

        pack_to_units.setdefault(pack, set()).add(unit)
        unit_to_packs.setdefault(unit, set()).add(pack)

    return published_file, pack_to_units, unit_to_packs


def get_state_map(estado_api_df: pd.DataFrame) -> Dict[str, dict]:
    result = {}
    if estado_api_df.empty:
        return result

    for _, row in estado_api_df.iterrows():
        sku = clean_sku(row.get("sku", ""))
        if not sku:
            continue
        result[sku] = row.to_dict()

    return result


def get_manually_published_skus(estado_api_df: pd.DataFrame) -> Dict[str, str]:
    if estado_api_df.empty:
        return {}

    df = estado_api_df.copy()
    if "estado" not in df.columns or "sku" not in df.columns:
        return {}

    mask = df["estado"].isin(["PUBLICADO", "PRODUCTO NUEVO PUBLICADO"])
    result = {}

    for _, row in df[mask].iterrows():
        sku = clean_sku(row.get("sku", ""))
        if sku:
            result[sku] = str(row.get("link_publicacion", "") or "")

    return result


def previous_stock_map_from_inventory(inv_df: pd.DataFrame) -> Dict[str, float]:
    if inv_df.empty:
        return {}
    return {clean_sku(r["SKU"]): float(r.get("Stock", 0) or 0) for _, r in inv_df.iterrows()}


def build_universe(maestro: pd.DataFrame, inv_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    Regla operacional:
    Si un SKU aparece en LibroInventario y no existe en el maestro de la repo,
    se agrega como PRODUCTO NUEVO.
    """
    if inv_df is None or inv_df.empty:
        inv_base = pd.DataFrame(columns=["SKU", "Articulo", "Familia", "Stock", "CostoPromedio", "SaldoValor"])
    else:
        inv_base = inv_df.copy()

    inv_base["SKU"] = inv_base["SKU"].map(clean_sku)

    universe = pd.merge(
        maestro,
        inv_base,
        on="SKU",
        how="outer",
        suffixes=("_maestro", "_inv")
    )

    def source(row):
        in_maestro = (
            pd.notna(row.get("DescripcionMaestro")) or
            pd.notna(row.get("FamiliaMaestro")) or
            pd.notna(row.get("EAN"))
        )
        in_inv = pd.notna(row.get("Articulo")) or pd.notna(row.get("Stock"))

        if in_maestro and in_inv:
            return "MAESTRO + INVENTARIO"
        if in_maestro:
            return "MAESTRO"
        return "PRODUCTO NUEVO"

    universe["Origen"] = universe.apply(source, axis=1)

    universe["Descripcion"] = universe.get("DescripcionMaestro", "").fillna("")
    mask_desc_empty = universe["Descripcion"].astype(str).str.strip() == ""
    universe.loc[mask_desc_empty, "Descripcion"] = universe.loc[mask_desc_empty, "Articulo"].fillna("")

    universe["Familia"] = universe.get("FamiliaMaestro", "").fillna("")
    mask_fam_empty = universe["Familia"].astype(str).str.strip() == ""
    if "Familia_inv" in universe.columns:
        universe.loc[mask_fam_empty, "Familia"] = universe.loc[mask_fam_empty, "Familia_inv"].fillna("")
    elif "Familia" in inv_base.columns:
        pass

    universe["EAN"] = universe.get("EAN", "").fillna("")
    universe["Stock"] = universe.get("Stock", 0).fillna(0).map(to_number)
    universe["SKU"] = universe["SKU"].map(clean_sku)

    return universe[universe["SKU"] != ""].drop_duplicates(subset=["SKU"], keep="first")


def build_work_queue(
    maestro: pd.DataFrame,
    publicaciones: pd.DataFrame,
    packs: pd.DataFrame,
    inv_df: Optional[pd.DataFrame],
    previous_stock_map: Dict[str, float],
    estado_api_df: pd.DataFrame,
) -> pd.DataFrame:
    published_file, pack_to_units, unit_to_packs = relation_maps(packs, publicaciones)
    manual_published = get_manually_published_skus(estado_api_df)
    published_all = set(published_file) | set(manual_published.keys())
    state_map = get_state_map(estado_api_df)

    universe = build_universe(maestro, inv_df)
    rows = []

    for _, row in universe.iterrows():
        sku = clean_sku(row["SKU"])
        stock_actual = float(row.get("Stock", 0) or 0)
        stock_anterior = float(previous_stock_map.get(sku, 0))
        origen = row.get("Origen", "")
        is_product_new = origen == "PRODUCTO NUEVO"
        is_published = sku in published_all
        existing_state = state_map.get(sku)

        # Los SKUs publicados desde el archivo de publicaciones no entran a cola
        # salvo que tengan un estado manual guardado.
        if is_published and not existing_state:
            if is_product_new:
                estado_sugerido = "PRODUCTO NUEVO PUBLICADO"
            else:
                continue
        else:
            estado_sugerido = "SIN STOCK"

        published_pack = sorted(list(unit_to_packs.get(sku, set()) & published_all))
        published_units = sorted(list(pack_to_units.get(sku, set()) & published_all))

        relacion = ""
        sku_relacionado = ""

        if is_published and is_product_new:
            estado_sugerido = "PRODUCTO NUEVO PUBLICADO"
        elif is_product_new:
            estado_sugerido = "PRODUCTO NUEVO CON STOCK" if stock_actual > 0 else "PRODUCTO NUEVO SIN STOCK"
        elif published_pack:
            estado_sugerido = "CUBIERTO POR PACK"
            relacion = "Unidad/componente ya cubierto por pack publicado"
            sku_relacionado = ", ".join(published_pack[:5])
        elif published_units:
            estado_sugerido = "CUBIERTO POR UNIDAD"
            relacion = "Pack ya cubierto por unidad/componente publicado"
            sku_relacionado = ", ".join(published_units[:5])
        elif stock_actual > 0 and stock_anterior <= 0:
            estado_sugerido = "LLEGÓ STOCK"
        elif stock_actual > 0:
            estado_sugerido = "PENDIENTE VERIFICAR FÍSICO"
        else:
            estado_sugerido = "SIN STOCK"

        estado_final = estado_sugerido
        motivo = ""
        observacion = ""
        responsable = ""
        link_publicacion = manual_published.get(sku, "")

        if existing_state:
            old_estado = str(existing_state.get("estado", "") or "")

            if old_estado in ESTADOS_MANUALES_PROTEGIDOS:
                estado_final = old_estado
            elif old_estado == "LLEGÓ STOCK" and stock_actual > 0:
                estado_final = old_estado
            elif old_estado == "PRODUCTO NUEVO CON STOCK" and is_product_new and stock_actual > 0:
                estado_final = old_estado
            elif old_estado in {"CUBIERTO POR PACK", "CUBIERTO POR UNIDAD"} and estado_sugerido.startswith("CUBIERTO"):
                estado_final = old_estado
            else:
                estado_final = estado_sugerido

            motivo = str(existing_state.get("motivo", "") or "")
            observacion = str(existing_state.get("observacion", "") or "")
            responsable = str(existing_state.get("responsable", "") or "")
            link_publicacion = str(existing_state.get("link_publicacion", "") or link_publicacion)

        rows.append({
            "SKU": sku,
            "Descripcion": row.get("Descripcion", ""),
            "Familia": row.get("Familia", ""),
            "EAN": row.get("EAN", ""),
            "Origen": origen,
            "StockAnterior": stock_anterior,
            "StockSistema": stock_actual,
            "EstadoSugerido": estado_sugerido,
            "Estado": estado_final,
            "RelacionPackUnidad": relacion,
            "SKURelacionadoPublicado": sku_relacionado,
            "Motivo": motivo,
            "Responsable": responsable,
            "Observacion": observacion,
            "LinkPublicacion": link_publicacion,
        })

    return pd.DataFrame(rows)


def build_auto_alerts(queue_df: pd.DataFrame, estado_api_df: pd.DataFrame) -> List[dict]:
    """
    Guarda en Google Sheets solo alertas útiles.
    No guarda todos los SIN STOCK para no llenar la hoja.
    """
    state_map = get_state_map(estado_api_df)
    items = []

    if queue_df.empty:
        return items

    alert_states = {"LLEGÓ STOCK", "PRODUCTO NUEVO CON STOCK"}

    for _, row in queue_df.iterrows():
        sku = row["SKU"]
        estado = row["Estado"]

        if estado not in alert_states:
            continue

        existing = state_map.get(sku)
        if existing and str(existing.get("estado", "")) in ESTADOS_MANUALES_PROTEGIDOS:
            continue

        items.append({
            "fecha": now_iso(),
            "sku": sku,
            "estado": estado,
            "origen": row.get("Origen", ""),
            "descripcion": row.get("Descripcion", ""),
            "familia": row.get("Familia", ""),
            "ean": row.get("EAN", ""),
            "stock_anterior": float(row.get("StockAnterior", 0) or 0),
            "stock_actual": float(row.get("StockSistema", 0) or 0),
            "responsable": "SISTEMA",
            "motivo": "Alerta automática",
            "observacion": "",
            "link_publicacion": "",
            "accion": "ESTADO AUTOMÁTICO",
        })

    return items


# ============================================================
# UI
# ============================================================

def page_config():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)


def validate_base_files():
    missing = []
    for p in [MAESTRO_FILE, PUBLICACIONES_FILE, PACKS_FILE]:
        if not safe_file_exists(p):
            missing.append(str(p.relative_to(BASE_DIR)))

    if missing:
        st.error("Faltan archivos base en la carpeta /data:")
        for item in missing:
            st.code(item)
        st.info("Coloca los archivos con los nombres exactos y vuelve a ejecutar la app.")
        st.stop()


def dashboard(queue_df: pd.DataFrame):
    total = len(queue_df)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("En gestión", total)
    c2.metric("Llegó stock", int((queue_df["Estado"] == "LLEGÓ STOCK").sum()) if total else 0)
    c3.metric("Pend. verificar", int((queue_df["Estado"] == "PENDIENTE VERIFICAR FÍSICO").sum()) if total else 0)
    c4.metric("Productos nuevos", int((queue_df["Origen"] == "PRODUCTO NUEVO").sum()) if total else 0)
    c5.metric("Nuevo c/stock", int((queue_df["Estado"] == "PRODUCTO NUEVO CON STOCK").sum()) if total else 0)

    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("Confirmados", int((queue_df["Estado"] == "STOCK FÍSICO CONFIRMADO").sum()) if total else 0)
    c7.metric("En proceso", int((queue_df["Estado"] == "EN PROCESO DE PUBLICACIÓN").sum()) if total else 0)
    c8.metric("No encontrados", int((queue_df["Estado"] == "NO ENCONTRADO FÍSICO").sum()) if total else 0)
    c9.metric("No publicables", int((queue_df["Estado"] == "NO PUBLICABLE").sum()) if total else 0)
    c10.metric("Publicados", int((queue_df["Estado"].isin(["PUBLICADO", "PRODUCTO NUEVO PUBLICADO"])).sum()) if total else 0)


def inventory_upload_ui(maestro, publicaciones, packs, estado_df, inv_current_df):
    st.sidebar.header("Actualizar inventario")

    usuario = st.sidebar.text_input("Responsable carga inventario", value="")
    uploaded_inventory = st.sidebar.file_uploader("Subir LibroInventario.xlsx", type=["xlsx"])

    if uploaded_inventory is None:
        return None, False

    if st.sidebar.button("Procesar y guardar inventario"):
        if not usuario.strip():
            st.sidebar.error("Debes indicar responsable antes de guardar.")
            return None, False

        try:
            inv_new_df = load_inventory_from_upload(uploaded_inventory)
            prev_map = previous_stock_map_from_inventory(inv_current_df)

            queue_tmp = build_work_queue(
                maestro=maestro,
                publicaciones=publicaciones,
                packs=packs,
                inv_df=inv_new_df,
                previous_stock_map=prev_map,
                estado_api_df=estado_df,
            )

            productos_nuevos = int((queue_tmp["Origen"] == "PRODUCTO NUEVO").sum()) if not queue_tmp.empty else 0
            llegaron_stock = int((queue_tmp["Estado"] == "LLEGÓ STOCK").sum()) if not queue_tmp.empty else 0

            api_replace_inventory(
                inv_df=inv_new_df,
                usuario=usuario.strip(),
                productos_nuevos=productos_nuevos,
                llegaron_stock=llegaron_stock,
            )

            auto_alerts = build_auto_alerts(queue_tmp, estado_df)
            api_bulk_upsert_products(auto_alerts)

            st.sidebar.success(
                f"Inventario guardado: {len(inv_new_df):,} SKUs | "
                f"Nuevos: {productos_nuevos:,} | Llegó stock: {llegaron_stock:,}"
            )

            st.cache_data.clear()
            st.rerun()

        except Exception as e:
            st.sidebar.error(f"Error procesando inventario: {e}")
            return None, False

    return None, False


def marketing_queue_ui(queue_df: pd.DataFrame):
    st.subheader("Cola de Marketing")

    if queue_df.empty:
        st.info("No hay productos para mostrar.")
        return

    col_a, col_b, col_c, col_d = st.columns([2, 2, 2, 2])
    estados_disponibles = ["TODOS"] + sorted(queue_df["Estado"].dropna().unique().tolist())
    origenes_disponibles = ["TODOS"] + sorted(queue_df["Origen"].dropna().unique().tolist())

    estado_filter = col_a.selectbox("Filtrar por estado", estados_disponibles)
    origen_filter = col_b.selectbox("Filtrar por origen", origenes_disponibles)
    familia_filter = col_c.selectbox("Filtrar por familia", ["TODAS"] + sorted(queue_df["Familia"].fillna("").unique().tolist()))
    search = col_d.text_input("Buscar SKU o descripción")

    df = queue_df.copy()

    if estado_filter != "TODOS":
        df = df[df["Estado"] == estado_filter]
    if origen_filter != "TODOS":
        df = df[df["Origen"] == origen_filter]
    if familia_filter != "TODAS":
        df = df[df["Familia"] == familia_filter]
    if search.strip():
        s = search.strip().lower()
        df = df[
            df["SKU"].str.lower().str.contains(s, na=False) |
            df["Descripcion"].str.lower().str.contains(s, na=False)
        ]

    priority = {
        "LLEGÓ STOCK": 1,
        "PRODUCTO NUEVO CON STOCK": 2,
        "STOCK FÍSICO CONFIRMADO": 3,
        "PENDIENTE VERIFICAR FÍSICO": 4,
        "FALTA FOTO": 5,
        "FALTA INFORMACIÓN": 6,
        "EN PROCESO DE PUBLICACIÓN": 7,
        "CUBIERTO POR PACK": 8,
        "CUBIERTO POR UNIDAD": 9,
        "SIN STOCK": 10,
        "PRODUCTO NUEVO SIN STOCK": 11,
        "NO ENCONTRADO FÍSICO": 12,
        "NO PUBLICABLE": 13,
        "PUBLICADO": 14,
        "PRODUCTO NUEVO PUBLICADO": 15,
    }

    df["_orden"] = df["Estado"].map(priority).fillna(99)
    df = df.sort_values(["_orden", "StockSistema"], ascending=[True, False]).drop(columns=["_orden"])

    display_cols = [
        "SKU", "Descripcion", "Familia", "EAN", "Origen", "StockAnterior", "StockSistema",
        "Estado", "EstadoSugerido", "RelacionPackUnidad", "SKURelacionadoPublicado",
        "Motivo", "Responsable", "Observacion", "LinkPublicacion"
    ]

    st.dataframe(df[display_cols], use_container_width=True, hide_index=True)

    st.download_button(
        "Descargar cola filtrada en Excel",
        data=export_excel(df),
        file_name=f"cola_marketing_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("Actualizar estado de un producto")

    sku_options = df["SKU"].tolist()

    if not sku_options:
        st.warning("No hay productos en el filtro actual.")
        return

    with st.form("form_update"):
        selected_sku = st.selectbox("SKU", sku_options)
        selected_row = df[df["SKU"] == selected_sku].iloc[0]

        c1, c2, c3 = st.columns(3)

        estado_actual = selected_row["Estado"]
        estado_index = ESTADOS.index(estado_actual) if estado_actual in ESTADOS else 0

        nuevo_estado = c1.selectbox("Nuevo estado", ESTADOS, index=estado_index)
        responsable = c2.text_input("Responsable", value=str(selected_row.get("Responsable", "") or ""))
        motivo = c3.selectbox(
            "Motivo",
            MOTIVOS,
            index=MOTIVOS.index(selected_row.get("Motivo", "")) if selected_row.get("Motivo", "") in MOTIVOS else 0
        )

        observacion = st.text_area("Observación", value=str(selected_row.get("Observacion", "") or ""))
        link_publicacion = st.text_input("Link publicación ML", value=str(selected_row.get("LinkPublicacion", "") or ""))

        submitted = st.form_submit_button("Guardar cambio")

    if submitted:
        if not responsable.strip():
            st.error("Debes indicar responsable para guardar el cambio.")
            return

        payload = {
            "fecha": now_iso(),
            "sku": selected_sku,
            "estado": nuevo_estado,
            "origen": selected_row.get("Origen", ""),
            "descripcion": selected_row.get("Descripcion", ""),
            "familia": selected_row.get("Familia", ""),
            "ean": selected_row.get("EAN", ""),
            "stock_anterior": float(selected_row.get("StockAnterior", 0) or 0),
            "stock_actual": float(selected_row.get("StockSistema", 0) or 0),
            "responsable": responsable.strip(),
            "motivo": motivo,
            "observacion": observacion,
            "link_publicacion": link_publicacion,
            "accion": "CAMBIO MANUAL",
        }

        try:
            api_upsert_product(payload)
            st.success(f"SKU {selected_sku} actualizado a: {nuevo_estado}")
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo guardar el cambio: {e}")


def historial_ui():
    st.subheader("Historial de cambios")

    try:
        data = api_call("get_history", {"limit": 500}, timeout=60)
        rows = data.get("historial_cambios", [])
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"No se pudo cargar historial: {e}")


def status_ui(estado_df: pd.DataFrame, inv_df: pd.DataFrame):
    st.subheader("Estado de conexión y base central")

    url, token = get_api_config()

    c1, c2, c3 = st.columns(3)
    c1.metric("URL Apps Script", "Configurada" if url else "Falta")
    c2.metric("Token", "Configurado" if token else "Falta")
    c3.metric("Inventario central", f"{len(inv_df):,} SKUs")

    st.write("Últimos estados guardados en Google Sheets:")
    st.dataframe(estado_df.tail(100), use_container_width=True, hide_index=True)


def main():
    page_config()
    validate_base_files()

    try:
        maestro = load_maestro()
        publicaciones = load_publicaciones()
        packs = load_packs()
    except Exception as e:
        st.error(f"Error leyendo archivos base: {e}")
        st.stop()

    try:
        estado_df, inv_api_df = api_get_data()
    except Exception as e:
        st.error(f"No se pudo conectar con Apps Script / Google Sheets: {e}")
        st.info("Revisa los Secrets de Streamlit y que el Web App de Apps Script esté desplegado.")
        st.stop()

    inv_current_df = normalize_inventory_from_api(inv_api_df)

    inventory_upload_ui(
        maestro=maestro,
        publicaciones=publicaciones,
        packs=packs,
        estado_df=estado_df,
        inv_current_df=inv_current_df,
    )

    previous_map = previous_stock_map_from_inventory(inv_current_df)

    queue_df = build_work_queue(
        maestro=maestro,
        publicaciones=publicaciones,
        packs=packs,
        inv_df=inv_current_df,
        previous_stock_map=previous_map,
        estado_api_df=estado_df,
    )

    tabs = st.tabs(["Resumen", "Cola Marketing", "Historial", "Base central"])

    with tabs[0]:
        dashboard(queue_df)
        st.divider()
        st.subheader("Prioridad de trabajo")

        if not queue_df.empty:
            preview = queue_df[queue_df["Estado"].isin([
                "LLEGÓ STOCK",
                "PRODUCTO NUEVO CON STOCK",
                "PENDIENTE VERIFICAR FÍSICO",
                "STOCK FÍSICO CONFIRMADO",
                "FALTA FOTO",
                "FALTA INFORMACIÓN",
            ])].copy()
            st.dataframe(preview.head(100), use_container_width=True, hide_index=True)
        else:
            st.info("Carga el inventario desde la barra lateral para comenzar.")

    with tabs[1]:
        marketing_queue_ui(queue_df)

    with tabs[2]:
        historial_ui()

    with tabs[3]:
        status_ui(estado_df, inv_current_df)


if __name__ == "__main__":
    main()
