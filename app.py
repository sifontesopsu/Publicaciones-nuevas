import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests
import streamlit as st


APP_TITLE = "Gestión de Publicaciones Pendientes - Aurora"
APP_VERSION = "V6.9.1 - fix import sync"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

MAESTRO_FILE = DATA_DIR / "maestro_sku_ean.xlsx"
PUBLICACIONES_FILE = DATA_DIR / "publicaciones_mercado_libre.xlsx"
PACKS_FILE = DATA_DIR / "packs.xlsx"

# URL de Apps Script ya configurada.
# En Streamlit Cloud puedes sobrescribirla desde Secrets si cambia el despliegue.
DEFAULT_APP_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw8k8UkeHtHdAcAFUKvBtHfELH7byRdM0hXao5-OjqeCbI1KL3JxaQfFebgq7_4fzoy/exec"
DEFAULT_APP_SCRIPT_TOKEN = "aurora_publicaciones_2026"
DEFAULT_ADMIN_PASSWORD = "aurora_admin_2026"
APP_USER_OPERACION = "OPERACION"
APP_USER_ADMIN = "ADMINISTRADOR"
APP_USER_SISTEMA = "SISTEMA"
ESTADO_API_COLUMNS = [
    "sku", "estado", "origen", "descripcion", "familia", "ean",
    "stock_anterior", "stock_actual", "responsable", "motivo",
    "observacion", "link_publicacion", "accion", "estado_anterior",
    "fecha_actualizacion"
]

INVENTARIO_API_COLUMNS = [
    "sku", "descripcion", "familia", "stock_actual",
    "costo_promedio", "saldo_valor", "fecha_carga"
]

SYNC_EXECUTOR = ThreadPoolExecutor(max_workers=3)
SYNC_LOCK = threading.Lock()


ESTADOS = [
    "SIN STOCK",
    "LLEGÓ STOCK",
    "PENDIENTE PUBLICAR",
    "PICKEADO PARA PUBLICAR",
    "EN PROCESO DE PUBLICACIÓN",
    "PUBLICADO",
    "FALTANTE FÍSICO CON STOCK KAME",
    "NO PUBLICABLE",
    "CUBIERTO POR PACK",
    "CUBIERTO POR UNIDAD",
    "PRODUCTO NUEVO CON STOCK",
    "PRODUCTO NUEVO SIN STOCK",
    "PRODUCTO NUEVO PUBLICADO",
]

LEGACY_ESTADOS_MAP = {
    "PENDIENTE VERIFICAR FÍSICO": "PENDIENTE PUBLICAR",
    "STOCK FÍSICO CONFIRMADO": "PENDIENTE PUBLICAR",
    "NO ENCONTRADO FÍSICO": "NO PUBLICABLE",
    "FALTA FOTO": "NO PUBLICABLE",
    "FALTA INFORMACIÓN": "NO PUBLICABLE",
}

ESTADOS_MANUALES_PROTEGIDOS = {
    "PENDIENTE PUBLICAR",
    "PICKEADO PARA PUBLICAR",
    "EN PROCESO DE PUBLICACIÓN",
    "PUBLICADO",
    "FALTANTE FÍSICO CON STOCK KAME",
    "NO PUBLICABLE",
    "PRODUCTO NUEVO PUBLICADO",
}

MOTIVOS_NO_PUBLICABLE = [
    "Producto duplicado / ya publicado",
    "Producto cubierto por pack",
    "Producto descontinuado",
    "Margen insuficiente / no rentable",
    "Producto a pedido / stock inestable",
    "Producto frágil o riesgoso para despacho",
    "Peso o medidas problemáticas",
    "Producto incompleto",
    "Producto dañado",
    "Producto no identificado",
    "Producto restringido por Mercado Libre",
    "Marca/modelo fuera de norma",
    "Otro",
]

MOTIVOS_NO_PUBLICABLE_OPERATIVO = MOTIVOS_NO_PUBLICABLE


MOTIVOS_GENERALES = [
    "",
    "Inicio de publicación",
    "Pickeado para publicar",
    "Publicado correctamente",
    "Faltante físico con stock en Kame",
    "Corrección de estado",
    "Alerta automática",
    "Producto nuevo detectado en inventario",
    "Otro",
]

MOTIVOS = MOTIVOS_GENERALES + [m for m in MOTIVOS_NO_PUBLICABLE if m and m not in MOTIVOS_GENERALES]


def normalize_estado_operativo(estado: str) -> str:
    estado = str(estado or "").strip()
    return LEGACY_ESTADOS_MAP.get(estado, estado)


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
    if hasattr(path_or_file, "seek"):
        path_or_file.seek(0)
    return pd.read_excel(path_or_file, dtype=str)


def make_unique_columns(cols: List[str]) -> List[str]:
    seen = {}
    result = []

    for col in cols:
        name = str(col).strip()
        if not name or name.lower() == "nan":
            name = "Unnamed"

        if name not in seen:
            seen[name] = 0
            result.append(name)
        else:
            seen[name] += 1
            result.append(f"{name}_{seen[name]}")

    return result


def row_has_candidate(row_values: List[str], candidates: List[str]) -> bool:
    normalized_candidates = [norm_col(c) for c in candidates]

    for cell in row_values:
        cell_norm = norm_col(cell)

        for cand in normalized_candidates:
            if not cand:
                continue

            if cell_norm == cand:
                return True

            if cand in cell_norm:
                return True

    return False


def read_excel_detect_header(path_or_file, required_groups: List[List[str]], max_scan_rows: int = 80) -> pd.DataFrame:
    """
    Lee un Excel donde los encabezados no necesariamente están en la primera fila.

    Caso esperado:
    Fila 1: LIBRO MAYOR AUXILIAR DE INVENTARIO
    Fila real de encabezados: Artículo | SKU | Familia | Q. Saldo Consolidado | $ Saldo | Costo promedio
    """
    if hasattr(path_or_file, "seek"):
        path_or_file.seek(0)

    raw = pd.read_excel(path_or_file, dtype=str, header=None)

    header_row_idx = None
    rows_to_scan = min(max_scan_rows, len(raw))

    for i in range(rows_to_scan):
        row_values = ["" if pd.isna(v) else str(v) for v in raw.iloc[i].tolist()]
        matches_all_required_groups = True

        for group in required_groups:
            if not row_has_candidate(row_values, group):
                matches_all_required_groups = False
                break

        if matches_all_required_groups:
            header_row_idx = i
            break

    if header_row_idx is None:
        preview = []
        for i in range(min(10, len(raw))):
            preview.append([str(v) for v in raw.iloc[i].fillna("").tolist()])

        raise ValueError(
            "No pude detectar la fila de encabezados del LibroInventario. "
            "Busqué una fila que contenga SKU y una columna de stock/saldo. "
            f"Primeras filas detectadas: {preview}"
        )

    headers = make_unique_columns(raw.iloc[header_row_idx].fillna("").astype(str).tolist())
    df = raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)

    return df


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


def api_get_data(force_refresh: bool = False) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Modo rápido:
    - Primera carga: lee desde Apps Script / Google Sheets.
    - Luego usa memoria de sesión.
    - Cada acción actualiza la memoria local.
    - Solo vuelve a consultar Google Sheets si el usuario presiona "Actualizar datos".
    """
    if (
        not force_refresh
        and "estado_df_cache" in st.session_state
        and "inv_df_cache" in st.session_state
    ):
        return (
            st.session_state["estado_df_cache"].copy(),
            st.session_state["inv_df_cache"].copy()
        )

    data = api_call("get_data", {}, timeout=90)
    estado_rows = data.get("estado_actual", [])
    inventario_rows = data.get("inventario_actual", [])

    estado_df = pd.DataFrame(estado_rows)
    inv_df = pd.DataFrame(inventario_rows)

    if estado_df.empty:
        estado_df = pd.DataFrame(columns=ESTADO_API_COLUMNS)
    else:
        for col in ESTADO_API_COLUMNS:
            if col not in estado_df.columns:
                estado_df[col] = ""
        estado_df = estado_df[ESTADO_API_COLUMNS]

    if inv_df.empty:
        inv_df = pd.DataFrame(columns=INVENTARIO_API_COLUMNS)
    else:
        for col in INVENTARIO_API_COLUMNS:
            if col not in inv_df.columns:
                inv_df[col] = ""
        inv_df = inv_df[INVENTARIO_API_COLUMNS]

    st.session_state["estado_df_cache"] = estado_df.copy()
    st.session_state["inv_df_cache"] = inv_df.copy()
    st.session_state["data_cache_loaded_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return estado_df, inv_df


def clear_session_data_cache():
    st.session_state.pop("estado_df_cache", None)
    st.session_state.pop("inv_df_cache", None)
    st.session_state.pop("data_cache_loaded_at", None)


def payload_to_estado_row(payload: dict, estado_anterior: str = "") -> dict:
    return {
        "sku": clean_sku(payload.get("sku", "")),
        "estado": str(payload.get("estado", "") or ""),
        "origen": str(payload.get("origen", "") or ""),
        "descripcion": str(payload.get("descripcion", "") or ""),
        "familia": str(payload.get("familia", "") or ""),
        "ean": str(payload.get("ean", "") or ""),
        "stock_anterior": payload.get("stock_anterior", 0) or 0,
        "stock_actual": payload.get("stock_actual", 0) or 0,
        "responsable": str(payload.get("responsable", "") or ""),
        "motivo": str(payload.get("motivo", "") or ""),
        "observacion": str(payload.get("observacion", "") or ""),
        "link_publicacion": str(payload.get("link_publicacion", "") or ""),
        "accion": str(payload.get("accion", "") or ""),
        "estado_anterior": estado_anterior or str(payload.get("estado_anterior", "") or ""),
        "fecha_actualizacion": str(payload.get("fecha", "") or now_iso()),
    }


def update_estado_cache_from_payload(payload: dict):
    sku = clean_sku(payload.get("sku", ""))

    if not sku:
        return

    if "estado_df_cache" not in st.session_state:
        return

    df = st.session_state["estado_df_cache"].copy()

    for col in ESTADO_API_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    if df.empty:
        estado_anterior = ""
        new_row = payload_to_estado_row(payload, estado_anterior)
        df = pd.DataFrame([new_row], columns=ESTADO_API_COLUMNS)
    else:
        df["sku"] = df["sku"].map(clean_sku)
        mask = df["sku"] == sku

        if mask.any():
            estado_anterior = str(df.loc[mask, "estado"].iloc[0] or "")
            new_row = payload_to_estado_row(payload, estado_anterior)

            for col in ESTADO_API_COLUMNS:
                df.loc[mask, col] = new_row.get(col, "")
        else:
            new_row = payload_to_estado_row(payload, "")
            df = pd.concat([df, pd.DataFrame([new_row], columns=ESTADO_API_COLUMNS)], ignore_index=True)

    st.session_state["estado_df_cache"] = df[ESTADO_API_COLUMNS].copy()


def update_estado_cache_from_payloads(items: List[dict]):
    for payload in items:
        update_estado_cache_from_payload(payload)


def update_inventory_cache_from_rows(rows: List[dict], fecha_carga: str):
    out_rows = []

    for item in rows:
        out_rows.append({
            "sku": clean_sku(item.get("sku", "")),
            "descripcion": str(item.get("descripcion", "") or ""),
            "familia": str(item.get("familia", "") or ""),
            "stock_actual": item.get("stock_actual", 0) or 0,
            "costo_promedio": item.get("costo_promedio", 0) or 0,
            "saldo_valor": item.get("saldo_valor", 0) or 0,
            "fecha_carga": fecha_carga,
        })

    inv_df = pd.DataFrame(out_rows, columns=INVENTARIO_API_COLUMNS)
    st.session_state["inv_df_cache"] = inv_df.copy()



def replace_inventory_worker(payload: dict):
    api_call("replace_inventory", payload, timeout=240)


def api_replace_inventory(
    inv_df: pd.DataFrame,
    usuario: str,
    productos_nuevos: int,
    llegaron_stock: int,
    background: bool = True,
) -> None:
    rows = []
    fecha_carga = now_iso()

    for _, row in inv_df.iterrows():
        rows.append({
            "sku": row["SKU"],
            "descripcion": row.get("Articulo", ""),
            "familia": row.get("Familia", ""),
            "stock_actual": float(row.get("Stock", 0) or 0),
            "costo_promedio": float(row.get("CostoPromedio", 0) or 0),
            "saldo_valor": float(row.get("SaldoValor", 0) or 0),
        })

    payload = {
        "usuario": usuario,
        "fecha_carga": fecha_carga,
        "productos_nuevos": int(productos_nuevos),
        "llegaron_stock": int(llegaron_stock),
        "items": rows,
    }

    update_inventory_cache_from_rows(rows, fecha_carga)

    if not background:
        api_call("replace_inventory", payload, timeout=240)
        return

    init_sync_state()
    sync_status_increment_pending(1)

    future = SYNC_EXECUTOR.submit(replace_inventory_worker, payload)

    def _done_callback(fut):
        try:
            fut.result()
            sync_status_mark_ok(1)
        except Exception as e:
            sync_status_mark_error(e, 1)

    future.add_done_callback(_done_callback)


def init_sync_state():
    if "sync_pending_count" not in st.session_state:
        st.session_state["sync_pending_count"] = 0
    if "sync_ok_count" not in st.session_state:
        st.session_state["sync_ok_count"] = 0
    if "sync_error_count" not in st.session_state:
        st.session_state["sync_error_count"] = 0
    if "sync_errors" not in st.session_state:
        st.session_state["sync_errors"] = []
    if "sync_last_ok" not in st.session_state:
        st.session_state["sync_last_ok"] = ""


def sync_status_increment_pending(amount: int = 1):
    init_sync_state()
    st.session_state["sync_pending_count"] += amount


def sync_status_mark_ok(amount: int = 1):
    with SYNC_LOCK:
        # Los callbacks pueden ejecutarse fuera del flujo normal de Streamlit.
        # Por eso esta función solo se usa cuando el estado de sesión está disponible.
        try:
            st.session_state["sync_pending_count"] = max(0, st.session_state.get("sync_pending_count", 0) - amount)
            st.session_state["sync_ok_count"] = st.session_state.get("sync_ok_count", 0) + amount
            st.session_state["sync_last_ok"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass


def sync_status_mark_error(error_message: str, amount: int = 1):
    with SYNC_LOCK:
        try:
            st.session_state["sync_pending_count"] = max(0, st.session_state.get("sync_pending_count", 0) - amount)
            st.session_state["sync_error_count"] = st.session_state.get("sync_error_count", 0) + amount
            errors = st.session_state.get("sync_errors", [])
            errors.append({
                "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "error": str(error_message),
            })
            st.session_state["sync_errors"] = errors[-20:]
        except Exception:
            pass


def sync_upsert_worker(payload: dict):
    api_call("upsert_product", payload, timeout=90)


def sync_bulk_worker(items: List[dict], chunk_size: int = 250):
    chunks = chunk_list(items, chunk_size)

    for chunk in chunks:
        api_call(
            "bulk_upsert_products",
            {"items": chunk},
            timeout=180
        )


def api_upsert_product(payload: dict, background: bool = True) -> None:
    """
    Guarda un cambio individual.
    En modo background:
    - actualiza la memoria local al instante,
    - envía a Google Sheets en segundo plano.
    """
    update_estado_cache_from_payload(payload)

    if not background:
        api_call("upsert_product", payload, timeout=90)
        return

    init_sync_state()
    sync_status_increment_pending(1)

    future = SYNC_EXECUTOR.submit(sync_upsert_worker, payload)

    def _done_callback(fut):
        try:
            fut.result()
            sync_status_mark_ok(1)
        except Exception as e:
            sync_status_mark_error(e, 1)

    future.add_done_callback(_done_callback)


def chunk_list(items: List[dict], chunk_size: int) -> List[List[dict]]:
    return [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]


def api_bulk_upsert_products(items: List[dict], chunk_size: int = 250, background: bool = True) -> None:
    """
    Guarda cambios masivos.
    En modo background:
    - actualiza cache local al instante,
    - manda los bloques a Google Sheets en segundo plano.
    """
    if not items:
        return

    update_estado_cache_from_payloads(items)

    chunks = chunk_list(items, chunk_size)
    total_chunks = len(chunks)

    if not background:
        progress = st.sidebar.progress(0, text=f"Guardando cambios 0/{len(items)}")
        saved = 0

        for chunk in chunks:
            api_call("bulk_upsert_products", {"items": chunk}, timeout=180)
            saved += len(chunk)
            progress.progress(
                min(saved / len(items), 1.0),
                text=f"Guardando cambios {saved}/{len(items)}"
            )

        progress.empty()
        return

    init_sync_state()
    sync_status_increment_pending(total_chunks)

    def _send_chunk(chunk: List[dict]):
        api_call("bulk_upsert_products", {"items": chunk}, timeout=180)
        return len(chunk)

    for chunk in chunks:
        future = SYNC_EXECUTOR.submit(_send_chunk, chunk)

        def _done_callback(fut):
            try:
                fut.result()
                sync_status_mark_ok(1)
            except Exception as e:
                sync_status_mark_error(e, 1)

        future.add_done_callback(_done_callback)




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
    """
    Lector robusto para LibroInventario.
    Fuerza lectura con header=None y detecta la fila real de encabezados.
    Esto evita que Streamlit/Pandas tome como encabezado:
    'LIBRO MAYOR AUXILIAR DE INVENTARIO'.
    """
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        source = BytesIO(file_bytes)
    else:
        source = uploaded_file

    raw = pd.read_excel(source, dtype=str, header=None)

    header_row_idx = None
    rows_to_scan = min(100, len(raw))

    for i in range(rows_to_scan):
        row_values = ["" if pd.isna(v) else str(v) for v in raw.iloc[i].tolist()]

        has_sku = row_has_candidate(row_values, ["SKU"])
        has_stock = row_has_candidate(
            row_values,
            ["Q. Saldo Consolidado", "Saldo Consolidado", "Stock", "Cantidad", "Existencia"]
        )

        if has_sku and has_stock:
            header_row_idx = i
            break

    if header_row_idx is None:
        preview = []
        for i in range(min(12, len(raw))):
            preview.append([str(v) for v in raw.iloc[i].fillna("").tolist()])

        raise ValueError(
            "V5.4: No pude detectar la fila real de encabezados del LibroInventario. "
            "Busqué una fila que contenga SKU y Q. Saldo Consolidado/Stock. "
            f"Primeras filas detectadas: {preview}"
        )

    headers = make_unique_columns(raw.iloc[header_row_idx].fillna("").astype(str).tolist())
    df = raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)

    # Guardar diagnóstico para mostrarlo en pantalla después de procesar.
    st.session_state["ultimo_inventario_header_excel"] = header_row_idx + 1
    st.session_state["ultimo_inventario_columnas"] = headers

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

    st.session_state["ultimo_inventario_filas"] = len(out)

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
            estado_sugerido = "PENDIENTE PUBLICAR"
        else:
            estado_sugerido = "SIN STOCK"

        estado_final = estado_sugerido
        motivo = ""
        observacion = ""
        responsable = ""
        link_publicacion = manual_published.get(sku, "")

        if existing_state:
            old_estado = normalize_estado_operativo(str(existing_state.get("estado", "") or ""))

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
            original_estado = str(existing_state.get("estado", "") or "")
            if not motivo and original_estado in LEGACY_ESTADOS_MAP:
                motivo = original_estado
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


def inject_operational_css():
    st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e8ebf0;
        padding: 10px 12px;
        border-radius: 12px;
    }
    .aurora-card {
        border: 1px solid #dfe3ea;
        border-radius: 14px;
        padding: 14px 16px;
        margin-bottom: 10px;
        background: #ffffff;
        box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
    }
    .aurora-sku {
        font-size: 0.90rem;
        font-weight: 700;
        color: #111827;
    }
    .aurora-desc {
        font-size: 1.02rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 4px;
    }
    .aurora-meta {
        font-size: 0.83rem;
        color: #4b5563;
    }
    .estado-pill {
        display: inline-block;
        padding: 4px 9px;
        border-radius: 999px;
        background: #f3f4f6;
        color: #111827;
        font-size: 0.80rem;
        font-weight: 700;
    }
    .stock-pill {
        display: inline-block;
        padding: 4px 9px;
        border-radius: 999px;
        background: #ecfdf3;
        color: #027a48;
        font-size: 0.80rem;
        font-weight: 700;
    }
    .section-soft {
        background: #f8fafc;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 12px 14px;
        margin-bottom: 12px;
    }
    </style>
    """, unsafe_allow_html=True)


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
    c3.metric("Pendiente publicar", int((queue_df["Estado"] == "PENDIENTE PUBLICAR").sum()) if total else 0)
    c4.metric("Productos nuevos c/stock", int((queue_df["Estado"] == "PRODUCTO NUEVO CON STOCK").sum()) if total else 0)
    c5.metric("Pickeados", int((queue_df["Estado"].isin(["PICKEADO PARA PUBLICAR", "EN PROCESO DE PUBLICACIÓN"])).sum()) if total else 0)

    c6, c7, c8, c9, c10 = st.columns(5)
    c6.metric("Faltante físico", int((queue_df["Estado"] == "FALTANTE FÍSICO CON STOCK KAME").sum()) if total else 0)
    c7.metric("No publicables", int((queue_df["Estado"] == "NO PUBLICABLE").sum()) if total else 0)
    c8.metric("Publicados", int((queue_df["Estado"].isin(["PUBLICADO", "PRODUCTO NUEVO PUBLICADO"])).sum()) if total else 0)
    c9.metric("Sin stock Kame", int((queue_df["Estado"].isin(["SIN STOCK", "PRODUCTO NUEVO SIN STOCK"])).sum()) if total else 0)
    c10.metric("Cubierto pack/unidad", int((queue_df["Estado"].isin(["CUBIERTO POR PACK", "CUBIERTO POR UNIDAD"])).sum()) if total else 0)

def inventory_upload_ui(maestro, publicaciones, packs, estado_df, inv_current_df):
    st.sidebar.header("Actualizar inventario")

    usuario = APP_USER_SISTEMA
    uploaded_inventory = st.sidebar.file_uploader("Subir LibroInventario.xlsx", type=["xlsx"])

    if uploaded_inventory is None:
        return None, False

    if st.sidebar.button("Procesar y guardar inventario"):
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

            st.sidebar.info("Enviando inventario a Google Sheets en segundo plano...")
            api_replace_inventory(
                inv_df=inv_new_df,
                usuario=usuario.strip(),
                productos_nuevos=productos_nuevos,
                llegaron_stock=llegaron_stock,
            )

            auto_alerts = build_auto_alerts(queue_tmp, estado_df)
            if auto_alerts:
                st.sidebar.info(f"Guardando {len(auto_alerts):,} alertas automáticas por bloques...")
            api_bulk_upsert_products(auto_alerts)

            st.sidebar.success(
                f"Inventario guardado: {len(inv_new_df):,} SKUs | "
                f"Nuevos: {productos_nuevos:,} | Llegó stock: {llegaron_stock:,}"
            )

            if "ultimo_inventario_header_excel" in st.session_state:
                st.sidebar.info(
                    f"Encabezados detectados en fila Excel: {st.session_state['ultimo_inventario_header_excel']}"
                )
                st.sidebar.caption(
                    "Columnas detectadas: " + ", ".join(st.session_state.get("ultimo_inventario_columnas", []))
                )

            st.cache_data.clear()
            st.rerun()

        except Exception as e:
            st.sidebar.error(f"Error procesando inventario: {e}")
            return None, False

    return None, False


def priority_sort(df: pd.DataFrame) -> pd.DataFrame:
    priority = {
        "LLEGÓ STOCK": 1,
        "PRODUCTO NUEVO CON STOCK": 2,
        "PENDIENTE PUBLICAR": 3,
        "PICKEADO PARA PUBLICAR": 4,
        "EN PROCESO DE PUBLICACIÓN": 5,
        "FALTANTE FÍSICO CON STOCK KAME": 6,
        "CUBIERTO POR PACK": 7,
        "CUBIERTO POR UNIDAD": 8,
        "SIN STOCK": 9,
        "PRODUCTO NUEVO SIN STOCK": 10,
        "NO PUBLICABLE": 11,
        "PUBLICADO": 12,
        "PRODUCTO NUEVO PUBLICADO": 13,
    }

    out = df.copy()
    out["_orden"] = out["Estado"].map(priority).fillna(99)
    out = out.sort_values(["_orden", "StockSistema"], ascending=[True, False]).drop(columns=["_orden"])
    return out


def save_status_change(row: pd.Series, nuevo_estado: str, responsable: str = APP_USER_OPERACION, motivo: str = "", observacion: str = "", link_publicacion: str = ""):
    payload = {
        "fecha": now_iso(),
        "sku": row.get("SKU", ""),
        "estado": nuevo_estado,
        "origen": row.get("Origen", ""),
        "descripcion": row.get("Descripcion", ""),
        "familia": row.get("Familia", ""),
        "ean": row.get("EAN", ""),
        "stock_anterior": float(row.get("StockAnterior", 0) or 0),
        "stock_actual": float(row.get("StockSistema", 0) or 0),
        "responsable": APP_USER_OPERACION,
        "motivo": motivo,
        "observacion": observacion,
        "link_publicacion": link_publicacion,
        "accion": "CAMBIO OPERATIVO",
    }
    api_upsert_product(payload)


def operar_productos_ui(queue_df: pd.DataFrame):
    inject_operational_css()

    st.subheader("Picking de productos para publicar")
    st.caption("Pantalla rápida con acciones según etapa: revisar/pickear, reportar faltantes o cerrar como publicado.")

    if queue_df.empty:
        st.info("No hay productos para operar. Primero carga el LibroInventario desde la barra lateral.")
        return

    pendiente_count = int((queue_df["Estado"].isin(["LLEGÓ STOCK", "PRODUCTO NUEVO CON STOCK", "PENDIENTE PUBLICAR"])).sum())
    pickeado_count = int((queue_df["Estado"].isin(["PICKEADO PARA PUBLICAR", "EN PROCESO DE PUBLICACIÓN"])).sum())
    faltante_count = int((queue_df["Estado"] == "FALTANTE FÍSICO CON STOCK KAME").sum())
    no_pub_count = int((queue_df["Estado"] == "NO PUBLICABLE").sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Por pickear/revisar", pendiente_count)
    k2.metric("Pickeados para publicar", pickeado_count)
    k3.metric("Faltante físico", faltante_count)
    k4.metric("No publicables", no_pub_count)

    responsable = APP_USER_OPERACION

    with st.container(border=True):
        c1, c2, c3 = st.columns([1.6, 1.4, 2.2])
        modo = c1.selectbox(
            "Cola",
            [
                "Pendientes por pickear/revisar",
                "Pickeados para publicar",
                "Faltante físico con stock Kame",
                "No publicables",
                "Cubiertos por pack/unidad",
                "Sin stock Kame",
                "Todos",
            ],
            key="operar_modo",
        )
        familias = ["TODAS"] + sorted(queue_df["Familia"].fillna("").astype(str).unique().tolist())
        familia = c2.selectbox("Familia", familias, key="operar_familia")
        search = c3.text_input("Buscar", key="operar_busqueda", placeholder="SKU o descripción")

    df = queue_df.copy()

    if modo == "Pendientes por pickear/revisar":
        df = df[df["Estado"].isin(["LLEGÓ STOCK", "PRODUCTO NUEVO CON STOCK", "PENDIENTE PUBLICAR"])]
    elif modo == "Pickeados para publicar":
        df = df[df["Estado"].isin(["PICKEADO PARA PUBLICAR", "EN PROCESO DE PUBLICACIÓN"])]
    elif modo == "Faltante físico con stock Kame":
        df = df[df["Estado"] == "FALTANTE FÍSICO CON STOCK KAME"]
    elif modo == "No publicables":
        df = df[df["Estado"] == "NO PUBLICABLE"]
    elif modo == "Cubiertos por pack/unidad":
        df = df[df["Estado"].isin(["CUBIERTO POR PACK", "CUBIERTO POR UNIDAD"])]
    elif modo == "Sin stock Kame":
        df = df[df["Estado"].isin(["SIN STOCK", "PRODUCTO NUEVO SIN STOCK"])]

    if familia != "TODAS":
        df = df[df["Familia"] == familia]

    if search.strip():
        s = search.strip().lower()
        df = df[
            df["SKU"].astype(str).str.lower().str.contains(s, na=False) |
            df["Descripcion"].astype(str).str.lower().str.contains(s, na=False)
        ]

    df = priority_sort(df)

    left, right = st.columns([2, 1])
    left.markdown(f"**Productos en esta cola:** {len(df):,}")
    cantidad = right.selectbox("Mostrar", [10, 20, 50, 100], index=1, key="operar_cantidad")

    if modo == "Faltante físico con stock Kame":
        reporte_cols = [
            "SKU", "Descripcion", "Familia", "EAN", "StockSistema", "Estado",
            "Motivo", "Observacion", "FechaUltimaGestion"
        ]
        reporte_cols = [c for c in reporte_cols if c in df.columns]
        if not df.empty:
            st.download_button(
                "Descargar informe para bodega",
                data=export_excel(df[reporte_cols]),
                file_name=f"informe_faltante_fisico_stock_kame_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if df.empty:
        st.info("No hay productos en esta cola con los filtros actuales.")
        return

    if "no_publicable_sku_abierto" not in st.session_state:
        st.session_state["no_publicable_sku_abierto"] = ""

    for idx, row in df.head(int(cantidad)).iterrows():
        sku = str(row.get("SKU", ""))
        estado = str(row.get("Estado", ""))
        descripcion = str(row.get("Descripcion", ""))
        familia_actual = str(row.get("Familia", ""))
        stock = row.get("StockSistema", 0)
        origen = str(row.get("Origen", ""))
        ean = str(row.get("EAN", "") or "")
        relacion = str(row.get("RelacionPackUnidad", "") or "")
        relacionado = str(row.get("SKURelacionadoPublicado", "") or "")

        st.markdown('<div class="aurora-card">', unsafe_allow_html=True)

        top_left, top_mid, top_right = st.columns([1.3, 4.7, 1.7])
        with top_left:
            st.markdown(f'<div class="aurora-sku">SKU: <code>{sku}</code></div>', unsafe_allow_html=True)
            stock_text = f"{stock:g}" if isinstance(stock, (int, float)) else str(stock)
            st.markdown(f'<span class="stock-pill">Stock Kame: {stock_text}</span>', unsafe_allow_html=True)
        with top_mid:
            st.markdown(f'<div class="aurora-desc">{descripcion}</div>', unsafe_allow_html=True)
            st.markdown(
                f'<div class="aurora-meta">Familia: {familia_actual} &nbsp; | &nbsp; Origen: {origen} &nbsp; | &nbsp; EAN: {ean}</div>',
                unsafe_allow_html=True
            )
            if relacion:
                st.warning(f"{relacion}: {relacionado}")
        with top_right:
            st.markdown(f'<span class="estado-pill">{estado}</span>', unsafe_allow_html=True)

        es_pendiente_pickear = estado in ["LLEGÓ STOCK", "PRODUCTO NUEVO CON STOCK", "PENDIENTE PUBLICAR"]
        es_pickeado_publicar = estado in ["PICKEADO PARA PUBLICAR", "EN PROCESO DE PUBLICACIÓN"]

        if es_pendiente_pickear:
            b1, b2, b3 = st.columns(3)

            if b1.button("Pickeado para publicar", key=f"pickeado_{sku}_{idx}", use_container_width=True):
                try:
                    save_status_change(row, "PICKEADO PARA PUBLICAR", responsable, "Pickeado para publicar", "", "")
                    st.success(f"{sku} pasó a PICKEADO PARA PUBLICAR")
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo guardar: {e}")

            if b2.button("Faltante físico", key=f"faltante_fisico_{sku}_{idx}", use_container_width=True):
                try:
                    save_status_change(
                        row,
                        "FALTANTE FÍSICO CON STOCK KAME",
                        responsable,
                        "Faltante físico con stock en Kame",
                        "",
                        ""
                    )
                    st.warning(f"{sku} enviado a cola: FALTANTE FÍSICO CON STOCK KAME")
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo guardar: {e}")

            if b3.button("No publicable", key=f"abrir_nopub_{sku}_{idx}", use_container_width=True):
                st.session_state["no_publicable_sku_abierto"] = sku
                st.rerun()

        elif es_pickeado_publicar:
            b1, b2 = st.columns(2)

            if b1.button("Publicado", key=f"publicado_{sku}_{idx}", use_container_width=True):
                try:
                    estado_publicado = "PRODUCTO NUEVO PUBLICADO" if origen == "PRODUCTO NUEVO" else "PUBLICADO"
                    save_status_change(row, estado_publicado, responsable, "Publicado correctamente", "", "")
                    st.success(f"{sku} marcado como {estado_publicado}")
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo guardar: {e}")

            if b2.button("No publicable", key=f"abrir_nopub_{sku}_{idx}", use_container_width=True):
                st.session_state["no_publicable_sku_abierto"] = sku
                st.rerun()

        else:
            st.caption("Esta cola es solo de consulta. Para correcciones usa el módulo Administrador.")

        if st.session_state.get("no_publicable_sku_abierto") == sku:
            st.markdown('<div class="section-soft">', unsafe_allow_html=True)
            st.warning("Selecciona el motivo. La lista es cerrada para mantener los informes ordenados.")

            motivo_categoria = st.selectbox(
                "Motivo no publicable",
                MOTIVOS_NO_PUBLICABLE,
                key=f"motivo_np_categoria_{sku}_{idx}",
            )

            csave, ccancel = st.columns(2)

            if csave.button("Confirmar no publicable", key=f"confirmar_nopub_{sku}_{idx}", use_container_width=True):
                try:
                    save_status_change(row, "NO PUBLICABLE", responsable, motivo_categoria, "", "")
                    st.session_state["no_publicable_sku_abierto"] = ""
                    st.success(f"{sku} marcado como NO PUBLICABLE")
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo guardar: {e}")

            if ccancel.button("Cancelar", key=f"cancelar_nopub_{sku}_{idx}", use_container_width=True):
                st.session_state["no_publicable_sku_abierto"] = ""
                st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)


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

    df = priority_sort(df)



    display_cols = [
        "SKU", "Descripcion", "Familia", "EAN", "Origen", "StockAnterior", "StockSistema",
        "Estado", "EstadoSugerido", "RelacionPackUnidad", "SKURelacionadoPublicado",
        "Motivo", "Observacion", "LinkPublicacion"
    ]

    st.dataframe(df[display_cols], use_container_width=True, hide_index=True)

    st.download_button(
        "Descargar cola filtrada en Excel",
        data=export_excel(df),
        file_name=f"cola_marketing_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("Edición manual avanzada")

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
            "responsable": responsable.strip() or APP_USER_ADMIN,
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


def reportes_ui(queue_df: pd.DataFrame):
    inject_operational_css()

    st.subheader("Informes")
    st.caption("Esta sección es solo para descargar informes útiles. No es para operar productos.")

    if queue_df.empty:
        st.info("No hay datos para generar informes.")
        return

    base_cols = [
        "SKU", "Descripcion", "Familia", "EAN", "Origen",
        "StockSistema", "Estado", "Motivo", "Responsable",
        "FechaUltimaGestion", "Observacion", "RelacionPackUnidad",
        "SKURelacionadoPublicado"
    ]
    base_cols = [c for c in base_cols if c in queue_df.columns]

    reportes = [
        {
            "titulo": "Productos pickeados para publicar",
            "descripcion": "Lista que debe tomar la siguiente etapa para crear o continuar publicaciones.",
            "estados": ["PICKEADO PARA PUBLICAR", "EN PROCESO DE PUBLICACIÓN"],
            "archivo": "productos_pickeados_para_publicar",
        },
        {
            "titulo": "Faltante físico con stock Kame",
            "descripcion": "Productos que Kame muestra con stock, pero no fueron encontrados físicamente. Informe para bodega.",
            "estados": ["FALTANTE FÍSICO CON STOCK KAME"],
            "archivo": "faltante_fisico_con_stock_kame",
        },
        {
            "titulo": "No publicables",
            "descripcion": "Productos descartados por criterio del operador, con motivo escrito.",
            "estados": ["NO PUBLICABLE"],
            "archivo": "productos_no_publicables",
        },
        {
            "titulo": "Pendientes por pickear/revisar",
            "descripcion": "Productos que todavía deben revisarse físicamente antes de pasar a publicación.",
            "estados": ["LLEGÓ STOCK", "PRODUCTO NUEVO CON STOCK", "PENDIENTE PUBLICAR"],
            "archivo": "pendientes_por_pickear_revisar",
        },
        {
            "titulo": "Sin stock Kame",
            "descripcion": "Productos que no pasan a operación porque el sistema no muestra stock disponible.",
            "estados": ["SIN STOCK", "PRODUCTO NUEVO SIN STOCK"],
            "archivo": "sin_stock_kame",
        },
        {
            "titulo": "Cubiertos por pack o unidad",
            "descripcion": "Productos que no necesariamente requieren publicación individual porque ya están cubiertos por otra publicación.",
            "estados": ["CUBIERTO POR PACK", "CUBIERTO POR UNIDAD"],
            "archivo": "cubiertos_por_pack_o_unidad",
        },
    ]

    for rep in reportes:
        df_rep = queue_df[queue_df["Estado"].isin(rep["estados"])].copy()
        df_rep = priority_sort(df_rep) if not df_rep.empty else df_rep

        with st.container(border=True):
            c1, c2, c3 = st.columns([3.2, 1, 1.2])
            c1.markdown(f"### {rep['titulo']}")
            c1.caption(rep["descripcion"])
            c2.metric("Productos", len(df_rep))

            if not df_rep.empty:
                c3.download_button(
                    "Descargar Excel",
                    data=export_excel(df_rep[base_cols]),
                    file_name=f"{rep['archivo']}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"reporte_descargar_{rep['archivo']}",
                )
            else:
                c3.button(
                    "Sin datos",
                    disabled=True,
                    use_container_width=True,
                    key=f"reporte_sin_datos_{rep['archivo']}",
                )

            if not df_rep.empty:
                with st.expander("Ver muestra"):
                    st.dataframe(df_rep[base_cols].head(50), use_container_width=True, hide_index=True)


def auditoria_ui():
    inject_operational_css()

    st.subheader("Auditoría")
    st.caption("Aquí se revisa quién cambió cada producto, cuándo y por qué. No muestra la cola actual, solo movimientos.")

    try:
        data = api_call("get_history", {"limit": 1000}, timeout=60)
        rows = data.get("historial_cambios", [])
        df = pd.DataFrame(rows)

        if df.empty:
            st.info("Todavía no hay movimientos registrados.")
            return

        c1, c2 = st.columns([1.2, 2.8])
        accion_filter = "TODAS"
        search = ""

        if "Accion" in df.columns:
            accion_filter = c1.selectbox(
                "Acción",
                ["TODAS"] + sorted(df["Accion"].fillna("").astype(str).unique().tolist()),
                key="aud_accion",
            )

        search = c2.text_input("Buscar SKU o descripción", key="aud_busqueda")

        if accion_filter != "TODAS" and "Accion" in df.columns:
            df = df[df["Accion"].fillna("").astype(str) == accion_filter]

        if search.strip():
            s = search.strip().lower()
            mask = pd.Series(False, index=df.index)

            for col in ["SKU", "Descripcion", "Motivo", "Observacion"]:
                if col in df.columns:
                    mask = mask | df[col].fillna("").astype(str).str.lower().str.contains(s, na=False)

            df = df[mask]

        st.write(f"Movimientos encontrados: **{len(df):,}**")

        if "Fecha" in df.columns:
            df = df.sort_values("Fecha", ascending=False)

        st.dataframe(df.head(500), use_container_width=True, hide_index=True)

        st.download_button(
            "Descargar auditoría filtrada",
            data=export_excel(df),
            file_name=f"auditoria_movimientos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"No se pudo cargar auditoría: {e}")



def status_ui(estado_df: pd.DataFrame, inv_df: pd.DataFrame):
    st.subheader("Estado de conexión y base central")

    url, token = get_api_config()

    c1, c2, c3 = st.columns(3)
    c1.metric("URL Apps Script", "Configurada" if url else "Falta")
    c2.metric("Token", "Configurado" if token else "Falta")
    c3.metric("Inventario central", f"{len(inv_df):,} SKUs")

    st.write("Últimos estados guardados en Google Sheets:")
    st.dataframe(estado_df.tail(100), use_container_width=True, hide_index=True)



# ============================================================
# Módulo administrador
# ============================================================

def get_admin_password() -> str:
    try:
        return st.secrets.get("ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD)
    except Exception:
        return DEFAULT_ADMIN_PASSWORD


def admin_login_ui() -> bool:
    if st.session_state.get("admin_autenticado"):
        cols = st.columns([3, 1])
        cols[0].success("Administrador activo")
        if cols[1].button("Cerrar sesión admin", use_container_width=True):
            st.session_state["admin_autenticado"] = False
            st.rerun()
        return True

    st.warning("Módulo restringido. Ingresa la clave de administrador.")
    clave = st.text_input("Clave administrador", type="password", key="admin_password_input")

    if st.button("Entrar como administrador", use_container_width=True):
        if clave == get_admin_password():
            st.session_state["admin_autenticado"] = True
            st.rerun()
        else:
            st.error("Clave incorrecta.")

    return False


def payload_from_queue_row(
    row: pd.Series,
    nuevo_estado: str,
    responsable: str,
    motivo: str = "",
    observacion: str = "",
    link_publicacion: str = "",
    accion: str = "CAMBIO ADMINISTRADOR",
) -> dict:
    return {
        "fecha": now_iso(),
        "sku": clean_sku(row.get("SKU", "")),
        "estado": nuevo_estado,
        "origen": row.get("Origen", ""),
        "descripcion": row.get("Descripcion", ""),
        "familia": row.get("Familia", ""),
        "ean": row.get("EAN", ""),
        "stock_anterior": float(row.get("StockAnterior", 0) or 0),
        "stock_actual": float(row.get("StockSistema", 0) or 0),
        "responsable": responsable.strip() or APP_USER_ADMIN,
        "motivo": motivo,
        "observacion": observacion,
        "link_publicacion": link_publicacion,
        "accion": accion,
    }


def payload_manual_sku(
    sku: str,
    nuevo_estado: str,
    responsable: str,
    motivo: str = "",
    observacion: str = "",
    link_publicacion: str = "",
) -> dict:
    sku = clean_sku(sku)
    return {
        "fecha": now_iso(),
        "sku": sku,
        "estado": nuevo_estado,
        "origen": "REGISTRO MANUAL ADMIN",
        "descripcion": "",
        "familia": "",
        "ean": "",
        "stock_anterior": 0,
        "stock_actual": 0,
        "responsable": responsable.strip() or APP_USER_ADMIN,
        "motivo": motivo,
        "observacion": observacion,
        "link_publicacion": link_publicacion,
        "accion": "CAMBIO ADMINISTRADOR MANUAL",
    }


def parse_skus(text_skus: str) -> List[str]:
    raw = re.split(r"[\n,; \t]+", str(text_skus or ""))
    skus = []
    seen = set()

    for item in raw:
        sku = clean_sku(item)

        if sku and sku not in seen:
            seen.add(sku)
            skus.append(sku)

    return skus


def admin_kpis(queue_df: pd.DataFrame, estado_df: pd.DataFrame, inv_current_df: pd.DataFrame):
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Cola total", len(queue_df))
    c2.metric("Estados guardados", len(estado_df))
    c3.metric("Inventario SKUs", len(inv_current_df))
    c4.metric("Faltante físico", int((queue_df["Estado"] == "FALTANTE FÍSICO CON STOCK KAME").sum()) if not queue_df.empty else 0)
    c5.metric("No publicables", int((queue_df["Estado"] == "NO PUBLICABLE").sum()) if not queue_df.empty else 0)


def admin_filtros(queue_df: pd.DataFrame, prefix: str = "admin") -> pd.DataFrame:
    if queue_df.empty:
        return queue_df

    c1, c2, c3 = st.columns([1.4, 1.4, 2.2])

    estados = ["TODOS"] + ESTADOS
    estado = c1.selectbox("Estado", estados, key=f"{prefix}_estado")

    familias = ["TODAS"] + sorted(queue_df["Familia"].fillna("").astype(str).unique().tolist())
    familia = c2.selectbox("Familia", familias, key=f"{prefix}_familia")

    search = c3.text_input("Buscar SKU o descripción", key=f"{prefix}_buscar")

    df = queue_df.copy()

    if estado != "TODOS":
        df = df[df["Estado"] == estado]

    if familia != "TODAS":
        df = df[df["Familia"] == familia]

    if search.strip():
        s = search.strip().lower()
        df = df[
            df["SKU"].astype(str).str.lower().str.contains(s, na=False) |
            df["Descripcion"].astype(str).str.lower().str.contains(s, na=False)
        ]

    return priority_sort(df)


def admin_editar_un_sku(queue_df: pd.DataFrame):
    st.markdown("### Editar un SKU")
    st.caption("Permite corregir un producto puntual sin buscarlo en la operación normal.")

    if queue_df.empty:
        st.info("No hay cola disponible.")
        return

    busqueda = st.text_input("Buscar SKU o descripción", key="admin_single_search")
    responsable = APP_USER_ADMIN

    df = queue_df.copy()

    if busqueda.strip():
        s = busqueda.strip().lower()
        df = df[
            df["SKU"].astype(str).str.lower().str.contains(s, na=False) |
            df["Descripcion"].astype(str).str.lower().str.contains(s, na=False)
        ]

    df = priority_sort(df)

    if df.empty:
        st.info("No hay resultados.")
        return

    opciones = [
        f"{row['SKU']} | {row['Descripcion']} | {row['Estado']}"
        for _, row in df.head(300).iterrows()
    ]

    seleccionado = st.selectbox("Producto", opciones, key="admin_single_producto")
    sku = seleccionado.split("|")[0].strip()
    row = df[df["SKU"].astype(str) == sku].iloc[0]

    with st.container(border=True):
        st.write(f"**SKU:** `{row.get('SKU', '')}`")
        st.write(f"**Descripción:** {row.get('Descripcion', '')}")
        st.write(f"**Estado actual:** {row.get('Estado', '')}")
        st.write(f"**Stock Kame:** {row.get('StockSistema', 0)}")
        st.write(f"**Familia:** {row.get('Familia', '')}")

    c3, c4 = st.columns(2)
    nuevo_estado = c3.selectbox(
        "Nuevo estado",
        ESTADOS,
        index=ESTADOS.index(row.get("Estado", "")) if row.get("Estado", "") in ESTADOS else 0,
        key="admin_single_estado",
    )
    link_publicacion = c4.text_input(
        "Link publicación ML",
        value=str(row.get("LinkPublicacion", "") or ""),
        key="admin_single_link",
    )

    if nuevo_estado == "NO PUBLICABLE":
        motivo = st.selectbox(
            "Motivo no publicable",
            MOTIVOS_NO_PUBLICABLE,
            key="admin_single_motivo_np",
        )
    else:
        motivo = st.selectbox(
            "Motivo del cambio",
            MOTIVOS_GENERALES,
            key="admin_single_motivo_general",
        )

    if st.button("Guardar cambio administrador", use_container_width=True):
        try:
            payload = payload_from_queue_row(
                row,
                nuevo_estado,
                responsable,
                motivo or "Cambio administrador",
                "",
                link_publicacion.strip(),
            )
            api_upsert_product(payload)
            st.success(f"SKU {sku} actualizado a {nuevo_estado}.")
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo guardar el cambio: {e}")


def admin_cambios_masivos(queue_df: pd.DataFrame):
    st.markdown("### Cambios masivos por SKU")
    st.caption("Pega una lista de SKUs y cambia su estado en bloque. Los SKUs no encontrados se reportan y no se modifican.")

    c1, c2 = st.columns([2, 1])
    skus_text = c1.text_area(
        "SKUs a modificar",
        key="admin_bulk_skus",
        placeholder="Pega SKUs separados por salto de línea, coma o espacio.",
        height=160,
    )
    responsable = APP_USER_ADMIN
    nuevo_estado = c2.selectbox("Nuevo estado masivo", ESTADOS, key="admin_bulk_estado")

    if nuevo_estado == "NO PUBLICABLE":
        motivo = st.selectbox(
            "Motivo no publicable",
            MOTIVOS_NO_PUBLICABLE,
            key="admin_bulk_motivo_np",
        )
    else:
        motivo = st.selectbox(
            "Motivo del cambio masivo",
            MOTIVOS_GENERALES,
            key="admin_bulk_motivo_general",
        )

    observacion = ""

    skus = parse_skus(skus_text)
    st.write(f"SKUs detectados: **{len(skus):,}**")

    if skus:
        queue_map = {clean_sku(r["SKU"]): r for _, r in queue_df.iterrows()}
        encontrados = [sku for sku in skus if sku in queue_map]
        no_encontrados = [sku for sku in skus if sku not in queue_map]

        c3, c4 = st.columns(2)
        c3.metric("Encontrados", len(encontrados))
        c4.metric("No encontrados", len(no_encontrados))

        if no_encontrados:
            with st.expander("Ver SKUs no encontrados"):
                st.code("\n".join(no_encontrados))

    confirmar = st.checkbox(
        "Confirmo que quiero aplicar este cambio masivo",
        key="admin_bulk_confirm",
    )

    if st.button("Aplicar cambio masivo", use_container_width=True):
        if not skus:
            st.error("Pega al menos un SKU.")
        elif not confirmar:
            st.error("Marca la confirmación antes de aplicar el cambio masivo.")
        else:
            try:
                queue_map = {clean_sku(r["SKU"]): r for _, r in queue_df.iterrows()}
                items = []

                for sku in skus:
                    if sku not in queue_map:
                        continue

                    items.append(
                        payload_from_queue_row(
                            queue_map[sku],
                            nuevo_estado,
                            responsable,
                            motivo or "Cambio masivo administrador",
                            observacion,
                            "",
                            accion="CAMBIO MASIVO ADMINISTRADOR",
                        )
                    )

                if not items:
                    st.error("Ningún SKU fue encontrado en la cola actual.")
                else:
                    api_bulk_upsert_products(items, chunk_size=250)
                    st.success(f"Cambio masivo aplicado a {len(items):,} SKUs.")
                    st.rerun()
            except Exception as e:
                st.error(f"No se pudo aplicar el cambio masivo: {e}")


def admin_crear_o_forzar_sku():
    st.markdown("### Crear o forzar estado de SKU manual")
    st.caption("Úsalo solo cuando el SKU no aparece en la cola, pero necesitas dejar un estado guardado en la base central.")

    c1, c2 = st.columns([1.2, 1.3])
    sku = c1.text_input("SKU", key="admin_manual_sku")
    nuevo_estado = c2.selectbox("Estado", ESTADOS, key="admin_manual_estado")
    responsable = APP_USER_ADMIN

    if nuevo_estado == "NO PUBLICABLE":
        motivo = st.selectbox(
            "Motivo no publicable",
            MOTIVOS_NO_PUBLICABLE,
            key="admin_manual_motivo_np",
        )
    else:
        motivo = st.selectbox(
            "Motivo",
            MOTIVOS_GENERALES,
            key="admin_manual_motivo_general",
        )

    link_publicacion = st.text_input("Link publicación ML", key="admin_manual_link")

    if st.button("Guardar SKU manual", use_container_width=True):
        if not clean_sku(sku):
            st.error("Indica SKU.")
        else:
            try:
                payload = payload_manual_sku(
                    sku,
                    nuevo_estado,
                    responsable,
                    motivo or "Registro manual administrador",
                    "",
                    link_publicacion.strip(),
                )
                api_upsert_product(payload)
                st.success(f"SKU {clean_sku(sku)} guardado manualmente como {nuevo_estado}.")
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo guardar SKU manual: {e}")


def admin_descargas(queue_df: pd.DataFrame, estado_df: pd.DataFrame, inv_current_df: pd.DataFrame):
    st.markdown("### Descargas administrativas")
    st.caption("Bases completas para respaldo, revisión o corrección externa.")

    c1, c2, c3 = st.columns(3)

    c1.download_button(
        "Descargar cola completa",
        data=export_excel(queue_df),
        file_name=f"admin_cola_completa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=queue_df.empty,
        key="admin_download_cola_completa",
    )

    c2.download_button(
        "Descargar estados guardados",
        data=export_excel(estado_df),
        file_name=f"admin_estado_actual_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=estado_df.empty,
        key="admin_download_estado_actual",
    )

    c3.download_button(
        "Descargar inventario central",
        data=export_excel(inv_current_df),
        file_name=f"admin_inventario_actual_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=inv_current_df.empty,
        key="admin_download_inventario_actual",
    )

    st.divider()

    df_admin = admin_filtros(queue_df, prefix="admin_download")
    st.write(f"Registros filtrados: **{len(df_admin):,}**")
    st.dataframe(df_admin.head(500), use_container_width=True, hide_index=True)

    st.download_button(
        "Descargar vista filtrada",
        data=export_excel(df_admin),
        file_name=f"admin_vista_filtrada_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=df_admin.empty,
        key="admin_download_vista_filtrada",
    )



def admin_informes_personalizados(queue_df: pd.DataFrame):
    st.markdown("### Informes administrativos por estado y motivo")
    st.caption(
        "Genera informes filtrados para los estados críticos: No publicable, Faltante físico, Pickeado para publicar, Publicado y Sin stock."
    )

    if queue_df.empty:
        st.info("No hay datos disponibles para generar informes.")
        return

    estados_informe = {
        "NO PUBLICABLE": ["NO PUBLICABLE"],
        "FALTANTE FÍSICO CON STOCK KAME": ["FALTANTE FÍSICO CON STOCK KAME"],
        "PICKEADO PARA PUBLICAR": ["PICKEADO PARA PUBLICAR", "EN PROCESO DE PUBLICACIÓN"],
        "PUBLICADO": ["PUBLICADO", "PRODUCTO NUEVO PUBLICADO"],
        "SIN STOCK": ["SIN STOCK", "PRODUCTO NUEVO SIN STOCK"],
    }

    base_cols = [
        "SKU",
        "Descripcion",
        "Familia",
        "EAN",
        "Origen",
        "StockSistema",
        "Estado",
        "Motivo",
        "Observacion",
        "FechaUltimaGestion",
        "RelacionPackUnidad",
        "SKURelacionadoPublicado",
        "LinkPublicacion",
    ]
    base_cols = [c for c in base_cols if c in queue_df.columns]

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("No publicable", int((queue_df["Estado"] == "NO PUBLICABLE").sum()))
    k2.metric("Faltante físico", int((queue_df["Estado"] == "FALTANTE FÍSICO CON STOCK KAME").sum()))
    k3.metric("Pickeado", int((queue_df["Estado"].isin(["PICKEADO PARA PUBLICAR", "EN PROCESO DE PUBLICACIÓN"])).sum()))
    k4.metric("Publicado", int((queue_df["Estado"].isin(["PUBLICADO", "PRODUCTO NUEVO PUBLICADO"])).sum()))
    k5.metric("Sin stock", int((queue_df["Estado"].isin(["SIN STOCK", "PRODUCTO NUEVO SIN STOCK"])).sum()))

    st.divider()

    c1, c2, c3, c4 = st.columns([1.4, 1.6, 1.4, 2.2])

    tipo_informe = c1.selectbox(
        "Tipo de informe",
        list(estados_informe.keys()),
        key="admin_report_tipo",
    )

    df = queue_df[queue_df["Estado"].isin(estados_informe[tipo_informe])].copy()

    motivos = ["TODOS"]
    if "Motivo" in df.columns and not df.empty:
        motivos += sorted([m for m in df["Motivo"].fillna("").astype(str).unique().tolist() if m.strip()])

    motivo = c2.selectbox("Motivo", motivos, key="admin_report_motivo")

    familias = ["TODAS"]
    if "Familia" in df.columns and not df.empty:
        familias += sorted([f for f in df["Familia"].fillna("").astype(str).unique().tolist() if f.strip()])

    familia = c3.selectbox("Familia", familias, key="admin_report_familia")
    search = c4.text_input("Buscar SKU o descripción", key="admin_report_search")

    orden = st.selectbox(
        "Ordenar por",
        ["Más reciente", "Mayor stock Kame", "Familia", "SKU"],
        key="admin_report_orden",
    )

    if motivo != "TODOS" and "Motivo" in df.columns:
        df = df[df["Motivo"].fillna("").astype(str) == motivo]

    if familia != "TODAS" and "Familia" in df.columns:
        df = df[df["Familia"].fillna("").astype(str) == familia]

    if search.strip():
        s = search.strip().lower()
        mask = pd.Series(False, index=df.index)
        for col in ["SKU", "Descripcion", "Motivo", "Observacion"]:
            if col in df.columns:
                mask = mask | df[col].fillna("").astype(str).str.lower().str.contains(s, na=False)
        df = df[mask]

    if not df.empty:
        if orden == "Más reciente" and "FechaUltimaGestion" in df.columns:
            df = df.sort_values("FechaUltimaGestion", ascending=False)
        elif orden == "Mayor stock Kame" and "StockSistema" in df.columns:
            df = df.sort_values("StockSistema", ascending=False)
        elif orden == "Familia" and "Familia" in df.columns:
            df = df.sort_values(["Familia", "Descripcion"], ascending=True)
        elif orden == "SKU" and "SKU" in df.columns:
            df = df.sort_values("SKU", ascending=True)

    st.write(f"Registros del informe: **{len(df):,}**")

    if df.empty:
        st.info("No hay registros con los filtros seleccionados.")
        return

    resumen_cols = []
    if "Estado" in df.columns:
        resumen_cols.append("Estado")
    if "Motivo" in df.columns:
        resumen_cols.append("Motivo")

    if resumen_cols:
        with st.expander("Resumen agrupado", expanded=True):
            resumen = (
                df.groupby(resumen_cols, dropna=False)
                .size()
                .reset_index(name="Cantidad")
                .sort_values("Cantidad", ascending=False)
            )
            st.dataframe(resumen, use_container_width=True, hide_index=True)

            st.download_button(
                "Descargar resumen agrupado",
                data=export_excel(resumen),
                file_name=f"resumen_{tipo_informe.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"admin_resumen_agrupado_{tipo_informe}",
            )

    st.dataframe(df[base_cols].head(1000), use_container_width=True, hide_index=True)

    nombre_base = tipo_informe.lower()
    nombre_base = (
        nombre_base.replace(" ", "_")
        .replace("í", "i")
        .replace("í", "i")
        .replace("é", "e")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("á", "a")
        .replace("/", "_")
    )

    st.download_button(
        "Descargar informe detallado",
        data=export_excel(df[base_cols]),
        file_name=f"informe_{nombre_base}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=f"admin_informe_detallado_{tipo_informe}",
    )

    st.divider()

    st.markdown("#### Descarga rápida por cada estado")
    st.caption("Botones directos para bajar cada informe sin cambiar filtros.")

    quick_cols = st.columns(5)

    for i, (nombre, estados) in enumerate(estados_informe.items()):
        df_quick = queue_df[queue_df["Estado"].isin(estados)].copy()
        file_name = (
            nombre.lower()
            .replace(" ", "_")
            .replace("í", "i")
            .replace("é", "e")
            .replace("ó", "o")
            .replace("ú", "u")
            .replace("á", "a")
        )

        quick_cols[i].download_button(
            f"{nombre} ({len(df_quick):,})",
            data=export_excel(df_quick[base_cols]) if not df_quick.empty else export_excel(pd.DataFrame(columns=base_cols)),
            file_name=f"informe_{file_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=df_quick.empty,
            key=f"admin_quick_download_{file_name}",
        )


def administrador_ui(queue_df: pd.DataFrame, estado_df: pd.DataFrame, inv_current_df: pd.DataFrame):
    inject_operational_css()

    st.subheader("Administrador")
    st.caption("Módulo de superusuario para corregir estados, hacer cambios masivos y descargar bases.")

    if not admin_login_ui():
        st.info("Configura `ADMIN_PASSWORD` en los Secrets de Streamlit para cambiar la clave por defecto.")
        return

    admin_kpis(queue_df, estado_df, inv_current_df)

    tabs_admin = st.tabs([
        "Editar SKU",
        "Cambios masivos",
        "Forzar SKU manual",
        "Informes admin",
        "Descargas",
    ])

    with tabs_admin[0]:
        admin_editar_un_sku(queue_df)

    with tabs_admin[1]:
        admin_cambios_masivos(queue_df)

    with tabs_admin[2]:
        admin_crear_o_forzar_sku()

    with tabs_admin[3]:
        admin_informes_personalizados(queue_df)

    with tabs_admin[4]:
        admin_descargas(queue_df, estado_df, inv_current_df)


def main():
    page_config()
    st.sidebar.caption(f"Versión: {APP_VERSION}")
    st.sidebar.caption("Modo rápido + sync segundo plano")

    init_sync_state()
    pending_sync = st.session_state.get("sync_pending_count", 0)
    ok_sync = st.session_state.get("sync_ok_count", 0)
    error_sync = st.session_state.get("sync_error_count", 0)

    if pending_sync:
        st.sidebar.warning(f"Sincronizando con Google Sheets: {pending_sync} pendiente(s)")
    else:
        st.sidebar.success("Sincronización al día")

    if error_sync:
        st.sidebar.error(f"Errores de sincronización: {error_sync}")
        with st.sidebar.expander("Ver últimos errores"):
            st.write(st.session_state.get("sync_errors", []))

    if ok_sync:
        st.sidebar.caption(f"Sincronizaciones OK: {ok_sync}")

    validate_base_files()

    try:
        maestro = load_maestro()
        publicaciones = load_publicaciones()
        packs = load_packs()
    except Exception as e:
        st.error(f"Error leyendo archivos base: {e}")
        st.stop()

    cache_loaded_at = st.session_state.get("data_cache_loaded_at", "")
    if cache_loaded_at:
        st.sidebar.caption(f"Datos en memoria: {cache_loaded_at}")

    if st.sidebar.button("Actualizar datos desde Google Sheets"):
        clear_session_data_cache()
        st.rerun()

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

    tabs = st.tabs(["Operar productos", "Informes", "Auditoría", "Administrador"])

    with tabs[0]:
        operar_productos_ui(queue_df)

    with tabs[1]:
        reportes_ui(queue_df)

    with tabs[2]:
        auditoria_ui()

    with tabs[3]:
        administrador_ui(queue_df, estado_df, inv_current_df)


if __name__ == "__main__":
    main()
