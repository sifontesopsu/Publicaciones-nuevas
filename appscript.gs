/**
 * Apps Script - Base central para App Streamlit
 * Gestión de Publicaciones Pendientes - Ferretería Aurora
 *
 * Hojas usadas:
 * - estado_actual
 * - inventario_actual
 * - inventario_cargas
 * - historial_cambios
 * - sync_log
 *
 * Script Properties necesarias:
 * - SPREADSHEET_ID
 * - SECRET_TOKEN
 */

const SHEETS = {
  ESTADO: "estado_actual",
  INVENTARIO: "inventario_actual",
  CARGAS: "inventario_cargas",
  HISTORIAL: "historial_cambios",
  LOG: "sync_log"
};

const HEADERS = {
  ESTADO: [
    "sku",
    "estado",
    "origen",
    "descripcion",
    "familia",
    "ean",
    "stock_anterior",
    "stock_actual",
    "responsable",
    "motivo",
    "observacion",
    "link_publicacion",
    "accion",
    "estado_anterior",
    "fecha_actualizacion"
  ],
  INVENTARIO: [
    "sku",
    "descripcion",
    "familia",
    "stock_actual",
    "costo_promedio",
    "saldo_valor",
    "fecha_carga"
  ],
  CARGAS: [
    "fecha_carga",
    "usuario",
    "total_skus",
    "total_stock",
    "productos_nuevos",
    "llegaron_stock"
  ],
  HISTORIAL: [
    "fecha",
    "sku",
    "accion",
    "estado_anterior",
    "estado_nuevo",
    "origen",
    "descripcion",
    "familia",
    "ean",
    "stock_anterior",
    "stock_actual",
    "responsable",
    "motivo",
    "observacion",
    "link_publicacion"
  ],
  LOG: [
    "fecha",
    "action",
    "sku",
    "resultado",
    "error",
    "payload_json"
  ]
};

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: props.getProperty("SPREADSHEET_ID"),
    secretToken: props.getProperty("SECRET_TOKEN")
  };
}

function getSpreadsheet_() {
  const config = getConfig_();
  if (!config.spreadsheetId) {
    throw new Error("Falta Script Property SPREADSHEET_ID");
  }
  return SpreadsheetApp.openById(config.spreadsheetId);
}

function getSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }

  return sh;
}

function ensureAllSheets_() {
  const ss = getSpreadsheet_();
  getSheet_(ss, SHEETS.ESTADO, HEADERS.ESTADO);
  getSheet_(ss, SHEETS.INVENTARIO, HEADERS.INVENTARIO);
  getSheet_(ss, SHEETS.CARGAS, HEADERS.CARGAS);
  getSheet_(ss, SHEETS.HISTORIAL, HEADERS.HISTORIAL);
  getSheet_(ss, SHEETS.LOG, HEADERS.LOG);
  return ss;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheetObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0].map(h => String(h).trim());
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const obj = {};
    let hasData = false;

    for (let j = 0; j < headers.length; j++) {
      const value = values[i][j];
      obj[headers[j]] = value;
      if (value !== "" && value !== null) hasData = true;
    }

    if (hasData) rows.push(obj);
  }

  return rows;
}

function appendLog_(ss, action, sku, resultado, error, payload) {
  const sh = getSheet_(ss, SHEETS.LOG, HEADERS.LOG);
  sh.appendRow([
    new Date(),
    action || "",
    sku || "",
    resultado || "",
    error || "",
    JSON.stringify(payload || {})
  ]);
}

function doGet(e) {
  ensureAllSheets_();
  return jsonResponse_({
    ok: true,
    app: "Gestión Publicaciones Aurora",
    message: "Web App activo"
  });
}

function doPost(e) {
  const config = getConfig_();

  try {
    if (!config.secretToken) {
      return jsonResponse_({ ok: false, error: "Falta Script Property SECRET_TOKEN" });
    }

    const body = JSON.parse(e.postData.contents || "{}");

    if (body.token !== config.secretToken) {
      return jsonResponse_({ ok: false, error: "Token inválido" });
    }

    const action = body.action || "";
    const payload = body.payload || {};
    const ss = ensureAllSheets_();

    if (action === "get_data") {
      return jsonResponse_(handleGetData_(ss));
    }

    if (action === "get_history") {
      return jsonResponse_(handleGetHistory_(ss, payload));
    }

    // Escrituras con lock para evitar choques entre usuarios.
    const lock = LockService.getScriptLock();
    const gotLock = lock.tryLock(30000);

    if (!gotLock) {
      return jsonResponse_({ ok: false, error: "No se pudo obtener lock de escritura. Intenta nuevamente." });
    }

    try {
      if (action === "upsert_product") {
        handleUpsertProduct_(ss, payload);
        appendLog_(ss, action, payload.sku, "OK", "", payload);
        return jsonResponse_({ ok: true });
      }

      if (action === "bulk_upsert_products") {
        const items = payload.items || [];
        for (let i = 0; i < items.length; i++) {
          handleUpsertProduct_(ss, items[i]);
        }
        appendLog_(ss, action, "", "OK", "", { total: items.length });
        return jsonResponse_({ ok: true, total: items.length });
      }

      if (action === "replace_inventory") {
        handleReplaceInventory_(ss, payload);
        appendLog_(ss, action, "", "OK", "", {
          total: (payload.items || []).length,
          usuario: payload.usuario || ""
        });
        return jsonResponse_({ ok: true, total: (payload.items || []).length });
      }

      appendLog_(ss, action, "", "IGNORADO", "Action no reconocido", payload);
      return jsonResponse_({ ok: false, error: "Action no reconocido: " + action });

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    try {
      const ss = ensureAllSheets_();
      appendLog_(ss, "", "", "ERROR", String(err), {});
    } catch (ignore) {}

    return jsonResponse_({
      ok: false,
      error: String(err),
      stack: err && err.stack ? String(err.stack) : ""
    });
  }
}

function handleGetData_(ss) {
  const shEstado = getSheet_(ss, SHEETS.ESTADO, HEADERS.ESTADO);
  const shInv = getSheet_(ss, SHEETS.INVENTARIO, HEADERS.INVENTARIO);

  return {
    ok: true,
    estado_actual: readSheetObjects_(shEstado),
    inventario_actual: readSheetObjects_(shInv)
  };
}

function handleGetHistory_(ss, payload) {
  const limit = Number(payload.limit || 500);
  const sh = getSheet_(ss, SHEETS.HISTORIAL, HEADERS.HISTORIAL);
  let rows = readSheetObjects_(sh);

  if (rows.length > limit) {
    rows = rows.slice(rows.length - limit);
  }

  rows.reverse();

  return {
    ok: true,
    historial_cambios: rows
  };
}

function findSkuRowMap_(sh) {
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const sku = String(values[i][0] || "").trim();
    if (sku) map[sku] = i + 2;
  }

  return map;
}

function handleUpsertProduct_(ss, payload) {
  const shEstado = getSheet_(ss, SHEETS.ESTADO, HEADERS.ESTADO);
  const shHist = getSheet_(ss, SHEETS.HISTORIAL, HEADERS.HISTORIAL);

  const sku = String(payload.sku || "").trim();
  if (!sku) throw new Error("SKU vacío en upsert_product");

  const rowMap = findSkuRowMap_(shEstado);
  const existingRow = rowMap[sku] || -1;

  let estadoAnterior = payload.estado_anterior || "";

  if (existingRow > 0) {
    estadoAnterior = String(shEstado.getRange(existingRow, 2).getValue() || "");
  }

  const fecha = payload.fecha || new Date();

  const estadoRow = [
    sku,
    payload.estado || "",
    payload.origen || "",
    payload.descripcion || "",
    payload.familia || "",
    payload.ean || "",
    payload.stock_anterior || 0,
    payload.stock_actual || 0,
    payload.responsable || "",
    payload.motivo || "",
    payload.observacion || "",
    payload.link_publicacion || "",
    payload.accion || "",
    estadoAnterior || "",
    fecha
  ];

  if (existingRow > 0) {
    shEstado.getRange(existingRow, 1, 1, HEADERS.ESTADO.length).setValues([estadoRow]);
  } else {
    shEstado.appendRow(estadoRow);
  }

  shHist.appendRow([
    fecha,
    sku,
    payload.accion || "",
    estadoAnterior || "",
    payload.estado || "",
    payload.origen || "",
    payload.descripcion || "",
    payload.familia || "",
    payload.ean || "",
    payload.stock_anterior || 0,
    payload.stock_actual || 0,
    payload.responsable || "",
    payload.motivo || "",
    payload.observacion || "",
    payload.link_publicacion || ""
  ]);
}

function handleReplaceInventory_(ss, payload) {
  const shInv = getSheet_(ss, SHEETS.INVENTARIO, HEADERS.INVENTARIO);
  const shCargas = getSheet_(ss, SHEETS.CARGAS, HEADERS.CARGAS);

  const items = payload.items || [];
  const fecha = payload.fecha_carga || new Date();

  shInv.clearContents();

  const values = [HEADERS.INVENTARIO];

  let totalStock = 0;

  for (let i = 0; i < items.length; i++) {
    const item = items[i] || {};
    const stock = Number(item.stock_actual || 0);
    totalStock += stock;

    values.push([
      String(item.sku || "").trim(),
      item.descripcion || "",
      item.familia || "",
      stock,
      Number(item.costo_promedio || 0),
      Number(item.saldo_valor || 0),
      fecha
    ]);
  }

  if (values.length > 0) {
    shInv.getRange(1, 1, values.length, HEADERS.INVENTARIO.length).setValues(values);
    shInv.setFrozenRows(1);
  }

  shCargas.appendRow([
    fecha,
    payload.usuario || "",
    items.length,
    totalStock,
    payload.productos_nuevos || 0,
    payload.llegaron_stock || 0
  ]);
}
