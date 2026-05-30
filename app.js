const state = {
  rawRows: [],
  deliveries: [],
  pickups: [],
  previousNotes: new Map(),
};

const requiredColumns = [
  "TIPO PEDIDO",
  "FECHA DE ENTREGA",
  "FECHA DE RECOGIDA",
  "#PEDIDO",
  "NOMBRE",
  "APELLIDO",
  "TELEFONO",
  "DIRECCION ENTREGA",
  "CIUDAD ENTREGA",
  "DEPARTAMENTO",
  "DIRECCION RECOGIDA",
  "CIUDAD RECOGIDA",
  "DPTO RECOGIDA",
  "COMBO",
  "CAJAS",
  "CARRITO",
  "ROLLO PAPEL BURBUJA",
  "SOBRES x10",
  "VINIPEL",
  "VERDES CNT40",
  "PRECINTOS",
  "TAGS",
  "OBSERVACIONES",
];

const deliveriesTable = document.getElementById("deliveriesTable");
const pickupsTable = document.getElementById("pickupsTable");
const fileInput = document.getElementById("fileInput");
const previousRouteInput = document.getElementById("previousRouteInput");
const generateRouteButton = document.getElementById("generateRoute");
const downloadExcelButton = document.getElementById("downloadExcel");
const downloadPdfButton = document.getElementById("downloadPdf");

document.querySelectorAll(".module-tab").forEach((tab) => {
  tab.addEventListener("click", () => switchPanel(tab.dataset.panel));
});
fileInput.addEventListener("change", readWorkbook);
previousRouteInput.addEventListener("change", readPreviousRoute);
generateRouteButton.addEventListener("click", generateRoute);
downloadExcelButton.addEventListener("click", downloadExcel);
downloadPdfButton.addEventListener("click", downloadPdf);

function normalize(value) {
  return String(value ?? "").trim();
}

function normalizeUpper(value) {
  return normalize(value).toUpperCase();
}

function normalizeHeader(value) {
  return normalizeUpper(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function switchPanel(panelId) {
  document.querySelectorAll(".module-tab").forEach((tab) => {
    tab.classList.toggle("active", tab.dataset.panel === panelId);
  });
  document.querySelectorAll(".module-panel").forEach((panel) => {
    const isActive = panel.id === panelId;
    panel.classList.toggle("active", isActive);
    panel.hidden = !isActive;
  });
}

function isAllowedOrderType(row) {
  return ["ALQUILER", "COMPRA"].includes(normalizeUpper(row["TIPO PEDIDO"]));
}

function normalizeExcelRows(rows) {
  const aliases = {
    "TIPO PEDIDO": ["TIPO PEDIDO"],
    "FECHA DE ENTREGA": ["FECHA DE ENTREGA"],
    "FECHA DE RECOGIDA": ["FECHA DE RECOGIDA"],
    "#PEDIDO": ["#PEDIDO", "PEDIDO", "NUMERO PEDIDO"],
    NOMBRE: ["NOMBRE"],
    APELLIDO: ["APELLIDO"],
    TELEFONO: ["TELEFONO", "TEL", "TEL."],
    "DIRECCION ENTREGA": [
      "DIRECCION DE FACTURACION / ENTREGA",
      "DIRECCION DE FACTURACIÓN / ENTREGA",
      "DIRECCIÓN DE FACTURACIÓN / ENTREGA",
      "DIRECCION ENTREGA",
    ],
    "CIUDAD ENTREGA": ["CIUDAD ENTREGA"],
    DEPARTAMENTO: ["DEPARTAMENTO", "DPTO ENTREGA"],
    "DIRECCION RECOGIDA": ["DIRECCION DE RECOGIDA", "DIRECCIÓN DE RECOGIDA", "DIRECCION RECOGIDA"],
    "CIUDAD RECOGIDA": ["CIUDAD RECOGIDA"],
    "DPTO RECOGIDA": ["DPTO RECOGIDA", "DEPARTAMENTO RECOGIDA"],
    COMBO: ["COMBO"],
    CAJAS: ["CAJAS", "CANT", "CANTIDAD"],
    CARRITO: ["CARRITO", "CARRITO "],
    "ROLLO PAPEL BURBUJA": ["ROLLO PAPEL BURBUJA"],
    "SOBRES x10": ["SOBRES x10", "SOBRES X10", "SOBRES PAPEL BURBUJA X 10 UNIDADES"],
    VINIPEL: ["VINIPEL"],
    "VERDES CNT40": ["VERDES CNT40", "CNT40"],
    PRECINTOS: ["PRECINTOS"],
    TAGS: ["TAGS"],
    OBSERVACIONES: ["OBSERVACIONES", "OBS."],
  };

  return rows.map((row) => {
    const sourceByHeader = {};
    Object.keys(row).forEach((key) => {
      sourceByHeader[normalizeHeader(key)] = key;
    });

    const normalizedRow = {};
    Object.entries(aliases).forEach(([canonical, options]) => {
      const sourceKey = options.map(normalizeHeader).map((option) => sourceByHeader[option]).find(Boolean);
      normalizedRow[canonical] = sourceKey ? row[sourceKey] : "";
    });
    return normalizedRow;
  });
}

function matchKey(pedido, cliente) {
  const cleanPedido = normalizeUpper(pedido);
  const cleanCliente = normalizeUpper(cliente).replace(/\s+/g, " ");
  return cleanPedido ? `${cleanPedido}||${cleanCliente}` : cleanCliente;
}

function pedidoKey(pedido) {
  return normalizeUpper(pedido);
}

function applyPreviousNote(item) {
  const exact = state.previousNotes.get(matchKey(item.pedido, item.cliente));
  const byPedido = state.previousNotes.get(pedidoKey(item.pedido));
  const byClient = state.previousNotes.get(matchKey("", item.cliente));
  const note = exact || byPedido || byClient;

  if (!note) return item;
  return {
    ...item,
    observaciones: note.observaciones || item.observaciones,
    estado: note.estado || item.estado,
  };
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function excelDateToIso(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value)) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }

  const text = normalize(value);
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return `${match[1]}-${match[2]}-${match[3]}`;
  match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return `${match[3]}-${String(match[2]).padStart(2, "0")}-${String(match[1]).padStart(2, "0")}`;

  const date = new Date(text);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDateLabel(value) {
  if (!value) return "-";
  const date = new Date(`${value}T00:00:00`);
  return new Intl.DateTimeFormat("es-CO", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
  }).format(date);
}

async function readWorkbook(event) {
  const file = event.target.files[0];
  if (!file) return;
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { cellDates: true });
  const sheetName = workbook.SheetNames.find((name) => normalizeUpper(name) === "DATOS");

  if (!sheetName) {
    alert("El archivo Excel debe tener una hoja llamada DATOS.");
    return;
  }

  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const rows = normalizeExcelRows(rawRows);
  const missing = requiredColumns.filter((column) => !(column in (rows[0] || {})));

  if (missing.length) {
    alert(`Al Excel le faltan estas columnas: ${missing.join(", ")}`);
    return;
  }

  state.rawRows = rows;
  generateRoute();
}

async function readPreviousRoute(event) {
  const file = event.target.files[0];
  if (!file) return;

  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  state.previousNotes = extractPreviousNotes(rows);

  if (!state.previousNotes.size) {
    alert("No encontre observaciones o estados en la ruta anterior.");
    return;
  }

  generateRoute();
  alert(`Ruta anterior cargada. Se encontraron ${state.previousNotes.size} registros con observaciones o estado.`);
}

function extractPreviousNotes(rows) {
  const notes = new Map();
  let header = null;

  rows.forEach((row) => {
    const upperRow = row.map((value) => normalizeHeader(value));
    const hasPedido = upperRow.includes("PEDIDO");
    const hasCliente = upperRow.includes("CLIENTE");
    const hasObservaciones = upperRow.includes("OBSERVACIONES") || upperRow.includes("OBS.");

    if (hasPedido && hasCliente && hasObservaciones) {
      header = {
        pedido: upperRow.findIndex((value) => value === "PEDIDO"),
        cliente: upperRow.findIndex((value) => value === "CLIENTE"),
        observaciones: upperRow.findIndex((value) => value === "OBSERVACIONES" || value === "OBS."),
        estado: upperRow.findIndex((value) => value === "ESTADO"),
      };
      return;
    }

    if (!header) return;

    const pedido = normalize(row[header.pedido]);
    const cliente = normalize(row[header.cliente]);
    const observaciones = header.observaciones >= 0 ? normalize(row[header.observaciones]) : "";
    const estado = header.estado >= 0 ? normalize(row[header.estado]) : "";

    if ((!pedido && !cliente) || (!observaciones && !estado)) return;

    const note = { observaciones, estado };
    notes.set(matchKey(pedido, cliente), note);
    if (pedido) notes.set(pedidoKey(pedido), note);
    if (cliente) notes.set(matchKey("", cliente), note);
  });

  return notes;
}

function mapItem(row, type) {
  const isDelivery = type === "ENTREGA";
  return {
    type,
    combo: normalize(row["COMBO"]),
    cant: toNumber(row["CAJAS"]),
    pedido: normalize(row["#PEDIDO"]),
    cliente: `${normalize(row["NOMBRE"])} ${normalize(row["APELLIDO"])}`.trim(),
    cnt30: "",
    cnt40: toNumber(row["VERDES CNT40"]),
    cnt40b: "",
    precintos: toNumber(row["PRECINTOS"]),
    tags: toNumber(row["TAGS"]),
    carro: toNumber(row["CARRITO"]),
    rollo: toNumber(row["ROLLO PAPEL BURBUJA"]),
    vinipel: toNumber(row["VINIPEL"]),
    sobres: toNumber(row["SOBRES x10"]),
    direccion: normalize(row[isDelivery ? "DIRECCION ENTREGA" : "DIRECCION RECOGIDA"]),
    ciudad: normalize(row[isDelivery ? "CIUDAD ENTREGA" : "CIUDAD RECOGIDA"]),
    telefono: normalize(row["TELEFONO"]),
    observaciones: normalize(row["OBSERVACIONES"]),
    estado: "",
  };
}

function generateRoute() {
  if (!state.rawRows.length) {
    renderEmpty();
    return;
  }

  const selectedDate = document.getElementById("routeDate").value;
  const department = normalizeUpper(document.getElementById("departmentFilter").value);

  state.deliveries = state.rawRows
    .filter((row) => excelDateToIso(row["FECHA DE ENTREGA"]) === selectedDate)
    .filter((row) => normalize(row["COMBO"]))
    .filter((row) => isAllowedOrderType(row))
    .filter((row) => !department || normalizeUpper(row["DEPARTAMENTO"]) === department)
    .map((row) => applyPreviousNote(mapItem(row, "ENTREGA")));

  state.pickups = state.rawRows
    .filter((row) => excelDateToIso(row["FECHA DE RECOGIDA"]) === selectedDate)
    .filter((row) => normalize(row["COMBO"]))
    .filter((row) => isAllowedOrderType(row))
    .filter((row) => !department || normalizeUpper(row["DPTO RECOGIDA"]) === department)
    .map((row) => applyPreviousNote(mapItem(row, "RECOGIDA")));

  document.getElementById("deliveryDateLabel").textContent = formatDateLabel(selectedDate);
  document.getElementById("pickupDateLabel").textContent = formatDateLabel(selectedDate);
  document.getElementById("deliveriesCount").textContent = state.deliveries.length;
  document.getElementById("pickupsCount").textContent = state.pickups.length;
  document.getElementById("totalCount").textContent = state.deliveries.length + state.pickups.length;

  renderTable(deliveriesTable, state.deliveries, true);
  renderTable(pickupsTable, state.pickups, false);
  const hasResults = state.deliveries.length + state.pickups.length > 0;
  downloadExcelButton.disabled = !hasResults;
  downloadPdfButton.disabled = !hasResults;
}

function renderEmpty() {
  deliveriesTable.innerHTML = `<tr><td class="empty-state">Sube tu base de datos Excel para generar entregas.</td></tr>`;
  pickupsTable.innerHTML = `<tr><td class="empty-state">Sube tu base de datos Excel para generar recogidas.</td></tr>`;
}

function renderTable(table, items, includeSupplies) {
  const headers = includeSupplies
    ? ["No.", "COMBO", "CANT", "PEDIDO", "CLIENTE", "CNT30", "CNT40", "CNT40", "PRECINTOS", "TAGS", "CARRO", "ROLLO PAPEL BURBUJA", "VINIPEL", "SOBRES PAPEL BURBUJA X 10 UNIDADES", "OBSERVACIONES"]
    : ["No.", "COMBO", "CANT", "PEDIDO", "CLIENTE", "CNT30", "CNT40", "CNT40", "CARRO", "OBSERVACIONES", "Estado"];

  if (!items.length) {
    table.innerHTML = `<tr><td class="empty-state">No hay registros para esta fecha y departamento.</td></tr>`;
    return;
  }

  table.innerHTML = `
    <thead>
      <tr>${headers.map((header) => `<th class="${headerClass(header)}">${header}</th>`).join("")}</tr>
    </thead>
    <tbody>
      ${items.map((item, index) => routeRows(item, index + 1, includeSupplies)).join("")}
    </tbody>
  `;
}

function headerClass(header) {
  if (header === "CNT30" || header === "CNT40") return "blue";
  if (header === "PRECINTOS" || header === "TAGS") return "green";
  if (["CARRO", "ROLLO PAPEL BURBUJA", "VINIPEL", "SOBRES PAPEL BURBUJA X 10 UNIDADES", "Estado"].includes(header)) return "yellow";
  return "";
}

function routeRows(item, number, includeSupplies) {
  const main = includeSupplies
    ? [number, item.combo, item.cant, item.pedido, item.cliente, item.cnt30, item.cnt40, item.cnt40b, item.precintos, item.tags, item.carro, item.rollo, item.vinipel, item.sobres, item.observaciones]
    : [number, item.combo, item.cant, item.pedido, item.cliente, item.cnt30, item.cnt40, item.cnt40b, item.carro, item.observaciones, item.estado];
  const span = includeSupplies ? 15 : 11;
  return `
    <tr>${main.map((value, index) => editableCell(value, index, includeSupplies, number - 1)).join("")}</tr>
    <tr class="address-row">
      <td class="label">DIRECCION:</td>
      <td colspan="${includeSupplies ? 4 : 4}">${item.direccion}</td>
      <td class="label">CIUDAD:</td>
      <td>${item.ciudad}</td>
      <td class="label">TEL:</td>
      <td colspan="${Math.max(span - 8, 1)}">${item.telefono}</td>
    </tr>
  `;
}

function editableCell(value, index, includeSupplies, itemIndex) {
  const obsIndex = includeSupplies ? 14 : 9;
  const estadoIndex = includeSupplies ? -1 : 10;
  const field = index === obsIndex ? "observaciones" : index === estadoIndex ? "estado" : "";
  const editable = field ? `contenteditable="true" data-field="${field}" data-index="${itemIndex}" data-section="${includeSupplies ? "deliveries" : "pickups"}"` : "";
  return `<td ${editable} class="${cellClass(index, includeSupplies)} ${index === 4 ? "client" : ""} ${field ? "editable" : ""}">${value}</td>`;
}

function cellClass(index, includeSupplies) {
  if ([5, 7].includes(index)) return "blue";
  if (includeSupplies && [6, 8, 9].includes(index)) return "green";
  if (includeSupplies && index >= 10 && index <= 13) return "yellow";
  if (!includeSupplies && [6].includes(index)) return "green";
  if (!includeSupplies && [8, 10].includes(index)) return "yellow";
  return "";
}

function syncEditsFromTables() {
  document.querySelectorAll("[contenteditable][data-field]").forEach((cell) => {
    const section = cell.dataset.section;
    const index = Number(cell.dataset.index);
    const field = cell.dataset.field;
    if (!state[section] || !state[section][index]) return;
    state[section][index][field] = cell.textContent.trim();
  });
}

async function downloadExcel() {
  syncEditsFromTables();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Ruta");
  worksheet.views = [{ showGridLines: false }];
  worksheet.properties.defaultRowHeight = 30;
  worksheet.pageSetup = {
    orientation: "landscape",
    paperSize: 5,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.2, right: 0.2, top: 0.25, bottom: 0.25, header: 0.1, footer: 0.1 },
  };
  const widths = [8, 11, 9, 12, 34, 9, 9, 9, 12, 11, 12, 18, 14, 22, 24];
  worksheet.columns = widths.map((width) => ({ width }));

  let rowNumber = addExcelSection(worksheet, 1, "PROGRAMACION DE ENTREGAS", state.deliveries, true);
  addExcelSection(worksheet, rowNumber + 3, "PROGRAMACION DE RECOGIDAS", state.pickups, false);

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), `programacion-rutas-${document.getElementById("routeDate").value}.xlsx`);
}

function addExcelSection(worksheet, startRow, title, items, includeSupplies) {
  const maxCol = includeSupplies ? 15 : 11;
  const dateLabel = formatDateLabel(document.getElementById("routeDate").value);
  const headerRow = startRow + 2;
  const firstItemRow = startRow + 3;

  worksheet.mergeCells(startRow, 1, startRow, maxCol);
  worksheet.getCell(startRow, 1).value = title;
  worksheet.getRow(startRow).height = 28;
  styleBlock(worksheet, startRow, 1, startRow, maxCol, {
    fill: "8BD14F",
    bold: true,
    size: 13,
    horizontal: "center",
  });

  worksheet.getCell(startRow + 1, 1).value = "FECHA:";
  worksheet.getCell(startRow + 1, 2).value = dateLabel;
  worksheet.mergeCells(startRow + 1, 2, startRow + 1, maxCol);
  worksheet.getRow(startRow + 1).height = 30;
  styleBlock(worksheet, startRow + 1, 1, startRow + 1, maxCol, {
    fill: "FFFFFF",
    bold: true,
    size: 12,
    horizontal: "left",
  });
  worksheet.getCell(startRow + 1, 2).font = { bold: true, italic: true, size: 12 };

  const headers = includeSupplies
    ? ["No.", "COMBO", "CANT", "PEDIDO", "CLIENTE", "CNT30", "CNT40", "CNT40", "PRECINTOS", "TAGS", "CARRO", "ROLLO PAPEL BURBUJA", "VINIPEL", "SOBRES PAPEL BURBUJA X 10 UNIDADES", "OBSERVACIONES"]
    : ["No.", "COMBO", "CANT", "PEDIDO", "CLIENTE", "CNT30", "CNT40", "CNT40", "CARRO", "OBSERVACIONES", "Estado"];
  worksheet.getRow(headerRow).values = headers;
  worksheet.getRow(headerRow).height = 52;
  styleBlock(worksheet, headerRow, 1, headerRow, maxCol, {
    fill: "D0CECE",
    bold: true,
    size: 10,
    horizontal: "center",
  });
  headers.forEach((header, index) => {
    const column = index + 1;
    const fill = excelHeaderFill(header);
    if (fill) worksheet.getCell(headerRow, column).fill = solidFill(fill);
  });

  let currentRow = firstItemRow;
  items.forEach((item, index) => {
    const row = includeSupplies
      ? [index + 1, item.combo, item.cant, item.pedido, item.cliente, item.cnt30, item.cnt40, item.cnt40b, item.precintos, item.tags, item.carro, item.rollo, item.vinipel, item.sobres, item.observaciones]
      : [index + 1, item.combo, item.cant, item.pedido, item.cliente, item.cnt30, item.cnt40, item.cnt40b, item.carro, item.observaciones, item.estado];
    worksheet.getRow(currentRow).values = row;
    worksheet.getRow(currentRow).height = 42;
    styleBlock(worksheet, currentRow, 1, currentRow, maxCol, {
      fill: "FFFFFF",
      bold: true,
      size: 10,
      horizontal: "center",
    });
    row.forEach((_, cellIndex) => {
      const fill = excelBodyFill(cellIndex, includeSupplies);
      if (fill) worksheet.getCell(currentRow, cellIndex + 1).fill = solidFill(fill);
    });

    worksheet.getCell(currentRow + 1, 1).value = "DIRECCION:";
    worksheet.getCell(currentRow + 1, 2).value = item.direccion;
    worksheet.mergeCells(currentRow + 1, 2, currentRow + 1, 5);
    worksheet.getCell(currentRow + 1, 6).value = "CIUDAD:";
    worksheet.getCell(currentRow + 1, 7).value = item.ciudad;
    worksheet.getCell(currentRow + 1, 8).value = "TEL:";
    worksheet.getCell(currentRow + 1, 9).value = item.telefono;
    worksheet.mergeCells(currentRow + 1, 9, currentRow + 1, maxCol);
    worksheet.getRow(currentRow + 1).height = 42;
    styleBlock(worksheet, currentRow + 1, 1, currentRow + 1, maxCol, {
      fill: "FFFFFF",
      bold: true,
      size: 10,
      horizontal: "left",
    });
    [1, 6, 8].forEach((column) => {
      worksheet.getCell(currentRow + 1, column).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    });
    currentRow += 2;
  });

  if (!items.length) {
    worksheet.mergeCells(firstItemRow, 1, firstItemRow + 1, maxCol);
    worksheet.getCell(firstItemRow, 1).value = "No hay registros para esta fecha y departamento.";
    worksheet.getRow(firstItemRow).height = 54;
    styleBlock(worksheet, firstItemRow, 1, firstItemRow + 1, maxCol, {
      fill: "FFFFFF",
      size: 10,
      horizontal: "center",
    });
    currentRow = firstItemRow + 2;
  }

  return currentRow;
}

function solidFill(color) {
  const argb = color.length === 6 ? `FF${color}` : color;
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function thinBorder() {
  return {
    top: { style: "thin", color: { argb: "7A7A7A" } },
    left: { style: "thin", color: { argb: "7A7A7A" } },
    bottom: { style: "thin", color: { argb: "7A7A7A" } },
    right: { style: "thin", color: { argb: "7A7A7A" } },
  };
}

function styleBlock(worksheet, startRow, startCol, endRow, endCol, options = {}) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      const cell = worksheet.getCell(row, col);
      cell.fill = solidFill(options.fill || "FFFFFF");
      cell.font = {
        bold: Boolean(options.bold),
        italic: Boolean(options.italic),
        size: options.size || 10,
        color: { argb: "FF102033" },
      };
      cell.alignment = {
        horizontal: options.horizontal || "center",
        vertical: "middle",
        wrapText: true,
      };
      cell.border = thinBorder();
    }
  }
}

function excelHeaderFill(header) {
  if (header === "CNT30" || header === "CNT40") return "BDD7EE";
  if (header === "PRECINTOS" || header === "TAGS") return "A9D18E";
  if (["CARRO", "ROLLO PAPEL BURBUJA", "VINIPEL", "SOBRES PAPEL BURBUJA X 10 UNIDADES", "Estado"].includes(header)) return "FFE699";
  return "D0CECE";
}

function excelBodyFill(index, includeSupplies) {
  if ([5, 7].includes(index)) return "BDD7EE";
  if (includeSupplies && [6, 8, 9].includes(index)) return "A9D18E";
  if (includeSupplies && index >= 10 && index <= 13) return "FFE699";
  if (!includeSupplies && index === 6) return "A9D18E";
  if (!includeSupplies && [8, 10].includes(index)) return "FFE699";
  return "";
}

async function downloadPdf() {
  syncEditsFromTables();
  const { jsPDF } = window.jspdf;
  const source = document.querySelector(".sheet-preview");
  const canvas = await html2canvas(source, {
    backgroundColor: "#ffffff",
    scale: 2,
    width: source.scrollWidth,
    height: source.scrollHeight,
    windowWidth: Math.max(document.documentElement.clientWidth, source.scrollWidth),
    windowHeight: Math.max(document.documentElement.clientHeight, source.scrollHeight),
  });
  const imgData = canvas.toDataURL("image/png");
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "legal" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 18;
  const imgWidth = pageWidth - margin * 2;
  const imgHeight = (canvas.height * imgWidth) / canvas.width;
  let heightLeft = imgHeight;
  let position = margin;

  doc.addImage(imgData, "PNG", margin, position, imgWidth, imgHeight);
  heightLeft -= pageHeight - margin * 2;

  while (heightLeft > 0) {
    doc.addPage();
    position = margin - (imgHeight - heightLeft);
    doc.addImage(imgData, "PNG", margin, position, imgWidth, imgHeight);
    heightLeft -= pageHeight - margin * 2;
  }

  doc.save(`programacion-rutas-${document.getElementById("routeDate").value}.pdf`);
}

renderEmpty();
