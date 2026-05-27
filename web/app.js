const form = document.querySelector("#upload-form");
const fileInput = document.querySelector("#files");
const sortInput = document.querySelector("#sort");
const minimalInput = document.querySelector("#minimal");
const correctionsInput = document.querySelector("#corrections");
const outputNameInput = document.querySelector("#output-name");
const previewButton = document.querySelector("#preview-button");
const exportButton = document.querySelector("#export-button");
const clearFilesButton = document.querySelector("#clear-files-button");
const fileCount = document.querySelector("#file-count");
const rowCount = document.querySelector("#row-count");
const statusText = document.querySelector("#status");
const results = document.querySelector("#results");
const selectedFilesList = document.querySelector("#selected-files");

const selectedFileStore = [];
let previewCache = null;

function selectedFiles() {
  return selectedFileStore;
}

function fileKey(file) {
  return `${file.name}-${file.size}-${file.lastModified}`;
}

function baseName(filename) {
  return filename.replace(/\.[^.]+$/, "");
}

function safeOutputPart(value) {
  return value
    .normalize("NFC")
    .replace(/[<>:"/\\|?*\u0000-\u001f]+/g, "_")
    .replace(/^[ ._]+|[ ._]+$/g, "") || "lich_tv360";
}

function outputNameForFile(file) {
  return `output_${safeOutputPart(baseName(file.name))}.xlsx`;
}

function outputNameForFiles(files) {
  if (!files.length) {
    return "output_lich_tv360.xlsx";
  }
  if (files.length === 1) {
    return outputNameForFile(files[0]);
  }
  return `${files.length} file export riêng`;
}

function formatFileSize(size) {
  if (size < 1024) {
    return `${size} B`;
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`;
  }
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function setStatus(message, tone = "") {
  statusText.textContent = message;
  statusText.className = tone;
}

function currentPreviewKey(files = selectedFiles()) {
  return [
    files.map(fileKey).join("|"),
    sortInput.checked ? "sort" : "input-order",
    correctionsInput.checked ? "corrections" : "raw-title",
  ].join("::");
}

function makeFormData(files = selectedFiles()) {
  const data = new FormData();
  for (const file of files) {
    data.append("files", file);
  }
  data.append("sort", sortInput.checked ? "true" : "false");
  data.append("minimal", minimalInput.checked ? "true" : "false");
  data.append("corrections", correctionsInput.checked ? "true" : "false");
  return data;
}

function updateFileCount() {
  fileCount.textContent = String(selectedFiles().length);
  outputNameInput.value = outputNameForFiles(selectedFiles());
  clearFilesButton.disabled = !selectedFiles().length;
}

function renderSelectedFiles() {
  selectedFilesList.textContent = "";
  if (!selectedFiles().length) {
    const item = document.createElement("li");
    item.className = "empty-file";
    item.textContent = "Chưa chọn file nào.";
    selectedFilesList.appendChild(item);
    return;
  }

  selectedFiles().forEach((file, index) => {
    const item = document.createElement("li");

    const details = document.createElement("span");
    details.className = "file-details";

    const name = document.createElement("strong");
    name.textContent = file.name;

    const meta = document.createElement("span");
    meta.textContent = formatFileSize(file.size);

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "file-remove";
    removeButton.textContent = "Xóa";
    removeButton.addEventListener("click", () => {
      selectedFileStore.splice(index, 1);
      resetParsedState();
      renderSelectedFiles();
      updateFileCount();
      setStatus(selectedFiles().length ? "Danh sách file đã được cập nhật." : "Chưa có dữ liệu.");
    });

    details.append(name, meta);
    item.append(details, removeButton);
    selectedFilesList.appendChild(item);
  });
}

function addFiles(files) {
  const allowedExtensions = /\.(txt|xlsx)$/i;
  const existingKeys = new Set(selectedFileStore.map(fileKey));
  let added = 0;

  for (const file of files) {
    if (!allowedExtensions.test(file.name)) {
      continue;
    }
    const key = fileKey(file);
    if (existingKeys.has(key)) {
      continue;
    }
    selectedFileStore.push(file);
    existingKeys.add(key);
    added += 1;
  }

  fileInput.value = "";
  resetParsedState();
  renderSelectedFiles();
  updateFileCount();

  if (added) {
    setStatus(`Đã thêm ${added} file.`, "ok");
  } else if (files.length) {
    setStatus("Không có file .txt hoặc .xlsx mới để thêm.", "error");
  } else {
    setStatus("Chưa có dữ liệu.");
  }
}

function resetParsedState() {
  previewCache = null;
  rowCount.textContent = "0";
  results.textContent = "";
  const row = document.createElement("tr");
  const cell = document.createElement("td");
  cell.colSpan = 6;
  cell.className = "empty";
  cell.textContent = "Chọn file rồi bấm Xem trước.";
  row.appendChild(cell);
  results.appendChild(row);
}

function renderRows(rows) {
  results.textContent = "";
  if (!rows.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 6;
    cell.className = "empty";
    cell.textContent = "Không parse được chương trình nào từ file đã chọn.";
    row.appendChild(cell);
    results.appendChild(row);
    return;
  }

  for (const item of rows) {
    const row = document.createElement("tr");
    const values = [item.stt, item.schedule_day || "", item.title, item.airtime, item.source, item.source_line];
    for (const value of values) {
      const cell = document.createElement("td");
      cell.textContent = String(value);
      row.appendChild(cell);
    }
    results.appendChild(row);
  }
}

async function responseErrorMessage(response, fallback) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const payload = await response.json();
    return payload.error || fallback;
  }

  const text = await response.text();
  return text.trim() || fallback;
}

async function loadPreviewRows(files = selectedFiles()) {
  const cacheKey = currentPreviewKey(files);
  if (previewCache && previewCache.key === cacheKey) {
    return previewCache.payload;
  }

  const response = await fetch("/api/preview", {
    method: "POST",
    body: makeFormData(files),
  });
  if (!response.ok) {
    throw new Error(await responseErrorMessage(response, "Không thể xem trước dữ liệu."));
  }

  const payload = await response.json();
  previewCache = { key: cacheKey, payload };
  return payload;
}

async function preview() {
  const files = selectedFiles();
  if (!files.length) {
    setStatus("Hãy chọn ít nhất một file .txt hoặc .xlsx.", "error");
    return;
  }

  previewButton.disabled = true;
  exportButton.disabled = true;
  setStatus("Đang đọc và chuẩn hóa dữ liệu...");

  try {
    const payload = await loadPreviewRows(files);
    rowCount.textContent = String(payload.count);
    renderRows(payload.rows);
    setStatus(`Đã parse ${payload.count} chương trình từ ${files.length} file.`, "ok");
  } catch (error) {
    rowCount.textContent = "0";
    renderRows([]);
    setStatus(error.message, "error");
  } finally {
    previewButton.disabled = false;
    exportButton.disabled = false;
  }
}

function downloadBlob(blob, filename) {
  if (window.navigator.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(blob, filename);
    return;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.rel = "noopener";
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function columnName(index) {
  let result = "";
  let current = index;
  while (current) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

function parseScheduleDate(value) {
  const match = String(value || "").match(/(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?/);
  if (!match) {
    return null;
  }
  let year = match[3] ? Number(match[3]) : new Date().getFullYear();
  if (year < 100) {
    year += 2000;
  }
  const date = new Date(year, Number(match[2]) - 1, Number(match[1]));
  if (date.getFullYear() !== year || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[1])) {
    return null;
  }
  return date;
}

function parseAirtime(value) {
  const match = String(value || "").trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!match) {
    return null;
  }
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour > 23 || minute > 59) {
    return null;
  }
  return { hour, minute };
}

function rowScheduleDates(records) {
  const firstDate = records.map((item) => parseScheduleDate(item.schedule_day)).find(Boolean) || new Date();
  let currentDate = firstDate;
  return records.map((item) => {
    const parsed = parseScheduleDate(item.schedule_day);
    if (parsed) {
      currentDate = parsed;
    }
    return new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());
  });
}

function sameDate(left, right) {
  return left.getFullYear() === right.getFullYear() && left.getMonth() === right.getMonth() && left.getDate() === right.getDate();
}

function addDays(value, days) {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate() + days, value.getHours(), value.getMinutes(), value.getSeconds());
}

function metadataStartEndDates(records) {
  const dates = rowScheduleDates(records);
  const starts = [];
  const explicitEnds = [];

  records.forEach((item, index) => {
    const [startText, endText] = String(item.airtime || "").split("-", 2).map((part) => part.trim());
    const startParts = parseAirtime(startText) || { hour: 0, minute: 0 };
    const start = new Date(dates[index].getFullYear(), dates[index].getMonth(), dates[index].getDate(), startParts.hour, startParts.minute, 0);
    starts.push(start);

    const endParts = parseAirtime(endText);
    if (endParts) {
      let end = new Date(dates[index].getFullYear(), dates[index].getMonth(), dates[index].getDate(), endParts.hour, endParts.minute, 0);
      if (end <= start) {
        end = addDays(end, 1);
      }
      explicitEnds.push(end);
    } else {
      explicitEnds.push(null);
    }
  });

  return records.map((item, index) => {
    const start = starts[index];
    let end = explicitEnds[index];
    if (!end && starts[index + 1] && sameDate(dates[index + 1], dates[index]) && starts[index + 1] > start) {
      end = starts[index + 1];
    }
    if (!end) {
      end = new Date(dates[index].getFullYear(), dates[index].getMonth(), dates[index].getDate(), 23, 59, 59);
    }
    return { start, end };
  });
}

function tv360DateTime(value) {
  const pad = (number, length = 2) => String(number).padStart(length, "0");
  return `${value.getFullYear()}${pad(value.getMonth() + 1)}${pad(value.getDate())}${pad(value.getHours())}${pad(value.getMinutes())}${pad(value.getSeconds())}`;
}

function makeCell(row, col, value, style = 0) {
  const ref = `${columnName(col)}${row}`;
  const styleAttr = style ? ` s="${style}"` : "";
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}"${styleAttr}><v>${value}</v></c>`;
  }
  return `<c r="${ref}" t="inlineStr"${styleAttr}><is><t>${xmlEscape(value)}</t></is></c>`;
}

function makeSheetXml(rows, widths) {
  const xmlRows = rows.map((values, rowIndex) => {
    const rowNumber = rowIndex + 1;
    const cells = values.map((value, colIndex) => makeCell(rowNumber, colIndex + 1, value, rowNumber === 1 ? 1 : 0));
    return `<row r="${rowNumber}">${cells.join("")}</row>`;
  });
  const cols = widths.map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`);
  const lastCol = columnName(widths.length);
  const lastRow = Math.max(rows.length, 1);
  const autoFilter = lastRow > 1 ? `<autoFilter ref="A1:${lastCol}${lastRow}"/>` : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <cols>${cols.join("")}</cols>
  <sheetData>${xmlRows.join("")}</sheetData>
  ${autoFilter}
</worksheet>`;
}

function workbookRows(records, minimal) {
  if (minimal) {
    const rows = [["STT", "Thứ ngày", "Tiêu đề chương trình", "Thời gian phát sóng"]];
    records.forEach((item, index) => {
      rows.push([index + 1, item.schedule_day || "", item.title || "", item.airtime || ""]);
    });
    return rows;
  }

  const headers = [
    "Main Title",
    "Main Language",
    "Start Time",
    "End Time",
    "Main Synopsis",
    "Rating",
    "Video Type",
    "Director",
    "Actor",
    "Price",
    "Fx Point",
    "Series Key",
    "Episode Key",
    "Is Last Episode",
    "Poster Url",
    "VOD AssetID",
    "ProductID",
    "CPIP",
  ];
  const rows = [headers];
  for (let index = 0; index < 17; index += 1) {
    rows.push(Array(headers.length).fill(""));
  }
  const ranges = metadataStartEndDates(records);
  records.forEach((item, index) => {
    const title = item.title || "";
    rows.push([
      title,
      "vie",
      tv360DateTime(ranges[index].start),
      tv360DateTime(ranges[index].end),
      title,
      "0",
      "HD",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
    ]);
  });
  return rows;
}

function xlsxFiles(rows, minimal) {
  const widths = minimal ? [8, 22, 46, 22] : [36, 16, 20, 20, 46, 10, 12, 20, 20, 10, 12, 18, 18, 16, 28, 16, 16, 12];
  const sheetName = minimal ? "Lich phat song" : "Metadata";
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");

  return {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
    "docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>TV360 Schedule Formatter</Application>
</Properties>`,
    "docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>TV360 Schedule Formatter</dc:creator>
  <dc:title>Lịch phát sóng TV360</dc:title>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`,
    "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${sheetName}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    "xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`,
    "xl/worksheets/sheet1.xml": makeSheetXml(rows, widths),
  };
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date) {
  const year = Math.max(date.getFullYear(), 1980);
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}

function writeUint16(output, value) {
  output.push(value & 0xff, (value >>> 8) & 0xff);
}

function writeUint32(output, value) {
  output.push(value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff);
}

function concatBytes(parts) {
  const totalLength = parts.reduce((sum, part) => sum + part.length, 0);
  const output = new Uint8Array(totalLength);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function zipBytes(files) {
  const encoder = new TextEncoder();
  const now = dosDateTime(new Date());
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const dataBytes = encoder.encode(content);
    const checksum = crc32(dataBytes);
    const localHeader = [];
    writeUint32(localHeader, 0x04034b50);
    writeUint16(localHeader, 20);
    writeUint16(localHeader, 0x0800);
    writeUint16(localHeader, 0);
    writeUint16(localHeader, now.time);
    writeUint16(localHeader, now.date);
    writeUint32(localHeader, checksum);
    writeUint32(localHeader, dataBytes.length);
    writeUint32(localHeader, dataBytes.length);
    writeUint16(localHeader, nameBytes.length);
    writeUint16(localHeader, 0);

    const centralHeader = [];
    writeUint32(centralHeader, 0x02014b50);
    writeUint16(centralHeader, 20);
    writeUint16(centralHeader, 20);
    writeUint16(centralHeader, 0x0800);
    writeUint16(centralHeader, 0);
    writeUint16(centralHeader, now.time);
    writeUint16(centralHeader, now.date);
    writeUint32(centralHeader, checksum);
    writeUint32(centralHeader, dataBytes.length);
    writeUint32(centralHeader, dataBytes.length);
    writeUint16(centralHeader, nameBytes.length);
    writeUint16(centralHeader, 0);
    writeUint16(centralHeader, 0);
    writeUint16(centralHeader, 0);
    writeUint16(centralHeader, 0);
    writeUint32(centralHeader, 0);
    writeUint32(centralHeader, offset);

    const localHeaderBytes = Uint8Array.from(localHeader);
    const centralHeaderBytes = Uint8Array.from(centralHeader);
    localParts.push(localHeaderBytes, nameBytes, dataBytes);
    centralParts.push(centralHeaderBytes, nameBytes);
    offset += localHeaderBytes.length + nameBytes.length + dataBytes.length;
  }

  const centralDirectory = concatBytes(centralParts);
  const endRecord = [];
  writeUint32(endRecord, 0x06054b50);
  writeUint16(endRecord, 0);
  writeUint16(endRecord, 0);
  writeUint16(endRecord, Object.keys(files).length);
  writeUint16(endRecord, Object.keys(files).length);
  writeUint32(endRecord, centralDirectory.length);
  writeUint32(endRecord, offset);
  writeUint16(endRecord, 0);

  return concatBytes([...localParts, centralDirectory, Uint8Array.from(endRecord)]);
}

function buildXlsxBlob(records, minimal) {
  const rows = workbookRows(records, minimal);
  const bytes = zipBytes(xlsxFiles(rows, minimal));
  return new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

async function exportExcel() {
  const files = selectedFiles();
  if (!files.length) {
    setStatus("Hãy chọn ít nhất một file .txt hoặc .xlsx.", "error");
    return;
  }

  exportButton.disabled = true;
  previewButton.disabled = true;
  outputNameInput.value = outputNameForFiles(files);
  setStatus(files.length === 1 ? `Đang tạo ${outputNameForFile(files[0])}...` : `Đang tạo ${files.length} file export riêng...`);

  try {
    const allRows = [];
    let totalCount = 0;

    for (const file of files) {
      const payload = await loadPreviewRows([file]);
      totalCount += payload.count;
      allRows.push(...payload.rows);

      const filename = outputNameForFile(file);
      const blob = buildXlsxBlob(payload.rows, minimalInput.checked);
      downloadBlob(blob, filename);
    }

    rowCount.textContent = String(totalCount);
    renderRows(allRows.map((row, index) => ({ ...row, stt: index + 1 })));
    setStatus(files.length === 1 ? `Đã tạo ${outputNameForFile(files[0])}.` : `Đã tạo ${files.length} file export riêng.`, "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    exportButton.disabled = false;
    previewButton.disabled = false;
  }
}

fileInput.addEventListener("change", () => {
  addFiles(Array.from(fileInput.files || []));
});

const dropZone = document.querySelector(".drop-zone");
dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("is-dragging");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("is-dragging");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("is-dragging");
  addFiles(Array.from(event.dataTransfer.files || []));
});

clearFilesButton.addEventListener("click", () => {
  selectedFileStore.length = 0;
  fileInput.value = "";
  resetParsedState();
  renderSelectedFiles();
  updateFileCount();
  setStatus("Chưa có dữ liệu.");
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  preview();
});

exportButton.addEventListener("click", exportExcel);
renderSelectedFiles();
updateFileCount();
