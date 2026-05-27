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
    .replace(/[^A-Za-z0-9_. -]+/g, "_")
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
  return "output_lich_tv360.xlsx";
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
    const response = await fetch("/api/preview", {
      method: "POST",
      body: makeFormData(),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Không thể xem trước dữ liệu.");
    }
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

async function responseErrorMessage(response, fallback) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const payload = await response.json();
    return payload.error || fallback;
  }

  const text = await response.text();
  return text.trim() || fallback;
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

async function exportExcel() {
  const files = selectedFiles();
  if (!files.length) {
    setStatus("Hãy chọn ít nhất một file .txt hoặc .xlsx.", "error");
    return;
  }

  exportButton.disabled = true;
  previewButton.disabled = true;
  const requestedName = outputNameForFiles(files);
  outputNameInput.value = requestedName;
  setStatus(`Đang tạo ${requestedName}...`);

  try {
    const response = await fetch(`/api/export?name=${encodeURIComponent(requestedName)}`, {
      method: "POST",
      body: makeFormData(files),
    });

    if (!response.ok) {
      throw new Error(await responseErrorMessage(response, "Không thể tạo file Excel."));
    }

    const blob = await response.blob();
    downloadBlob(blob, requestedName);
    setStatus(`Đã tạo ${requestedName}.`, "ok");
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
