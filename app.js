const fileInput = document.querySelector('#file-input');
const dropZone = document.querySelector('.drop-zone');
const fileList = document.querySelector('#file-list');
const mergeButton = document.querySelector('#merge-button');
const resetButton = document.querySelector('#reset-button');
const exportButton = document.querySelector('#export-button');
const statusText = document.querySelector('#status');
const statsTableBody = document.querySelector('#stats-table tbody');
const previewTable = document.querySelector('#preview-table');
const statFiles = document.querySelector('#stat-files');
const statRows = document.querySelector('#stat-rows');
const statColumns = document.querySelector('#stat-columns');
const statNumeric = document.querySelector('#stat-numeric');

let selectedFiles = [];
let mergedRows = [];
let summaryRows = [];
let allHeaders = [];

const SOURCE_FILE = 'Source File';
const SOURCE_SHEET = 'Source Sheet';

fileInput.addEventListener('change', (event) => {
  setSelectedFiles([...event.target.files]);
});

['dragenter', 'dragover'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add('drag-over');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove('drag-over');
  });
});

dropZone.addEventListener('drop', (event) => {
  const files = [...event.dataTransfer.files].filter(isSpreadsheetFile);
  fileInput.files = createFileList(files);
  setSelectedFiles(files);
});

mergeButton.addEventListener('click', mergeFiles);
resetButton.addEventListener('click', resetApp);
exportButton.addEventListener('click', exportWorkbook);

function setSelectedFiles(files) {
  selectedFiles = files.filter(isSpreadsheetFile);
  fileList.innerHTML = '';

  selectedFiles.forEach((file) => {
    const item = document.createElement('li');
    item.innerHTML = `<span>${escapeHtml(file.name)}</span><span class="file-size">${formatBytes(file.size)}</span>`;
    fileList.appendChild(item);
  });

  mergeButton.disabled = selectedFiles.length === 0;
  resetButton.disabled = selectedFiles.length === 0 && mergedRows.length === 0;
  statusText.textContent = selectedFiles.length
    ? `${selectedFiles.length} file${selectedFiles.length === 1 ? '' : 's'} ready to merge.`
    : 'Waiting for files.';
}

async function mergeFiles() {
  if (!window.XLSX) {
    statusText.textContent = 'SheetJS could not be loaded. Check your internet connection or vendor xlsx.full.min.js locally.';
    return;
  }

  mergeButton.disabled = true;
  exportButton.disabled = true;
  statusText.textContent = 'Reading workbooks locally in your browser...';

  try {
    const parsedReports = (await Promise.all(selectedFiles.map(readWorkbook))).flat();
    const headers = new Set([SOURCE_FILE, SOURCE_SHEET]);
    const rows = [];

    parsedReports.forEach(({ fileName, sheetName, data }) => {
      data.forEach((row) => {
        Object.keys(row).forEach((header) => headers.add(header));
        rows.push({ [SOURCE_FILE]: fileName, [SOURCE_SHEET]: sheetName, ...row });
      });
    });

    allHeaders = [...headers];
    mergedRows = rows.map((row) => normalizeRow(row, allHeaders));
    summaryRows = buildSummaryRows(mergedRows, allHeaders);

    renderStats(summaryRows);
    renderPreview(mergedRows, allHeaders);
    updateStatCards();

    exportButton.disabled = mergedRows.length === 0;
    statusText.textContent = mergedRows.length
      ? `Merged ${mergedRows.length.toLocaleString()} row${mergedRows.length === 1 ? '' : 's'} from ${selectedFiles.length} file${selectedFiles.length === 1 ? '' : 's'}.`
      : 'No data rows were found in the selected workbooks.';
  } catch (error) {
    console.error(error);
    statusText.textContent = `Unable to merge files: ${error.message}`;
  } finally {
    mergeButton.disabled = selectedFiles.length === 0;
    resetButton.disabled = selectedFiles.length === 0 && mergedRows.length === 0;
  }
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true });
        if (workbook.SheetNames.length === 0) {
          resolve([{ fileName: file.name, sheetName: 'No sheets', data: [] }]);
          return;
        }

        const sheets = workbook.SheetNames.map((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
          return { fileName: file.name, sheetName, data: dedupeObjectKeys(data) };
        });

        resolve(sheets);
      } catch (error) {
        reject(new Error(`${file.name}: ${error.message}`));
      }
    };

    reader.onerror = () => reject(new Error(`${file.name}: ${reader.error?.message || 'file read failed'}`));
    reader.readAsArrayBuffer(file);
  });
}

function buildSummaryRows(rows, headers) {
  return headers.map((header) => {
    const values = rows.map((row) => row[header]);
    const nonEmptyValues = values.filter((value) => String(value).trim() !== '');
    const numericValues = nonEmptyValues
      .map((value) => Number(String(value).replace(/[$,%\s,]/g, '')))
      .filter((value) => Number.isFinite(value));
    const isNumeric = numericValues.length > 0 && numericValues.length === nonEmptyValues.length;
    const sum = numericValues.reduce((total, value) => total + value, 0);

    return {
      Column: header,
      Type: isNumeric ? 'Numeric' : 'Text/Mixed',
      'Non-empty': nonEmptyValues.length,
      Blank: values.length - nonEmptyValues.length,
      Unique: new Set(nonEmptyValues.map((value) => String(value).trim())).size,
      Sum: isNumeric ? round(sum) : '',
      Average: isNumeric ? round(sum / numericValues.length) : '',
      Min: isNumeric ? round(Math.min(...numericValues)) : '',
      Max: isNumeric ? round(Math.max(...numericValues)) : '',
    };
  });
}

function renderStats(rows) {
  statsTableBody.innerHTML = '';

  if (rows.length === 0) {
    statsTableBody.innerHTML = '<tr><td colspan="9" class="empty-state">Merge files to calculate statistics.</td></tr>';
    return;
  }

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    Object.values(row).forEach((value) => {
      const td = document.createElement('td');
      td.textContent = value;
      tr.appendChild(td);
    });
    statsTableBody.appendChild(tr);
  });
}

function renderPreview(rows, headers) {
  previewTable.innerHTML = '';

  if (rows.length === 0) {
    previewTable.innerHTML = '<tbody><tr><td class="empty-state">No merged rows yet.</td></tr></tbody>';
    return;
  }

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  headers.forEach((header) => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement('tbody');
  rows.slice(0, 100).forEach((row) => {
    const tr = document.createElement('tr');
    headers.forEach((header) => {
      const td = document.createElement('td');
      td.textContent = row[header];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  previewTable.append(thead, tbody);
}

function exportWorkbook() {
  if (!window.XLSX || mergedRows.length === 0) return;

  const workbook = XLSX.utils.book_new();
  const mergedSheet = XLSX.utils.json_to_sheet(mergedRows, { header: allHeaders });
  const statsSheet = XLSX.utils.json_to_sheet(summaryRows);

  XLSX.utils.book_append_sheet(workbook, mergedSheet, 'Merged Data');
  XLSX.utils.book_append_sheet(workbook, statsSheet, 'Summary Stats');
  XLSX.writeFile(workbook, 'summary.xlsx');
}

function resetApp() {
  selectedFiles = [];
  mergedRows = [];
  summaryRows = [];
  allHeaders = [];
  fileInput.value = '';
  fileList.innerHTML = '';
  mergeButton.disabled = true;
  resetButton.disabled = true;
  exportButton.disabled = true;
  statusText.textContent = 'Waiting for files.';
  renderStats([]);
  renderPreview([], []);
  updateStatCards();
}

function updateStatCards() {
  statFiles.textContent = selectedFiles.length.toLocaleString();
  statRows.textContent = mergedRows.length.toLocaleString();
  statColumns.textContent = allHeaders.length.toLocaleString();
  statNumeric.textContent = summaryRows.filter((row) => row.Type === 'Numeric').length.toLocaleString();
}

function normalizeRow(row, headers) {
  return headers.reduce((normalized, header) => {
    normalized[header] = row[header] ?? '';
    return normalized;
  }, {});
}

function dedupeObjectKeys(rows) {
  return rows.map((row) => {
    const deduped = {};
    Object.entries(row).forEach(([key, value]) => {
      const cleanKey = String(key).trim() || 'Unnamed Column';
      let finalKey = cleanKey;
      let index = 2;

      while (Object.prototype.hasOwnProperty.call(deduped, finalKey)) {
        finalKey = `${cleanKey} ${index}`;
        index += 1;
      }

      deduped[finalKey] = value;
    });
    return deduped;
  });
}

function createFileList(files) {
  const dataTransfer = new DataTransfer();
  files.forEach((file) => dataTransfer.items.add(file));
  return dataTransfer.files;
}

function isSpreadsheetFile(file) {
  return /\.(xlsx|xls|csv)$/i.test(file.name);
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / 1024 ** index).toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

function round(value) {
  return Number.isFinite(value) ? Math.round((value + Number.EPSILON) * 100) / 100 : '';
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (character) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;',
  }[character]));
}
