// Bank Statement Spreadsheet Ledger Updater Client JS

// UI Elements
const lastUpdateDateEl = document.getElementById('last-update-date');
const uploaderCard = document.getElementById('uploader-card');
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const fileInfoBox = document.getElementById('file-info-box');
const fileNameEl = document.getElementById('file-name');
const fileSizeEl = document.getElementById('file-size');
const btnRemoveFile = document.getElementById('btn-remove-file');
const logBox = document.getElementById('log-box');
const previewCard = document.getElementById('preview-card');
const previewStartDateEl = document.getElementById('preview-start-date');
const previewEndDateEl = document.getElementById('preview-end-date');
const previewInflowTotalEl = document.getElementById('preview-inflow-total');
const previewOutflowTotalEl = document.getElementById('preview-outflow-total');
const previewTableBody = document.getElementById('preview-table-body');
const btnUpdateSheet = document.getElementById('btn-update-sheet');
const googleSheetBtn = document.getElementById('google-sheet-btn');

// Modal Elements
const warningModal = document.getElementById('warning-modal');
const modalLastUpdate = document.getElementById('modal-last-update');
const modalStmtStart = document.getElementById('modal-stmt-start');
const modalCutoffDate = document.getElementById('modal-cutoff-date');
const btnModalCancel = document.getElementById('btn-modal-cancel');
const btnModalContinue = document.getElementById('btn-modal-continue');

// Application State
let appState = {
  isConnected: false,
  lastUpdateDate: null, // "YYYY-MM-DD" or null
  initialBalance: 0.0,
  uploadedFile: null,
  transactions: [] // Array of parsed transactions
};

// Initialize app on load
window.addEventListener('DOMContentLoaded', () => {
  log('System Console initialized.', 'info');
  
  // Attach event listeners
  btnRemoveFile.addEventListener('click', removeFile);
  btnUpdateSheet.addEventListener('click', handleUpdateClick);
  
  // Modal buttons
  btnModalCancel.addEventListener('click', () => {
    warningModal.classList.remove('active');
    log('Upload cancelled by user during overlap warning.', 'warn');
  });
  
  btnModalContinue.addEventListener('click', () => {
    warningModal.classList.remove('active');
    performUpload(appState.lastUpdateDate);
  });
  
  setupFileUploader();
  
  // Fetch initial connection status from local server
  fetchStatus();
});

// Logger Function
function log(message, type = 'info') {
  const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const entry = document.createElement('div');
  entry.className = `log-entry ${type}`;
  entry.innerHTML = `<span class="log-time">[${time}]</span> <span>${escapeHtml(message)}</span>`;
  logBox.appendChild(entry);
  logBox.scrollTop = logBox.scrollHeight;
}

function escapeHtml(text) {
  const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

// Fetch status from server internally
async function fetchStatus() {
  log('Connecting to Google Sheet internally...', 'info');
  try {
    const response = await fetch('/api/status');
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    
    if (data.status === 'success') {
      appState.isConnected = true;
      appState.lastUpdateDate = data.lastUpdateDate;
      appState.initialBalance = parseFloat(data.initialBalance) || 0.0;
      
      // Update UI Status
      lastUpdateDateEl.textContent = data.lastUpdateDate ? formatDateString(data.lastUpdateDate) : 'No transactions uploaded yet';
      
      log(`Connected successfully! Last update date: ${data.lastUpdateDate || 'N/A'}. Initial Balance: ${appState.initialBalance.toFixed(2)}`, 'success');
    } else {
      throw new Error(data.message || 'Unknown response structure from server status API.');
    }
  } catch (err) {
    appState.isConnected = false;
    lastUpdateDateEl.textContent = 'Connection Error';
    log(`Connection failed: ${err.message}. Ensure your local backend is running and the script in config.json is deployed correctly.`, 'error');
  }
}

// Drag and drop file upload setup
function setupFileUploader() {
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length > 0) {
      handleFileSelection(files[0]);
    }
  }, false);

  fileInput.addEventListener('change', (e) => {
    if (fileInput.files.length > 0) {
      handleFileSelection(fileInput.files[0]);
    }
  });
}

function handleFileSelection(file) {
  const extension = file.name.split('.').pop().toLowerCase();
  const validExtensions = ['xlsx', 'xls', 'csv'];
  
  if (!validExtensions.includes(extension)) {
    log(`Error: Selected file is not supported (${file.name}). Please upload an Excel (.xlsx, .xls) or CSV (.csv) file.`, 'error');
    alert('Invalid file format. Please upload .xlsx, .xls or .csv statement.');
    return;
  }
  
  appState.uploadedFile = file;
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatBytes(file.size);
  
  // Show file info, hide dropzone contents
  dropzone.style.display = 'none';
  fileInfoBox.style.display = 'flex';
  
  log(`Selected statement file: ${file.name} (${formatBytes(file.size)})`, 'info');
  parseStatementFile(file);
}

function removeFile() {
  appState.uploadedFile = null;
  appState.transactions = [];
  fileInput.value = '';
  
  dropzone.style.display = 'flex';
  fileInfoBox.style.display = 'none';
  previewCard.style.display = 'none';
  
  log('Removed statement file.', 'info');
}

// Parse file using SheetJS
function parseStatementFile(file) {
  log(`Parsing bank statement data...`, 'info');
  const reader = new FileReader();
  
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });
      
      // Load the first sheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Parse as 2D array of values
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
      
      if (rows.length < 2) {
        throw new Error('Spreadsheet appears to be empty or contains too few rows.');
      }
      
      // 1. Find Header row & columns index
      let headerRowIdx = -1;
      let colIndexes = { date: -1, description: -1, ref_no: -1, amount: -1, dr_cr: -1, balance: -1 };
      
      for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const row = rows[i];
        if (!row) continue;
        
        const cleanRow = row.map(cell => cell !== null && cell !== undefined ? String(cell).trim().toLowerCase() : "");
        
        // Find columns
        const dateIdx = cleanRow.findIndex(h => h.includes('value date') || h === 'date' || h === 'txn date' || h === 'transaction date');
        const descIdx = cleanRow.findIndex(h => h.includes('description') || h.includes('particulars') || h.includes('narration'));
        const amtIdx = cleanRow.findIndex(h => h === 'amount' || h === 'withdrawal amt.' || h === 'deposit amt.' || (h.includes('amount') && !h.includes('date')));
        
        // Required columns to identify header row
        if (dateIdx !== -1 && descIdx !== -1 && amtIdx !== -1) {
          headerRowIdx = i;
          colIndexes.date = dateIdx;
          colIndexes.description = descIdx;
          colIndexes.amount = amtIdx;
          colIndexes.ref_no = cleanRow.findIndex(h => h.includes('chq') || h.includes('ref') || h.includes('cheque'));
          colIndexes.dr_cr = cleanRow.findIndex(h => h === 'dr / cr' || h === 'dr/cr' || h === 'type' || h === 'cr/dr' || h === 'transaction type' || h === 'debit/credit' || (h.includes('dr') && h.includes('cr')));
          colIndexes.balance = cleanRow.findIndex(h => h === 'balance' || h === 'closing balance' || (h.includes('bal') && !h.includes('debit') && !h.includes('credit')));
          break;
        }
      }
      
      if (headerRowIdx === -1) {
        throw new Error("Could not find required columns. Statement must contain headers for 'Value Date' (or 'Date'), 'Description' (or 'Particulars'), and 'Amount'.");
      }
      
      // Perform strict validation of remaining required columns
      let missingRequired = [];
      if (colIndexes.dr_cr === -1) missingRequired.push("'Dr / Cr' (or 'Type')");
      if (colIndexes.balance === -1) missingRequired.push("'Balance'");
      
      if (missingRequired.length > 0) {
        throw new Error("Invalid Format: Missing required columns: " + missingRequired.join(" and ") + ". Make sure your sheet includes all expected bank book fields.");
      }
      
      log(`Detected headers at row ${headerRowIdx + 1}. Mapping columns: Date (Col ${colIndexes.date + 1}), Description (Col ${colIndexes.description + 1}), Amount (Col ${colIndexes.amount + 1}), Dr/Cr (Col ${colIndexes.dr_cr + 1}), Balance (Col ${colIndexes.balance + 1})`, 'info');
      
      // 2. Loop and Parse transactions with strict formatting validation
      let statementTxs = [];
      let totalSkipped = 0;
      let totalOutsideFY = 0;
      let formatWarnings = 0;
      
      for (let r = headerRowIdx + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        
        // Skip empty rows
        const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === "");
        if (isRowEmpty) continue;
        
        const rawDate = row[colIndexes.date];
        if (rawDate === null || rawDate === undefined || String(rawDate).trim() === "") {
          log(`Warning [Row ${r + 1}]: Skip - Date field is empty.`, 'warn');
          totalSkipped++;
          formatWarnings++;
          continue;
        }
        
        const dateStr = parseBankDate(rawDate);
        if (!dateStr) {
          log(`Error [Row ${r + 1}]: Malformed Date format: "${rawDate}". Skipping row.`, 'error');
          totalSkipped++;
          formatWarnings++;
          continue;
        }
        
        // Filter: Keep only transactions within the FY 2026 Apr to 2027 Mar
        if (dateStr < "2026-04-01" || dateStr > "2027-03-31") {
          totalOutsideFY++;
          continue;
        }
        
        const desc = row[colIndexes.description] !== null && row[colIndexes.description] !== undefined ? String(row[colIndexes.description]).trim() : "";
        if (desc === "" || desc === "-") {
          log(`Warning [Row ${r + 1}]: Empty or placeholder transaction description.`, 'warn');
          formatWarnings++;
        }
        
        const refNo = colIndexes.ref_no !== -1 && row[colIndexes.ref_no] !== null ? String(row[colIndexes.ref_no]).trim() : "-";
        
        const rawAmount = row[colIndexes.amount];
        if (rawAmount === null || rawAmount === undefined || String(rawAmount).trim() === "") {
          log(`Error [Row ${r + 1}]: Empty Amount field. Skipping row.`, 'error');
          totalSkipped++;
          formatWarnings++;
          continue;
        }
        
        const cleanAmtStr = String(rawAmount).replace(/,/g, "").trim();
        const amount = parseFloat(cleanAmtStr);
        if (isNaN(amount) || amount < 0) {
          log(`Error [Row ${r + 1}]: Malformed numeric Amount: "${rawAmount}". Skipping row.`, 'error');
          totalSkipped++;
          formatWarnings++;
          continue;
        }
        
        let drCr = "DR";
        const rawDrCr = row[colIndexes.dr_cr];
        if (rawDrCr === null || rawDrCr === undefined || String(rawDrCr).trim() === "") {
          log(`Warning [Row ${r + 1}]: Transaction Type (Dr/Cr) is blank. Defaulting to DR.`, 'warn');
          formatWarnings++;
        } else {
          const typeVal = String(rawDrCr).trim().toUpperCase();
          if (typeVal === "CR" || typeVal === "CREDIT" || typeVal === "IN" || typeVal === "INFLOW" || typeVal === "C") {
            drCr = "CR";
          } else if (typeVal === "DR" || typeVal === "DEBIT" || typeVal === "OUT" || typeVal === "OUTFLOW" || typeVal === "D") {
            drCr = "DR";
          } else {
            log(`Warning [Row ${r + 1}]: Unrecognized Transaction Type "${rawDrCr}". Defaulting to DR.`, 'warn');
            formatWarnings++;
          }
        }
        
        const rawBalance = row[colIndexes.balance];
        if (rawBalance === null || rawBalance === undefined || String(rawBalance).trim() === "") {
          log(`Error [Row ${r + 1}]: Empty Balance field. Skipping row.`, 'error');
          totalSkipped++;
          formatWarnings++;
          continue;
        }
        
        const cleanBalStr = String(rawBalance).replace(/,/g, "").trim();
        const balance = parseFloat(cleanBalStr);
        if (isNaN(balance)) {
          log(`Error [Row ${r + 1}]: Malformed numeric Balance: "${rawBalance}". Skipping row.`, 'error');
          totalSkipped++;
          formatWarnings++;
          continue;
        }
        
        statementTxs.push({
          date: dateStr,
          description: desc,
          ref_no: refNo,
          amount: amount,
          dr_cr: drCr,
          balance: balance,
          row_index: r
        });
      }
      
      if (statementTxs.length === 0) {
        let msg = "No valid transactions found in the bank statement file.";
        if (totalOutsideFY > 0) {
          msg += ` Note: ${totalOutsideFY} transactions were found but skipped because they fall outside the ledger's financial year (April 1st, 2026 to March 31st, 2027).`;
        }
        throw new Error(msg);
      }
      
      // 3. Chronological sorting
      statementTxs.sort((a, b) => a.date.localeCompare(b.date) || a.row_index - b.row_index);
      
      appState.transactions = statementTxs;
      
      // Log details
      let logMsg = `Successfully parsed ${statementTxs.length} transactions for FY 2026-2027.`;
      if (totalOutsideFY > 0) logMsg += ` Skip ${totalOutsideFY} out-of-year transactions.`;
      if (totalSkipped > 0) logMsg += ` Skip ${totalSkipped} malformed/empty rows.`;
      if (formatWarnings > 0) logMsg += ` Found ${formatWarnings} format warnings/errors (see log history above).`;
      log(logMsg, formatWarnings > 0 ? 'warn' : 'success');
      
      // Render Preview Card
      renderStatementPreview();
      
    } catch (err) {
      log(`Error parsing spreadsheet: ${err.message}`, 'error');
      alert(`Parsing failed: ${err.message}`);
      removeFile();
    }
  };
  
  reader.readAsArrayBuffer(file);
}

// Convert cell dates to YYYY-MM-DD
function parseBankDate(val) {
  if (val === undefined || val === null || val === '') return null;
  
  if (val instanceof Date && !isNaN(val.valueOf())) {
    return formatDate(val);
  }
  
  if (typeof val === 'number') {
    // Math.floor of Excel serial date
    const utcDays = Math.floor(val);
    // Excel epoch begins Dec 30, 1899
    const d = new Date(1899, 11, 30 + utcDays);
    return formatDate(d);
  }
  
  const valStr = String(val).trim();
  if (/^\d+(\.\d+)?$/.test(valStr)) {
    const utcDays = Math.floor(parseFloat(valStr));
    const d = new Date(1899, 11, 30 + utcDays);
    return formatDate(d);
  }
  
  // Standard JS Date parsing
  const parsed = Date.parse(valStr);
  if (!isNaN(parsed)) {
    return formatDate(new Date(parsed));
  }
  
  // Custom Regex for DD/MM/YYYY or DD-MM-YYYY
  const dmyMatch = valStr.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (dmyMatch) {
    let day = parseInt(dmyMatch[1]);
    let month = parseInt(dmyMatch[2]);
    const year = parseInt(dmyMatch[3]);
    
    // Swap if month seems to be day (often standard in DD/MM/YYYY Indian bank statements)
    if (month > 12 && day <= 12) {
      const temp = day;
      day = month;
      month = temp;
    }
    
    const d = new Date(year, month - 1, day);
    if (!isNaN(d.getTime())) {
      return formatDate(d);
    }
  }
  
  return null;
}

function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function formatDateString(str) {
  const date = new Date(str);
  return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

// Display Preview Information in the Card
function renderStatementPreview() {
  const txs = appState.transactions;
  const startD = txs[0].date;
  const endD = txs[txs.length - 1].date;
  
  let inflowTotal = 0;
  let outflowTotal = 0;
  
  previewTableBody.innerHTML = '';
  
  txs.forEach(tx => {
    if (tx.dr_cr === 'CR') {
      inflowTotal += tx.amount;
    } else {
      outflowTotal += tx.amount;
    }
    
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${formatDateString(tx.date)}</td>
      <td class="table-desc" title="${escapeHtml(tx.description)}">${escapeHtml(tx.description)}</td>
      <td class="${tx.dr_cr === 'CR' ? 'table-inflow' : ''}">${tx.dr_cr === 'CR' ? tx.amount.toFixed(2) : '-'}</td>
      <td class="${tx.dr_cr === 'DR' ? 'table-outflow' : ''}">${tx.dr_cr === 'DR' ? tx.amount.toFixed(2) : '-'}</td>
      <td class="table-ref">${escapeHtml(tx.ref_no)}</td>
    `;
    previewTableBody.appendChild(row);
  });
  
  previewStartDateEl.textContent = formatDateString(startD);
  previewEndDateEl.textContent = formatDateString(endD);
  previewInflowTotalEl.textContent = inflowTotal.toFixed(2);
  previewOutflowTotalEl.textContent = outflowTotal.toFixed(2);
  
  previewCard.style.display = 'block';
}

// Handle trigger action to save statement records to Sheets
function handleUpdateClick() {
  if (!appState.isConnected) {
    alert('Please connect to the Google Sheet first.');
    return;
  }
  if (appState.transactions.length === 0) {
    alert('No statement file loaded.');
    return;
  }
  
  const sheetLastDate = appState.lastUpdateDate;
  const stmtStartDate = appState.transactions[0].date;
  
  // Overlap Warning Trigger: If statement starts on or before sheet's last update date
  if (sheetLastDate && stmtStartDate <= sheetLastDate) {
    modalLastUpdate.textContent = formatDateString(sheetLastDate);
    modalStmtStart.textContent = formatDateString(stmtStartDate);
    modalCutoffDate.textContent = formatDateString(sheetLastDate);
    warningModal.classList.add('active');
    log('Warning: statement dates overlap with ledger records. Showing confirmation dialogue.', 'warn');
  } else {
    // No overlap: upload all transactions directly
    performUpload(null);
  }
}

// Upload Data payload via local server API (which relays to Apps Script internally)
async function performUpload(cutoffDateStr) {
  let filteredTxs = [...appState.transactions];
  
  if (cutoffDateStr) {
    filteredTxs = appState.transactions.filter(tx => tx.date >= cutoffDateStr);
    log(`Filtering transactions: kept ${filteredTxs.length} items on/after ${cutoffDateStr} (discarded ${appState.transactions.length - filteredTxs.length} older items).`, 'info');
    
    if (filteredTxs.length === 0) {
      log('Cancel upload: all uploaded transactions are older than the last update date.', 'warn');
      alert('Upload Cancelled: No transactions found that are newer than or equal to the last update date.');
      return;
    }
  }

  btnUpdateSheet.disabled = true;
  btnUpdateSheet.textContent = 'Updating Spreadsheet...';
  log('Sending data blocks to server...', 'info');
  
  // Dynamically estimate starting balance of the year if the Google Sheet currently has no starting balance
  let startBalance = undefined;
  if (appState.initialBalance === 0.0 && appState.transactions.length > 0) {
    const firstTx = appState.transactions[0];
    startBalance = firstTx.dr_cr === 'DR' ? firstTx.balance + firstTx.amount : firstTx.balance - firstTx.amount;
    log(`Calculated dynamic opening balance for the financial year from statement: ${startBalance.toFixed(2)}`, 'info');
  }
  
  try {
    const response = await fetch('/api/update', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        transactions: filteredTxs,
        lastUpdateDate: cutoffDateStr,
        startBalance: startBalance
      })
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    
    if (result.status === 'success') {
      log('Google Sheet updated successfully!', 'success');
      
      // Render success screen on preview card area
      showSuccessScreen(filteredTxs.length);
      
      // Refresh Sheet update metadata
      log('Refreshing sheet synchronization state...', 'info');
      await syncLastUpdateDate();
      
    } else {
      throw new Error(result.message || 'Unknown error during sheet write operation.');
    }
  } catch (err) {
    log(`Upload failed: ${err.message}`, 'error');
    alert(`Failed to update Google Sheet: ${err.message}`);
  } finally {
    btnUpdateSheet.disabled = false;
    btnUpdateSheet.textContent = 'Update Google Sheet';
  }
}

// Silently fetch and sync update date from server
async function syncLastUpdateDate() {
  try {
    const response = await fetch('/api/status');
    const data = await response.json();
    if (data.status === 'success') {
      appState.lastUpdateDate = data.lastUpdateDate;
      appState.initialBalance = parseFloat(data.initialBalance) || 0.0;
      lastUpdateDateEl.textContent = data.lastUpdateDate ? formatDateString(data.lastUpdateDate) : 'No transactions uploaded yet';
      log(`Sync: last update date is now ${data.lastUpdateDate || 'N/A'}.`, 'info');
    }
  } catch (err) {
    log(`Failed to refresh sync state: ${err.message}`, 'warn');
  }
}

// Display Success Overlay inside the Preview Card area
function showSuccessScreen(count) {
  previewCard.innerHTML = `
    <div class="success-screen">
      <div class="success-icon">✓</div>
      <h2 style="font-family: var(--font-title); font-size: 1.5rem; margin-top: 1rem;">Upload Complete!</h2>
      <p style="color: var(--text-muted); font-size: 0.9rem;">
        Successfully uploaded <strong>${count}</strong> transactions to the Google Sheet ledger. All formulas, monthly totals, and preceding balance linkages have been recalculated.
      </p>
      <button onclick="resetPreviewToInitial()" class="btn-primary" style="margin-top: 1rem;">
        Upload Another File
      </button>
    </div>
  `;
}

// Reset right preview column to accept another upload
window.resetPreviewToInitial = function() {
  // Restore preview card HTML template
  previewCard.innerHTML = `
    <div class="card-title">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="2.5"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><line x1="9" y1="3" x2="9" y2="21"></line><line x1="15" y1="3" x2="15" y2="21"></line><line x1="3" y1="9" x2="21" y2="9"></line><line x1="3" y1="15" x2="21" y2="15"></line></svg>
      Statement Preview
    </div>
    
    <div class="preview-summary">
      <div class="summary-item">
        <div class="summary-item-label">Start Date</div>
        <div id="preview-start-date" class="summary-item-value">-</div>
      </div>
      <div class="summary-item">
        <div class="summary-item-label">End Date</div>
        <div id="preview-end-date" class="summary-item-value">-</div>
      </div>
      <div class="summary-item">
        <div class="summary-item-label">Inflows Total</div>
        <div id="preview-inflow-total" class="summary-item-value value-inflow">0.00</div>
      </div>
      <div class="summary-item">
        <div class="summary-item-label">Outflows Total</div>
        <div id="preview-outflow-total" class="summary-item-value value-outflow">0.00</div>
      </div>
    </div>
 
    <div class="table-wrapper">
      <table id="preview-table">
        <thead>
          <tr>
            <th>Date</th>
            <th>Description</th>
            <th>Inflow</th>
            <th>Outflow</th>
            <th>Ref No.</th>
          </tr>
        </thead>
        <tbody id="preview-table-body">
        </tbody>
      </table>
    </div>

    <div class="action-bar">
      <button id="btn-update-sheet" class="btn-primary">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline><polyline points="7 3 7 8 15 8"></polyline></svg>
        Update Google Sheet
      </button>
    </div>
  `;
  
  // Re-bind click event to new button element
  document.getElementById('btn-update-sheet').addEventListener('click', handleUpdateClick);
  
  // Reset uploader state
  removeFile();
};
