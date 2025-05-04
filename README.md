# report
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Enhanced Invoice Mismatch Tool (Fixed)</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
  
  .heading-gradient {
  font-size: 36px;
  font-weight: 800;
  background: linear-gradient(to right, #0066ff, #00c6ff);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  text-shadow: 0px 2px 4px rgba(0,0,0,0.2);
  margin-bottom: 30px;
  font-family: 'Segoe UI', 'Helvetica Neue', sans-serif;
}

  body {
  background-color: #f8f9fa;
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}

.container {
  margin-top: 40px;
  max-width: 1200px;
}

.card {
  background: white;
  border-radius: 15px;
  padding: 30px;
  box-shadow: 0px 8px 20px rgba(0,0,0,0.1);
  margin-bottom: 20px;
}

.form-label {
  font-weight: 600;
}

input[type="file"] {
  border-radius: 8px;
}

.btn {
  font-weight: 600;
  font-size: 16px;
}

.progress {
  height: 25px;
  border-radius: 30px;
  overflow: hidden;
}

.progress-bar {
  font-size: 14px;
  line-height: 25px;
  background-color: #0d6efd;
}

#result {
  margin-top: 30px;
}

.summary-count {
  font-weight: bold;
  font-size: 18px;
  color: #d63384;
  margin-bottom: 10px;
  text-align: right;
}

.styled-table {
  border-collapse: collapse;
  margin: 0 auto;
  font-size: 14px;
  min-width: 100%;
  background-color: white;
  border-radius: 12px;
  overflow: hidden;
  box-shadow: 0 5px 20px rgba(0,0,0,0.05);
}

.styled-table thead tr {
  background-color: #0d6efd;
  color: #ffffff;
  text-align: center;
  font-weight: bold;
}

.styled-table th, .styled-table td {
  padding: 12px 16px;
  text-align: center;
}

.styled-table tbody tr {
  border-bottom: 1px solid #dee2e6;
}

.styled-table tbody tr:nth-child(even) {
  background-color: #f9f9f9;
}

.styled-table tbody tr:hover {
  background-color: #e2f0ff;
}

.credit-highlight {
  background-color: #ffe6e6 !important;
  font-weight: bold;
}


</style>

</head>
<body>
<div class="container">
  <div class="card">
    <h2 class="mb-4 text-center heading-gradient">Invoice Mismatch Tool</h2>

    <div class="mb-3">
      <label class="form-label">Upload FBR Excel File(s)</label>
      <input type="file" id="fbrMultipleFiles" class="form-control" accept=".xlsx,.xls" multiple>
    </div>
    <div class="mb-3">
      <label class="form-label">Upload POS Excel File(s)</label>
      <input type="file" id="posMultipleFiles" class="form-control" accept=".xlsx,.xls" multiple>
    </div>
    <button class="btn btn-primary w-100 mb-3" onclick="processBulkFiles()">Process Bulk Files</button>
    <button class="btn btn-success w-100 mb-3" onclick="downloadExcel()">Download Mismatches in Excel</button>
    <div class="progress d-none" id="processingProgressContainer">
      <div id="processingProgress" class="progress-bar progress-bar-striped progress-bar-animated" role="progressbar" style="width: 0%"></div>
    </div>
    <div id="result"></div>
  </div>
</div>
<script>
let finalMismatches = [];
let fbrFiles = [], posFiles = [];

function convertExcelDate(serial) {
  if (!serial || isNaN(serial)) return '-';
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);
  const seconds = total_seconds % 60;
  total_seconds -= seconds;
  const hours = Math.floor(total_seconds / 3600);
  const minutes = Math.floor(total_seconds / 60) % 60;
  return `${date_info.getFullYear()}-${String(date_info.getMonth() + 1).padStart(2, '0')}-${String(date_info.getDate()).padStart(2, '0')} ${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);
        resolve(jsonData);
      } catch (error) {
        console.error('Excel read error:', file.name, error);
        resolve([]);
      }
    };
    reader.onerror = (error) => {
      console.error('File read error:', file.name, error);
      resolve([]);
    };
    reader.readAsArrayBuffer(file);
  });
}

async function processBulkFiles() {
  finalMismatches = [];
  fbrFiles = Array.from(document.getElementById('fbrMultipleFiles').files);
  posFiles = Array.from(document.getElementById('posMultipleFiles').files);

  if (fbrFiles.length === 0 || posFiles.length === 0) {
    alert('Please upload both FBR and POS files.');
    return;
  }

  const progress = document.getElementById('processingProgress');
  document.getElementById('processingProgressContainer').classList.remove('d-none');
  progress.style.width = '0%';
  progress.textContent = '0%';

  let allFbrData = [], allPosData = [];
  for (let i = 0; i < fbrFiles.length; i++) {
    const data = await readExcel(fbrFiles[i]);
    allFbrData.push(...data);
    progress.style.width = `${Math.round(((i + 1) / (fbrFiles.length + posFiles.length)) * 50)}%`;
    progress.textContent = progress.style.width;
    await new Promise(r => setTimeout(r, 50));
  }

  for (let i = 0; i < posFiles.length; i++) {
    const data = await readExcel(posFiles[i]);
    allPosData.push(...data);
    progress.style.width = `${50 + Math.round(((i + 1) / posFiles.length) * 40)}%`;
    progress.textContent = progress.style.width;
    await new Promise(r => setTimeout(r, 50));
  }

  const fbrMap = new Map();
  const posMap = new Map();

  allFbrData.forEach(row => {
    const inv = (row['FBRInvoiceNumber'] || '').toString().trim();
    if (inv) fbrMap.set(inv, row);
  });
  allPosData.forEach(row => {
    const inv = (row['FBR#'] || '').toString().trim();
    if (inv) posMap.set(inv, row);
  });

  fbrMap.forEach((row, inv) => {
    if (!posMap.has(inv)) {
      finalMismatches.push({
        InvoiceNo: inv,
        MissingIn: 'POS',
        SaleValue: row['Sale_Value'] || 0,
        Tax: row['Tax_Charged'] || 0,
        Discount: row['Discount'] || 0,
        Total: row['Total_Balance'] || 0,
        Type: row['InvoiceType'] || row['Type'] || '-',
        DateTime: convertExcelDate(row['DateTime'])
      });
    }
  });

  posMap.forEach((row, inv) => {
    if (!fbrMap.has(inv)) {
      finalMismatches.push({
        InvoiceNo: inv,
        MissingIn: 'FBR',
        SaleValue: row['Sale_Value'] || 0,
        Tax: row['Tax_Charged'] || 0,
        Discount: row['Discount'] || 0,
        Total: row['Total_Balance'] || 0,
        Type: row['InvoiceType'] || row['Type'] || '-',
        DateTime: convertExcelDate(row['DateTime'])
      });
    }
  });

  progress.style.width = '100%';
  progress.textContent = '100%';
  displayResults();
}

function displayResults() {
  const container = document.getElementById('result');
  if (finalMismatches.length === 0) {
    container.innerHTML = '<div class="alert alert-success">No mismatches found.</div>';
    return;
  }

  let summary = `<div class="summary-count">Total Mismatched Invoices: ${finalMismatches.length}</div>`;

  let html = `<table class="styled-table"><thead><tr><th>Invoice No</th><th>Missing In</th><th>Type</th><th>Date Time</th><th>Sale Value</th><th>Tax</th><th>Discount</th><th>Total</th></tr></thead><tbody>`;
  
  finalMismatches.forEach(row => {
  const highlightClass = row.Type && row.Type.toUpperCase() === 'CREDIT' ? 'credit-highlight' : '';
  html += `
    <tr class="${highlightClass}">
      <td>${row.InvoiceNo}</td>
      <td>${row.MissingIn}</td>
      <td>${row.Type}</td>
      <td>${row.DateTime}</td>
      <td>${row.SaleValue}</td>
      <td>${row.Tax}</td>
      <td>${row.Discount}</td>
      <td>${row.Total}</td>
    </tr>`;
});


  html += `</tbody></table>`;
  container.innerHTML = summary + html;
}


function downloadExcel() {
  if (finalMismatches.length === 0) {
    alert('No mismatches to download.');
    return;
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(finalMismatches);
  XLSX.utils.book_append_sheet(wb, ws, "Mismatches");
  XLSX.writeFile(wb, "Mismatches.xlsx");
}
</script>
</body>
</html>
