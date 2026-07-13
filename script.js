// --- UI INTERACTION ---
const inputExcel = document.getElementById('inputExcel');
const dropZone = document.getElementById('dropZone');
const fileNameDisplay = document.getElementById('fileName');
const dropTitle = document.getElementById('dropTitle');

dropZone.addEventListener('click', () => inputExcel.click());
inputExcel.addEventListener('change', function() { handleFiles(this.files); });
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('active'); });
dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('active'); });
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('active');
    handleFiles(e.dataTransfer.files);
    inputExcel.files = e.dataTransfer.files;
});

function handleFiles(files) {
    if (files.length > 0) {
        const file = files[0];
        dropTitle.innerText = "File Siap!";
        fileNameDisplay.innerHTML = `<span class="text-success fw-bold"><i class="fa-solid fa-check me-1"></i> ${file.name}</span>`;
    }
}

// --- FUNGSI VALIDASI FIELD WAJIB ---
function validateForm() {
    const clientName = document.getElementById('clientName').value.trim();
    const period = document.getElementById('period').value.trim();
    const materialitas = document.getElementById('materialitas').value.trim();
    const preparedBy = document.getElementById('preparedBy').value.trim();
    const preparedDate = document.getElementById('preparedDate').value;
    const reviewedBy = document.getElementById('reviewedBy').value.trim();
    const reviewedDate = document.getElementById('reviewedDate').value;
    const sortingOption = document.querySelector('input[name="sortingOption"]:checked');

    const errors = [];
    if (!clientName) errors.push('Nama Klien / Entitas harus diisi');
    if (!period) errors.push('Periode Audit harus diisi');
    if (!materialitas) errors.push('Batas Nilai Materialitas harus diisi');
    if (!preparedBy) errors.push('Dibuat Oleh harus diisi');
    if (!preparedDate) errors.push('Tanggal Persiapan harus diisi');
    if (!reviewedBy) errors.push('Direview Oleh harus diisi');
    if (!reviewedDate) errors.push('Tanggal Review harus diisi');
    if (!sortingOption) errors.push('Opsi Pengurutan Data harus dipilih');
    if (!inputExcel.files.length) errors.push('File Excel harus di-upload');

    if (errors.length > 0) {
        Swal.fire({
            icon: 'warning',
            title: 'Data Tidak Lengkap',
            html: '<ul style="text-align: left; margin: 0;">' +
                  errors.map(err => `<li>${err}</li>`).join('') +
                  '</ul>',
            confirmButtonColor: '#0d6efd'
        });
        return false;
    }
    return true;
}

// --- FUNGSI HELPER ---
function getNilaiMaterialitas() {
    const val = document.getElementById('materialitas').value;
    return parseInt(val.replace(/\./g, ''), 10) || 0;
}

function getSortingOption() {
    const opt = document.querySelector('input[name="sortingOption"]:checked');
    return opt ? opt.value : 'date';
}

function parseToTimestamp(val) {
    if (!val) return 0;
    if (val instanceof Date) return val.getTime();
    if (typeof val === 'number') {
        return val < 50000 ? new Date(1900, 0, 1 + val).getTime() : val;
    }
    if (typeof val === 'string') {
        if (val.includes('/') && !val.includes('-')) {
            const parts = val.split('/');
            if (parts.length === 3) {
                const [d, m, y] = parts.map(Number);
                if (!isNaN(d) && !isNaN(m) && !isNaN(y)) return new Date(y, m - 1, d).getTime();
            }
        }
        const parsed = new Date(val);
        if (!isNaN(parsed.getTime())) return parsed.getTime();
    }
    return 0;
}

function formatTanggal(item) {
    if (!item.tgl) return '';
    if (typeof item.tgl === 'string') return item.tgl;
    let dateObj = null;
    if (item.tgl instanceof Date) dateObj = item.tgl;
    else if (typeof item.tgl === 'number' && item.tglTimestamp > 0) dateObj = new Date(item.tglTimestamp);
    if (!dateObj) return '';
    const dd = dateObj.getDate().toString().padStart(2, '0');
    const mm = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    return `${dd}/${mm}/${dateObj.getFullYear()}`;
}

function formatRupiahGL(num) {
    if (typeof num !== 'number') return String(num || '');
    const isNeg = num < 0;
    const parts = Math.abs(num).toFixed(2).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (isNeg ? '-' : '') + parts[0] + ',' + parts[1];
}

// --- LOGIKA SAMPLING (SHARED) ---
function determineSampleCounts() {
    const topInput = document.getElementById('topSamples').value;
    const randomInput = document.getElementById('randomSamples').value;

    // Setiap field bersifat independen: jika dikosongkan, gunakan default 15
    // sesuai dengan placeholder "Default: 15" yang ditampilkan di UI.
    let topCount = topInput === "" ? 15 : parseInt(topInput, 10);
    let randomCount = randomInput === "" ? 15 : parseInt(randomInput, 10);

    if (isNaN(topCount) || topCount < 0) topCount = 0;
    if (isNaN(randomCount) || randomCount < 0) randomCount = 0;

    return { topCount, randomCount };
}

function extractTopSamples(data, count) {
    if (count <= 0) return [];

    return data.slice(0, count);
}

function extractRandomSamples(data, count) {
    if (count <= 0) return [];
    let sisaData = [...data];
    for (let i = sisaData.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [sisaData[i], sisaData[j]] = [sisaData[j], sisaData[i]];
    }
    return sisaData.slice(0, count);
}

function getSampledItems(inputSheet, batasMaterialitas, sortingOption, sampleCounts) {
    let dataRaw = [];
    inputSheet.eachRow((row, rowNum) => {
        if (rowNum > 1) {
            let nominal = row.getCell(4).value;
            if (typeof nominal !== 'number') nominal = parseFloat(nominal) || 0;
            const tglValue = row.getCell(1).value;
            dataRaw.push({
                idBaris: rowNum,
                tgl: tglValue,
                tglTimestamp: parseToTimestamp(tglValue),
                voucher: row.getCell(2).value || '',
                ket: row.getCell(3).value || '',
                nominal: nominal
            });
        }
    });

    // Urutkan berdasarkan NILAI ABSOLUT (bukan nilai bertanda/signed) agar transaksi
    // kredit/retur besar (nominal negatif) tetap dapat masuk sebagai "Sampel Teratas",
    // konsisten dengan logika filter materialitas di bawah yang juga memakai Math.abs().
    dataRaw.sort((a, b) => Math.abs(b.nominal) - Math.abs(a.nominal));
    const dataFiltered = dataRaw.filter(item => Math.abs(item.nominal) >= batasMaterialitas);

    const topSamples = extractTopSamples(dataFiltered, sampleCounts.topCount);
    const sisaData = dataFiltered.slice(sampleCounts.topCount);

    const randomSamples = extractRandomSamples(sisaData, sampleCounts.randomCount);

    const finalSamples = [...topSamples, ...randomSamples];
    if (sortingOption === 'date') {
        finalSamples.sort((a, b) => a.tglTimestamp - b.tglTimestamp);
    } else {
        finalSamples.sort((a, b) => b.nominal - a.nominal);
    }
    return finalSamples;
}

// --- LOGIKA UTAMA ---
async function processAuditWorkpaper(mode = 'excel') {
    if (!validateForm()) return;

    const clientName = document.getElementById('clientName').value || "";
    const period = document.getElementById('period').value || "";
    const preparedBy = document.getElementById('preparedBy').value || "";
    const batasMaterialitas = getNilaiMaterialitas();
    const sortingOption = getSortingOption();
    const sampleCounts = determineSampleCounts();

    const formatDate = (dateStr) => {
        if (!dateStr) return "";
        const d = new Date(dateStr);
        return `${d.getDate().toString().padStart(2,'0')}/${(d.getMonth()+1).toString().padStart(2,'0')}/${d.getFullYear()}`;
    };

    const preparedDate = formatDate(document.getElementById('preparedDate').value);
    const reviewedBy = document.getElementById('reviewedBy').value || "";
    const reviewedDate = formatDate(document.getElementById('reviewedDate').value);

    Swal.fire({ title: 'Sedang Memproses...', html: 'Menyiapkan data...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });

    try {
        const workbookInput = new ExcelJS.Workbook();
        await workbookInput.xlsx.load(await inputExcel.files[0].arrayBuffer());

        if (mode === 'pdf') {
            // --- PDF EXPORT ---
            const allSheetsData = [];
            workbookInput.eachSheet((inputSheet) => {
                const samples = getSampledItems(inputSheet, batasMaterialitas, sortingOption, sampleCounts);
                allSheetsData.push({ sheetName: inputSheet.name, samples });
            });

            await generatePDF(allSheetsData, clientName, period);
            Swal.fire({ icon: 'success', title: 'Berhasil!', text: 'File PDF Test of Detail telah diunduh.', confirmButtonColor: '#dc3545' });

        } else {
            // --- EXCEL EXPORT ---
            const response = await fetch('assets/Template_Output.xlsx');
            if (!response.ok) throw new Error("Gagal load Template_Output.xlsx");

            const workbookTemplate = new ExcelJS.Workbook();
            await workbookTemplate.xlsx.load(await response.arrayBuffer());
            const masterSheet = workbookTemplate.worksheets[0];

            const HEADER_ROWS = 12;
            const START_DATA_ROW = 13;
            const FOOTER_START_ROW = 45;
            const FOOTER_GAP = 5;

            const workbookResult = new ExcelJS.Workbook();

            workbookInput.eachSheet((inputSheet) => {
                const sheetName = inputSheet.name;
                const newSheet = workbookResult.addWorksheet(sheetName);

                masterSheet.columns.forEach((col, index) => {
                    const newCol = newSheet.getColumn(index + 1);
                    newCol.width = col.width;
                    if (col.style) newCol.style = col.style;
                });
                newSheet.getColumn(4).width = 60;

                copyRows(masterSheet, newSheet, 1, HEADER_ROWS, 1, sheetName);

                newSheet.getCell('B7').value = clientName;
                newSheet.getCell('B8').value = period;
                newSheet.getCell('V7').value = preparedBy;
                newSheet.getCell('V8').value = preparedDate;
                newSheet.getCell('Y7').value = reviewedBy;
                newSheet.getCell('Y8').value = reviewedDate;

                ['B7', 'B8', 'V7', 'V8', 'Y7', 'Y8'].forEach(addr => {
                    newSheet.getCell(addr).font = { name: 'Arial', size: 10, bold: true };
                });

                const finalSamples = getSampledItems(inputSheet, batasMaterialitas, sortingOption, sampleCounts);

                let currentRowIdx = START_DATA_ROW;
                const templateDataRow = masterSheet.getRow(START_DATA_ROW);

                finalSamples.forEach((item, index) => {
                    const row = newSheet.getRow(currentRowIdx);
                    if (templateDataRow.height) row.height = templateDataRow.height;

                    row.getCell(1).value = index + 1;
                    row.getCell(2).value = formatTanggal(item);
                    row.getCell(3).value = item.voucher;
                    row.getCell(4).value = item.ket;
                    row.getCell(5).value = item.nominal;

                    for (let c = 10; c <= 21; c++) row.getCell(c).value = "N/A";

                    for (let c = 22; c <= 25; c++) {
                        row.getCell(c).value = "";
                        row.getCell(c).dataValidation = {
                            type: 'list', allowBlank: true, formulae: ['"V,X"'],
                            showErrorMessage: true, error: 'Pilih V atau X'
                        };
                    }

                    for (let c = 1; c <= 31; c++) {
                        const cell = row.getCell(c);
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.font = { name: 'Arial', size: 10 };
                        if (c <= 2 || (c >= 10 && c <= 25)) {
                            cell.alignment = { vertical: 'top', horizontal: 'center', wrapText: true };
                        } else if (c === 5) {
                            cell.alignment = { vertical: 'top', horizontal: 'right' };
                        } else {
                            cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
                        }
                    }

                    row.getCell(5).numFmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
                    currentRowIdx++;
                });

                const footerDestRow = currentRowIdx + FOOTER_GAP + 1;
                const footerRowCount = masterSheet.rowCount - FOOTER_START_ROW + 1;
                if (footerRowCount > 0) {
                    copyRows(masterSheet, newSheet, FOOTER_START_ROW, masterSheet.rowCount, footerDestRow, sheetName);
                }
            });

            const buffer = await workbookResult.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            saveAs(blob, `Kertas_Kerja_${clientName.replace(/ /g, "_") || 'Audit'}.xlsx`);

            Swal.fire({ icon: 'success', title: 'Berhasil!', text: 'Kertas Kerja Audit telah diunduh.', confirmButtonColor: '#198754' });
        }

    } catch (error) {
        console.error(error);
        Swal.fire({ icon: 'error', title: 'Error', text: error.message });
    }
}

// --- GENERATE PDF ---
async function generatePDF(allSheetsData, clientName, period) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    // A4 usable width: 210 - 14 - 14 = 182mm
    const colWidths = { no: 8, tgl: 22, voucher: 45, ket: 72, nominal: 35 };

    allSheetsData.forEach((sheetData, idx) => {
        if (idx > 0) doc.addPage();

        const { sheetName, samples } = sheetData;

        // Header blok (bold, 9pt)
        doc.setFontSize(9);
        doc.setFont('helvetica', 'bold');
        doc.text(`Klien    : ${clientName}`, 14, 14);
        doc.text(`Periode  : ${period}`, 14, 19);
        doc.text(`Akun     : ${sheetName}`, 14, 24);

        const tableBody = samples.map((item, i) => [
            i + 1,
            formatTanggal(item),
            item.voucher || '',
            item.ket || '',
            formatRupiahGL(item.nominal)
        ]);

        doc.autoTable({
            startY: 29,
            head: [['No', 'Tgl', 'No. Voucher', 'Nama Transaksi', 'Jumlah Menurut GL']],
            body: tableBody,
            styles: {
                font: 'helvetica',
                fontSize: 8,
                cellPadding: { top: 1.5, right: 2, bottom: 1.5, left: 2 },
                overflow: 'linebreak',
                textColor: [0, 0, 0],
                lineColor: [0, 0, 0],
                lineWidth: 0.25
            },
            headStyles: {
                fillColor: [255, 255, 255],
                textColor: [0, 0, 0],
                fontStyle: 'bold',
                halign: 'center'
            },
            bodyStyles: {
                valign: 'top'
            },
            columnStyles: {
                0: { cellWidth: colWidths.no,      halign: 'center' },
                1: { cellWidth: colWidths.tgl,     halign: 'center' },
                2: { cellWidth: colWidths.voucher  },
                3: { cellWidth: colWidths.ket      },
                4: { cellWidth: colWidths.nominal,  halign: 'right'  }
            },
            margin: { left: 14, right: 14 },
        });
    });

    const safeName = clientName.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
    doc.save(`Test_of_Detail_${safeName}.pdf`);
}

// --- HELPER UNTUK EXCELJS ---
function copyRows(srcSheet, destSheet, srcStartRow, srcEndRow, destStartRow, sheetNameForReplace) {
    const rowOffset = destStartRow - srcStartRow;
    for (let r = srcStartRow; r <= srcEndRow; r++) {
        const srcRow = srcSheet.getRow(r);
        const destRow = destSheet.getRow(r + rowOffset);
        if (srcRow.height) destRow.height = srcRow.height;
        srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const destCell = destRow.getCell(colNumber);
            if (cell.value && sheetNameForReplace && cell.value.toString().includes('<<NamaAkun>>')) {
                destCell.value = cell.value.toString().replace('<<NamaAkun>>', sheetNameForReplace);
            } else {
                destCell.value = cell.value;
            }
            if (cell.style) destCell.style = JSON.parse(JSON.stringify(cell.style));
            if (cell.dataValidation) destCell.dataValidation = cell.dataValidation;
        });
    }
    if (srcSheet.model.merges) {
        srcSheet.model.merges.forEach(mergeRange => {
            const range = parseRangeString(mergeRange);
            if (range && range.top >= srcStartRow && range.bottom <= srcEndRow) {
                try { destSheet.mergeCells(range.top + rowOffset, range.left, range.bottom + rowOffset, range.right); } catch(e) {}
            }
        });
    }
}

function parseRangeString(rangeStr) {
    try {
        const parts = rangeStr.split(':');
        if (parts.length !== 2) return null;
        const decode = (ref) => {
            const match = ref.match(/([A-Z]+)(\d+)/);
            if (!match) return null;
            let colNum = 0;
            for (let i = 0; i < match[1].length; i++) colNum = colNum * 26 + (match[1].charCodeAt(i) - 64);
            return { row: parseInt(match[2]), col: colNum };
        };
        const start = decode(parts[0]);
        const end = decode(parts[1]);
        if (!start || !end) return null;
        return { top: Math.min(start.row, end.row), left: Math.min(start.col, end.col), bottom: Math.max(start.row, end.row), right: Math.max(start.col, end.col) };
    } catch (e) { return null; }
}
