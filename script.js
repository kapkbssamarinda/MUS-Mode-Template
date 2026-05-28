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

// --- FUNGSI HELPER MATERIALITAS ---
// Fungsi ini memanggil input materialitas dari HTML dan mengubahnya menjadi angka murni (Integer)
function getNilaiMaterialitas() {
    const inputElement = document.getElementById('materialitas');
    if (!inputElement) return 0; // Jika input tidak ditemukan (fallback aman)
    
    // Hapus semua titik untuk mendapatkan angka murni
    const val = inputElement.value;
    const angkaMurni = val.replace(/\./g, '');
    
    return parseInt(angkaMurni, 10) || 0;
}

// --- LOGIKA UTAMA ---
async function processAuditWorkpaper() {
    // 1. Validasi Input File
    if (!inputExcel.files.length) {
        Swal.fire({ icon: 'warning', title: 'File Kosong', text: 'Upload file input dulu ya!', confirmButtonColor: '#0d6efd' });
        return;
    }

    // 2. Ambil Data Form Informasi
    const clientName = document.getElementById('clientName').value || "";
    const period = document.getElementById('period').value || "";
    const preparedBy = document.getElementById('preparedBy').value || "";
    
    // Ambil Batas Materialitas
    const batasMaterialitas = getNilaiMaterialitas();
    
    // Format Tanggal (YYYY-MM-DD ke DD/MM/YYYY)
    const formatDate = (dateStr) => {
        if(!dateStr) return "";
        const d = new Date(dateStr);
        return `${d.getDate().toString().padStart(2,'0')}/${(d.getMonth()+1).toString().padStart(2,'0')}/${d.getFullYear()}`;
    };
    
    const preparedDate = formatDate(document.getElementById('preparedDate').value);
    const reviewedBy = document.getElementById('reviewedBy').value || "";
    const reviewedDate = formatDate(document.getElementById('reviewedDate').value);

    // Loading
    Swal.fire({ title: 'Sedang Memproses...', html: 'Menyiapkan kertas kerja audit...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });

    try {
        // LOAD FILES
        const workbookInput = new ExcelJS.Workbook();
        await workbookInput.xlsx.load(await inputExcel.files[0].arrayBuffer());

        const response = await fetch('assets/Template_Output.xlsx');
        if (!response.ok) throw new Error("Gagal load Template_Output.xlsx");
        
        const workbookTemplate = new ExcelJS.Workbook();
        await workbookTemplate.xlsx.load(await response.arrayBuffer());
        const masterSheet = workbookTemplate.worksheets[0];

        // CONFIG
        const HEADER_ROWS = 12;
        const START_DATA_ROW = 13;
        const FOOTER_START_ROW = 45;
        const FOOTER_GAP = 5;

        const workbookResult = new ExcelJS.Workbook();
        
        // LOOP SHEETS
        workbookInput.eachSheet((inputSheet, sheetId) => {
            const sheetName = inputSheet.name;
            const newSheet = workbookResult.addWorksheet(sheetName);

            // A. Copy Width
            masterSheet.columns.forEach((col, index) => {
                const newCol = newSheet.getColumn(index + 1);
                newCol.width = col.width;
                if(col.style) newCol.style = col.style;
            });

            newSheet.getColumn(4).width = 60;

            // B. Copy Header Template (Baris 1-12)
            copyRows(masterSheet, newSheet, 1, HEADER_ROWS, 1, sheetName);

            // --- INJEKSI DATA FORMULIR KE HEADER ---
            // Berdasarkan Template Output Anda:
            // Klien: B7, Periode: B8
            newSheet.getCell('B7').value = clientName;
            newSheet.getCell('B8').value = period;

            // Dibuat Oleh: V7, Tanggal: V8
            newSheet.getCell('V7').value = preparedBy;
            newSheet.getCell('V8').value = preparedDate;

            // Direview Oleh: Y7, Tanggal: Y8
            newSheet.getCell('Y7').value = reviewedBy;
            newSheet.getCell('Y8').value = reviewedDate;

            // Styling Font Header Injeksi (Opsional: Bold)
            ['B7', 'B8', 'V7', 'V8', 'Y7', 'Y8'].forEach(addr => {
                newSheet.getCell(addr).font = { name: 'Arial', size: 10, bold: true };
            });


            // C. Sampling Logic (REVISED: Sort DESC → Filter Materialitas → Top 15 → Random 15)
            let dataRaw = [];
            inputSheet.eachRow((row, rowNum) => {
                if (rowNum > 1) { 
                    let nominal = row.getCell(4).value;
                    if (typeof nominal !== 'number') nominal = parseFloat(nominal) || 0;
                    dataRaw.push({
                        idBaris: rowNum,
                        tgl: row.getCell(1).value || '',
                        voucher: row.getCell(2).value || '',
                        ket: row.getCell(3).value || '',
                        nominal: nominal
                    });
                }
            });

            // 1. SORT: Terbesar ke Terkecil
            dataRaw.sort((a, b) => b.nominal - a.nominal);
            
            // 2. FILTER: Hanya ambil yang nominal >= materialitas
            let dataFiltered = dataRaw.filter(item => Math.abs(item.nominal) >= batasMaterialitas);
            
            // 3. TOP 15: Ambil 15 teratas dari data yang sudah difilter
            const top15 = dataFiltered.slice(0, 15);
            
            // 4. SISA DATA: Ambil yang tidak termasuk top 15
            let sisaDataFiltered = dataFiltered.slice(15);
            
            // 5. RANDOM 15: Acak dan ambil maksimal 15 dari sisa data
            let random15 = [];
            if (sisaDataFiltered.length > 0) {
                // Fisher-Yates Shuffle untuk randomisasi yang fair
                for (let i = sisaDataFiltered.length - 1; i > 0; i--) {
                    const j = Math.floor(Math.random() * (i + 1));
                    [sisaDataFiltered[i], sisaDataFiltered[j]] = [sisaDataFiltered[j], sisaDataFiltered[i]];
                }
                // Ambil maksimal 15 setelah diacak
                random15 = sisaDataFiltered.slice(0, 15);
            }
            
            // 6. FINAL: Gabungkan Top 15 + Random 15
            const finalSamples = [...top15, ...random15];
            
            // D. Tulis Data
            let currentRowIdx = START_DATA_ROW; 
            const templateDataRow = masterSheet.getRow(START_DATA_ROW);

            finalSamples.forEach((item, index) => {
                const row = newSheet.getRow(currentRowIdx);
                if(templateDataRow.height) row.height = templateDataRow.height;

                row.getCell(1).value = index + 1;
                row.getCell(2).value = item.tgl;
                row.getCell(3).value = item.voucher;
                row.getCell(4).value = item.ket;
                row.getCell(5).value = item.nominal;

                // N/A (J-U)
                for (let c = 10; c <= 21; c++) row.getCell(c).value = "N/A";

                // Dropdown Asersi (V-Y)
                for (let c = 22; c <= 25; c++) {
                    row.getCell(c).value = ""; 
                    row.getCell(c).dataValidation = {
                        type: 'list', allowBlank: true, formulae: ['"V,X"'],
                        showErrorMessage: true, error: 'Pilih V atau X'
                    };
                }

                // Styling
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

            // F. Footer Keterangan
            const footerDestRow = currentRowIdx + FOOTER_GAP + 1; 
            const footerRowCount = masterSheet.rowCount - FOOTER_START_ROW + 1;
            if (footerRowCount > 0) {
                copyRows(masterSheet, newSheet, FOOTER_START_ROW, masterSheet.rowCount, footerDestRow, sheetName);
            }
        });

        // 3. DOWNLOAD
        const buffer = await workbookResult.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        saveAs(blob, `Kertas_Kerja_${clientName.replace(/ /g,"_") || 'Audit'}.xlsx`);
        
        Swal.fire({ icon: 'success', title: 'Berhasil!', text: 'Kertas Kerja Audit telah diunduh.', confirmButtonColor: '#198754' });

    } catch (error) {
        console.error(error);
        Swal.fire({ icon: 'error', title: 'Error', text: error.message });
    }
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
        const parts = rangeStr.split(':'); if (parts.length !== 2) return null;
        const decode = (ref) => {
            const match = ref.match(/([A-Z]+)(\d+)/); if (!match) return null;
            let colStr = match[1], colNum = 0;
            for (let i = 0; i < colStr.length; i++) colNum = colNum * 26 + (colStr.charCodeAt(i) - 64);
            return { row: parseInt(match[2]), col: colNum };
        };
        const start = decode(parts[0]); const end = decode(parts[1]);
        if (!start || !end) return null;
        return { top: Math.min(start.row, end.row), left: Math.min(start.col, end.col), bottom: Math.max(start.row, end.row), right: Math.max(start.col, end.col) };
    } catch (e) { return null; }
}