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
    
    // Validasi setiap field
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
    
    // Jika ada error, tampilkan
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

// --- FUNGSI HELPER MATERIALITAS ---
function getNilaiMaterialitas() {
    const inputElement = document.getElementById('materialitas');
    if (!inputElement) return 0;
    
    const val = inputElement.value;
    const angkaMurni = val.replace(/\./g, '');
    
    return parseInt(angkaMurni, 10) || 0;
}

// --- FUNGSI HELPER SORTING OPTION ---
function getSortingOption() {
    const sortingOption = document.querySelector('input[name="sortingOption"]:checked');
    return sortingOption ? sortingOption.value : 'date';
}

// --- LOGIKA UTAMA ---
async function processAuditWorkpaper() {
    // 1. Validasi Input
    if (!validateForm()) {
        return;
    }

    // 2. Ambil Data Form Informasi
    const clientName = document.getElementById('clientName').value || "";
    const period = document.getElementById('period').value || "";
    const preparedBy = document.getElementById('preparedBy').value || "";
    const batasMaterialitas = getNilaiMaterialitas();
    const sortingOption = getSortingOption();
    
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
            newSheet.getCell('B7').value = clientName;
            newSheet.getCell('B8').value = period;

            newSheet.getCell('V7').value = preparedBy;
            newSheet.getCell('V8').value = preparedDate;

            newSheet.getCell('Y7').value = reviewedBy;
            newSheet.getCell('Y8').value = reviewedDate;

            ['B7', 'B8', 'V7', 'V8', 'Y7', 'Y8'].forEach(addr => {
                newSheet.getCell(addr).font = { name: 'Arial', size: 10, bold: true };
            });


            // C. Sampling Logic (REVISED: Sort DESC → Filter Materialitas → Top 15 → Random 15)
            let dataRaw = [];
            inputSheet.eachRow((row, rowNum) => {
                if (rowNum > 1) { 
                    let nominal = row.getCell(4).value;
                    if (typeof nominal !== 'number') nominal = parseFloat(nominal) || 0;
                    
                    let tglValue = row.getCell(1).value;
                    let tglTimestamp = null;
                    
                    if (rowNum <= 10) {
                        console.log(`[Sheet: ${inputSheet.name}] Row ${rowNum} - tglValue:`, tglValue, `| Type:`, typeof tglValue, `| Raw:`, JSON.stringify(tglValue));
                    }
                    
                    const parseToTimestamp = (val) => {
                        if (!val) return 0;
                        
                        if (val instanceof Date) {
                            const ts = val.getTime();
                            if (rowNum <= 10) console.log(`  → Parsed as Date: ${ts}`);
                            return ts;
                        }
                        
                        if (typeof val === 'number') {
                            let ts;
                            if (val < 50000) {
                                ts = new Date(1900, 0, 1 + val).getTime();
                                if (rowNum <= 10) console.log(`  → Parsed as Excel serial: ${ts}`);
                            } else {
                                ts = val;
                                if (rowNum <= 10) console.log(`  → Parsed as timestamp: ${ts}`);
                            }
                            return ts;
                        }
                        
                        if (typeof val === 'string') {
                            if (val.includes('/') && !val.includes('-')) {
                                const parts = val.split('/');
                                if (parts.length === 3) {
                                    const day = parseInt(parts[0], 10);
                                    const month = parseInt(parts[1], 10);
                                    const year = parseInt(parts[2], 10);
                                    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
                                        const ts = new Date(year, month - 1, day).getTime();
                                        if (rowNum <= 10) console.log(`  → Parsed as dd/mm/yyyy: ${ts}`);
                                        return ts;
                                    }
                                }
                            }
                            
                            if (val.includes('-')) {
                                const parsed = new Date(val);
                                if (!isNaN(parsed.getTime())) {
                                    const ts = parsed.getTime();
                                    if (rowNum <= 10) console.log(`  → Parsed as ISO/yyyy-mm-dd: ${ts}`);
                                    return ts;
                                }
                            }
                            
                            const parsed = new Date(val);
                            if (!isNaN(parsed.getTime())) {
                                const ts = parsed.getTime();
                                if (rowNum <= 10) console.log(`  → Parsed as generic string: ${ts}`);
                                return ts;
                            }
                        }
                        
                        if (rowNum <= 10) console.log(`  → FAILED to parse`);
                        return 0;
                    };
                    
                    tglTimestamp = parseToTimestamp(tglValue);
                    
                    dataRaw.push({
                        idBaris: rowNum,
                        tgl: tglValue,
                        tglTimestamp: tglTimestamp,
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
                for (let i = sisaDataFiltered.length - 1; i > 0; i--) {
                    const j = Math.floor(Math.random() * (i + 1));
                    [sisaDataFiltered[i], sisaDataFiltered[j]] = [sisaDataFiltered[j], sisaDataFiltered[i]];
                }
                random15 = sisaDataFiltered.slice(0, 15);
            }
            
            // 6. FINAL: Gabungkan Top 15 + Random 15
            const finalSamples = [...top15, ...random15];
            
            // 7. SORT BERDASARKAN OPSI YANG DIPILIH
            if (sortingOption === 'date') {
                // SORT TANGGAL PAKSA: Urutkan dari tanggal awal ke akhir (Januari - Desember)
                finalSamples.sort((a, b) => {
                    return a.tglTimestamp - b.tglTimestamp;
                });
                
                console.log(`\n[Sheet: ${inputSheet.name}] SETELAH SORTING PER TANGGAL:`);
                finalSamples.slice(0, 10).forEach((item, idx) => {
                    console.log(`  ${idx + 1}. ${item.tgl} (timestamp: ${item.tglTimestamp})`);
                });
                console.log('...\n');
            } else if (sortingOption === 'nominal') {
                // SORT NOMINAL: Urutkan dari nominal terbesar ke terkecil
                finalSamples.sort((a, b) => b.nominal - a.nominal);
                
                console.log(`\n[Sheet: ${inputSheet.name}] SETELAH SORTING PER NOMINAL:`);
                finalSamples.slice(0, 10).forEach((item, idx) => {
                    console.log(`  ${idx + 1}. ${item.tgl} - Rp ${item.nominal.toLocaleString('id-ID')}`);
                });
                console.log('...\n');
            }
            
            // D. Tulis Data
            let currentRowIdx = START_DATA_ROW; 
            const templateDataRow = masterSheet.getRow(START_DATA_ROW);

            finalSamples.forEach((item, index) => {
                const row = newSheet.getRow(currentRowIdx);
                if(templateDataRow.height) row.height = templateDataRow.height;

                row.getCell(1).value = index + 1;
                
                // Format Tanggal ke dd/mm/yyyy
                let tglFormatted = '';
                if (item.tgl) {
                    if (typeof item.tgl === 'string') {
                        tglFormatted = item.tgl;
                    } 
                    else if (item.tgl instanceof Date) {
                        const day = item.tgl.getDate().toString().padStart(2, '0');
                        const month = (item.tgl.getMonth() + 1).toString().padStart(2, '0');
                        const year = item.tgl.getFullYear();
                        tglFormatted = `${day}/${month}/${year}`;
                    }
                    else if (typeof item.tgl === 'number' && item.tglTimestamp > 0) {
                        const dateObj = new Date(item.tglTimestamp);
                        const day = dateObj.getDate().toString().padStart(2, '0');
                        const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
                        const year = dateObj.getFullYear();
                        tglFormatted = `${day}/${month}/${year}`;
                    }
                }
                
                row.getCell(2).value = tglFormatted;
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