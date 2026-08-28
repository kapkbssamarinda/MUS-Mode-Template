// script.js — AuditWorkpaper Pro Core Engine & Simulation Suite

// --- STATE & CACHE ---
let cachedTemplateBuffer = null;
let deferredPwaPrompt = null;
let activeDemoScenario = null;

// --- DOM ELEMENTS ---
const htmlRoot = document.documentElement;
const themeToggleBtn = document.getElementById('themeToggleBtn');
const themeIcon = document.getElementById('themeIcon');
const metaThemeColor = document.getElementById('metaThemeColor');

const inputExcel = document.getElementById('inputExcel');
const dropZone = document.getElementById('dropZone');
const dropTitle = document.getElementById('dropTitle');
const dropDesc = document.getElementById('dropDesc');
const dropIcon = document.getElementById('dropIcon');

const materialitasInput = document.getElementById('materialitas');
const btnExportExcel = document.getElementById('btnExportExcel');
const btnExportPdf = document.getElementById('btnExportPdf');
const btnResetForm = document.getElementById('btnResetForm');
const btnClearDemo = document.getElementById('btnClearDemo');
const demoActiveBanner = document.getElementById('demoActiveBanner');
const demoBannerScenarioTitle = document.getElementById('demoBannerScenarioTitle');
const demoBannerScenarioDesc = document.getElementById('demoBannerScenarioDesc');

const btnOpenDemoModal = document.getElementById('btnOpenDemoModal');
const btnApplyDemo = document.getElementById('btnApplyDemo');
const pwaBanner = document.getElementById('pwaBanner');
const btnPwaInstall = document.getElementById('btnPwaInstall');
const btnPwaDismiss = document.getElementById('btnPwaDismiss');

// --- THEME MANAGEMENT ---
function initTheme() {
    const savedTheme = localStorage.getItem('mus_theme');
    const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    const initialTheme = savedTheme || (prefersDark ? 'dark' : 'light');
    applyTheme(initialTheme);
}

function applyTheme(theme) {
    htmlRoot.setAttribute('data-theme', theme);
    localStorage.setItem('mus_theme', theme);

    if (theme === 'dark') {
        if (themeIcon) {
            themeIcon.classList.remove('fa-moon');
            themeIcon.classList.add('fa-sun');
        }
        if (metaThemeColor) metaThemeColor.setAttribute('content', '#090d16');
    } else {
        if (themeIcon) {
            themeIcon.classList.remove('fa-sun');
            themeIcon.classList.add('fa-moon');
        }
        if (metaThemeColor) metaThemeColor.setAttribute('content', '#f8fafc');
    }
}

if (themeToggleBtn) {
    themeToggleBtn.addEventListener('click', () => {
        const currentTheme = htmlRoot.getAttribute('data-theme') || 'light';
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        applyTheme(newTheme);
    });
}

// --- INITIALIZE DATES, MODAL & YEAR ---
document.addEventListener('DOMContentLoaded', () => {
    initTheme();

    const preparedDateInput = document.getElementById('preparedDate');
    const reviewedDateInput = document.getElementById('reviewedDate');
    const yearSpan = document.getElementById('year');

    const today = new Date();
    if (preparedDateInput && !preparedDateInput.value) preparedDateInput.valueAsDate = today;
    if (reviewedDateInput && !reviewedDateInput.value) reviewedDateInput.valueAsDate = today;
    if (yearSpan) yearSpan.textContent = today.getFullYear().toString();

    // Check PWA Standalone Mode
    const isStandalone = window.matchMedia('(display-mode: standalone)').matches || 
                         window.navigator.standalone ||
                         document.referrer.includes('android-app://');
    if (isStandalone) {
        document.body.classList.add('pwa-mode');
    }

    initScenarioCardSelection();
    initDemoModalTrigger();
});

// --- OPEN DEMO MODAL TRIGGER ---
function initDemoModalTrigger() {
    if (btnOpenDemoModal) {
        btnOpenDemoModal.addEventListener('click', (e) => {
            e.preventDefault();
            const demoModalEl = document.getElementById('demoModal');
            if (demoModalEl && window.bootstrap) {
                const modal = bootstrap.Modal.getOrCreateInstance(demoModalEl);
                modal.show();
            }
        });
    }
}

// --- NUMBER FORMATTING (MATERIALITAS) ---
function formatRibuanString(val) {
    let angka = val.replace(/\D/g, '');
    if (angka !== '') {
        angka = parseInt(angka, 10).toString();
        let sisa = angka.length % 3;
        let rupiah = angka.substr(0, sisa);
        let ribuan = angka.substr(sisa).match(/\d{3}/g);
        
        if (ribuan) {
            let separator = sisa ? '.' : '';
            rupiah += separator + ribuan.join('.');
        }
        return rupiah;
    }
    return '';
}

if (materialitasInput) {
    materialitasInput.addEventListener('input', (e) => {
        const cursorPos = e.target.selectionStart;
        const prevLength = e.target.value.length;
        e.target.value = formatRibuanString(e.target.value);
        const newLength = e.target.value.length;
        const newPos = Math.max(0, cursorPos + (newLength - prevLength));
        e.target.setSelectionRange(newPos, newPos);
    });
}

const customMaterialityInput = document.getElementById('customMateriality');
if (customMaterialityInput) {
    customMaterialityInput.addEventListener('input', (e) => {
        e.target.value = formatRibuanString(e.target.value);
    });
}

// --- RADIO BUTTON INTERACTION ---
const radioSortByDate = document.getElementById('sortByDate');
const radioSortByNominal = document.getElementById('sortByNominal');
const labelSortDate = document.getElementById('labelSortDate');
const labelSortNominal = document.getElementById('labelSortNominal');

function updateRadioStyles() {
    if (radioSortByDate && labelSortDate) {
        labelSortDate.classList.toggle('active', radioSortByDate.checked);
    }
    if (radioSortByNominal && labelSortNominal) {
        labelSortNominal.classList.toggle('active', radioSortByNominal.checked);
    }
}

if (radioSortByDate) radioSortByDate.addEventListener('change', updateRadioStyles);
if (radioSortByNominal) radioSortByNominal.addEventListener('change', updateRadioStyles);

// --- DROPZONE & FILE UPLOAD ---
if (dropZone && inputExcel) {
    dropZone.addEventListener('click', () => inputExcel.click());
    
    // Keyboard Accessibility (Enter & Space)
    dropZone.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            inputExcel.click();
        }
    });

    inputExcel.addEventListener('change', function() {
        handleFiles(this.files);
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer && e.dataTransfer.files.length > 0) {
            handleFiles(e.dataTransfer.files);
            inputExcel.files = e.dataTransfer.files;
        }
    });
}

function handleFiles(files) {
    if (files && files.length > 0) {
        const file = files[0];
        dropZone.classList.add('ready');
        dropTitle.innerText = "File Siap Diproses";
        dropDesc.innerHTML = `<span class="text-success fw-bold"><i class="fa-solid fa-circle-check me-1" aria-hidden="true"></i> ${escapeHtml(file.name)}</span>`;
        if (dropIcon) {
            dropIcon.className = "fa-solid fa-file-circle-check fa-2x";
        }
    }
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// --- RESET FORM & CLEAR DEMO ---
function resetFormToDefault() {
    document.getElementById('clientName').value = '';
    document.getElementById('period').value = '31 Desember 2024';
    document.getElementById('materialitas').value = '20.000';
    document.getElementById('topSamples').value = '';
    document.getElementById('randomSamples').value = '';
    document.getElementById('preparedBy').value = '';
    document.getElementById('reviewedBy').value = '';
    document.getElementById('preparedDate').valueAsDate = new Date();
    document.getElementById('reviewedDate').valueAsDate = new Date();
    
    if (radioSortByDate) {
        radioSortByDate.checked = true;
        updateRadioStyles();
    }

    inputExcel.value = '';
    dropZone.classList.remove('ready');
    dropTitle.innerText = "Unggah File Input (.xlsx)";
    dropDesc.innerText = "Klik atau seret file Excel populasi ke area ini";
    if (dropIcon) dropIcon.className = "fa-solid fa-cloud-arrow-up fa-2x";

    if (demoActiveBanner) demoActiveBanner.classList.add('d-none');
    activeDemoScenario = null;
}

if (btnResetForm) btnResetForm.addEventListener('click', resetFormToDefault);
if (btnClearDemo) btnClearDemo.addEventListener('click', resetFormToDefault);

// --- SCENARIO CARD SELECTION IN MODAL ---
function initScenarioCardSelection() {
    const scenarioCards = document.querySelectorAll('.scenario-card');
    scenarioCards.forEach(card => {
        card.addEventListener('click', () => {
            scenarioCards.forEach(c => {
                c.classList.remove('active');
                c.setAttribute('aria-pressed', 'false');
                const check = c.querySelector('.scenario-check');
                if (check) {
                    check.className = 'fa-regular fa-circle text-muted scenario-check';
                }
            });
            card.classList.add('active');
            card.setAttribute('aria-pressed', 'true');
            const check = card.querySelector('.scenario-check');
            if (check) {
                check.className = 'fa-solid fa-check-circle text-primary scenario-check';
            }
        });

        card.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                card.click();
            }
        });
    });
}

// --- AUDIT SIMULATION ENGINE (COMPLEX DEMO SUITE) ---
const AuditSimulationEngine = {
    descriptions: {
        kas: [
            "Penerimaan Pelunasan Piutang Pelanggan",
            "Pembayaran Faktur Tagihan Pemasok",
            "Pembayaran Biaya Gaji dan Upah Karyawan",
            "Penyetoran Pajak Penghasilan PPh 21/23",
            "Pembayaran Biaya Utilitas Listrik dan Air",
            "Pembayaran Premi Asuransi Gedung dan Aset",
            "Pencairan Cek Giro Operasional",
            "Pembayaran Sewa Kantor dan Gudang",
            "Penerimaan Pendapatan Bunga Deposito",
            "Biaya Administrasi Bank dan Buku Cek"
        ],
        piutang: [
            "Faktur Penjualan Barang Dagang Batch A",
            "Faktur Penjualan Komersial Ekspor",
            "Tagihan Termin Proyek Tahap Penyelesaian",
            "Penjualan Produk Grosir Pelanggan Utama",
            "Nota Kredit Retur Penjualan Produk Rusak",
            "Tagihan Jasa Pemeliharaan Sistem",
            "Penjualan Kredit Distributor Wilayah Barat",
            "Penyesuaian Diskon Pelunasan Dini",
            "Faktur Pengiriman Barang Cabang",
            "Tagihan Penggantian Biaya Operasional"
        ],
        persediaan_hpp: [
            "Pembelian Bahan Baku Impor Batch 1",
            "Pembelian Komponen Produksi Lokal",
            "Alokasi Biaya Tenaga Kerja Langsung",
            "Biaya Overhead Pabrik dan Pengemasan",
            "Retur Pembelian Bahan Baku Cacat",
            "Penerimaan Barang Dagang dari Vendor",
            "Penyesuaian Nilai Bersih Realisasi Persediaan",
            "Biaya Pengiriman dan Asuransi Pengangkutan",
            "Pemakaian Bahan Penolong Produksi",
            "Pembelian Perlengkapan Operasional Gudang"
        ]
    },

    getRandomDate(year = 2024) {
        const start = new Date(year, 0, 1).getTime();
        const end = new Date(year, 11, 31).getTime();
        const d = new Date(start + Math.random() * (end - start));
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        return `${dd}/${mm}/${year}`;
    },

    getRandomVoucher(prefix, index, year = 2024) {
        const month = String(Math.floor(Math.random() * 12) + 1).padStart(2, '0');
        const num = String(index + 1).padStart(4, '0');
        return `${prefix}-${year}/${month}/${num}`;
    },

    getRandomAmount(baseScale, returnRate = 0.08) {
        // Distribusi Log-normal untuk mencerminkan transaksi buku besar akuntansi nyata
        const isNegative = Math.random() < returnRate;
        const u1 = Math.random();
        const u2 = Math.random();
        const randStdNormal = Math.sqrt(-2.0 * Math.log(u1 || 0.0001)) * Math.cos(2.0 * Math.PI * u2);
        
        let amount = Math.exp(12 + randStdNormal * 1.5) * (baseScale / 500000);
        amount = Math.round(amount / 1000) * 1000; // Round to ribuan

        // Pastikan tidak nol
        if (amount < 100000) amount = Math.floor(Math.random() * 900000) + 100000;
        return isNegative ? -Math.abs(amount) : Math.abs(amount);
    },

    async generateWorkbook(scenarioConfig) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "AuditWorkpaper Pro Simulation Engine";
        workbook.created = new Date();

        let totalRows = 0;
        let totalSum = 0;

        for (const sheetConf of scenarioConfig.sheets) {
            const sheet = workbook.addWorksheet(sheetConf.name);
            sheet.columns = [
                { header: 'Tanggal', key: 'tgl', width: 14 },
                { header: 'Voucher', key: 'voucher', width: 22 },
                { header: 'Keterangan', key: 'ket', width: 48 },
                { header: 'Nominal', key: 'nominal', width: 22 }
            ];

            const headerRow = sheet.getRow(1);
            headerRow.font = { name: 'Arial', size: 10, bold: true };
            headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

            const categoryKey = sheetConf.category || 'kas';
            const descList = this.descriptions[categoryKey] || this.descriptions.kas;

            for (let i = 0; i < sheetConf.rowCount; i++) {
                const dateStr = this.getRandomDate(2024);
                const voucherStr = this.getRandomVoucher(sheetConf.voucherPrefix || 'JV', i);
                const descStr = descList[Math.floor(Math.random() * descList.length)] + ` #${i + 1}`;
                const nominalVal = this.getRandomAmount(sheetConf.baseScale || 10000000, scenarioConfig.returnRate || 0.08);

                const row = sheet.addRow({
                    tgl: dateStr,
                    voucher: voucherStr,
                    ket: descStr,
                    nominal: nominalVal
                });

                row.getCell(4).numFmt = '#,##0.00;(#,##0.00);"-"';
                totalRows++;
                totalSum += nominalVal;
            }
        }

        const buffer = await workbook.xlsx.writeBuffer();
        return { buffer, totalRows, totalSum };
    }
};

// --- APPLY SIMULATION DEMO DATA ---
if (btnApplyDemo) {
    btnApplyDemo.addEventListener('click', async () => {
        const activeTab = document.querySelector('#demoTab .nav-link.active');
        const isCustom = activeTab && activeTab.id === 'custom-tab';

        Swal.fire({
            title: 'Menyiapkan Data Simulasi...',
            html: 'Menyusun transaksi buku besar multi-sheet di memori...',
            allowOutsideClick: false,
            didOpen: () => Swal.showLoading()
        });

        try {
            let config = null;

            if (!isCustom) {
                const activeCard = document.querySelector('.scenario-card.active');
                const scenarioType = activeCard ? activeCard.getAttribute('data-scenario') : 'manufaktur';

                if (scenarioType === 'ritel') {
                    config = {
                        clientName: "PT Nusantara Retail Mart",
                        period: "31 Desember 2024",
                        materialitas: "15.000.000",
                        preparedBy: "Siti Rahma (Auditor)",
                        reviewedBy: "Ahmad Fauzi, CA (Manager Audit)",
                        scenarioTitle: "Skenario Ritel & Barang Konsumsi",
                        scenarioDesc: "PT Nusantara Retail Mart (2 Akun, 350 Transaksi Populasi)",
                        returnRate: 0.06,
                        sheets: [
                            { name: "1101 - Kas Kasir & Induk", category: "kas", voucherPrefix: "BKK", rowCount: 180, baseScale: 8000000 },
                            { name: "1105 - Persediaan Barang Dagang", category: "persediaan_hpp", voucherPrefix: "GRN", rowCount: 170, baseScale: 18000000 }
                        ]
                    };
                } else if (scenarioType === 'jasa_it') {
                    config = {
                        clientName: "PT Sinergi Solusi Digital",
                        period: "31 Desember 2024",
                        materialitas: "35.000.000",
                        preparedBy: "Reza Pratama (Senior Associate)",
                        reviewedBy: "Diana Kusuma, CPA (Partner)",
                        scenarioTitle: "Skenario Jasa IT & Konsultan",
                        scenarioDesc: "PT Sinergi Solusi Digital (2 Akun, 160 Transaksi Populasi)",
                        returnRate: 0.04,
                        sheets: [
                            { name: "1101 - Kas Bank BCA Operasional", category: "kas", voucherPrefix: "BKK", rowCount: 80, baseScale: 45000000 },
                            { name: "1103 - Piutang Termin Proyek", category: "piutang", voucherPrefix: "INV", rowCount: 80, baseScale: 60000000 }
                        ]
                    };
                } else {
                    // Default Manufaktur
                    config = {
                        clientName: "PT Mega Perkasa Abadi Tbk",
                        period: "31 Desember 2024",
                        materialitas: "50.000.000",
                        preparedBy: "Viany Ramadhany (Senior Auditor)",
                        reviewedBy: "Budi Santoso, CPA (Audit Partner)",
                        scenarioTitle: "Skenario Manufaktur & Perdagangan",
                        scenarioDesc: "PT Mega Perkasa Abadi Tbk (3 Akun, 250 Transaksi Populasi)",
                        returnRate: 0.08,
                        sheets: [
                            { name: "1102 - Kas Bank Mandiri", category: "kas", voucherPrefix: "BKK", rowCount: 90, baseScale: 75000000 },
                            { name: "1104 - Piutang Usaha", category: "piutang", voucherPrefix: "INV", rowCount: 80, baseScale: 85000000 },
                            { name: "5101 - Beban Pokok Penjualan", category: "persediaan_hpp", voucherPrefix: "JV", rowCount: 80, baseScale: 95000000 }
                        ]
                    };
                }
            } else {
                // Custom Generator
                const client = document.getElementById('customClientName').value.trim() || "PT Simulasi Audit Nusantara";
                const mat = document.getElementById('customMateriality').value.trim() || "25.000.000";
                const accCount = parseInt(document.getElementById('customAccountCount').value, 10) || 3;
                const rowsPerAcc = parseInt(document.getElementById('customRowCount').value, 10) || 100;
                const retPercent = (parseInt(document.getElementById('customReturnPercent').value, 10) || 10) / 100;

                const sampleAccountNames = [
                    { name: "1101 - Kas dan Setara Kas", cat: "kas", pref: "BKK" },
                    { name: "1103 - Piutang Usaha Pihak Ketiga", cat: "piutang", pref: "INV" },
                    { name: "1105 - Persediaan Barang Dagang", cat: "persediaan_hpp", pref: "GRN" },
                    { name: "5101 - Beban Pokok Penjualan", cat: "persediaan_hpp", pref: "JV" },
                    { name: "6102 - Beban Operasional Umum", cat: "kas", pref: "BKK" }
                ];

                const selectedSheets = [];
                for (let i = 0; i < accCount; i++) {
                    const acc = sampleAccountNames[i] || { name: `Akun Uji ${i + 1}`, cat: "kas", pref: "JV" };
                    selectedSheets.push({
                        name: acc.name,
                        category: acc.cat,
                        voucherPrefix: acc.pref,
                        rowCount: rowsPerAcc,
                        baseScale: 30000000
                    });
                }

                config = {
                    clientName: client,
                    period: "31 Desember 2024",
                    materialitas: mat,
                    preparedBy: "Auditor Simulasi",
                    reviewedBy: "Manajer Audit",
                    scenarioTitle: "Generator Kustom",
                    scenarioDesc: `${client} (${accCount} Akun, ${accCount * rowsPerAcc} Transaksi)`,
                    returnRate: retPercent,
                    sheets: selectedSheets
                };
            }

            // Generate the in-memory Excel file
            const { buffer, totalRows, totalSum } = await AuditSimulationEngine.generateWorkbook(config);

            const fileName = `Populasi_GL_${config.clientName.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`;
            const file = new File([buffer], fileName, {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            // Set file to inputExcel via DataTransfer
            const dt = new DataTransfer();
            dt.items.add(file);
            inputExcel.files = dt.files;
            handleFiles([file]);

            // Auto-fill assignment form
            document.getElementById('clientName').value = config.clientName;
            document.getElementById('period').value = config.period;
            document.getElementById('materialitas').value = config.materialitas;
            document.getElementById('preparedBy').value = config.preparedBy;
            document.getElementById('reviewedBy').value = config.reviewedBy;
            document.getElementById('preparedDate').valueAsDate = new Date();
            document.getElementById('reviewedDate').valueAsDate = new Date();

            // Show active demo banner
            if (demoActiveBanner) {
                demoBannerScenarioTitle.innerText = `Mode Simulasi: ${config.scenarioTitle}`;
                demoBannerScenarioDesc.innerText = `${config.scenarioDesc} | Total GL: ${formatRupiahGL(totalSum)}`;
                demoActiveBanner.classList.remove('d-none');
            }

            activeDemoScenario = config;

            // Close modal
            const modalEl = document.getElementById('demoModal');
            if (modalEl && window.bootstrap) {
                const modalInstance = bootstrap.Modal.getInstance(modalEl) || bootstrap.Modal.getOrCreateInstance(modalEl);
                if (modalInstance) modalInstance.hide();
            }

            Swal.fire({
                icon: 'success',
                title: 'Data Simulasi Berhasil Dimuat',
                html: `<div style="text-align: left; font-size: 0.9rem; line-height: 1.6;">` +
                      `<div><strong>Entitas:</strong> ${escapeHtml(config.clientName)}</div>` +
                      `<div><strong>Jumlah Akun:</strong> ${config.sheets.length} Akun Sheet</div>` +
                      `<div><strong>Total Populasi:</strong> ${totalRows} Baris Transaksi</div>` +
                      `<div><strong>Total Nilai Buku Besar:</strong> Rp ${formatRupiahGL(totalSum)}</div>` +
                      `<div><strong>Batas Materialitas:</strong> Rp ${config.materialitas}</div>` +
                      `</div>`,
                confirmButtonColor: '#1d4ed8',
                confirmButtonText: 'Tutup'
            });

        } catch (err) {
            console.error(err);
            Swal.fire({
                icon: 'error',
                title: 'Gagal Menghasilkan Data Simulasi',
                text: err.message || 'Terjadi kendala saat menyusun data di memori.',
                confirmButtonColor: '#1d4ed8'
            });
        }
    });
}

// --- FORM VALIDATION ---
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
    if (!inputExcel.files || !inputExcel.files.length) errors.push('File Excel Populasi harus diunggah');

    if (errors.length > 0) {
        Swal.fire({
            icon: 'warning',
            title: 'Parameter Belum Lengkap',
            html: '<ul style="text-align: left; margin: 0; padding-left: 1.25rem;">' +
                  errors.map(err => `<li style="margin-bottom: 4px;">${escapeHtml(err)}</li>`).join('') +
                  '</ul>',
            confirmButtonColor: '#1d4ed8'
        });
        return false;
    }
    return true;
}

// --- HELPER FUNCTIONS ---
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

// --- SAMPLING LOGIC ---
function determineSampleCounts() {
    const topInput = document.getElementById('topSamples').value;
    const randomInput = document.getElementById('randomSamples').value;

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

// --- TEMPLATE BUFFER LOADER (WITH IN-MEMORY CACHING) ---
async function getTemplateArrayBuffer() {
    if (cachedTemplateBuffer) {
        return cachedTemplateBuffer.slice(0); // return copy
    }
    const response = await fetch('assets/Template_Output.xlsx');
    if (!response.ok) throw new Error("Gagal memuat template Excel (Template_Output.xlsx).");
    cachedTemplateBuffer = await response.arrayBuffer();
    return cachedTemplateBuffer.slice(0);
}

// --- CORE EXPORT DRIVER ---
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

    Swal.fire({
        title: 'Sedang Memproses Sampling...',
        html: 'Menghitung sampel Monetary Unit Sampling...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

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
            Swal.fire({
                icon: 'success',
                title: 'Berhasil!',
                text: 'Dokumen PDF Test of Details telah berhasil diunduh.',
                confirmButtonColor: '#b91c1c'
            });

        } else {
            // --- EXCEL EXPORT ---
            const templateBuffer = await getTemplateArrayBuffer();
            const workbookTemplate = new ExcelJS.Workbook();
            await workbookTemplate.xlsx.load(templateBuffer);
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

            Swal.fire({
                icon: 'success',
                title: 'Berhasil!',
                text: 'Kertas Kerja Audit Excel telah berhasil diunduh.',
                confirmButtonColor: '#15803d'
            });
        }

    } catch (error) {
        console.error(error);
        Swal.fire({
            icon: 'error',
            title: 'Terjadi Kesalahan',
            text: error.message || 'Gagal memproses kertas kerja audit.',
            confirmButtonColor: '#1d4ed8'
        });
    }
}

// --- GENERATE PDF ---
async function generatePDF(allSheetsData, clientName, period) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    const colWidths = { no: 8, tgl: 22, voucher: 45, ket: 72, nominal: 35 };

    allSheetsData.forEach((sheetData, idx) => {
        if (idx > 0) doc.addPage();

        const { sheetName, samples } = sheetData;

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
                fillColor: [240, 240, 240],
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
    doc.save(`Test_of_Detail_${safeName || 'Audit'}.pdf`);
}

// --- HELPER EXCELJS ROW COPYING ---
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

// --- ATTACH BUTTON LISTENERS ---
if (btnExportExcel) {
    btnExportExcel.addEventListener('click', () => processAuditWorkpaper('excel'));
}

if (btnExportPdf) {
    btnExportPdf.addEventListener('click', () => processAuditWorkpaper('pdf'));
}

// --- PWA INSTALL PROMPT & SERVICE WORKER ---
window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPwaPrompt = e;
    if (pwaBanner) pwaBanner.classList.remove('d-none');
});

if (btnPwaInstall) {
    btnPwaInstall.addEventListener('click', () => {
        if (deferredPwaPrompt) {
            deferredPwaPrompt.prompt();
            deferredPwaPrompt.userChoice.then((choiceResult) => {
                if (choiceResult.outcome === 'accepted') {
                    console.log('PWA terpasang');
                }
                deferredPwaPrompt = null;
                if (pwaBanner) pwaBanner.classList.add('d-none');
            });
        }
    });
}

if (btnPwaDismiss) {
    btnPwaDismiss.addEventListener('click', () => {
        if (pwaBanner) pwaBanner.classList.add('d-none');
    });
}

window.addEventListener('appinstalled', () => {
    if (pwaBanner) pwaBanner.classList.add('d-none');
    console.log('PWA berhasil dipasang');
});

// Service Worker Registration
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('service-worker.js')
        .then((registration) => {
            console.log('Service Worker registered with scope:', registration.scope);
            
            registration.addEventListener('updatefound', () => {
                const newWorker = registration.installing;
                if (newWorker) {
                    newWorker.addEventListener('statechange', () => {
                        if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
                            Swal.fire({
                                title: 'Pembaruan Tersedia',
                                text: 'Versi baru AuditWorkpaper Pro telah siap. Muat ulang sekarang?',
                                icon: 'info',
                                showCancelButton: true,
                                confirmButtonText: 'Muat Ulang',
                                cancelButtonText: 'Nanti',
                                confirmButtonColor: '#1d4ed8'
                            }).then((res) => {
                                if (res.isConfirmed) {
                                    window.location.reload();
                                }
                            });
                        }
                    });
                }
            });
        })
        .catch((error) => {
            console.error('Service Worker registration failed:', error);
        });
    });
}
