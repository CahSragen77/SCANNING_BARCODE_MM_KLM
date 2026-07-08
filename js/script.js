// script.js
// =================================================================
// KONFIGURASI: GANTI DENGAN ID GOOGLE SHEETS KAMU
// =================================================================
const SPREADSHEET_ID = "18aui0vuuOVN4nxmLmUHqqe-y4PSrsr9ezfn7qjuOhb8";
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwREXUsIuYRLCJjY1kwY_YR3fmjFOx-_WTZh1_PBjlUyunFD-BYvak2jbgQSe3RRj5P/exec";

// =================================================================
// STATE APLIKASI
// =================================================================
let masterData = [];
let returList = [];
let currentActiveProduct = null;
let dashboardScanner = null;
let isScannerActive = false;

// =================================================================
// ELEMEN DOM
// =================================================================
const excelInput = document.getElementById('excel-file');
const statusMasterLoad = document.getElementById('status-master-load');
const totalItems = document.getElementById('total-items');
const interactionSection = document.getElementById('interaction-section');
const appStatus = document.getElementById('app-status');
const btnClearAll = document.getElementById('btn-clear-all');
const manualBarcode = document.getElementById('manual-barcode');
const btnLookup = document.getElementById('btn-lookup');
const returCount = document.getElementById('retur-count');
const scannerContainer = document.getElementById('scanner-container');
const cameraStatusMsg = document.getElementById('camera-status-msg');
const entryFormSection = document.getElementById('entry-form-section');
const entryDescp = document.getElementById('entry-descp');
const entryPlu = document.getElementById('entry-plu');
const entryBarcode = document.getElementById('entry-barcode');
const entryKategori = document.getElementById('entry-kategori');
const entryHarga = document.getElementById('entry-harga');
const returQty = document.getElementById('retur-qty');
const returExp = document.getElementById('retur-exp');
const btnSaveEntry = document.getElementById('btn-save-entry');
const btnCancelEntry = document.getElementById('btn-cancel-entry');
const returTableBody = document.getElementById('retur-table-body');
const btnExport = document.getElementById('btn-export');

// =================================================================
// FUNGSI UTAMA
// =================================================================

// --- CEK PROTOKOL HTTPS ---
function checkProtocol() {
    if (window.location.protocol !== 'https:' && window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
        const warningMsg = document.createElement('div');
        warningMsg.className = 'camera-error-box mb-4';
        warningMsg.innerHTML = `
            <strong>⚠️ Akses Kamera Diblokir</strong><br>
            <p class="text-sm mt-1">Browser memblokir kamera karena situs diakses melalui <strong>HTTP</strong>.<br>
            Silakan access melalui <strong>HTTPS</strong> (misal: GitHub Pages, Vercel, Netlify) agar kamera berfungsi.</p>
            <p class="text-xs mt-2 text-gray-600">Untuk pengujian lokal, gunakan <strong>localhost</strong>.</p>
        `;
        const readerDiv = document.getElementById('reader-dashboard');
        if (readerDiv) readerDiv.parentNode.insertBefore(warningMsg, readerDiv);
        cameraStatusMsg.innerHTML = `
            <span class="text-4xl mb-2">🔒</span>
            <span class="font-bold">Kamera Tidak Tersedia</span>
            <span class="text-xs mt-1 opacity-75">Gunakan HTTPS untuk mengakses kamera</span>
        `;
        return false;
    }
    return true;
}

// --- AMBIL DATA OTOMATIS DARI GOOGLE SHEETS ---
window.addEventListener('DOMContentLoaded', () => {
    checkProtocol();

    if (SPREADSHEET_ID === "MASUKKAN_ID_GOOGLE_SHEETS_KAMU_DISINI" || SPREADSHEET_ID === "") {
        appStatus.innerText = "⚠️ Belum Ada Master Data (Ganti ID Sheets)";
        appStatus.className = "text-red-500 font-bold";
        return;
    }

    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=csv`;

    fetch(url)
        .then(response => {
            if (!response.ok) throw new Error("Gagal mengambil data dari Google Sheets");
            return response.arrayBuffer();
        })
        .then(buffer => {
            const data = new Uint8Array(buffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const rawData = XLSX.utils.sheet_to_json(worksheet);
            prosesDataSukses(rawData, "Cloud (Google Sheets)");
        })
        .catch(error => {
            console.error(error);
            appStatus.innerText = "⚠️ Gagal Sinkronisasi Cloud (Offline Mode)";
            appStatus.className = "text-red-500 font-bold";
            alert("Gagal memuat otomatis dari Google Sheets. Silakan upload file Excel manual.");
        });
});

// --- PROSES BACA EXCEL MANUAL ---
excelInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const rawData = XLSX.utils.sheet_to_json(worksheet);
            prosesDataSukses(rawData, "Lokal (Upload Manual)");
        } catch (error) {
            console.error(error);
            alert('Terjadi kesalahan saat membaca file Excel: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
});

function prosesDataSukses(rawData, sumber) {
    if(rawData.length === 0) {
        alert('Data produk kosong!');
        return;
    }

    masterData = rawData.map(row => {
        const normalizedRow = {};
        for (let key in row) {
            const lowerKey = key.trim().toLowerCase();
            if (lowerKey === 'plu' || lowerKey === 'kode plu' || lowerKey === 'kode_plu') {
                normalizedRow.plu = row[key];
            } else if (lowerKey === 'barcode' || lowerKey === 'no_barcode' || lowerKey === 'no barcode') {
                normalizedRow.barcode = row[key];
            } else if (lowerKey === 'descp' || lowerKey === 'deskripsi' || lowerKey === 'nama barang' || lowerKey === 'nama_barang') {
                normalizedRow.descp = row[key];
            } else if (lowerKey === 'price1' || lowerKey === 'harga' || lowerKey === 'price') {
                normalizedRow.price1 = row[key];
            } else if (lowerKey === 'kategori' || lowerKey === 'category') {
                normalizedRow.kategori = row[key];
            }
        }
        return normalizedRow;
    });

    if(masterData.length > 0 && (masterData[0].barcode || masterData[0].plu)) {
        totalItems.innerText = masterData.length;
        statusMasterLoad.classList.remove('hidden');
        interactionSection.classList.remove('opacity-40', 'pointer-events-none');
        excelInput.disabled = true;
        appStatus.innerText = `✅ Dashboard Aktif [${sumber}]`;
        appStatus.className = "text-green-600 font-bold";
        btnClearAll.classList.remove('hidden');
    } else {
        alert('Gagal memetakan data. Pastikan kolom judul bernamakan PLU, Barcode, dan Deskripsi.');
    }
}

// =====================================================
// SCANNER UNTUK DASHBOARD (DENGAN TOMBOL START)
// =====================================================
function startDashboardScanner() {
    cameraStatusMsg.innerHTML = `
        <span class="text-4xl mb-2">📱</span>
        <span class="font-bold">Tekan Tombol "Mulai Scan"</span>
        <span class="text-xs mt-1 opacity-75">Di bawah ini untuk mengaktifkan kamera</span>
    `;
    cameraStatusMsg.style.display = 'flex';
    console.log('📱 Siap scan. Tekan tombol "Mulai Scan" di bawah.');
}

document.getElementById('btn-start-scanner')?.addEventListener('click', function() {
    if (dashboardScanner) {
        try {
            dashboardScanner.stop().catch(() => {});
            dashboardScanner.clear();
        } catch(e) {}
        isScannerActive = false;
    }
    
    cameraStatusMsg.innerHTML = `
        <span class="text-4xl mb-2">📷</span>
        <span>Memulai kamera... Mohon izinkan akses kamera</span>
    `;
    cameraStatusMsg.style.display = 'flex';
    document.getElementById('reader-dashboard').style.display = 'block';
    
    startDashboardScannerWithRetry();
});

function startDashboardScannerWithRetry() {
    let retryCount = 0;
    const maxRetries = 3;
    
    function attemptStart() {
        if (retryCount >= maxRetries) {
            showCameraError('Gagal mengakses kamera setelah beberapa percobaan. Pastikan izin kamera diberikan.');
            return;
        }
        retryCount++;
        
        const config = {
            fps: 15,
            qrbox: { width: 250, height: 200 },
            aspectRatio: 1.0
        };
        
        dashboardScanner = new Html5Qrcode("reader-dashboard");
        
        Html5Qrcode.getCameras().then(devices => {
            if (devices && devices.length > 0) {
                let targetId = devices[0].id;
                for (let device of devices) {
                    if (device.label.toLowerCase().includes('back') || 
                        device.label.toLowerCase().includes('rear') || 
                        device.label.toLowerCase().includes('belakang')) {
                        targetId = device.id;
                        break;
                    }
                }
                
                dashboardScanner.start(
                    targetId, 
                    config, 
                    onScanSuccess, 
                    onScanFailure
                ).then(() => {
                    isScannerActive = true;
                    cameraStatusMsg.style.display = 'none';
                    console.log('✅ Scanner HP berhasil diaktifkan');
                    document.getElementById('btn-start-scanner').textContent = '📷 Scan Aktif';
                }).catch(err => {
                    console.warn(`Percobaan ${retryCount} gagal:`, err);
                    attemptStart();
                });
            } else {
                attemptStart();
            }
        }).catch(err => {
            console.warn(`Percobaan ${retryCount} gagal:`, err);
            attemptStart();
        });
    }
    
    attemptStart();
}

function showCameraError(errorMessage) {
    isScannerActive = false;
    const readerDiv = document.getElementById('reader-dashboard');
    if (readerDiv) readerDiv.style.display = 'none';
    
    cameraStatusMsg.innerHTML = `
        <span class="text-4xl mb-2">🚫</span>
        <span class="font-bold text-red-400">Kamera Gagal Diakses</span>
        <span class="text-xs mt-1 opacity-75">${errorMessage}</span>
        <button id="retry-camera-btn" class="retry-btn mt-3">
            🔄 Coba Lagi
        </button>
        <span class="text-xs mt-2 opacity-50">Pastikan izin kamera diberikan di browser</span>
    `;
    cameraStatusMsg.style.display = 'flex';

    document.getElementById('retry-camera-btn')?.addEventListener('click', function() {
        const readerDiv = document.getElementById('reader-dashboard');
        if (readerDiv) readerDiv.style.display = 'block';
        cameraStatusMsg.innerHTML = `
            <span class="text-4xl mb-2">📷</span>
            <span>Mengaktifkan kamera...</span>
        `;
        startDashboardScannerWithRetry();
    });
}

function onScanSuccess(decodedText) {
    handleProductLookup(decodedText);
    if (navigator.vibrate) navigator.vibrate(100);
}

function onScanFailure(error) {}

// =====================================================
// LOOKUP PRODUK & FUNGSI LAINNYA
// =====================================================
function handleProductLookup(barcodeOrPlu) {
    const searchKey = String(barcodeOrPlu).trim().replace(/^0+/, '');

    const product = masterData.find(item => {
        const itemBarcode = item.barcode ? String(item.barcode).trim().replace(/^0+/, '') : '';
        const itemPlu = item.plu ? String(item.plu).trim().replace(/^0+/, '') : '';
        return itemBarcode === searchKey || itemPlu === searchKey;
    });

    if (product) {
        currentActiveProduct = product;
        entryFormSection.classList.remove('hidden');
        
        entryDescp.innerText = product.descp || 'Tanpa Deskripsi';
        entryPlu.innerText = product.plu || '-';
        entryBarcode.innerText = product.barcode || '-';
        if (entryKategori) entryKategori.innerText = product.kategori || 'Umum';
        
        const hargaTerformat = !isNaN(product.price1) && product.price1 !== "" 
            ? 'Rp ' + Number(product.price1).toLocaleString('id-ID') 
            : 'Rp 0';
        if (entryHarga) entryHarga.innerText = hargaTerformat;
        
        returQty.value = '';
        returExp.value = '';
        
        entryFormSection.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } else {
        alert(`❌ Produk dengan Barcode/PLU "${barcodeOrPlu}" tidak ditemukan.`);
    }
}

btnLookup.addEventListener('click', () => {
    if(manualBarcode.value.trim() !== "") handleProductLookup(manualBarcode.value);
});
manualBarcode.addEventListener('keypress', (e) => {
    if(e.key === 'Enter' && manualBarcode.value.trim() !== "") handleProductLookup(manualBarcode.value);
});
btnCancelEntry.addEventListener('click', () => {
    entryFormSection.classList.add('hidden');
    currentActiveProduct = null;
    manualBarcode.value = '';
});

// =====================================================
// MANAJEMEN RETUR
// =====================================================

// --- 4. DATA ENTRY: SIMPAN ITEM & OTOMATIS SINKRON KE CLOUD ---
btnSaveEntry.addEventListener('click', () => {
    const qty = parseInt(returQty.value);
    const expiredDate = returExp.value;

    if(!qty || qty <= 0) { alert('Isi Qty minimal 1!'); returQty.focus(); return; }
    if(!expiredDate) { alert('Isi tanggal kedaluwarsa!'); returExp.focus(); return; }

    // Membungkus muatan data
    const itemRetur = {
        plu: currentActiveProduct.plu,
        barcode: currentActiveProduct.barcode,
        descp: currentActiveProduct.descp,
        qty_retur: qty,
        tgl_expired: expiredDate,
        timestamp: new Date().toISOString()
    };

    // 1. Simpan ke dalam array lokal antrean aplikasi
    returList.push(itemRetur);
    renderReturTable();

    // Ubah status tombol penanda proses
    const originalBtnText = btnSaveEntry.innerText;
    btnSaveEntry.innerText = "Mengirim ke Sheets...";
    btnSaveEntry.disabled = true;

    // 2. UTAMA: Panggil kurir FETCH dengan menyamakan struktur payload Apps Script backend
    fetch(WEB_APP_URL, {
        method: "POST",
        body: JSON.stringify({ action: 'saveRetur', data: itemRetur }),
        headers: {
            "Content-Type": "text/plain;charset=utf-8"
        }
    })
    .then(res => res.json())
    .then(response => {
        if(response.status === "success") {
            console.log("Data berhasil masuk ke Google Sheets!");
        } else {
            alert("Aplikasi mencatat lokal, tapi gagal sinkron ke Sheets: " + response.message);
        }
    })
    .catch(err => {
        console.error(err);
        alert("Gagal otomatis sinkron ke Google Sheets (Kendala Jaringan). Namun data tetap tersimpan di tabel bawah aplikasi.");
    })
    .finally(() => {
        btnSaveEntry.innerText = originalBtnText;
        btnSaveEntry.disabled = false;
        
        entryFormSection.classList.add('hidden');
        manualBarcode.value = '';
        currentActiveProduct = null;
    });
});

// --- RENDER VISUAL TABEL KERJA ---
function renderReturTable() {
    if (returCount) returCount.textContent = returList.length;
    
    if(returList.length === 0) {
        returTableBody.innerHTML = `
            <tr>
                <td colspan="6" class="p-8 text-center text-gray-400 italic bg-gray-50/50">
                    <span class="text-4xl block mb-2">📭</span>
                    Belum ada item ditambahkan.<br>
                    <span class="text-xs">Silakan scan produk untuk memulai</span>
                </td>
            </tr>
        `;
        btnExport.classList.add('hidden');
        return;
    }
    btnExport.classList.remove('hidden');
    returTableBody.innerHTML = '';
    returList.forEach((item, index) => {
        const tr = document.createElement('tr');
        tr.className = "row-hover transition-all duration-200";
        tr.style.animation = `slideIn 0.3s ease-out forwards`;
        tr.style.animationDelay = `${index * 0.05}s`;
        tr.innerHTML = `
            <td class="p-3 font-mono text-gray-700">${item.plu || '-'}</td>
            <td class="p-3 font-mono text-gray-600">${item.barcode || '-'}</td>
            <td class="p-3 font-medium text-gray-900">${item.descp || '-'}</td>
            <td class="p-3 text-center font-bold text-indigo-600 text-sm">${item.qty_retur}</td>
            <td class="p-3 text-gray-600">${item.tgl_expired}</td>
            <td class="p-3 text-center">
                <button onclick="deleteReturItem(${index})" 
                    class="text-red-500 hover:text-red-700 font-semibold transition-all duration-300 hover:scale-110 inline-flex items-center gap-1">
                    🗑️ Hapus
                </button>
            </td>
        `;
        returTableBody.appendChild(tr);
    });
}

// Fungsi hapus global
window.deleteReturItem = function(index) { 
    returList.splice(index, 1); 
    renderReturTable(); 
};

// =====================================================
// EKSPORT EXCEL MANUAL
// =====================================================
btnExport.addEventListener('click', () => {
    if(returList.length === 0) return;
    const dataToExport = returList.map((item, idx) => ({
        "No": idx + 1, 
        "Kode PLU": item.plu, 
        "Barcode": item.barcode, 
        "Deskripsi Produk": item.descp, 
        "Qty Retur": item.qty_retur, 
        "Tanggal Kedaluwarsa (Expired)": item.tgl_expired
    }));
    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Laporan Retur Supplier");
    const max_width = dataToExport.reduce((w, r) => Math.max(w, String(r["Deskripsi Produk"]).length), 20);
    worksheet["!cols"] = [{ wch: 5 }, { wch: 12 }, { wch: 18 }, { wch: max_width }, { wch: 12 }, { wch: 25 }];
    const namaFile = `Laporan_Kerja_Retur_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, namaFile);
});

// =====================================================
// FUNGSI CLOUD (SAVE KE GOOGLE SHEETS CADANGAN)
// =====================================================
async function saveReturToCloud(returData) {
    try {
        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ action: 'saveRetur', data: returData })
        });
        return true;
    } catch (error) {
        console.error('❌ Gagal menyimpan ke cloud:', error);
        return false;
    }
}

// =====================================================
// TOMBOL SYNC MANUAL KE CLOUD
// =====================================================
function addSyncAllButton() {
    const headerDiv = document.querySelector('header .flex.items-center.gap-3.flex-wrap');
    if (headerDiv) {
        const syncBtn = document.createElement('button');
        syncBtn.id = 'btn-sync-cloud';
        syncBtn.className = 'bg-gradient-to-r from-blue-500 to-cyan-600 hover:from-blue-600 hover:to-cyan-700 text-white px-5 py-2 rounded-full text-xs font-bold transition-all duration-300 shadow-md hover:shadow-lg transform hover:scale-105 flex items-center gap-2 btn-glow';
        syncBtn.innerHTML = '☁️ Sync ke Cloud';
        syncBtn.onclick = async function() {
            if (returList.length === 0) {
                alert('Tidak ada data retur untuk disync.');
                return;
            }
            let successCount = 0;
            for (const item of returList) {
                const status = await saveReturToCloud(item);
                if (status) successCount++;
            }
            alert(`✅ Sinkronisasi massal selesai.`);
        };
        headerDiv.appendChild(syncBtn);
    }
}

// =====================================================
// RESET DATA
// =====================================================
btnClearAll.addEventListener('click', () => {
    if(confirm('🔄 Reset semua data entry dari awal?')) {
        location.reload();
    }
});

// =====================================================
// INISIALISASI HALAMAN
// =====================================================
document.addEventListener('DOMContentLoaded', function() {
    startDashboardScanner();
    addSyncAllButton();
    renderReturTable(); // Inisialisasi visual tabel kosong
    console.log("🔄 Script utama telah diperbaiki sempurna!");
});
