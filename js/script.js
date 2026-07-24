// script.js
// =================================================================
// KONFIGURASI GOOGLE SHEETS & APPS SCRIPT
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
const btnStartScanner = document.getElementById('btn-start-scanner');

// =================================================================
// FUNGSI UTAMA & SINKRONISASI MASTER DATA
// =================================================================

function checkProtocol() {
    if (window.location.protocol !== 'https:' && window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
        const warningMsg = document.createElement('div');
        warningMsg.className = 'camera-error-box mb-4';
        warningMsg.innerHTML = `
            <strong>⚠️ Akses Kamera Diblokir</strong><br>
            <p class="text-sm mt-1">Browser memblokir kamera karena situs diakses melalui <strong>HTTP</strong>.<br>
            Silakan akses melalui <strong>HTTPS</strong> agar kamera berfungsi.</p>
        `;
        const readerDiv = document.getElementById('reader-dashboard');
        if (readerDiv) readerDiv.parentNode.insertBefore(warningMsg, readerDiv);
        if (cameraStatusMsg) {
            cameraStatusMsg.innerHTML = `
                <span class="text-4xl mb-2">🔒</span>
                <span class="font-bold">Kamera Tidak Tersedia</span>
                <span class="text-xs mt-1 opacity-75">Gunakan HTTPS untuk mengakses kamera</span>
            `;
        }
        return false;
    }
    return true;
}

// Ambil Master Data Otomatis dari Google Sheets (CSV Export) secara Asynchronous
window.addEventListener('DOMContentLoaded', () => {
    checkProtocol();

    if (!SPREADSHEET_ID || SPREADSHEET_ID === "MASUKKAN_ID_GOOGLE_SHEETS_KAMU_DISINI") {
        if (appStatus) {
            appStatus.innerText = "⚠️ Belum Ada Master Data (Ganti ID Sheets)";
            appStatus.className = "text-red-500 font-bold";
        }
        return;
    }

    if (appStatus) {
        appStatus.innerText = "⏳ Memuat Master Data Cloud...";
        appStatus.className = "text-yellow-600 font-bold animate-pulse";
    }

    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=csv`;

    // Fetch asynchronously agar UI tidak membeku saat loading pertama
    setTimeout(() => {
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
                if (appStatus) {
                    appStatus.innerText = "⚠️ Gagal Sinkronisasi Cloud (Offline Mode)";
                    appStatus.className = "text-red-500 font-bold";
                }
            });
    }, 100);
});

// Upload Excel Manual (Backup)
if (excelInput) {
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
}

function prosesDataSukses(rawData, sumber) {
    if(!rawData || rawData.length === 0) {
        alert('Data produk kosong!');
        return;
    }

    masterData = rawData.map(row => {
        const normalizedRow = {};
        for (let key in row) {
            const lowerKey = key.trim().toLowerCase();
            if (['plu', 'kode plu', 'kode_plu'].includes(lowerKey)) normalizedRow.plu = row[key];
            else if (['barcode', 'no_barcode', 'no barcode'].includes(lowerKey)) normalizedRow.barcode = row[key];
            else if (['descp', 'deskripsi', 'nama barang', 'nama_barang'].includes(lowerKey)) normalizedRow.descp = row[key];
            else if (['price1', 'harga', 'price'].includes(lowerKey)) normalizedRow.price1 = row[key];
            else if (['kategori', 'category'].includes(lowerKey)) normalizedRow.kategori = row[key];
        }
        return normalizedRow;
    });

    if(masterData.length > 0) {
        if (totalItems) totalItems.innerText = masterData.length;
        if (statusMasterLoad) statusMasterLoad.classList.remove('hidden');
        if (interactionSection) interactionSection.classList.remove('opacity-40', 'pointer-events-none');
        if (excelInput) excelInput.disabled = true;
        if (appStatus) {
            appStatus.innerText = `✅ Dashboard Siap [${sumber}]`;
            appStatus.className = "text-green-600 font-bold";
        }
        if (btnClearAll) btnClearAll.classList.remove('hidden');
    }
}

// =====================================================
// FITUR KONTROL KAMERA (START / STOP / TOGGLE)
// =====================================================

function updateButtonToStart() {
    isScannerActive = false;
    if (btnStartScanner) {
        btnStartScanner.textContent = '📷 Mulai Scan';
        btnStartScanner.className = 'w-full bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-600 hover:to-teal-700 text-white font-bold py-3 px-4 rounded-xl shadow-md transition-all duration-200 transform active:scale-95 flex items-center justify-center gap-2';
    }
}

function updateButtonToStop() {
    isScannerActive = true;
    if (btnStartScanner) {
        btnStartScanner.textContent = '🛑 Stop Kamera';
        btnStartScanner.className = 'w-full bg-gradient-to-r from-red-500 to-rose-600 hover:from-red-600 hover:to-rose-700 text-white font-bold py-3 px-4 rounded-xl shadow-md transition-all duration-200 transform active:scale-95 flex items-center justify-center gap-2';
    }
}

function stopDashboardScanner() {
    if (dashboardScanner) {
        dashboardScanner.stop().then(() => {
            dashboardScanner.clear();
            dashboardScanner = null;
            updateButtonToStart();
            showCameraOffMessage();
        }).catch(err => {
            console.warn("Gagal stop scanner secara normal:", err);
            try { dashboardScanner.clear(); } catch(e){}
            dashboardScanner = null;
            updateButtonToStart();
            showCameraOffMessage();
        });
    } else {
        updateButtonToStart();
        showCameraOffMessage();
    }
}

function showCameraOffMessage() {
    const readerDiv = document.getElementById('reader-dashboard');
    if (readerDiv) readerDiv.style.display = 'none';
    if (cameraStatusMsg) {
        cameraStatusMsg.innerHTML = `
            <span class="text-4xl mb-2">📷</span>
            <span class="font-bold">Kamera Nonaktif</span>
            <span class="text-xs mt-1 opacity-75">Tekan "Mulai Scan" untuk menggunakan kamera</span>
        `;
        cameraStatusMsg.style.display = 'flex';
    }
}

// Event Listener Tombol On/Off Scanner
btnStartScanner?.addEventListener('click', function() {
    if (isScannerActive) {
        // Jika kamera aktif, maka MATIKAN
        stopDashboardScanner();
    } else {
        // Jika kamera mati, maka NYALAKAN
        if (cameraStatusMsg) {
            cameraStatusMsg.innerHTML = `
                <span class="text-4xl mb-2">📷</span>
                <span>Memulai kamera... Mohon izinkan akses</span>
            `;
            cameraStatusMsg.style.display = 'flex';
        }
        const readerDiv = document.getElementById('reader-dashboard');
        if (readerDiv) readerDiv.style.display = 'block';
        
        startDashboardScannerWithRetry();
    }
});

function startDashboardScannerWithRetry() {
    let retryCount = 0;
    const maxRetries = 3;
    
    function attemptStart() {
        if (retryCount >= maxRetries) {
            showCameraError('Gagal mengakses kamera. Pastikan izin kamera telah diberikan.');
            return;
        }
        retryCount++;
        
        const config = { fps: 15, qrbox: { width: 250, height: 200 }, aspectRatio: 1.0 };
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
                
                dashboardScanner.start(targetId, config, onScanSuccess, onScanFailure)
                .then(() => {
                    updateButtonToStop();
                    if (cameraStatusMsg) cameraStatusMsg.style.display = 'none';
                }).catch(() => attemptStart());
            } else {
                attemptStart();
            }
        }).catch(() => attemptStart());
    }
    attemptStart();
}

function showCameraError(errorMessage) {
    updateButtonToStart();
    const readerDiv = document.getElementById('reader-dashboard');
    if (readerDiv) readerDiv.style.display = 'none';
    
    if (cameraStatusMsg) {
        cameraStatusMsg.innerHTML = `
            <span class="text-4xl mb-2">🚫</span>
            <span class="font-bold text-red-400">Kamera Gagal Diakses</span>
            <span class="text-xs mt-1 opacity-75">${errorMessage}</span>
            <button id="retry-camera-btn" class="retry-btn mt-3">🔄 Coba Lagi</button>
        `;
        cameraStatusMsg.style.display = 'flex';
    }

    document.getElementById('retry-camera-btn')?.addEventListener('click', function() {
        if (readerDiv) readerDiv.style.display = 'block';
        startDashboardScannerWithRetry();
    });
}

function onScanSuccess(decodedText) {
    if (navigator.vibrate) navigator.vibrate(100);
    
    // Temukan produk
    handleProductLookup(decodedText);
    
    // OTOMATIS STOP KAMERA setelah barang ditemukan agar HP tidak panas
    stopDashboardScanner();
}

function onScanFailure(error) {}

// =====================================================
// LOOKUP PRODUK
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
        if (entryFormSection) entryFormSection.classList.remove('hidden');
        
        if (entryDescp) entryDescp.innerText = product.descp || 'Tanpa Deskripsi';
        if (entryPlu) entryPlu.innerText = product.plu || '-';
        if (entryBarcode) entryBarcode.innerText = product.barcode || '-';
        if (entryKategori) entryKategori.innerText = product.kategori || 'Umum';
        
        const hargaTerformat = !isNaN(product.price1) && product.price1 !== "" 
            ? 'Rp ' + Number(product.price1).toLocaleString('id-ID') 
            : 'Rp 0';
        if (entryHarga) entryHarga.innerText = hargaTerformat;
        
        if (returQty) returQty.value = '';
        if (returExp) returExp.value = '';
        
        entryFormSection?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } else {
        alert(`❌ Produk dengan Barcode/PLU "${barcodeOrPlu}" tidak ditemukan.`);
    }
}

btnLookup?.addEventListener('click', () => {
    if(manualBarcode.value.trim() !== "") handleProductLookup(manualBarcode.value);
});
manualBarcode?.addEventListener('keypress', (e) => {
    if(e.key === 'Enter' && manualBarcode.value.trim() !== "") handleProductLookup(manualBarcode.value);
});
btnCancelEntry?.addEventListener('click', () => {
    if (entryFormSection) entryFormSection.classList.add('hidden');
    currentActiveProduct = null;
    if (manualBarcode) manualBarcode.value = '';
});

// =====================================================
// MANAJEMEN RETUR & OTOMATIS SYNC KE SHEETS
// =====================================================

btnSaveEntry?.addEventListener('click', () => {
    const qty = parseInt(returQty.value);
    const expiredDate = returExp.value;

    if(!qty || qty <= 0) { alert('Isi Qty minimal 1!'); returQty.focus(); return; }
    if(!expiredDate) { alert('Isi tanggal kedaluwarsa!'); returExp.focus(); return; }

    const itemRetur = {
        plu: currentActiveProduct.plu,
        barcode: currentActiveProduct.barcode,
        descp: currentActiveProduct.descp,
        qty_retur: qty,
        tgl_expired: expiredDate,
        timestamp: new Date().toISOString()
    };

    // 1. Simpan ke lokal
    returList.push(itemRetur);
    renderReturTable();

    // Tampilan Loading
    const originalBtnText = btnSaveEntry.innerText;
    btnSaveEntry.innerText = "Mengirim ke Sheets...";
    btnSaveEntry.disabled = true;

    // 2. Kirim Otomatis ke Google Sheets via Apps Script
    fetch(WEB_APP_URL, {
        method: "POST",
        body: JSON.stringify({ action: 'saveRetur', data: itemRetur }),
        headers: { "Content-Type": "text/plain;charset=utf-8" }
    })
    .then(res => res.json())
    .then(response => {
        if(response.status === "success") {
            console.log("Data berhasil masuk ke Google Sheets!");
        }
    })
    .catch(err => {
        console.error(err);
        alert("Gagal sinkron otomatis ke Sheets (Kendala Jaringan). Data tetap tersimpan di tabel lokal.");
    })
    .finally(() => {
        btnSaveEntry.innerText = originalBtnText;
        btnSaveEntry.disabled = false;
        
        if (entryFormSection) entryFormSection.classList.add('hidden');
        if (manualBarcode) manualBarcode.value = '';
        currentActiveProduct = null;
    });
});

function renderReturTable() {
    if (returCount) returCount.textContent = returList.length;
    if(!returTableBody) return;

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
        if (btnExport) btnExport.classList.add('hidden');
        return;
    }
    
    if (btnExport) btnExport.classList.remove('hidden');
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

window.deleteReturItem = function(index) { 
    returList.splice(index, 1); 
    renderReturTable(); 
};

// Export ke Excel Manual
btnExport?.addEventListener('click', () => {
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

// Reset
btnClearAll?.addEventListener('click', () => {
    if(confirm('🔄 Reset semua data entry dari awal?')) {
        location.reload();
    }
});

// Inisialisasi
document.addEventListener('DOMContentLoaded', function() {
    showCameraOffMessage();
    renderReturTable();
});
