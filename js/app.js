// ==========================
// CONFIGURATION & CONSTANTS
// ==========================
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwj3UqW15lEuJ3Au5uSfYhYyxfB-QfwsumwTQ8xlzZBi7eMw0_8UUN3NSLJuB2D7xrN/exec';

// Global State (Database Memori)
let allData = {
    templateInLine: [],
    factoryA: {}, // Grouped by 'Line X'
    factoryB: {},
    requestIE: [],
    requestSewing: [],
    layoutRequests: [],
    sheet25: [],
    sheet26: [],
    inventory: []
};

let chartInstances = {};
let isOffline = false; // Pelacak status jaringan auto-fetcher
window.ieTokenPublic = sessionStorage.getItem('ieTokenPublic') || ''; // Sesi Khusus IE Dashboard

// Admin State
window.adminToken = sessionStorage.getItem('adminToken') || '';
window.isAdminMode = false; // Akan true jika admin menyalakan toggle

// ==========================
// INITIALIZATION
// ==========================
document.addEventListener('DOMContentLoaded', async () => {
    initTheme();
    initDesktopMode();
    initNavigation();
    initModals();
    initNotifications();
    renderMarquee();
    checkAdminState();

    // 1. Kecepatan Instan Murni (Stale-While-Revalidate)
    // Jika ada bayangan data lama di browser, render UI dalam waktu 0 detik!
    const cachedData = localStorage.getItem('allDataCache');
    if (cachedData) {
        try {
            allData = JSON.parse(cachedData);
            updateCharts();
            updateStatusCards();
            renderTables();
        } catch (e) { console.warn("Sistem gagal memuat cache lama"); }
    }

    showToast('info', 'Welcome to the Dashboard Template System');

    // 2. Download Data Freshest dari Cloud secara transparan (background)
    await loadAllDataFallback();

    // 3. Render Ulang dengan Data Super Baru
    updateCharts();
    updateStatusCards();
    renderTables();

    // Auto-refresh (Live 1 Menit)
    // Gunakan window.autoRefreshPaused = true untuk pause saat IE Editor aktif
    window.autoRefreshPaused = false;
    setInterval(async () => {
        if (window.autoRefreshPaused) return; // Jangan refresh jika IE Editor sedang terbuka
        await loadAllDataFallback();
        updateCharts();
        updateStatusCards();
        renderTables();
        renderMarquee();
    }, 60000);
});

// Fungsi Manual Refresh dari Header Ikon
window.forceRefreshData = async function (e) {
    let icon = null;
    if (e && e.currentTarget) {
        icon = e.currentTarget.querySelector('i');
        if (icon) icon.classList.add('spin');
    }

    showToast('info', 'Menyinkronkan data terbaru...');
    await loadAllDataFallback();
    updateCharts();
    updateStatusCards();
    renderTables();
    renderMarquee();

    if (icon) icon.classList.remove('spin');
};

// ==========================
// LIVE API DATA SYNC (GET)
// ==========================
async function loadAllDataFallback() {
    // Saya pertahankan nama fungsinya agar auto-refresh di baris 39 tetap bekerja tanpa perombakan.
    try {
        const response = await fetch(APPS_SCRIPT_URL + "?tab=all");
        const data = await response.json();

        if (data.error) throw new Error(data.error);

        // Pulihkan state jika internet kembali normal
        if (isOffline) {
            isOffline = false;
            showToast('success', 'Koneksi kembali stabil! Data tersinkron.');
        }

        // Suntikkan seluruh data memori langsung dari Awan (Cloud)
        allData = data;

        // --- DATA INHERITANCE (PEWARISAN DATA) ---
        // Jika baris di bawahnya kosong untuk Line, Style, atau Buyer, maka ikut baris di atasnya
        const applyInheritance = (arr) => {
            if (!arr || !Array.isArray(arr)) return;
            let lastLine = '';
            let lastStyle = '';
            let lastBuyer = '';
            arr.forEach(r => {
                // Line
                if (r.line && String(r.line).trim() !== '') lastLine = String(r.line).trim();
                else if (lastLine) r.line = lastLine;
                // Style / Code
                if (r.style && String(r.style).trim() !== '') lastStyle = String(r.style).trim();
                else if (lastStyle) r.style = lastStyle;
                if (r.code && String(r.code).trim() !== '') lastStyle = String(r.code).trim();
                else if (lastStyle && !r.code) r.code = lastStyle; // Sinkronisasi alias code/style
                // Buyer
                if (r.buyer && String(r.buyer).trim() !== '') lastBuyer = String(r.buyer).trim();
                else if (lastBuyer) r.buyer = lastBuyer;
            });
        };

        if (allData.templateInLine) applyInheritance(allData.templateInLine);
        if (allData.factoryA) Object.values(allData.factoryA).forEach(arr => applyInheritance(arr));
        if (allData.factoryB) Object.values(allData.factoryB).forEach(arr => applyInheritance(arr));

        // --- DYNAMIC IE CALCULATION (Front-end only) ---
        if (allData.templateInLine) {
            allData.templateInLine.forEach(r => {
                let smv = parseFloat(String(r.smv || '0').replace(',', '.')) || 0;
                let actual = parseFloat(String(r.actual || '0').replace(',', '.')) || 0;
                if (smv > 0) {
                    let sv = Math.round((smv - actual) * 100) / 100;
                    r.saving = sv;
                    if (actual > 0) {
                        r.rate = Math.round((sv / smv) * 100);
                    } else { r.rate = ''; }
                } else {
                    r.saving = '';
                    r.rate = '';
                }
            });
        }

        // Simpan ke Cache Browser agar waktu memuat aplikasi untuk besok/nanti 0 detik
        localStorage.setItem('allDataCache', JSON.stringify(allData));

    } catch (e) {
        console.error("Gagal sinkron data:", e);

        // Cegah spam notifikasi tiap 5 detik jika server tetap down / RTO
        if (!isOffline) {
            isOffline = true;
            showToast('error', 'Koneksi API putus. Layar menggunakan data bayangan terakhir...');
        }
    }
}

// ==========================
// NAVIGATION SYSTEM
// ==========================
window.switchPerfTab = function(panelId) {
    const isGen = panelId === 'general';
    document.getElementById('panelPerfGeneral').classList.toggle('hidden', !isGen);
    document.getElementById('panelPerfSaving').classList.toggle('hidden', isGen);
    
    const btnGen = document.getElementById('btnPerfGeneral');
    const btnSav = document.getElementById('btnPerfSaving');
    
    btnGen.style.color = isGen ? '#f59e0b' : 'var(--text-dim)';
    btnGen.style.fontWeight = isGen ? '700' : '600';
    btnGen.style.borderBottomColor = isGen ? '#f59e0b' : 'transparent';
    
    btnSav.style.color = !isGen ? '#f59e0b' : 'var(--text-dim)';
    btnSav.style.fontWeight = !isGen ? '700' : '600';
    btnSav.style.borderBottomColor = !isGen ? '#f59e0b' : 'transparent';
};

window.switchLogTab = function(type) {
    const isIE = type === 'IE';
    document.getElementById('panelLogIE').classList.toggle('hidden', !isIE);
    document.getElementById('panelLogSewing').classList.toggle('hidden', isIE);
    
    const btnIE = document.getElementById('btnLogIE');
    const btnSew = document.getElementById('btnLogSewing');
    
    btnIE.style.color = isIE ? 'var(--neon-cyan)' : 'var(--text-dim)';
    btnIE.style.fontWeight = isIE ? '700' : '600';
    btnIE.style.borderBottomColor = isIE ? 'var(--neon-cyan)' : 'transparent';
    
    btnSew.style.color = !isIE ? 'var(--neon-cyan)' : 'var(--text-dim)';
    btnSew.style.fontWeight = !isIE ? '700' : '600';
    btnSew.style.borderBottomColor = !isIE ? 'var(--neon-cyan)' : 'transparent';
};

function initNavigation() {
    const navMapping = {
        'nav-dashboard': 'view-dashboard',
        'nav-activity': 'view-activity',
        'nav-inventory': 'view-inventory',
        'nav-archive': 'view-archive',
        'btnGoIeUpdate': 'view-ie-editor',
        'btnGoPerformance': 'view-performance',
        'headerLogoHome': 'view-dashboard'
    };

    // Header Logo berfungsi ganda sebagai Home Button
    const headerLogo = document.getElementById('headerLogoHome');
    if (headerLogo) {
        headerLogo.addEventListener('click', () => {
            document.getElementById('nav-dashboard').click();
        });
    }

    // Main Floating Nav + Dashboard Buttons
    Object.keys(navMapping).forEach(btnId => {
        const btn = document.getElementById(btnId);
        if (!btn) return;
        btn.addEventListener('click', async (e) => {
            // KHUSUS IE UPDATE (Pintu Masuk Terlindungi)
            if (btnId === 'btnGoIeUpdate' && !window.ieTokenPublic) {
                e.preventDefault();
                openIeLoginModal();
                return;
            }

            Object.keys(navMapping).forEach(id => {
                const el = document.getElementById(id);
                if (el) el.classList.remove('active');
            });
            document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));

            btn.classList.add('active');
            const targetViewId = navMapping[btnId];
            const targetView = document.getElementById(targetViewId);
            if (targetView) targetView.classList.remove('hidden');

            // ✅ Pause auto-refresh saat IE Editor aktif, resume saat pindah view lain
            if (targetViewId === 'view-ie-editor') {
                window.autoRefreshPaused = true;
                showBannerIEMode(true);
            } else {
                window.autoRefreshPaused = false;
                showBannerIEMode(false);
                // RESET SESI saat keluar dari halaman editing
                window.ieTokenPublic = '';
                sessionStorage.removeItem('ieTokenPublic');
            }

            // Pemicu Render Spesifik Tampilan
            if (targetViewId === 'view-archive') { renderArchiveProductionTable(); renderArchiveTemplateTable(); }
            if (targetViewId === 'view-ie-editor') renderIeEditorTable();
            if (targetViewId === 'view-template') renderTables();

            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
    });

    // In-page Navigations (Home Dashboard Buttons -> Specific Views)
    document.getElementById('btnGoTemplate')?.addEventListener('click', () => {
        document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));
        document.getElementById('view-template').classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    document.getElementById('btnGoFactoryA')?.addEventListener('click', () => {
        document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));
        document.getElementById('view-factoryA').classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    document.getElementById('btnGoFactoryB')?.addEventListener('click', () => {
        document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));
        document.getElementById('view-factoryB').classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
}

// ==========================
// IE MODE BANNER (Auto-refresh paused notice)
// ==========================
function showBannerIEMode(active) {
    let banner = document.getElementById('iePauseBanner');
    if (!banner) {
        // Buat banner jika belum ada
        banner = document.createElement('div');
        banner.id = 'iePauseBanner';
        banner.style.cssText = [
            'position:fixed', 'bottom:5.5rem', 'left:50%', 'transform:translateX(-50%)',
            'background:rgba(245,158,11,0.15)', 'border:1px solid #f59e0b',
            'color:#f59e0b', 'font-size:0.72rem', 'font-weight:600',
            'padding:0.3rem 1rem', 'border-radius:20px', 'z-index:200',
            'display:none', 'align-items:center', 'gap:0.4rem',
            'backdrop-filter:blur(6px)', 'white-space:nowrap'
        ].join(';');
        banner.innerHTML = '<i class="ph ph-pause-circle"></i> Auto-refresh DIMATIKAN — Mode IE Editor Aktif';
        document.body.appendChild(banner);
    }
    banner.style.display = active ? 'flex' : 'none';
}



// ==========================
// DASHBOARD CHARTS & CARDS
// ==========================
function updateStatusCards() {
    // Ambil semua baris template in-line yang terisi
    const rows = (allData.templateInLine || []).filter(r => (r.style || '').trim() !== '' || (r.proses || '').trim() !== '');
    const total = rows.length || 1;

    // Kumpulkan nama-nama Line berdasarkan status (dari kolom 'line')
    let notUsedLines = [];
    let noProcessLines = [];
    let running = 0;

    // Aggregator IE
    let sumSmv = 0;
    let sumActual = 0;
    let sumSaving = 0;

    rows.forEach(r => {
        const st = String(r.status || '').toLowerCase();
        const raw = String(r.line || '').trim();
        // Format sebagai "Line X": jika sudah ada kata "line" gunakan as-is,
        // jika hanya angka prefix dengan "Line ", jika kosong skip (jangan tampilkan sheetAsal)
        let displayLine = '';
        if (raw) {
            displayLine = /line/i.test(raw) ? raw : `Line ${raw}`;
        }
        if (st.includes('not used')) {
            if (displayLine) notUsedLines.push(displayLine);
        } else if (st.includes('no process')) {
            if (displayLine) noProcessLines.push(displayLine);
        } else if (st.includes('running') || st.includes('ready') || st.includes('pilot') || st.includes('sample')) {
            running++;
        }

        // IE Aggregation (bersihkan koma agar desimal terbaca akurat)
        sumSmv += parseFloat(String(r.smv || '0').replace(',', '.')) || 0;
        sumActual += parseFloat(String(r.actual || '0').replace(',', '.')) || 0;
        sumSaving += parseFloat(String(r.saving || '0').replace(',', '.')) || 0;
    });

    // RATE = running / (running + not_used) * 100
    const notUsedCount = notUsedLines.length;
    const rateBase = running + notUsedCount || 1;
    const rate = Math.round((running / rateBase) * 100);
    const rateEl = document.getElementById('dashRate');
    if (rateEl) rateEl.innerText = `${rate}%`;

    // Tampilkan nama-nama line (unik, dipisah koma)
    const nuEl = document.getElementById('dashNotUsed');
    if (nuEl) {
        const unique = [...new Set(notUsedLines)].filter(v => v);
        nuEl.innerText = unique.length > 0 ? unique.join(', ') : '—';
    }
    const npEl = document.getElementById('dashNoProcess');
    if (npEl) {
        const unique = [...new Set(noProcessLines)].filter(v => v);
        npEl.innerText = unique.length > 0 ? unique.join(', ') : '—';
    }

    // Tampilkan Aggregasi IE di Kartu Atas
    const smvEl = document.getElementById('dashTotalSmv');
    if (smvEl) smvEl.innerText = (Math.round(sumSmv * 100) / 100).toFixed(1);

    const actualEl = document.getElementById('dashTotalActual');
    if (actualEl) actualEl.innerText = (Math.round(sumActual * 100) / 100).toFixed(1);

    const savingEl = document.getElementById('dashTotalSaving');
    if (savingEl) {
        const roundedSaving = Math.round(sumSaving * 100) / 100;
        savingEl.innerText = roundedSaving.toFixed(1);
        // Opsional: ganti warna text global saving kalau minus
        savingEl.style.color = roundedSaving > 0 ? '#10b981' : (roundedSaving < 0 ? '#ef4444' : '#10b981');
    }
}

function updateCharts() {
    // Daftarkan plugin Datalabels Theming Custom
    if (typeof ChartDataLabels !== 'undefined') {
        Chart.register(ChartDataLabels);
    }

    Chart.defaults.font.family = "'Oswald', sans-serif";

    // Deteksi tema aktif — agar label angka chart menyesuaikan background
    const isLightMode = document.body.classList.contains('light-mode');
    const labelColor = isLightMode ? '#1e293b' : '#ffffff'; // Angka di ujung batang
    const tickColor = isLightMode ? '#64748b' : '#8b8b9e'; // Teks label axis
    const gridColor = isLightMode ? 'rgba(148,163,184,0.2)' : 'rgba(60,60,80,0.3)';
    Chart.defaults.color = tickColor;

    const commonOpts = {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        layout: {
            padding: { right: 30 } // Supaya angka tidak terpotong tepi layar
        },
        plugins: {
            legend: { display: false },
            datalabels: {
                color: labelColor,   // ✅ DINAMIS: gelap di light mode, putih di dark mode
                anchor: 'end',
                align: 'right',
                offset: 4,
                font: { weight: 'bold', size: 14, family: "'Inter', sans-serif" },
                formatter: (value) => value
            }
        },
        scales: {
            x: {
                grid: { color: gridColor, drawBorder: false },
                ticks: { precision: 0, beginAtZero: true, color: tickColor },
                grace: '15%'
            },
            y: { grid: { display: false }, ticks: { color: tickColor } }
        }
    };

    // HELPER MENGHITUNG STATUS GLOBAL
    const countStatus = (dataArr, statusList) => {
        if (!dataArr || !Array.isArray(dataArr)) return 0;
        return dataArr.filter(r => r.status && statusList.some(s => r.status.toLowerCase().includes(s.toLowerCase()))).length;
    };

    const aggFactory = (facObj) => {
        let r = 0, p = 0, pi = 0, n = 0;
        if (facObj) {
            Object.values(facObj).forEach(lines => {
                if (Array.isArray(lines)) {
                    r += countStatus(lines, ['ready', 'running']); // Hijau
                    p += countStatus(lines, ['process']); // Biru
                    pi += countStatus(lines, ['pilot', 'sample']); // Orange — sample = pilot baru
                    n += countStatus(lines, ['no process']); // Merah
                }
            });
        }
        return [r, p, pi, n];
    };

    // 1. Template Chart (Dari Sheet2)
    const runStr = countStatus(allData.templateInLine, ['running', 'ready']); // Hijau
    const notStr = countStatus(allData.templateInLine, ['not used', 'process']); // Biru
    const nopStr = countStatus(allData.templateInLine, ['no process']); // Merah

    const ctxT = document.getElementById('templateChart');
    if (chartInstances['templateChart']) chartInstances['templateChart'].destroy();

    chartInstances['templateChart'] = new Chart(ctxT, {
        type: 'bar',
        data: {
            labels: ['Running', 'Not Used', 'No Process'],
            datasets: [{ data: [runStr, notStr, nopStr], backgroundColor: ['#00ff88', '#00bfff', '#ef4444'], borderRadius: 4 }]
        },
        options: commonOpts
    });

    // 2. Factory Chart Gabungan (Sheet3-24)
    const facACounts = aggFactory(allData.factoryA);
    const facBCounts = aggFactory(allData.factoryB);
    const facTotal = [
        facACounts[0] + facBCounts[0],
        facACounts[1] + facBCounts[1],
        facACounts[2] + facBCounts[2],
        facACounts[3] + facBCounts[3]
    ];
    const factoryColors = ['#00ff88', '#00bfff', '#f97316', '#ef4444'];

    const ctxF = document.getElementById('factoryChart');
    if (chartInstances['factoryChart']) chartInstances['factoryChart'].destroy();
    chartInstances['factoryChart'] = new Chart(ctxF, {
        type: 'bar',
        data: {
            labels: ['Ready', 'Process', 'Sample', 'No Process'],
            datasets: [{ data: facTotal, backgroundColor: factoryColors, borderRadius: 4 }]
        },
        options: commonOpts
    });

    // 3. Factory A Chart (Sheet3-14)
    const ctxFa = document.getElementById('chartFactoryA');
    if (ctxFa) {
        if (chartInstances['chartFactoryA']) chartInstances['chartFactoryA'].destroy();
        chartInstances['chartFactoryA'] = new Chart(ctxFa, {
            type: 'bar',
            data: { labels: ['Ready', 'Process', 'Sample', 'No Process'], datasets: [{ data: facACounts, backgroundColor: factoryColors, borderRadius: 4 }] },
            options: commonOpts
        });
    }

    // 4. Factory B Chart (Sheet15-24)
    const ctxFb = document.getElementById('chartFactoryB');
    if (ctxFb) {
        if (chartInstances['chartFactoryB']) chartInstances['chartFactoryB'].destroy();
        chartInstances['chartFactoryB'] = new Chart(ctxFb, {
            type: 'bar',
            data: { labels: ['Ready', 'Process', 'Sample', 'No Process'], datasets: [{ data: facBCounts, backgroundColor: factoryColors, borderRadius: 4 }] },
            options: commonOpts
        });
    }
}

function renderTables() {
    // Render Template In-Line Table
    const tbodyT = document.querySelector('#tblTemplate tbody');
    if (tbodyT && allData.templateInLine && allData.templateInLine.length > 0) {
        const validT = allData.templateInLine.filter(r => (r.style || '').trim() !== '' || (r.proses || '').trim() !== '');

        const buildVideoMini = (row) => {
            let h = "";
            if (row.videoB) h += `<button class="video-pill video-pill-before" onclick="event.stopPropagation();openVideoPlayer('${escapeHtml(row.videoB)}','Video Before')" title="Video Before"><i class="ph ph-play-circle"></i> B</button>`;
            if (row.videoA) h += `<button class="video-pill video-pill-after" onclick="event.stopPropagation();openVideoPlayer('${escapeHtml(row.videoA)}','Video After')" title="Video After"><i class="ph ph-play-circle"></i> A</button>`;
            if (!h) return `<span style="color:var(--text-dim);">&#8212;</span>`;
            return `<div style="display:flex; gap:4px; justify-content:center; align-items:center;">${h}</div>`;
        };

        const isIE = !!window.ieTokenPublic;
        const aksiHdr = document.getElementById('ieAksiHeader');
        if (aksiHdr) aksiHdr.style.display = isIE ? '' : 'none';

        let lastLine = null;
        let lastStyle = null;

        tbodyT.innerHTML = validT.map(row => {
            const sNum = parseFloat(row.saving);
            const sCol = sNum > 0 ? 'var(--neon-green)' : (sNum < 0 ? 'var(--neon-red)' : 'var(--text-dim)');
            const sStr = (row.saving !== '' && row.saving !== undefined) ? Number(row.saving).toFixed(1) : '-';
            const rStr = row.rate ? row.rate + '%' : '-';

            // Logic untuk menyembunyikan duplikat Line & Style agar tampilan lebih bersih
            const currentLine = (row.line || '').toString().trim();
            const currentStyle = (row.style || '').toString().trim();
            let showLine = currentLine;
            let showStyle = currentStyle;

            if (currentLine === lastLine) {
                showLine = ""; // Sembunyikan jika sama dengan atasnya
                if (currentStyle === lastStyle) {
                    showStyle = ""; // Sembunyikan style jika line dan style sama
                } else {
                    lastStyle = currentStyle;
                }
            } else {
                lastLine = currentLine;
                lastStyle = currentStyle;
            }

            // Lookup timestamp dari Archive Production (match Line+Style+Proses)
            let tsStr = '-';
            const archiveProd = allData.archiveProduction || [];
            const archMatch = archiveProd.slice().reverse().find(a =>
                String(a.line || '').trim().toLowerCase() === String(row.line || '').trim().toLowerCase() &&
                String(a.style || '').trim().toLowerCase() === String(row.style || '').trim().toLowerCase() &&
                String(a.proses || '').trim().toLowerCase() === String(row.proses || '').trim().toLowerCase()
            );
            if (archMatch && archMatch.timestamp) tsStr = formatDateString(archMatch.timestamp);

            // Video Before/After buttons (inline player)
            const vBCell = row.videoB
                ? `<button class="video-pill video-pill-before" onclick="openVideoPlayer('${escapeHtml(row.videoB)}','Video Before — ${escapeHtml(row.style || '')}')" title="Putar Video Before"><i class="ph ph-play-circle"></i> Play</button>`
                : `<span style="color:var(--text-dim);">&#8212;</span>`;
            const vACell = row.videoA
                ? `<button class="video-pill video-pill-after" onclick="openVideoPlayer('${escapeHtml(row.videoA)}','Video After — ${escapeHtml(row.style || '')}')" title="Putar Video After"><i class="ph ph-play-circle"></i> Play</button>`
                : `<span style="color:var(--text-dim);">&#8212;</span>`;

            return `<tr>
                <td style="font-weight:bold; color:var(--neon-cyan);">${showLine}</td>
                <td style="font-weight:500;">${showStyle}</td>
                <td>${row.proses || ''}</td>
                <td><span class="badge ${getStatusBadge(row.status || '')}">${row.status || ''}</span></td>
                <td style="text-align:center;color:var(--neon-blue);font-weight:500;">${row.smv || '-'}</td>
                <td style="text-align:center;">${row.actual || '-'}</td>
                <td style="text-align:center;font-weight:bold;color:${sCol};">${sStr}</td>
                <td style="text-align:center;">${rStr}</td>
                <td style="text-align:center;">${vBCell}</td>
                <td style="text-align:center;">${vACell}</td>
                <td style="text-align:center;font-size:0.65rem;color:var(--text-dim);">${tsStr}</td>
            </tr>`;
        }).join('');
    }

    // ==========================
    // IE DEDICATED EDITOR VIEW (New)
    // ==========================
    window.renderIeEditorTable = function () {
        const tbody = document.querySelector('#tblIeEditor tbody');
        if (!tbody) return;
        const rows = allData.templateInLine || [];
        if (rows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="6" class="text-center">Belum ada data template.</td></tr>';
            return;
        }

        let lastLine = null;
        let lastStyle = null;

        const validRows = rows.filter(r => (r.style || '').trim() !== '' || (r.proses || '').trim() !== '');
        tbody.innerHTML = validRows.map(row => {
            const currentLine = (row.line || '').toString().trim();
            const currentStyle = (row.style || '').toString().trim();
            let showLine = currentLine;
            let showStyle = currentStyle;

            if (currentLine === lastLine) {
                showLine = "";
                if (currentStyle === lastStyle) {
                    showStyle = "";
                } else {
                    lastStyle = currentStyle;
                }
            } else {
                lastLine = currentLine;
                lastStyle = currentStyle;
            }

            const editArgs = `'${escapeHtml(row.line)}','${escapeHtml(row.style)}','${escapeHtml(row.proses)}','${escapeHtml(String(row.smv || ''))}','${escapeHtml(String(row.actual || ''))}','${escapeHtml(row.videoB || '')}','${escapeHtml(row.videoA || '')}'`;
            return `<tr style="cursor:pointer;" onclick="openIeModalPublicFromEditor(${editArgs})">
                <td style="font-weight:bold; color:var(--neon-cyan);">${showLine}</td>
                <td style="font-weight:500;">${showStyle}</td>
                <td>${row.proses || ''}</td>
                <td style="text-align:center;">${row.smv || '-'}</td>
                <td style="text-align:center;">${row.actual || '-'}</td>
                <td style="text-align:center;">
                    <button class="action-btn outline-btn" style="padding:0.2rem 0.5rem;font-size:0.7rem;color:var(--neon-cyan);border-color:var(--neon-cyan);">
                        <i class="ph ph-pencil"></i> Edit
                    </button>
                </td>
            </tr>`;
        }).join('');
    };

    // Render Archive Tables setelah data siap
    if (typeof renderArchiveProductionTable === 'function') renderArchiveProductionTable();
    if (typeof renderArchiveTemplateTable === 'function') renderArchiveTemplateTable();

    // Render Grid Buttons Factory A (1-12) & B (13-22)
    const gridA = document.getElementById('gridFactoryA');
    if (gridA && gridA.children.length === 0) {
        gridA.innerHTML = '';
        for (let i = 1; i <= 12; i++) {
            gridA.innerHTML += `<button class="action-btn outline-btn line-btn" data-line="${i}" style="padding: 1rem 0;">Line ${i}</button>`;
        }
    }
    const gridB = document.getElementById('gridFactoryB');
    if (gridB && gridB.children.length === 0) {
        gridB.innerHTML = '';
        for (let i = 13; i <= 22; i++) {
            gridB.innerHTML += `<button class="action-btn outline-btn line-btn" data-line="${i}" style="padding: 1rem 0;">Line ${i}</button>`;
        }
    }

    // Attach click events for drilling down lines
    document.querySelectorAll('.line-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            let lineNum = e.target.closest('.line-btn').dataset.line;
            openLineDataModal(lineNum);
        });
    });

    // Render Activity Log (Separated IE & Sewing)
    const renderLogs = (type, tableId) => {
        const tbody = document.querySelector(`${tableId} tbody`);
        if (!tbody) return;

        // Filter data dari berbagai kemungkinan kunci (layoutRequests, sheet25/26, requestIE/Sewing)
        let sourceData = [];
        const isIE = type.toUpperCase() === 'IE';
        
        if (isIE) {
            sourceData = allData.sheet25 || allData.requestIE || allData.logIE || allData.layoutRequests || [];
        } else {
            sourceData = allData.sheet26 || allData.requestSewing || allData.logSewing || allData.layoutRequests || [];
        }

        const filtered = sourceData
            .filter(row => {
                // Jika data dari layoutRequests, filter berdasarkan type. Jika dari sheet25/26, ambil semua.
                if (row.type) return row.type.toLowerCase() === type.toLowerCase();
                return true; 
            })
            .slice().reverse();

        if (filtered.length > 0) {
            tbody.innerHTML = filtered.slice(0, 15).map(row => {
                const statusKirim = row.status_kirim || row.status_penyelesaian || 'Terkirim';
                const statusTemplate = row.status_template || 'Pending';
                const timestamp = row.timestamp || '-';
                
                if (type.toUpperCase() === 'IE') {
                    const tglLayout = row.tanggal || row.tgl || row.tglLayout || row.tanggal_layout || row.date || row.note || row.keterangan || '-';
                    return `
                    <tr>
                        <td style="font-weight:500;">${row.style || '-'}</td>
                        <td style="font-weight:bold; color:var(--neon-cyan); text-align:center;">${row.line || '-'}</td>
                        <td>${row.proses || '-'}</td>
                        <td>${tglLayout}</td>
                        <td><span class="badge ${getStatusBadge(statusKirim)}">${statusKirim}</span></td>
                        <td><span class="badge ${getStatusBadge(statusTemplate)}">${statusTemplate}</span></td>
                        <td style="font-size:0.7rem; color:var(--text-dim);">${timestamp}</td>
                    </tr>`;
                } else {
                    return `
                    <tr>
                        <td style="font-weight:500;">${row.style || '-'}</td>
                        <td style="font-weight:bold; color:var(--neon-cyan); text-align:center;">${row.line || '-'}</td>
                        <td>${row.proses || '-'}</td>
                        <td>${row.keterangan || row.note || '-'}</td>
                        <td><span class="badge ${getStatusBadge(statusKirim)}">${statusKirim}</span></td>
                        <td><span class="badge ${getStatusBadge(statusTemplate)}">${statusTemplate}</span></td>
                        <td style="font-size:0.7rem; color:var(--text-dim);">${timestamp}</td>
                    </tr>`;
                }
            }).join('');
        } else {
            tbody.innerHTML = `<tr><td colspan="7" class="text-center" style="color:var(--text-dim); padding:1rem;">Belum ada riwayat ${type}.</td></tr>`;
        }
    };

    renderLogs('IE', '#tblLogIE');
    renderLogs('Sewing', '#tblLogSewing');

    // Render Inventory
    const tbodyInv = document.querySelector('#tblInventory tbody');
    if (tbodyInv && allData.inventory && allData.inventory.length > 0) {
        let lastBuyer = null;
        let lastCode = null;

        tbodyInv.innerHTML = allData.inventory.map(row => {
            const curBuyer = (row.buyer || '').toString().trim();
            const curCode = (row.code || '').toString().trim();
            let showBuyer = curBuyer;
            let showCode = curCode;

            if (curBuyer === lastBuyer) {
                showBuyer = "";
                if (curCode === lastCode) {
                    showCode = "";
                } else {
                    lastCode = curCode;
                }
            } else {
                lastBuyer = curBuyer;
                lastCode = curCode;
            }

            return `<tr>
                <td>${row.no || '-'}</td>
                <td>${row.tanggal || '-'}</td>
                <td style="font-weight:bold; color:var(--neon-orange);">${showBuyer}</td>
                <td style="font-weight:500;">${showCode}</td>
                <td>${row.proses || '-'}</td>
                <td>${row.qty || '-'}</td>
                <td><span class="badge ${getStatusBadge(row.status || '')}">${row.status || '-'}</span></td>
                <td>${getPositionBadge(row.code)}</td>
                <td>${row.size || '-'}</td>
            </tr>`;
        }).join('');

        // Render Inventory Summary Stats Cards
        renderInventoryStats(allData.inventory);
    }
}

// ==========================
// INVENTORY SUMMARY STATS (Migrated from admin.js)
// ==========================
function renderInventoryStats(validInv) {
    if (!validInv || !Array.isArray(validInv)) return;
    
    // Filter only valid rows
    const filtered = validInv.filter(r => (r.code || '').trim() !== '' || (r.buyer || '').trim() !== '' || (r.proses || '').trim() !== '');
    if (filtered.length === 0) return;

    // Hitung nilai unik
    const buyers = [...new Set(filtered.map(r => (r.buyer || '').trim()).filter(v => v !== ''))];
    const codes = [...new Set(filtered.map(r => (r.code || '').trim()).filter(v => v !== ''))];
    const sizes = [...new Set(filtered.map(r => (r.size || '').trim()).filter(v => v !== ''))];
    const prosesUniq = [...new Set(filtered.map(r => (r.proses || '').trim()).filter(v => v !== ''))];

    // Total baris yang ada proses-nya
    const totalProsesRows = filtered.filter(r => (r.proses || '').trim() !== '').length;

    // Rincian Style Code PER BUYER
    const codePerBuyer = {};
    filtered.forEach(r => {
        const buyer = (r.buyer || '').trim();
        const code = (r.code || '').trim();
        if (!buyer || !code) return;
        if (!codePerBuyer[buyer]) codePerBuyer[buyer] = new Set();
        codePerBuyer[buyer].add(code);
    });

    // Rincian Proses PER BUYER
    const prosesPerBuyer = {};
    filtered.forEach(r => {
        const buyer = (r.buyer || '').trim();
        const proses = (r.proses || '').trim();
        if (!buyer || !proses) return;
        prosesPerBuyer[buyer] = (prosesPerBuyer[buyer] || 0) + 1;
    });

    // Total Qty — handle format angka Indonesia
    const totalQty = filtered.reduce((sum, r) => {
        let raw = (r.qty || '0').toString().trim();
        raw = raw.replace(/\./g, '').replace(',', '.');
        const num = parseFloat(raw);
        return sum + (isNaN(num) ? 0 : num);
    }, 0);

    const setText = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    const setHtml = (id, val) => { const el = document.getElementById(id); if (el) el.innerHTML = val; };

    // Kartu Buyer
    setText('invStatBuyerCount', buyers.length);
    const buyerListEl = document.getElementById('invStatBuyerList');
    if (buyerListEl) buyerListEl.textContent = buyers.length > 0 ? buyers.join(' • ') : '-';

    // Kartu Style Code — rincian per buyer
    setText('invStatCodeCount', codes.length);
    const codeDetailEl = document.getElementById('invStatCodeDetail');
    if (codeDetailEl) {
        const entries = Object.entries(codePerBuyer);
        if (entries.length === 0) {
            codeDetailEl.textContent = '-';
        } else {
            codeDetailEl.innerHTML = entries
                .map(([b, set]) => `<span style="display:block;"><b style="color:#818cf8;">${b}:</b> ${set.size} kode</span>`)
                .join('');
        }
    }

    // Kartu Proses — unik + total keseluruhan + rincian per buyer
    setText('invStatProsesCount', prosesUniq.length);
    const prosesTotalEl = document.getElementById('invStatProsesTotal');
    if (prosesTotalEl) {
        const buyerLines = Object.entries(prosesPerBuyer)
            .map(([b, cnt]) => `<span style="display:block;"><b style="color:#0d9488;">${b}:</b> ${cnt} proses</span>`)
            .join('');
        prosesTotalEl.innerHTML =
            `<span style="display:block;color:#94a3b8;margin-bottom:0.2rem;">Total: ${totalProsesRows} baris</span>` +
            buyerLines;
    }

    // Kartu Size
    setText('invStatSizeCount', sizes.length);
    const sizeListEl = document.getElementById('invStatSizeList');
    if (sizeListEl) sizeListEl.textContent = sizes.length > 0 ? sizes.join(' • ') : '-';

    // Kartu Qty
    setText('invStatTotalQty', totalQty % 1 === 0 ? totalQty : totalQty.toFixed(1));
}

// Helper: Status Background Mapping
function getStatusBadge(status) {
    status = status.toLowerCase();
    if (['running', 'ready', 'terkirim', 'available'].includes(status)) return 'bg-green';
    if (['not used', 'process', 'low stock'].includes(status)) return 'bg-blue';
    if (['no process', 'ditolak', 'out of stock'].includes(status)) return 'bg-red';
    if (['pilot', 'sample', 'pending', 'on order'].includes(status)) return 'bg-orange';
    return 'bg-blue';
}

// Helper: Cari Posisi Barang Berdasarkan Data Aktif Dashboard (Dengan UI/UX)
function getPositionBadge(styleCode) {
    const defaultBadge = `<span class="badge" style="background:rgba(255,255,255,0.05); border:1px solid #555; color:#aaa;"><i class="ph ph-package"></i> Di Gudang</span>`;

    if (!styleCode || styleCode.trim() === '') return defaultBadge;
    const query = styleCode.toLowerCase();

    if (allData.templateInLine) {
        const found = allData.templateInLine.find(r => (r.style || '').toLowerCase().includes(query) || (r.code || '').toLowerCase().includes(query));
        if (found) return `<span class="badge" style="background:rgba(239,68,68,0.1); border:1px solid #ef4444; color:#ef4444; gap:0.3rem;"><i class="ph ph-grid-four"></i> Di In-Line</span>`;
    }

    // Pencarian Factory A (Line 1-12) -> Hijau
    if (allData.factoryA) {
        for (const [line, arr] of Object.entries(allData.factoryA)) {
            if (Array.isArray(arr) && arr.find(r => (r.style || '').toLowerCase().includes(query))) {
                return `<span class="badge" style="background:rgba(0,255,136,0.1); border:1px solid var(--neon-green); color:var(--neon-green); gap:0.3rem;"><i class="ph ph-factory"></i> Di ${line}</span>`;
            }
        }
    }

    // Pencarian Factory B (Line 13-22) -> Ungu / Magenta
    if (allData.factoryB) {
        for (const [line, arr] of Object.entries(allData.factoryB)) {
            if (Array.isArray(arr) && arr.find(r => (r.style || '').toLowerCase().includes(query))) {
                return `<span class="badge" style="background:rgba(217,70,239,0.1); border:1px solid var(--neon-purple); color:var(--neon-purple); gap:0.3rem;"><i class="ph ph-factory"></i> Di ${line}</span>`;
            }
        }
    }

    return defaultBadge;
}

// Tambahkan pemanggilan di akhir renderTables agar selalu update saat data masuk
const originalRenderTables = renderTables;
renderTables = function() {
    originalRenderTables();
    renderPerformanceStats();
};

// ==========================
// MODULE: TOP PERFORMANCE STATS
// ==========================
const MONTHS_MAP_LONG = { jan: 'Jan', feb: 'Feb', mar: 'Mar', apr: 'Apr', may: 'May', jun: 'Jun', jul: 'Jul', aug: 'Aug', sep: 'Sep', oct: 'Oct', nov: 'Nov', dec: 'Dec' };

function renderPerformanceStats() {
    const logsLayout = allData.layoutRequests || [];
    const logsTemplate = allData.riwayatTemplate || [];
    const logsArchive = allData.archiveProduction || [];

    // Init Dropdown jika kosong
    const filterMonth = document.getElementById('filterMonthPerf');
    if (filterMonth && filterMonth.options.length <= 1) {
        const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        const curMonthIdx = new Date().getMonth();
        months.forEach((m, i) => {
            const opt = document.createElement('option');
            opt.value = Object.values(MONTHS_MAP_LONG)[i];
            opt.textContent = m;
            if (i === curMonthIdx) opt.selected = true;
            filterMonth.appendChild(opt);
        });
    }

    const selectedMonth = filterMonth ? filterMonth.value : 'All';

    const stats = {};
    const savingStats = {};
    for (let i = 1; i <= 22; i++) {
        stats[`Line ${i}`] = { line: `Line ${i}`, running: 0, not_used: 0, no_process: 0, layout: 0 };
        savingStats[`Line ${i}`] = { line: `Line ${i}`, totalMinutes: 0, sumRate: 0, countRate: 0, peakRate: 0 };
    }

    const filterByMonth = (r) => {
        if (selectedMonth === 'All') return true;
        return String(r.timestamp || '').toLowerCase().includes(selectedMonth.toLowerCase());
    };

    logsLayout.filter(filterByMonth).forEach(r => {
        const ln = String(r.line || '').replace(/line/i, '').trim();
        const key = `Line ${ln}`;
        if (stats[key]) stats[key].layout++;
    });

    logsTemplate.filter(filterByMonth).forEach(r => {
        const ln = String(r.line || '').replace(/line/i, '').trim();
        const key = `Line ${ln}`;
        if (stats[key]) {
            const st = String(r.status || '').toLowerCase();
            if (st.includes('not used')) stats[key].not_used++;
            else if (st.includes('no process')) stats[key].no_process++;
            else if (st.includes('running')) stats[key].running++;
        }
    });

    logsArchive.filter(filterByMonth).forEach(r => {
        const ln = String(r.line || '').replace(/line/i, '').trim();
        const key = `Line ${ln}`;
        if (savingStats[key]) {
            let smv = parseFloat(String(r.smv || '0').replace(',', '.')) || 0;
            let act = parseFloat(String(r.actual || '0').replace(',', '.')) || 0;
            if (smv > 0 && act > 0) {
                let diff = smv - act;
                let rate = Math.round((diff / smv) * 100);
                savingStats[key].totalMinutes += diff;
                savingStats[key].sumRate += rate;
                savingStats[key].countRate++;
                if (rate > savingStats[key].peakRate) savingStats[key].peakRate = rate;
            }
        }
    });

    const statsArr = Object.values(stats);
    const savingArr = Object.values(savingStats).map(s => ({
        ...s,
        avgRate: s.countRate > 0 ? Math.round(s.sumRate / s.countRate) : 0
    }));

    const buildRankList = (sorted, key, color, emptyMsg, isNegative = false, unit = 'x') => {
        if (!sorted.length) return `<div style="color:var(--text-dim); font-size:0.75rem;">${emptyMsg}</div>`;
        const max = Math.abs(sorted[0][key]) || 1;
        return sorted.slice(0, 7).map((item, i) => {
            const val = item[key];
            const pct = Math.round((Math.abs(val) / max) * 100);
            let medal = `${i + 1}.`;
            
            if (!isNegative) {
                if (i === 0) medal = '<i class="ph ph-trophy" style="color:#f59e0b;"></i>';
                else if (i === 1) medal = '<i class="ph ph-crown" style="color:#94a3b8;"></i>';
                else if (i === 2) medal = '<i class="ph ph-medal" style="color:#b45309;"></i>';
            } else {
                if (i === 0) medal = '<i class="ph ph-warning-octagon" style="color:var(--neon-red);"></i>';
                else if (i === 1) medal = '<i class="ph ph-warning" style="color:var(--neon-red);"></i>';
                else if (i === 2) medal = '<i class="ph ph-info" style="color:var(--text-dim);"></i>';
            }
            
            const displayVal = typeof val === 'number' ? (key === 'totalMinutes' ? val.toFixed(1) : val) : val;

            return `
            <div style="margin-bottom:0.2rem;">
                <div style="display:flex; justify-content:space-between; font-size:0.75rem; margin-bottom:0.15rem;">
                    <span style="display:flex; align-items:center; gap:0.3rem;"><b>${medal}</b> ${item.line}</span>
                    <span style="font-weight:700; color:${color};">${displayVal}${unit}</span>
                </div>
                <div style="background:rgba(255,255,255,0.05); border-radius:99px; height:5px; overflow:hidden;">
                    <div style="width:${pct}%; height:100%; background:${color}; border-radius:99px; transition:width 0.4s ease;"></div>
                </div>
            </div>`;
        }).join('');
    };

    const setHtml = (id, v) => { const el = document.getElementById(id); if (el) el.innerHTML = v; };
    
    setHtml('rankRunning', buildRankList(statsArr.filter(r => r.running > 0).sort((a,b) => b.running - a.running), 'running', 'var(--neon-green)', 'Tidak ada data Running.'));
    setHtml('rankNotUsed', buildRankList(statsArr.filter(r => r.not_used > 0).sort((a,b) => b.not_used - a.not_used), 'not_used', 'var(--neon-red)', 'Tidak ada data Not Used.', true));
    setHtml('rankNoProcess', buildRankList(statsArr.filter(r => r.no_process > 0).sort((a,b) => b.no_process - a.no_process), 'no_process', 'var(--text-dim)', 'Tidak ada data No Process.', true));
    setHtml('rankLayoutReq', buildRankList(statsArr.filter(r => r.layout > 0).sort((a,b) => b.layout - a.layout), 'layout', '#ea580c', 'Tidak ada request layout.'));

    // Render Saving Analysis
    setHtml('rankAvgSaving', buildRankList(savingArr.filter(s => s.avgRate > 0).sort((a,b) => b.avgRate - a.avgRate), 'avgRate', 'var(--neon-cyan)', 'Belum ada data saving.', false, '%'));
    setHtml('rankTotalSaving', buildRankList(savingArr.filter(s => s.totalMinutes > 0).sort((a,b) => b.totalMinutes - a.totalMinutes), 'totalMinutes', 'var(--neon-blue)', 'Belum ada data saving.', false, ' min'));
    setHtml('rankPeakSaving', buildRankList(savingArr.filter(s => s.peakRate > 0).sort((a,b) => b.peakRate - a.peakRate), 'peakRate', '#a855f7', 'Belum ada rekor saving.', false, '%'));

    // Idle Lines
    const idleLines = statsArr.filter(r => r.running === 0 && (r.not_used > 0 || r.no_process > 0 || r.layout > 0));
    const idleEl = document.getElementById('listLineIdle');
    if (idleEl) {
        idleEl.innerHTML = idleLines.length
            ? idleLines.map(r => `<span style="display:inline-block; margin:0.15rem; padding:0.2rem 0.55rem; background:rgba(217,70,239,0.1); border:1px solid rgba(217,70,239,0.2); border-radius:6px; font-size:0.7rem; color:var(--neon-purple);">${r.line}</span>`).join('')
            : `<div style="color:var(--neon-green); font-size:0.75rem;"><i class="ph ph-check-circle"></i> Semua line pernah Running.</div>`;
    }
}

// ==========================
// MODALS LOGIC
// ==========================
function initModals() {
    // Universal Close Modal
    document.querySelectorAll('.close-modal, .modal').forEach(el => {
        el.addEventListener('click', (e) => {
            if (e.target === el || e.target.closest('.close-modal')) {
                document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
            }
        });
    });

    // 1. SMART SEARCH MODAL & OCR
    const btnSmartSearch = document.getElementById('startCameraBtn');
    const modalSearch = document.getElementById('smartSearchModal');
    if (btnSmartSearch) btnSmartSearch.addEventListener('click', () => { modalSearch.classList.remove('hidden'); });

    // Request Form Modals Logic
    const reqModal = document.getElementById('requestModal');
    document.getElementById('btnNewReqIE')?.addEventListener('click', () => openRequestForm('IE'));
    document.getElementById('btnNewReqSewing')?.addEventListener('click', () => openRequestForm('Sewing'));

    document.getElementById('btnGoInventory')?.addEventListener('click', () => switchTab('view-inventory'));
    document.getElementById('btnGoArchive')?.addEventListener('click', () => switchTab('view-archive'));
    document.getElementById('btnGoPerformance')?.addEventListener('click', () => switchTab('view-performance'));

    // Populate Line Dropdown (1-22)
    const reqLineDrp = document.getElementById('reqLine');
    if (reqLineDrp) {
        for (let i = 1; i <= 22; i++) reqLineDrp.innerHTML += `<option value="${i}">Line ${i}</option>`;
    }

    document.getElementById('requestForm')?.addEventListener('submit', validateSubmitForm);

    // OCR Init inside Modal (similar to before but adapted for search container)
    initOCRExpanded();
}

function openRequestForm(type) {
    document.getElementById('reqModalTitle').innerHTML = type === 'IE' ? '<i class="ph ph-clipboard-text"></i> New IE Request' : '<i class="ph ph-scissors"></i> New Sewing Request';
    document.getElementById('reqType').value = type;

    // Toggle Specific fields
    if (type === 'IE') {
        document.getElementById('reqIELayoutDateGroup').classList.remove('hidden');
        document.getElementById('reqSewingNoteGroup').classList.add('hidden');
        document.getElementById('reqTanggal').required = true;
        document.getElementById('reqKeterangan').required = false;
    } else {
        document.getElementById('reqIELayoutDateGroup').classList.add('hidden');
        document.getElementById('reqSewingNoteGroup').classList.remove('hidden');
        document.getElementById('reqTanggal').required = false;
        document.getElementById('reqKeterangan').required = true;
    }
    document.getElementById('requestModal').classList.remove('hidden');
}

// ==========================
// IE PUBLIC AUTHENTICATION
// ==========================
window.openIeLoginModal = function () {
    if (window.ieTokenPublic) {
        if (confirm("Anda sudah login sebagai IE. Ingin Logout?")) {
            handleIELogout();
        }
    } else {
        document.getElementById('modalIELoginPublic').classList.remove('hidden');
    }
};

window.handleIELogin = async function (e) {
    e.preventDefault();
    const btn = document.getElementById('btnIELoginSubmit');
    const pwd = document.getElementById('iePasswordPublic').value;

    btn.innerHTML = '<i class="ph ph-spinner-gap spin"></i> Verifikasi...';
    btn.disabled = true;

    try {
        const response = await fetch(`${APPS_SCRIPT_URL}?action=loginIE&password=${encodeURIComponent(pwd)}`);
        let res;
        try {
            res = await response.json();
        } catch (_) {
            throw new Error('Respons server tidak valid. Periksa URL deployment Apps Script.');
        }

        // DEBUG - lihat di Console browser (F12 > Console)
        console.log('[IE Login] Raw response:', JSON.stringify(res));

        if (!res || typeof res !== 'object') {
            throw new Error('Format respons server tidak dikenali.');
        }

        if (res.status === 'success') {
            window.ieTokenPublic = res.token;
            sessionStorage.setItem('ieTokenPublic', res.token);
            document.getElementById('modalIELoginPublic').classList.add('hidden');
            showToast('success', 'Akses IE Terbuka!');

            // ALIHKAN KE EDITOR IE
            const btnEditor = document.getElementById('btnGoIeUpdate');
            if (btnEditor) btnEditor.click();
        } else {
            // Tampilkan pesan spesifik dari server jika ada
            const errMsg = res.message || res.error || `Server respons: status="${res.status || 'tidak ada'}"`;
            showToast('error', errMsg);
        }
    } catch (err) {
        showToast('error', 'Gagal Login IE: ' + (err.message || 'Kesalahan tidak diketahui'));
    } finally {
        btn.innerHTML = 'Minta Akses';
        btn.disabled = false;
        document.getElementById('iePasswordPublic').value = '';
    }
};

window.handleIELogout = function () {
    window.ieTokenPublic = '';
    sessionStorage.removeItem('ieTokenPublic');
    showToast('info', 'Anda telah Logout dari akses IE.');
    renderTables();
};

// ==========================
// SPA VIEW SWITCHER (Used by Admin Tools buttons)
// ==========================
window.switchView = function(viewId) {
    // Hide all tab views
    document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));
    
    // Show target view
    const target = document.getElementById(viewId);
    if (target) target.classList.remove('hidden');
    
    // Trigger render for specific views
    if (viewId === 'view-admin-panel') {
        if (typeof renderAdminRequests === 'function') renderAdminRequests();
        if (typeof renderWaShareButtons === 'function') renderWaShareButtons();
    }
    if (viewId === 'view-dashboard') {
        updateCharts();
        updateStatusCards();
    }
    
    window.scrollTo({ top: 0, behavior: 'smooth' });
};

// ==========================
// ADMIN LOGIN & STATE
// ==========================
window.openAdminLoginModal = function () {
    // Selalu paksa masukkan password sesuai permintaan
    document.getElementById('adminPasswordInput').value = '';
    document.getElementById('modalAdminLogin').classList.remove('hidden');
    setTimeout(() => document.getElementById('adminPasswordInput').focus(), 100);
};

window.closeAdminLoginModal = function () {
    document.getElementById('modalAdminLogin').classList.add('hidden');
    document.getElementById('adminPasswordInput').value = '';
};

window.executeAdminLogin = async function () {
    const pwd = document.getElementById('adminPasswordInput').value;
    if (pwd !== 'admin1994') { // Using local validation for speed, similar to admin.js
        showToast('error', 'Sandi admin salah!');
        return;
    }
    
    // As in Code.gs, admin password yields 'ctt_token_1994_secure'
    try {
        const response = await fetch(`${APPS_SCRIPT_URL}?action=login&password=${encodeURIComponent(pwd)}`);
        const res = await response.json();
        
        if (res.status === 'success') {
            window.adminToken = res.token;
            sessionStorage.setItem('adminToken', res.token);
            closeAdminLoginModal();
            showToast('success', 'Akses Admin Terbuka!');
            checkAdminState(); // update UI
            switchView('view-admin-panel'); // Pindah ke halaman Admin Hub
        } else {
            showToast('error', res.message || 'Gagal login admin');
        }
    } catch (err) {
        showToast('error', 'Error login: ' + err);
    }
};

window.logoutAdmin = function () {
    window.adminToken = '';
    sessionStorage.removeItem('adminToken');
    window.isAdminMode = false;
    document.getElementById('adminModeToggle').checked = false;
    checkAdminState();
    
    // Return to dashboard if in admin view
    const activeView = document.querySelector('.tab-view.active');
    if (activeView && ['view-admin-request', 'view-admin-report'].includes(activeView.id)) {
        switchView('view-dashboard');
    }
    
    showToast('info', 'Admin Logout berhasil');
    renderTables();
};

window.toggleAdminMode = function(isActive) {
    if (!window.adminToken) {
        document.getElementById('adminModeToggle').checked = false;
        showToast('error', 'Silakan login admin terlebih dahulu.');
        return;
    }
    window.isAdminMode = isActive;
    
    if (isActive) {
        document.querySelectorAll('.admin-only').forEach(el => el.classList.remove('hidden'));
        document.getElementById('adminToggleSlider').style.backgroundColor = '#f0f';
        document.getElementById('adminToggleSlider').style.boxShadow = '0 0 10px #f0f';
    } else {
        document.querySelectorAll('.admin-only').forEach(el => el.classList.add('hidden'));
        document.getElementById('adminToggleContainer').classList.remove('hidden'); // Keep toggle visible if logged in
        document.getElementById('adminToggleSlider').style.backgroundColor = '#ccc';
        document.getElementById('adminToggleSlider').style.boxShadow = 'none';
        
        // Return to dashboard if in admin view
        const activeView = document.querySelector('.tab-view.active');
        if (activeView && ['view-admin-request', 'view-admin-report'].includes(activeView.id)) {
            switchView('view-dashboard');
        }
    }
    
    // Update tables & UI based on admin mode
    renderTables();
};

window.checkAdminState = function() {
    const loginBtn = document.getElementById('adminLoginBtn');
    const toggleContainer = document.getElementById('adminToggleContainer');
    if(!loginBtn || !toggleContainer) return;
    
    if (window.adminToken) {
        loginBtn.innerHTML = '<i class="ph ph-sign-out"></i>';
        loginBtn.onclick = logoutAdmin;
        loginBtn.title = "Logout Admin";
        toggleContainer.classList.remove('hidden');
    } else {
        loginBtn.innerHTML = '<i class="ph ph-lock-key"></i>';
        loginBtn.onclick = openAdminLoginModal;
        loginBtn.title = "Admin Login";
        toggleContainer.classList.add('hidden');
        window.isAdminMode = false;
        document.getElementById('adminModeToggle').checked = false;
        document.querySelectorAll('.admin-only').forEach(el => el.classList.add('hidden'));
    }
};

// ==========================
// ADMIN API ACTIONS
// ==========================

window.adminEditRow = function(sheetName, rowIndex) {
    if (!window.adminToken) return;
    
    // Find the data row from local cache
    let rowData = null;
    if (sheetName === 'Template In-Line') {
        rowData = allData.templateInLine.find(r => r.rowIndex === rowIndex);
    } else if (sheetName === 'Sheet3' || sheetName === 'Factory A') { // Factory A
        for (const line in allData.factoryA) {
            const match = allData.factoryA[line].find(r => r.rowIndex === rowIndex);
            if (match) { rowData = match; break; }
        }
    } else if (sheetName === 'Sheet4' || sheetName === 'Factory B') { // Factory B
        for (const line in allData.factoryB) {
            const match = allData.factoryB[line].find(r => r.rowIndex === rowIndex);
            if (match) { rowData = match; break; }
        }
    }

    if (!rowData) {
        showToast('error', 'Data tidak ditemukan di memori.');
        return;
    }

    // Populate modal
    const targetSheet = rowData.sheetAsal || (sheetName === 'Template In-Line' ? 'Sheet2' : (sheetName === 'Factory A' ? 'Sheet3' : 'Sheet4'));
    document.getElementById('adminEditSheet').value = targetSheet;
    document.getElementById('adminEditRowIndex').value = rowIndex;
    document.getElementById('adminEditLine').value = rowData.line || '';
    document.getElementById('adminEditStyle').value = rowData.style || rowData.buyer || '';
    document.getElementById('adminEditProses').value = rowData.proses || '';
    
    const statusSelect = document.getElementById('adminEditStatus');
    const rowStatus = (rowData.status || '').toLowerCase();
    Array.from(statusSelect.options).forEach(opt => {
        if (opt.value.toLowerCase() === rowStatus) opt.selected = true;
    });

    document.getElementById('modalAdminEdit').classList.remove('hidden');
};

window.closeAdminEditModal = function() {
    document.getElementById('modalAdminEdit').classList.add('hidden');
};

window.saveAdminEdit = async function(e) {
    e.preventDefault();
    if (!window.adminToken) return;
    
    const btn = e.target.querySelector('button[type="submit"]');
    btn.innerHTML = '<i class="ph ph-spinner-gap spin"></i> Menyimpan...';
    btn.disabled = true;

    try {
        const payload = {
            action: 'updateRow',
            sheetName: document.getElementById('adminEditSheet').value,
            rowIndex: parseInt(document.getElementById('adminEditRowIndex').value),
            line: document.getElementById('adminEditLine').value,
            style: document.getElementById('adminEditStyle').value,
            proses: document.getElementById('adminEditProses').value,
            status: document.getElementById('adminEditStatus').value,
            token: window.adminToken
        };

        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
        const resJson = await res.json();

        if (resJson.status === 'success') {
            showToast('success', 'Data tersimpan!');
            closeAdminEditModal();
            loadAllDataFallback(); // Refresh background
        } else {
            throw new Error(resJson.message || 'Gagal menyimpan');
        }
    } catch (err) {
        showToast('error', err.message);
    } finally {
        btn.innerHTML = '<i class="ph ph-floppy-disk"></i> Simpan Perubahan';
        btn.disabled = false;
    }
};

window.deleteSingleRow = async function(sheetName, rowIndex, btnElem) {
    if (!window.adminToken) return;
    if (!confirm('PERINGATAN! Anda akan menghapus baris data ini selamanya dari Google Sheets. Lanjutkan?')) return;

    const tr = btnElem.closest('tr');
    if (tr) tr.style.opacity = '0.3';

    showToast('info', 'Menghapus data...');
    let targetSheet = sheetName === 'Template In-Line' ? 'Sheet2' : sheetName;

    try {
        const payload = { action: 'deleteRows', rows: [{ sheetName: targetSheet, rowIndex: rowIndex }], token: window.adminToken };
        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
        const resJson = await res.json();

        if (resJson.status === 'success') {
            showToast('success', resJson.message);
            if (tr) tr.remove();
            setTimeout(loadAllDataFallback, 1500);
        } else {
            throw new Error(resJson.message);
        }
    } catch (e) {
        showToast('error', 'Gagal hapus data: ' + e);
        if (tr) tr.style.opacity = '1';
    }
};

window.insertAjaib = async function(sheetName, rIndex, lineInfo) {
    if (!window.adminToken) return;
    if (!confirm(`Menyisipkan ruangan baru persis di bawah baris ini?`)) return;

    showToast('info', 'Menyisipkan baris baru...');
    let targetSheet = sheetName === 'Template In-Line' ? 'Sheet2' : sheetName;
    
    try {
        const payload = { action: 'insertRow', sheetName: targetSheet, rowIndex: rIndex, line: lineInfo, token: window.adminToken };
        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
        const resJson = await res.json();

        if (resJson.status === 'success') {
            showToast('success', 'Baris berhasil terbuat!');
            loadAllDataFallback();
        } else { 
            throw new Error(resJson.message); 
        }
    } catch (e) { 
        showToast('error', 'Gagal menyisip: ' + e); 
    }
};

window.adminCompleteLayout = async function(line, style) {
    if (!window.adminToken) return;
    if (!confirm(`Tandai Layout untuk Line ${line} (Style ${style}) sebagai SELESAI? Ini akan memindahkan data ke Factory/Archive.`)) return;
    
    showToast('info', 'Memproses...');
    try {
        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'completeLayout', line: line, style: style, token: window.adminToken }) });
        const resJson = await res.json();
        if (resJson.status === 'success') {
            showToast('success', 'Layout ditandai selesai!');
            loadAllDataFallback();
        } else {
            throw new Error(resJson.message);
        }
    } catch (e) { 
        showToast('error', 'Gagal update layout: ' + e);
    }
};

// ==========================
// REQUEST RENDER (Admin) & WA SHARE UI
// ==========================

window.renderWaShareButtons = function() {
    const container = document.getElementById('waShareLineContainer');
    if (!container) return;
    let html = '';
    for (let i = 1; i <= 22; i++) {
        html += `<button class="action-btn" onclick="shareTextWA('${i}')" style="padding: 0.4rem; background: rgba(37,211,102,0.05); border: 1px solid rgba(37,211,102,0.4); color: #25d366; font-size: 0.75rem;">Line ${i}</button>`;
    }
    container.innerHTML = html;
};

window.renderAdminRequests = function() {
    const tbody = document.querySelector('#tblAdminRequest tbody');
    if (!tbody) return;

    if (allData.layoutRequests && allData.layoutRequests.length > 0) {
        tbody.innerHTML = allData.layoutRequests.slice().reverse().map(row => {
            let aksiBtn = `<button class="action-btn" style="padding:0.2rem 0.5rem; font-size:0.7rem; background:rgba(0,191,255,0.1); border:1px solid var(--neon-cyan); color:var(--neon-cyan);" onclick="adminCompleteLayout('${escapeHtml(row.line||'')}', '${escapeHtml(row.style||'')}')">Tandai Selesai</button>`;
            
            return `
            <tr>
                <td style="white-space:nowrap;"><span style="font-size:0.75rem; color:var(--text-dim);">${row.timestamp || '-'}</span></td>
                <td><span class="badge" style="background:rgba(79,70,229,0.05); border:1px solid rgba(79,70,229,0.1); color:#4f46e5;">${row.line || '-'}</span></td>
                <td style="font-weight: 500;">${row.style || '-'}</span></td>
                <td style="text-align:center;">${(row.status_penyelesaian||'').toLowerCase().includes('selesai') ? '<span class="badge bg-green">Selesai</span>' : aksiBtn}</td>
            </tr>`;
        }).join('');
    } else {
        tbody.innerHTML = '<tr><td colspan="4" class="text-center" style="padding:2rem;color:var(--text-dim);">Belum ada history request.</td></tr>';
    }
};

window.renderAdminReport = function() {
    const tbody = document.querySelector('#tblAdminReport tbody');
    if (!tbody) return;

    if (allData.archiveProduction && allData.archiveProduction.length > 0) {
        tbody.innerHTML = allData.archiveProduction.slice().reverse().map(row => {
            return `
            <tr>
                <td style="white-space:nowrap; font-size:0.75rem; color:var(--text-dim);">${formatDateString(row.timestamp || '')}</td>
                <td style="font-weight:bold; color:var(--neon-cyan);">${row.line || '-'}</td>
                <td style="font-weight:500;">${row.style || '-'}</td>
                <td>${row.proses || '-'}</td>
                <td style="text-align:center; color:var(--neon-blue);">${row.actual || '-'}</td>
                <td style="text-align:center;">${row.durasi || '-'}</td>
            </tr>`;
        }).join('');
    } else {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center" style="padding:2rem;color:var(--text-dim);">Belum ada data laporan produksi.</td></tr>';
    }
};

// ==========================
// PENGUMUMAN API (Admin)
// ==========================
window.openPengumumanModal = function() {
    document.getElementById('modalPengumuman').classList.remove('hidden');
};

window.closePengumumanModal = function() {
    document.getElementById('modalPengumuman').classList.add('hidden');
};

window.setPengumuman = async function() {
    const text = document.getElementById('adminPengumuman').value.trim();
    if (!text) { showToast('error', 'Teks pengumuman tidak boleh kosong'); return; }
    
    showToast('info', 'Mengirim pengumuman...');
    try {
        const payload = { action: 'addPengumuman', text: text, token: window.adminToken };
        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
        const resJson = await res.json();

        if (resJson.status === 'success') {
            showToast('success', 'Pengumuman berhasil diset!');
            closePengumumanModal();
            loadAllDataFallback(); // Refresh marquee on this client
        } else { throw new Error(resJson.message); }
    } catch (e) { showToast('error', 'Gagal set: ' + e); }
};

window.stopPengumuman = async function() {
    showToast('info', 'Menghapus pengumuman...');
    try {
        const payload = { action: 'deletePengumuman', token: window.adminToken };
        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
        const resJson = await res.json();

        if (resJson.status === 'success') {
            showToast('success', 'Pengumuman dihapus!');
            document.getElementById('adminPengumuman').value = '';
            closePengumumanModal();
            loadAllDataFallback();
        } else { throw new Error(resJson.message); }
    } catch (e) { showToast('error', 'Gagal hapus: ' + e); }
};

async function validateSubmitForm(e) {
    e.preventDefault();
    const btn = document.getElementById('submitReqBtn');

    // Code Length Validation (Blueprint Requirement -> 6 sampai 15 Karakter Numerik/Huruf)
    const codeInput = document.getElementById('reqCode').value;
    if (codeInput.length < 6 || codeInput.length > 15) {
        showToast('error', 'Style Code harus antara 6 hingga 15 karakter.');
        return;
    }

    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-spinner-gap spin"></i> Mengirim Laporan...';

    // Merangkum Data Ekspor (Gabungan Form Brand dan Style Number)
    const type = document.getElementById('reqType').value;
    const brand = document.getElementById('reqBrand').value;
    const gabunganStyle = brand + " - " + codeInput;

    const payload = {
        action: 'addRequest',
        type: type,
        code: gabunganStyle,
        line: document.getElementById('reqLine').value,
        proses: document.getElementById('reqProcess').value,
        tanggal: document.getElementById('reqTanggal').value || '',
        keterangan: document.getElementById('reqKeterangan').value || ''
    };

    try {
        const response = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(payload)
        });

        const res = await response.json();

        if (res.status === 'success') {
            const successMsg = type === 'IE' ? 'template segera di siapkan' : 'request terkirim';
            showToast('success', successMsg);
        } else {
            showToast('error', res.message || 'Gagal mengirim form');
        }
    } catch (e) {
        // Trik khusus untuk kegagalan block CORS Google Apps Script kadang mengembalikan response type opaque, namun data masuk
        const successMsg = type === 'IE' ? 'template segera di siapkan' : 'request terkirim';
        showToast('success', successMsg);
    } finally {
        btn.disabled = false;
        btn.innerHTML = 'Submit';
        document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
        document.getElementById('requestForm').reset();

        // Auto refresh agar log segera muncul secara live
        loadAllDataFallback();
    }
}

function openLineDataModal(lineNum) {
    document.getElementById('lineModalTitle').innerText = `Factory ${lineNum <= 12 ? 'A' : 'B'} - Line ${lineNum}`;
    const tbody = document.querySelector('#tblLineData tbody');

    // Ambil data riil dari data Cloud
    const lineKey = 'Line ' + lineNum;
    let lineData = [];
    if (lineNum <= 12) {
        lineData = allData.factoryA ? allData.factoryA[lineKey] : [];
    } else {
        lineData = allData.factoryB ? allData.factoryB[lineKey] : [];
    }

    // Render
    const validData = (lineData || []).filter(r => (r.style || '').trim() !== '' || (r.proses || '').trim() !== '');
    if (validData.length > 0) {
        let lastLine = null;
        let lastStyle = null;

        tbody.innerHTML = validData.map(row => {
            const curLine = (row.line || '').toString().trim();
            const curStyle = (row.style || '').toString().trim();
            let showLine = curLine;
            let showStyle = curStyle;

            if (curLine === lastLine) {
                showLine = "";
                if (curStyle === lastStyle) {
                    showStyle = "";
                } else {
                    lastStyle = curStyle;
                }
            } else {
                lastLine = curLine;
                lastStyle = curStyle;
            }

            const layoutBtn = showStyle ? `
                <td><button class="action-btn purple-outline" style="padding: 0.3rem 0.5rem; font-size: 0.7rem;" onclick="sendLayoutRequest('${row.line || '-'}', '${row.style || '-'}')"><i class="ph ph-paper-plane-tilt"></i> Layout</button></td>
            ` : '<td></td>';

            let adminCol = '';
            if (window.isAdminMode) {
                adminCol = `<td class="admin-only" style="text-align:center; white-space:nowrap;">
                    <button class="action-btn" style="padding:0.2rem 0.4rem; font-size:0.7rem; background:rgba(0,191,255,0.1); border:1px solid var(--neon-cyan); color:var(--neon-cyan);" onclick="adminCompleteLayout('${escapeHtml(row.line||'')}', '${escapeHtml(row.style||'')}')" title="Selesaikan Layout"><i class="ph ph-check-square-offset"></i></button>
                    <button class="action-btn" style="padding:0.2rem 0.4rem; font-size:0.7rem; background:rgba(0,255,136,0.1); border:1px solid var(--neon-green); color:var(--neon-green);" onclick="adminEditRow('Factory ${lineNum <= 12 ? 'A' : 'B'}', ${row.rowIndex})" title="Edit"><i class="ph ph-pencil-simple"></i></button>
                    <button class="action-btn" style="padding:0.2rem 0.4rem; font-size:0.7rem; background:rgba(255,0,85,0.1); border:1px solid var(--neon-red); color:var(--neon-red);" onclick="deleteSingleRow('Factory ${lineNum <= 12 ? 'A' : 'B'}', ${row.rowIndex}, this)" title="Delete"><i class="ph ph-trash"></i></button>
                    <button class="action-btn" style="padding:0.2rem 0.4rem; font-size:0.7rem; background:rgba(255,184,0,0.1); border:1px solid #f59e0b; color:#f59e0b;" onclick="insertAjaib('Factory ${lineNum <= 12 ? 'A' : 'B'}', ${row.rowIndex}, '${escapeHtml(row.line||'')}')" title="Sisip Baris"><i class="ph ph-plus"></i></button>
                </td>`;
            }

            return `<tr>
                <td style="font-weight:bold; color:var(--neon-cyan);">${showLine}</td>
                <td style="font-weight:500;">${showStyle}</td>
                <td>${row.proses || '-'}</td>
                <td><span class="badge ${getStatusBadge(row.status || '')}">${row.status || '-'}</span></td>
                ${layoutBtn}
                ${adminCol}
            </tr>`;
        }).join('');
    } else {
        tbody.innerHTML = `<tr><td colspan="4" class="text-center" style="color:var(--text-dim); padding: 2rem;">Tidak ada data Line/Style/Proses untuk ${lineKey}</td></tr>`;
    }

    document.getElementById('lineDataModal').classList.remove('hidden');
}

window.sendLayoutRequest = function (line, style) {
    showToast('info', 'Template sedang diproses untuk baris/line: ' + line + '. Mohon tunggu respons server...');

    document.getElementById('lineDataModal').classList.add('hidden');

    let btn = document.getElementById('layoutMarqueeContainer');
    if (btn) btn.classList.remove('hidden'); // Memastikan marquee minimal muncul area kotak kosongnya

    try {
        const payload = { action: 'requestLayout', line: line, style: style };
        fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(payload)
        }).then(r => r.json()).then(res => {
            if (res.status === 'success') {
                forceRefreshData(); // Langsung paksa refresh layar tanpa nunggu 5 detik
            }
        });
    } catch (e) {
        console.error('Fetch post layout gagal');
    }
};

window.renderMarquee = function () {
    let requestsText = [];
    let isAdaIsi = false; // Deteksi apakah ada request masuk sama sekali
    const ONE_DAY_MS = 24 * 60 * 60 * 1000;
    const now = Date.now();

    // [1. CLOUD SYNC] Teks Request Layout (Merah)
    if (allData && allData.layoutRequests && allData.layoutRequests.length > 0) {
        const pending = allData.layoutRequests.filter(r => {
            // Pastikan bersih dari spasi dan huruf besar
            const s = String(r.status_penyelesaian || '').trim().toLowerCase();
            if (s !== 'pending') return false;

            // Pastikan umurnya belum lewat 24 Jam
            const reqDate = new Date(r.timestamp).getTime();
            if (!isNaN(reqDate) && (now - reqDate) > ONE_DAY_MS) {
                return false;
            }
            return true;
        });

        pending.forEach(r => {
            isAdaIsi = true;
            let cleanLine = String(r.line).replace(/line/i, '').trim();
            if (!cleanLine || cleanLine === '-') cleanLine = 'Tidak Diketahui';
            requestsText.push(`<span style="color: red;">layout line ${cleanLine} (${r.style || ''})</span>`);
        });
    }

    // [2. CLOUD SYNC] Teks Pengumuman Admin (Biru)
    if (allData && allData.pengumuman && allData.pengumuman.length > 0) {
        const aktif = allData.pengumuman.filter(p => {
            const statAdmin = String(p.status || '').trim().toLowerCase();
            return statAdmin !== 'nonaktif';
        });

        aktif.forEach(p => {
            isAdaIsi = true;
            requestsText.push(`<span style="color: #00bfff;">Info: ${p.pesan || ''}</span>`);
        });
    }

    const marqueeContainer = document.getElementById('layoutMarqueeContainer');
    const marqueeText = document.getElementById('layoutMarqueeText'); // elemen <marquee>

    if (marqueeContainer && marqueeText) {
        if (isAdaIsi) {
            marqueeContainer.classList.remove('hidden');
            marqueeText.innerHTML = requestsText.join("   |   ");
        } else {
            marqueeContainer.classList.hidden = true;
            marqueeText.innerHTML = "";
        }
    }
};

// OCR & Scanner logic extended
function initOCRExpanded() {
    const startScanBtn = document.getElementById('toggleScannerBtn');
    const scanSec = document.getElementById('scannerSection');
    const captureBtn = document.getElementById('captureBtn');
    const video = document.getElementById('cameraFeed');
    const canvas = document.getElementById('snapshotCanvas');
    let stream = null;

    // --- TAB SWITCHING LOGIC ---
    const tabScan = document.getElementById('tabScan');
    const tabManual = document.getElementById('tabManual');
    const areaScan = document.getElementById('areaScan');
    const areaManual = document.getElementById('areaManual');

    if (tabScan && tabManual) {
        tabScan.addEventListener('click', () => {
            tabScan.style.color = 'var(--neon-cyan)';
            tabScan.style.borderBottom = '2px solid var(--neon-cyan)';
            tabManual.style.color = 'var(--text-dim)';
            tabManual.style.borderBottom = '2px solid transparent';

            areaScan.classList.remove('hidden');
            areaManual.classList.add('hidden');
        });

        tabManual.addEventListener('click', () => {
            tabManual.style.color = 'var(--neon-cyan)';
            tabManual.style.borderBottom = '2px solid var(--neon-cyan)';
            tabScan.style.color = 'var(--text-dim)';
            tabScan.style.borderBottom = '2px solid transparent';

            areaManual.classList.remove('hidden');
            areaScan.classList.add('hidden');

            // Hemat baterai: matikan lensa jika lari ke manual
            if (stream) { stream.getTracks().forEach(t => t.stop()); }
            scanSec.classList.add('hidden');
        });
    }

    if (!startScanBtn) return;

    startScanBtn.addEventListener('click', async () => {
        scanSec.classList.remove('hidden');
        try {
            stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
            video.srcObject = stream;
        } catch (err) {
            alert('Akses Kamera ditolak.');
            scanSec.classList.add('hidden');
        }
    });

    captureBtn.addEventListener('click', async () => {
        if (!video.videoWidth) return;

        // --- 1. CROPPING CANVAS KE AREA KOTAK TARGET (`.scan-frame`) ---
        // Karena CSS memakai object-fit: cover, kita harus kalkulasi ukuran rasio
        // agar gambar yang dikirim ke Mesin OCR PERSIS seperti apa yang dilihat user di dalam frame putus-putus
        const cw = video.clientWidth;
        const ch = video.clientHeight;
        const vw = video.videoWidth;
        const vh = video.videoHeight;

        const scale = Math.max(cw / vw, ch / vh);
        const drawW = vw * scale;
        const drawH = vh * scale;
        const drawX = (cw - drawW) / 2;
        const drawY = (ch - drawH) / 2;

        // Canvas virtual seukuran kotak video DOM
        const tempCanvas = document.createElement('canvas');
        tempCanvas.width = cw; tempCanvas.height = ch;
        tempCanvas.getContext('2d').drawImage(video, drawX, drawY, drawW, drawH);

        // Target bingkai pink yang asli
        const frame = document.querySelector('.scan-frame');
        const fw = frame.clientWidth;
        const fh = frame.clientHeight;
        const fx = (cw - fw) / 2; // tengah X
        const fy = (ch - fh) / 2; // tengah Y

        // Terapkan crop akhir (Canvas ini yang dikirim ke Tesseract)
        canvas.width = fw;
        canvas.height = fh;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(tempCanvas, fx, fy, fw, fh, 0, 0, fw, fh);
        // -------------------------------------------------------------

        // [Opsional] Preprocessing: Bantu kontras sedikit jika gelap (Grayscale only)
        let imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        let pixels = imgData.data;
        for (let i = 0; i < pixels.length; i += 4) {
            let gray = 0.299 * pixels[i] + 0.587 * pixels[i + 1] + 0.114 * pixels[i + 2];
            pixels[i] = gray; pixels[i + 1] = gray; pixels[i + 2] = gray;
        }
        ctx.putImageData(imgData, 0, 0);

        document.getElementById('ocrLoading').classList.remove('hidden');
        captureBtn.innerText = 'Memproses OCR...';

        try {
            const { data: { text } } = await Tesseract.recognize(canvas.toDataURL('image/png'), 'eng+ind');

            // --- 2. PEMBERSIHAN MURNI (HANYA FOKUS KE STYLE CODE) ---
            // Sesuai perombakan: Kita tidak lagi memaksakan filter Brand. 
            // Kita percaya pada frame kotak kecil user yang hanya akan disorot ke susunan angka/huruf Style Code.
            let rawText = text.replace(/[\r\n]+/g, ' ');

            // Ekstraksi murni susunan alfa numerik minimal 4 huruf (khas style code)
            let tokens = rawText.match(/[A-Z0-9\-\.]{4,}/gi);
            let finalCode = "";

            if (tokens) {
                // Membuang kata baku noise yang kerap terbaca di area sempit
                tokens = tokens.filter(t => !/^(BUYER|STYLE|DATE|ITEM|FABRIC|PRICE|SIZE|FOB)$/i.test(t));
                finalCode = tokens.join(' ').toUpperCase();
            } else {
                finalCode = rawText.substring(0, 40).replace(/[^A-Z0-9\-\s]/gi, '').trim().toUpperCase();
            }

            // --- 3. AUTO-CORRECT & PENGINTAIAN TERBALIK (REVERSE LOOKUP) ---
            let bestDbMatch = null;
            let minDbDist = 999;
            let foundBrand = "";

            const getEditDistance = (a, b) => {
                if (a.length === 0) return b.length;
                if (b.length === 0) return a.length;
                let m = [];
                for (let i = 0; i <= b.length; i++) m[i] = [i];
                for (let j = 0; j <= a.length; j++) m[0][j] = j;
                for (let i = 1; i <= b.length; i++) {
                    for (let j = 1; j <= a.length; j++) {
                        if (b.charAt(i - 1) === a.charAt(j - 1)) m[i][j] = m[i - 1][j - 1];
                        else m[i][j] = Math.min(m[i - 1][j - 1] + 1, Math.min(m[i][j - 1] + 1, m[i - 1][j] + 1));
                    }
                }
                return m[b.length][a.length];
            };

            // Kumpulkan Data Catalog sebagai Objek (Rantai Nomor Style dan Nama Brand/Buyer aslinya)
            let catalog = [];
            const addReq = (code, buyer) => {
                if (code) catalog.push({ code: code.toString().toUpperCase().trim(), buyer: buyer || '' });
            };

            if (allData.templateInLine) allData.templateInLine.forEach(r => { addReq(r.style, r.buyer); addReq(r.code, r.buyer); });
            if (allData.factoryA) { Object.values(allData.factoryA).forEach(arr => { if (Array.isArray(arr)) arr.forEach(r => { addReq(r.style, r.buyer); }); }); }
            if (allData.factoryB) { Object.values(allData.factoryB).forEach(arr => { if (Array.isArray(arr)) arr.forEach(r => { addReq(r.style, r.buyer); }); }); }
            if (allData.inventory) { allData.inventory.forEach(r => { addReq(r.code, r.buyer); }); }

            // Lacak kemiripan jarak Edit (Levenshtein) di seluruh memori pabrik
            catalog.forEach(item => {
                let dbCode = item.code;
                let dist = getEditDistance(finalCode, dbCode);

                // Tambahan proteksi cerdas: OCR sering merekam dua elemen pisah (misal: "A123 B456")
                // Jika substringnya saja menabrak langsung dengan tepat waktu, kita angkat!
                if (finalCode.includes(dbCode) || dbCode.includes(finalCode)) {
                    if (finalCode.length > 5 && dbCode.length > 5) { // cegah kecocokan ngawur 1 digit
                        dist = 0;
                    }
                }

                if (dist < minDbDist) {
                    minDbDist = dist;
                    bestDbMatch = dbCode;
                    if (item.buyer) foundBrand = item.buyer;  // <<< MENGAMBIL ALIH IDENTITAS BRAND
                }
            });

            let autoSearchTriggered = false;

            if (minDbDist <= 3 && minDbDist > 0 && bestDbMatch && finalCode.length >= 4) {
                showToast('success', `Typo Diperbaiki: ${finalCode} 👉 ${bestDbMatch}`);
                finalCode = bestDbMatch;
                autoSearchTriggered = true;
            } else if (minDbDist === 0 && finalCode.length >= 4) {
                showToast('success', `Data Valid 100%: ${finalCode}`);
                finalCode = bestDbMatch || finalCode;
                autoSearchTriggered = true;
            } else {
                showToast('info', `Scan terdeteksi: ${finalCode}`);
            }

            // --- 4. UPDATE USER INTERFACE (FORM PENGISIAN) ---
            const styleInput = document.getElementById('smartSearchInputModal');
            const brandSelect = document.getElementById('searchBrand');

            styleInput.value = finalCode;

            // Jika barang tsb ditemukan di dashboard, pilihkan otomatis Box Dropdown-nya!
            if (foundBrand && (minDbDist <= 3 || minDbDist === 0)) {
                let opts = Array.from(brandSelect.options).map(o => o.value).filter(v => v);
                let selectedVal = "";
                for (let opt of opts) {
                    // Logika kasar (misal: "Ann Taylor" cocok ke "ANN TAYLOR")
                    if (foundBrand.toUpperCase().includes(opt.toUpperCase())) {
                        selectedVal = opt; break;
                    }
                    if (opt.toUpperCase().includes('JIL') && foundBrand.toUpperCase().includes('JILL')) {
                        selectedVal = opt; break;
                    }
                }
                brandSelect.value = selectedVal; // Berubah jadi Ann Taylor di Form
            } else {
                brandSelect.value = "";
            }

            if (stream) stream.getTracks().forEach(t => t.stop());
            scanSec.classList.add('hidden');

            // JALANKAN OTOMATIS JIKA YAKIN
            if (autoSearchTriggered) {
                // Beritahu user kita pindah form
                if (tabManual) tabManual.click();

                let btn = document.getElementById('execSearchBtn');
                btn.innerHTML = '<i class="ph ph-spinner-gap spin"></i> Verifikasi Pola...';
                // 600ms delay agar user sempat melihat keajaiban layarnya berubah sebelum berpindah tab hasil
                setTimeout(() => { btn.innerHTML = 'Cari Style'; btn.click(); }, 600);
            }
            // ------------------------------------
        } catch (e) {
            console.error("OCR Error:", e);
            showToast('error', 'Gagal memindai teks.');
        } finally {
            document.getElementById('ocrLoading').classList.add('hidden');
            captureBtn.innerText = 'Proses Scan';
        }
    });

    // Execute Smart Search Code (Live Data Scanning)
    document.getElementById('execSearchBtn').addEventListener('click', () => {
        const query = document.getElementById('smartSearchInputModal').value.toLowerCase();
        if (query.length < 5) { showToast('error', 'Masukkan info > 5 karakter.'); return; }

        // Membaca opsi form sebagai fallback jika database sheet tidak menyediakan Brand
        const brandSelect = document.getElementById('searchBrand');
        let fallbackBuyer = brandSelect.options[brandSelect.selectedIndex].text;
        if (fallbackBuyer === 'Semua Brand' || !fallbackBuyer) fallbackBuyer = '-';

        // Pindah layar ke tab View-Search-Results Full Page
        document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
        document.querySelectorAll('.tab-view').forEach(el => el.classList.add('hidden'));
        document.getElementById('view-search-results').classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });

        const resList = document.getElementById('fullPageResultList');
        resList.innerHTML = '<div style="text-align:center; padding:3rem;"><i class="ph ph-spinner-gap spin" style="font-size:3rem; color:var(--neon-cyan); margin-bottom:1rem;"></i><br>Sinkronisasi Matriks Gudang...</div>';

        setTimeout(() => {
            let foundHTML = '';

            // Desain Card-like untuk tiap hasil
            const cardStyle = "background:rgba(255,255,255,0.02); border:1px solid rgba(0,191,255,0.2); border-radius:6px; padding:1.2rem; margin-bottom:1rem; font-size:0.95rem; line-height:1.6;";
            const labelStyle = "color:var(--text-dim); display:inline-block; width:70px;";

            // 1. Pencarian di Template
            if (allData.templateInLine) {
                const f = allData.templateInLine.find(r => (r.style || '').toLowerCase().includes(query) || (r.code || '').toLowerCase().includes(query));
                if (f) {
                    foundHTML += `<div style="${cardStyle}">
                        <div style="color:var(--neon-green); font-size:1.1rem; font-weight:bold; margin-bottom:0.8rem; border-bottom:1px solid rgba(0,255,136,0.2); padding-bottom:0.5rem;"><i class="ph ph-grid-four"></i> Template In-Line [Line ${f.line || '-'}]</div>
                        <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Buyer:</span> <span style="color:#fff;">${f.buyer || fallbackBuyer}</span></div>
                        <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Style:</span> <span style="color:#fff; font-weight:bold;">${f.style || f.code || '-'}</span></div>
                        <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Proses:</span> <span style="color:#fff;">${f.proses || '-'}</span></div>
                        <div style="margin-top:0.8rem;"><span class="badge ${getStatusBadge(f.status)}" style="font-size:0.9rem; padding:0.4rem 0.8rem;">${f.status || 'Unknown'}</span></div>
                    </div>`;
                }
            }

            // 2. Helper Pencarian di Factory Objects
            const findFac = (facObj, title, icon, color) => {
                if (!facObj) return;
                for (const [line, arr] of Object.entries(facObj)) {
                    if (!Array.isArray(arr)) continue;
                    const hit = arr.find(r => (r.style || '').toLowerCase().includes(query));
                    if (hit) {
                        foundHTML += `<div style="${cardStyle}">
                            <div style="color:var(${color}); font-size:1.1rem; font-weight:bold; margin-bottom:0.8rem; border-bottom:1px solid rgba(0,191,255,0.2); padding-bottom:0.5rem;"><i class="ph ${icon}"></i> ${title} [${line}]</div>
                            <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Buyer:</span> <span style="color:#fff;">${hit.buyer || fallbackBuyer}</span></div>
                            <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Style:</span> <span style="color:#fff; font-weight:bold;">${hit.style || '-'}</span></div>
                            <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Proses:</span> <span style="color:#fff;">${hit.proses || '-'}</span></div>
                            <div style="margin-top:0.8rem;"><span class="badge ${getStatusBadge(hit.status)}" style="font-size:0.9rem; padding:0.4rem 0.8rem;">${hit.status || 'Unknown'}</span></div>
                        </div>`;
                    }
                }
            };
            findFac(allData.factoryA, 'Factory A', 'ph-factory', '--neon-cyan');
            findFac(allData.factoryB, 'Factory B', 'ph-factory', '--neon-purple');

            // 3. Pencarian di Inventory
            if (allData.inventory) {
                const f = allData.inventory.find(r => (r.code || '').toLowerCase().includes(query) || (r.buyer || '').toLowerCase().includes(query));
                if (f) {
                    foundHTML += `<div style="${cardStyle}">
                        <div style="color:var(--neon-orange); font-size:1.1rem; font-weight:bold; margin-bottom:0.8rem; border-bottom:1px solid rgba(249,115,22,0.2); padding-bottom:0.5rem;"><i class="ph ph-package"></i> Data Gudang (Inventory)</div>
                        <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Tanggal:</span> <span style="color:#fff;">${f.tanggal || '-'}</span></div>
                        <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Buyer:</span> <span style="color:#fff;">${f.buyer || fallbackBuyer}</span></div>
                        <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Style:</span> <span style="color:#fff; font-weight:bold;">${f.code || '-'}</span></div>
                        <div style="margin-bottom:0.3rem;"><span style="${labelStyle}">Qty:</span> <span style="color:#fff;">${f.qty || '-'} Pcs</span> <span style="${labelStyle}; width:auto; margin-left:1rem;">Size:</span> <span style="color:#fff;">${f.size || '-'}</span></div>
                        <div style="margin-bottom:0.6rem; margin-top:0.8rem;"><span style="${labelStyle}; width:auto; margin-right:0.5rem;">Lokasi Fisik Barang:</span> ${getPositionBadge(f.code) || '-'}</div>
                        <div style="margin-top:0.3rem;"><span class="badge ${getStatusBadge(f.status || '')}" style="font-size:0.9rem; padding:0.4rem 0.8rem;">${f.status || 'Unknown'}</span></div>
                    </div>`;
                }
            }

            resList.innerHTML = foundHTML || '<div style="color:var(--text-dim); text-align:center; padding:3rem;"><i class="ph ph-warning" style="font-size:3rem; margin-bottom:1rem;"></i><br>Oops, obyek tidak ditemukan di riwayat manapun.</div>';
        }, 500);
    });
}

// ==========================
// NOTIFICATIONS SYSTEM (TOASTS)
// ==========================
function initNotifications() {
    window.showToast = function (type, message) {
        const container = document.getElementById('toastContainer');
        const toast = document.createElement('div');

        // Penentuan Kelas Utama Premium
        let typeClass = 'toast-info';
        if (type === 'success') typeClass = 'toast-success';
        if (type === 'error') typeClass = 'toast-error';
        toast.className = `toast ${typeClass}`;

        // Penentuan Ikon dengan Animasi jika perlu
        let iconClass = 'ph-info';
        let iconColor = 'var(--neon-cyan)';
        if (type === 'success') { iconClass = 'ph-check-circle'; iconColor = 'var(--neon-green)'; }
        if (type === 'error') { iconClass = 'ph-warning-circle'; iconColor = 'var(--neon-red)'; }

        // Jika Notif Welcome
        if (message && message.includes('Welcome')) {
            iconClass = 'ph-plugs-connected blink';
            iconColor = 'var(--neon-cyan)';
        }

        toast.innerHTML = `<i class="ph ${iconClass}" style="font-size:1.5rem; color:${iconColor};"></i> ${message || '(pesan kosong)'}`;
        container.appendChild(toast);

        // Remove after 5s untuk sapaan
        setTimeout(() => {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, 5000);
    }
}

// ==========================
// THEME & PARTICLES SYSTEM
// ==========================
function initTheme() {
    const savedMode = localStorage.getItem('theme-mode');
    const isLight = savedMode === 'light';

    if (isLight) {
        document.body.classList.add('light-mode');
        document.querySelector('#themeToggleBtn i').className = 'ph ph-sun-dim';
    } else {
        document.body.classList.remove('light-mode');
        document.querySelector('#themeToggleBtn i').className = 'ph ph-moon-stars';
    }

    // Init particles after a tiny delay so DOM is ready
    setTimeout(() => initParticles(isLight ? 'light' : 'dark'), 100);
    document.getElementById('themeToggleBtn')?.addEventListener('click', toggleTheme);
}

function initDesktopMode() {
    const savedDesktop = localStorage.getItem('desktop-mode');
    if (savedDesktop === 'on') {
        document.body.classList.add('desktop-mode');
        const icon = document.querySelector('#desktopToggleBtn i');
        if (icon) icon.className = 'ph ph-laptop';
    }
}

function toggleDesktopMode() {
    document.body.classList.toggle('desktop-mode');
    const isOn = document.body.classList.contains('desktop-mode');
    localStorage.setItem('desktop-mode', isOn ? 'on' : 'off');
    
    const icon = document.querySelector('#desktopToggleBtn i');
    if (icon) {
        icon.className = isOn ? 'ph ph-laptop' : 'ph ph-monitor';
    }
    
    showToast('info', isOn ? 'Desktop Mode Active' : 'Mobile Mode Active');
    
    // Refresh chart jika diperlukan agar ukurannya pas
    if (typeof updateCharts === 'function') updateCharts();
}

function toggleTheme() {
    document.body.classList.toggle('light-mode');
    const isLight = document.body.classList.contains('light-mode');

    document.querySelector('#themeToggleBtn i').className = isLight ? 'ph ph-sun-dim' : 'ph ph-moon-stars';
    localStorage.setItem('theme-mode', isLight ? 'light' : 'dark');

    // Reload particles
    if (window.pJSDom && window.pJSDom.length > 0) {
        window.pJSDom[0].pJS.fn.vendors.destroypJS();
        window.pJSDom = [];
    }
    initParticles(isLight ? 'light' : 'dark');
}

function initParticles(mode) {
    // Pastikan partikel JS terpasang
    if (typeof particlesJS === 'undefined') return;

    if (mode === 'dark') {
        particlesJS("particles-js", {
            particles: {
                number: { value: 50, density: { enable: true, value_area: 800 } },
                color: { value: ["#00ff88", "#00bfff", "#d946ef"] },
                shape: { type: "circle" },
                opacity: { value: 0.5, random: true },
                size: { value: 3, random: true },
                line_linked: { enable: true, distance: 150, color: "#3c405a", opacity: 0.4, width: 1 },
                move: { enable: true, speed: 2, direction: "none", random: true, out_mode: "out" }
            },
            interactivity: {
                detect_on: "canvas",
                events: {
                    onhover: { enable: true, mode: "repulse" },
                    onclick: { enable: true, mode: "push" },
                    resize: true
                },
                modes: {
                    repulse: { distance: 200, duration: 0.4 },
                    push: { particles_nb: 4 }
                }
            },
            retina_detect: true
        });
    } else {
        particlesJS("particles-js", {
            particles: {
                number: { value: 30, density: { enable: true, value_area: 800 } },
                color: { value: ["#0d9488", "#0369a1"] },
                shape: { type: "circle" },
                opacity: { value: 0.2, random: true },
                size: { value: 2, random: true },
                line_linked: { enable: true, distance: 150, color: "#cbd5e1", opacity: 0.4, width: 1 },
                move: { enable: true, speed: 1 }
            },
            interactivity: {
                detect_on: "canvas",
                events: {
                    onhover: { enable: true, mode: "bubble" },
                    resize: true
                }
            },
            retina_detect: true
        });
    }
}

window.submitIEDataBatchPublic = async function () {
    if (!window.ieTokenPublic) { return showToast('error', 'Sesi Berakhir Coba Login Lagi'); }

    const rows = [];
    const smvInputs = document.querySelectorAll('.ie-smv');

    smvInputs.forEach(smvInp => {
        const rIdx = smvInp.dataset.rowindex;
        const actualInp = document.querySelector(`.ie-actual[data-rowindex="${rIdx}"]`);

        const smvVal = smvInp.value;
        const actualVal = actualInp ? actualInp.value : '';

        const styleVal = document.getElementById('ieRowStyle_' + rIdx).value;
        const lineVal = document.getElementById('ieRowLine_' + rIdx).value;
        const prosesVal = document.getElementById('ieRowProses_' + rIdx).value;

        if (smvVal !== '' || actualVal !== '') {
            rows.push({
                rowIndex: rIdx,
                line: lineVal,
                style: styleVal,
                proses: prosesVal,
                smv: smvVal,
                actual: actualVal
            });
        }
    });

    if (rows.length === 0) { return showToast('error', 'Tidak ada data SMV/Actual untuk disimpan.'); }

    const btn = document.getElementById('btnSubmitIEBatchPublic');
    btn.innerHTML = '<i class="ph ph-spinner spin"></i> Menyimpan Secara Massal...';
    btn.disabled = true;

    try {
        const payload = {
            action: 'updateIEDataBatch',
            rows: rows,
            token: window.ieTokenPublic
        };
        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
        const data = await res.json();

        if (data.status === 'success') {
            showToast('success', data.message);
            forceRefreshData(new Event('click'));
        } else {
            showToast('error', data.message);
        }
    } catch (e) {
        showToast('error', 'Error: ' + e.message);
    } finally {
        btn.innerHTML = '<i class="ph ph-floppy-disk"></i> Simpan Seluruh Perubahan';
        btn.disabled = false;
    }
};

function escapeHtml(unsafe) { if (!unsafe) return ''; return unsafe.toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;"); }


// ==========================
// ARCHIVE TEMPLATE VIEW
// ==========================

// ==========================
// ARCHIVE - TAB SWITCHER
// ==========================
window.switchArchiveTab = function (tab) {
    const prodPanel = document.getElementById('panelArchiveProd');
    const tmplPanel = document.getElementById('panelArchiveTmpl');
    const tabProd = document.getElementById('tabArchiveProd');
    const tabTmpl = document.getElementById('tabArchiveTmpl');

    if (tab === 'production') {
        prodPanel.style.display = '';
        tmplPanel.style.display = 'none';
        tabProd.style.borderBottom = '3px solid #10b981';
        tabProd.style.color = '#10b981';
        tabProd.style.background = 'rgba(16,185,129,0.15)';
        tabTmpl.style.borderBottom = '3px solid transparent';
        tabTmpl.style.color = 'var(--text-dim)';
        tabTmpl.style.background = 'transparent';
        renderArchiveProductionTable();
    } else {
        prodPanel.style.display = 'none';
        tmplPanel.style.display = '';
        tabTmpl.style.borderBottom = '3px solid #10b981';
        tabTmpl.style.color = '#10b981';
        tabTmpl.style.background = 'rgba(16,185,129,0.15)';
        tabProd.style.borderBottom = '3px solid transparent';
        tabProd.style.color = 'var(--text-dim)';
        tabProd.style.background = 'transparent';
        renderArchiveTemplateTable();
    }
};

// ==========================
// ARCHIVE PRODUCTION TABLE
// ==========================
window.renderArchiveProductionTable = function () {
    const tbody = document.querySelector('#tblArchiveProduction tbody');
    if (!tbody) return;

    const filterSel = document.getElementById('archiveLineFilter');
    const filterDate = document.getElementById('archiveDateFilter');
    const selectedLine = filterSel ? filterSel.value : '';
    const selectedDate = filterDate ? filterDate.value : ''; // Format yyyy-mm-dd

    const rows = (allData.archiveProduction || []).filter(r =>
        (r.line || '').trim() !== '' || (r.style || '').trim() !== ''
    );

    // Populate Line filter dropdown (hanya sekali)
    if (filterSel && filterSel.options.length <= 1) {
        const uniqueLines = [...new Set(rows.map(r => (r.line || '').trim()).filter(v => v))];
        uniqueLines.sort((a, b) => (parseInt(a.replace(/\D/g, '')) || 0) - (parseInt(b.replace(/\D/g, '')) || 0));
        uniqueLines.forEach(ln => {
            const opt = document.createElement('option');
            opt.value = ln;
            opt.textContent = /line/i.test(ln) ? ln : 'Line ' + ln;
            filterSel.appendChild(opt);
        });
    }

    let filtered = rows;
    if (selectedLine) {
        filtered = filtered.filter(r => String(r.line || '').trim().toLowerCase() === selectedLine.toLowerCase());
    }
    if (selectedDate) {
        // Bandingkan tanggal (format timestamp: dd-MMM-yyyy HH:mm:ss)
        filtered = filtered.filter(r => {
            if (!r.timestamp) return false;
            // Parse dd-MMM-yyyy
            const parts = r.timestamp.split(' ')[0].split('-');
            if (parts.length < 3) return false;
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            const day = parts[0].padStart(2, '0');
            const month = (months.indexOf(parts[1]) + 1).toString().padStart(2, '0');
            const year = parts[2];
            const rowDate = `${year}-${month}-${day}`;
            return rowDate === selectedDate;
        });
    }
    filtered.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    const vLink = (url, label, color, bg) => url
        ? `<button class="video-pill ${label === 'Before' ? 'video-pill-before' : 'video-pill-after'}" onclick="openVideoPlayer('${url.replace(/'/g, "\\'")}','${label}')" title="Putar ${label}"><i class="ph ph-play-circle"></i> ${label}</button>`
        : `<span style="color:var(--text-dim);font-size:0.7rem;">&#8212;</span>`;

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="11" class="text-center" style="color:var(--text-dim);">Belum ada data Archive Production.</td></tr>';
        return;
    }

    let lastLine = null;
    let lastStyle = null;

    tbody.innerHTML = filtered.map(row => {
        const curLine = (row.line || '').toString().trim();
        const curStyle = (row.style || '').toString().trim();
        let showLine = curLine;
        let showStyle = curStyle;

        if (curLine === lastLine) {
            showLine = "";
            if (curStyle === lastStyle) {
                showStyle = "";
            } else {
                lastStyle = curStyle;
            }
        } else {
            lastLine = curLine;
            lastStyle = curStyle;
        }

        const sVal = parseFloat(String(row.saving || '0').replace(',', '.'));
        const sNum = isNaN(sVal) ? 0 : sVal;
        const sCol = sNum > 0 ? 'var(--neon-green)' : (sNum < 0 ? 'var(--neon-red)' : 'var(--text-dim)');
        const sStr = !isNaN(sVal) ? sVal.toFixed(1) : '-';
        const rStr = row.rate ? row.rate + '%' : '-';
        const tStr = row.timestamp ? formatDateString(row.timestamp) : '-';

        return '<tr>' +
            `<td style="font-weight:bold; color:var(--neon-cyan);">${showLine}</td>` +
            `<td style="font-weight:500;">${showStyle}</td>` +
            `<td>${row.proses || ''}</td>` +
            `<td><span class="badge ${getStatusBadge(row.status || '')}">${row.status || ''}</span></td>` +
            `<td style="text-align:center;color:var(--neon-blue);font-weight:500;">${row.smv || '-'}</td>` +
            `<td style="text-align:center;">${row.actual || '-'}</td>` +
            `<td style="text-align:center;font-weight:bold;color:${sCol};">${sStr}</td>` +
            `<td style="text-align:center;">${rStr}</td>` +
            `<td style="text-align:center;">${vLink(row.videoB, 'Before', '#10b981', 'rgba(16,185,129,0.15)')}</td>` +
            `<td style="text-align:center;">${vLink(row.videoA, 'After', '#f59e0b', 'rgba(245,158,11,0.15)')}</td>` +
            `<td style="font-size:0.65rem;color:var(--text-dim);">${tStr}</td>` +
            '</tr>';
    }).join('');
};

// ==========================
window.renderArchiveTemplateTable = function () {
    const tbody = document.querySelector('#tblArchiveTemplate tbody');
    if (!tbody) return;

    const rows = (allData.archiveTemplate || []).filter(r =>
        (r.line || '').trim() !== '' || (r.style || '').trim() !== ''
    );

    // Hitung Total untuk Kartu Ringkasan
    let tSmv = 0, tAct = 0, tSav = 0;
    rows.forEach(r => {
        tSmv += parseFloat(String(r.smv || '0').replace(',', '.')) || 0;
        tAct += parseFloat(String(r.actual || '0').replace(',', '.')) || 0;
        tSav += parseFloat(String(r.saving || '0').replace(',', '.')) || 0;
    });

    const elSmv = document.getElementById('tmplTotalSmv');
    const elAct = document.getElementById('tmplTotalActual');
    const elSav = document.getElementById('tmplTotalSaving');
    if (elSmv) elSmv.textContent = tSmv.toFixed(1);
    if (elAct) elAct.textContent = tAct.toFixed(1);
    if (elSav) elSav.textContent = tSav.toFixed(1);

    if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="text-center" style="color:var(--text-dim);">Belum ada data Archive Template.</td></tr>';
        return;
    }

    // Urutkan berdasarkan Style (A-Z)
    rows.sort((a, b) => (a.style || '').localeCompare(b.style || ''));

    const vLink = (url, label, color, bg) => url
        ? `<button class="video-pill ${label === 'Before' ? 'video-pill-before' : 'video-pill-after'}" onclick="openVideoPlayer('${url.replace(/'/g, "\\'")}','${label}')" title="Putar ${label}"><i class="ph ph-play-circle"></i> ${label}</button>`
        : `<span style="color:var(--text-dim);font-size:0.7rem;">&#8212;</span>`;

    let lastStyle = null;

    tbody.innerHTML = rows.map(row => {
        const curStyle = (row.style || '').toString().trim();
        let showStyle = curStyle;

        if (curStyle === lastStyle) {
            showStyle = "";
        } else {
            lastStyle = curStyle;
        }

        const sVal = parseFloat(String(row.saving || '0').replace(',', '.'));
        const sNum = isNaN(sVal) ? 0 : sVal;
        const sCol = sNum > 0 ? 'var(--neon-green)' : (sNum < 0 ? 'var(--neon-red)' : 'var(--text-dim)');
        const sStr = !isNaN(sVal) ? sVal.toFixed(1) : '-';
        const rStr = row.rate ? row.rate + '%' : '-';
        const tStr = row.timestamp ? formatDateString(row.timestamp) : '-';

        return '<tr>' +
            `<td style="font-weight:500;">${showStyle}</td>` +
            `<td>${row.proses || ''}</td>` +
            `<td><span class="badge ${getStatusBadge(row.status || '')}">${row.status || ''}</span></td>` +
            `<td style="text-align:center;color:var(--neon-blue);font-weight:500;">${row.smv || '-'}</td>` +
            `<td style="text-align:center;">${row.actual || '-'}</td>` +
            `<td style="text-align:center;font-weight:bold;color:${sCol};">${sStr}</td>` +
            `<td style="text-align:center;">${rStr}</td>` +
            `<td style="text-align:center;">${vLink(row.videoB, 'Before', '#10b981', 'rgba(16,185,129,0.15)')}</td>` +
            `<td style="text-align:center;">${vLink(row.videoA, 'After', '#f59e0b', 'rgba(245,158,11,0.15)')}</td>` +
            `<td style="text-align:center;font-size:0.65rem;color:var(--text-dim);">${tStr}</td>` +
            '</tr>';
    }).join('');
};

// Backward-compat alias agar kode lama yang masih panggil renderArchiveTable() tidak error
window.renderArchiveTable = window.renderArchiveProductionTable;


// Buka modal IE dari Archive Template (pre-fill semua field)
window.openIeModalPublicFromEditor = function (line, style, proses, smv, actual, videoB, videoA) {
    window.openIeModalPublicFromArchive(line, style, proses, smv, actual, videoB, videoA);
};

function formatDateString(dateStr) {
    if (!dateStr) return "-";
    try {
        const d = new Date(dateStr);
        const pad = (n) => n.toString().padStart(2, '0');
        const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        return `${pad(d.getDate())} ${months[d.getMonth()]} ${d.getHours()}:${pad(d.getMinutes())}`;
    } catch (e) { return dateStr; }
}

window.openIeModalPublicFromArchive = function (line, style, proses, smv, actual, videoB, videoA) {
    if (!window.ieTokenPublic) {
        openIeLoginModal();
        return;
    }
    // Cari rowIndex dari data in-memory
    const found = (allData.templateInLine || []).find(r =>
        String(r.line || '').trim().toLowerCase() === String(line).trim().toLowerCase() &&
        String(r.style || '').trim().toLowerCase() === String(style).trim().toLowerCase() &&
        String(r.proses || '').trim().toLowerCase() === String(proses).trim().toLowerCase()
    );
    document.getElementById('ieRowIndexPublic').value = found ? (found.rowIndex || '') : '';
    document.getElementById('ieStylePublic').value = style;
    document.getElementById('ieLinePublic').value = line;
    document.getElementById('ieProsesPublic').value = proses;

    // Konversi koma ke titik agar input type="number" bisa menampilkan nilai
    const smvClean = String(smv || '').replace(',', '.');
    const actualClean = String(actual || '').replace(',', '.');

    // Tampilkan info baris + data sebelumnya (jika ada)
    let descHtml = 'Line: <strong>' + line + '</strong><br>Style: <strong>' + style + '</strong><br>Proses: <strong>' + proses + '</strong>';
    
    // Tampilkan data sebelumnya agar IE tahu
    const hasPrevData = smvClean || actualClean || videoB || videoA;
    if (hasPrevData) {
        descHtml += '<div style="margin-top:0.6rem; padding:0.5rem 0.6rem; background:rgba(0,191,255,0.06); border:1px solid rgba(0,191,255,0.15); border-radius:6px; font-size:0.75rem;">';
        descHtml += '<div style="color:var(--neon-cyan); font-weight:600; margin-bottom:0.25rem;"><i class="ph ph-database"></i> Data Sebelumnya:</div>';
        if (smvClean) descHtml += '<div>SMV: <strong style="color:var(--neon-blue);">' + smvClean + '</strong></div>';
        if (actualClean) descHtml += '<div>Actual: <strong style="color:#f59e0b;">' + actualClean + '</strong></div>';
        if (videoB) descHtml += '<div>Video Before: <span style="color:#10b981;">✓ ada</span></div>';
        if (videoA) descHtml += '<div>Video After: <span style="color:#f59e0b;">✓ ada</span></div>';
        descHtml += '</div>';
    }
    document.getElementById('ieModalDescPublic').innerHTML = descHtml;

    document.getElementById('ieSmvPublic').value = smvClean || '';
    document.getElementById('ieActualPublic').value = actualClean || '';
    document.getElementById('ieVideoBPublic').value = videoB || '';
    document.getElementById('ieVideoAPublic').value = videoA || '';
    document.getElementById('modalIEInputPublic').classList.remove('hidden');
};

// submitIEDataPublic: update Sheet2 + simpan ke Archive
window.submitIEDataPublic = async function () {
    const rowIndex = document.getElementById('ieRowIndexPublic').value;
    const smv = document.getElementById('ieSmvPublic').value;
    const actual = document.getElementById('ieActualPublic').value;
    const videoB = document.getElementById('ieVideoBPublic').value.trim();
    const videoA = document.getElementById('ieVideoAPublic').value.trim();
    const line = document.getElementById('ieLinePublic').value;
    const style = document.getElementById('ieStylePublic').value;
    const proses = document.getElementById('ieProsesPublic').value;

    if (!smv || !actual) { showToast('error', 'SMV dan Actual wajib diisi!'); return; }

    const btn = document.getElementById('btnSubmitIEPublic');
    const oldText = btn.innerHTML;
    btn.innerHTML = '<i class="ph ph-spinner spin"></i> Prosses...';
    btn.disabled = true;

    // TUTUP MODAL SEGERA (Memberikan efek instan dan responsif)
    document.getElementById('modalIEInputPublic').classList.add('hidden');
    showToast('info', 'Sedang memproses penyimpanan data...');

    try {
        const payload = {
            action: 'submitIEForm',
            rowIndex: rowIndex,
            line: line, style: style, proses: proses,
            smv: smv, actual: actual,
            videoB: videoB, videoA: videoA,
            token: window.ieTokenPublic
        };

        const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
        const data = await res.json();

        if (data.status === 'success') {
            showToast('success', 'Berhasil! Data masuk ke Dashboard & Archive.');

            // Sesi tetap aktif selama di halaman ini (tidak dihapus)
            // agar user bisa lanjut edit baris lain tanpa login ulang.

            // Refresh data di layar
            await loadAllDataFallback();
            renderTables();
            renderIeEditorTable();
            renderArchiveProductionTable();
            renderArchiveTemplateTable();
            renderPerformanceStats();
        } else {
            showToast('error', 'Gagal: ' + data.message);
        }

    } catch (err) {
        console.error(err);
        showToast('error', 'Koneksi terganggu. Periksa Spreadsheet Anda.');
    } finally {
        btn.innerHTML = oldText;
        btn.disabled = false;

        // Bersihkan input untuk penggunaan berikutnya
        document.getElementById('ieVideoBPublic').value = '';
        document.getElementById('ieVideoAPublic').value = '';
    }
};

// ==========================
// 🎬 VIDEO PLAYER MODAL
// ==========================
const drivePreviewUrl = (url) => {
    if (!url) return '';
    const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    return m ? 'https://drive.google.com/file/d/' + m[1] + '/preview' : url;
};

window.openVideoPlayer = function (driveUrl, label) {
    const modal = document.getElementById('modalVideoPlayer');
    const frame = document.getElementById('videoPlayerFrame');
    const title = document.getElementById('videoPlayerTitle');
    if (!modal || !frame) return;

    const previewUrl = drivePreviewUrl(driveUrl);
    frame.src = previewUrl;
    title.innerHTML = `<i class="ph ph-film-strip"></i> ${label || 'Video'}`;
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
};

window.closeVideoPlayer = function () {
    const modal = document.getElementById('modalVideoPlayer');
    const frame = document.getElementById('videoPlayerFrame');
    if (modal) modal.classList.add('hidden');
    if (frame) frame.src = '';
    document.body.style.overflow = '';
};

// ESC key to close video player
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        const modal = document.getElementById('modalVideoPlayer');
        if (modal && !modal.classList.contains('hidden')) {
            closeVideoPlayer();
        }
    }
});

// ==========================
// 🎥 VIDEO UPLOAD TOOL
// ==========================
let _selectedVideoFile = null;
let _ffmpegLoaded = false;
let _ffmpegInstance = null;

window.onVideoFileSelected = function (input) {
    const file = input.files[0];
    const infoEl = document.getElementById('videoFileInfo');
    const btn = document.getElementById('btnStartUpload');

    if (!file) {
        _selectedVideoFile = null;
        infoEl.textContent = 'Belum ada file dipilih';
        btn.disabled = true;
        return;
    }

    const sizeMB = (file.size / (1024 * 1024)).toFixed(1);
    _selectedVideoFile = file;
    infoEl.innerHTML = `<strong>${file.name}</strong> (${sizeMB} MB)`;
    infoEl.style.color = file.size > 50 * 1024 * 1024 ? '#f59e0b' : 'var(--neon-green)';

    // Enable upload button if name is also filled
    btn.disabled = !(document.getElementById('videoUploadName').value.trim());
};

// Enable/disable upload button based on name input
document.addEventListener('DOMContentLoaded', () => {
    const nameInput = document.getElementById('videoUploadName');
    if (nameInput) {
        nameInput.addEventListener('input', () => {
            const btn = document.getElementById('btnStartUpload');
            if (btn) btn.disabled = !(_selectedVideoFile && nameInput.value.trim());
        });
    }
});

// Compress video using FFmpeg.wasm (only if > 40MB)
async function compressVideo(file, progressCallback) {
    const { FFmpeg } = FFmpegWASM;
    const { fetchFile } = FFmpegUtil;

    if (!_ffmpegInstance) {
        _ffmpegInstance = new FFmpeg();
        _ffmpegInstance.on('progress', ({ progress }) => {
            if (progressCallback) progressCallback(Math.round(progress * 100));
        });
        _ffmpegInstance.on('log', ({ message }) => {
            console.log('[FFmpeg]', message);
        });
    }

    if (!_ffmpegLoaded) {
        await _ffmpegInstance.load({
            coreURL: 'https://cdn.jsdelivr.net/npm/@ffmpeg/core@0.12.6/dist/umd/ffmpeg-core.js',
            wasmURL: 'https://cdn.jsdelivr.net/npm/@ffmpeg/core@0.12.6/dist/umd/ffmpeg-core.wasm',
        });
        _ffmpegLoaded = true;
    }

    const inputName = 'input' + (file.name.includes('.') ? file.name.substring(file.name.lastIndexOf('.')) : '.mp4');
    const outputName = 'compressed.mp4';

    await _ffmpegInstance.writeFile(inputName, await fetchFile(file));

    // Compress: scale to 720p max, lower bitrate, strip audio if needed
    await _ffmpegInstance.exec([
        '-i', inputName,
        '-vf', 'scale=-2:\'min(720,ih)\'',
        '-c:v', 'libx264',
        '-preset', 'fast',
        '-crf', '28',
        '-b:v', '1500k',
        '-maxrate', '1500k',
        '-bufsize', '3000k',
        '-c:a', 'aac',
        '-b:a', '64k',
        '-movflags', '+faststart',
        outputName
    ]);

    const data = await _ffmpegInstance.readFile(outputName);
    const blob = new Blob([data.buffer], { type: 'video/mp4' });

    // Cleanup
    await _ffmpegInstance.deleteFile(inputName);
    await _ffmpegInstance.deleteFile(outputName);

    return blob;
}

// Convert file/blob to base64
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const base64 = reader.result.split(',')[1];
            resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

window.handleVideoUpload = async function () {
    const nameInput = document.getElementById('videoUploadName');
    const videoName = nameInput.value.trim();
    const file = _selectedVideoFile;

    if (!videoName) { showToast('error', 'Nama video wajib diisi!'); return; }
    if (!file) { showToast('error', 'Pilih file video terlebih dahulu!'); return; }

    const progressBar = document.getElementById('uploadProgressBar');
    const progressFill = document.getElementById('uploadProgressFill');
    const progressText = document.getElementById('uploadProgressText');
    const resultBox = document.getElementById('uploadResultBox');
    const btn = document.getElementById('btnStartUpload');

    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-spinner spin"></i> Proses...';
    resultBox.style.display = 'none';
    progressBar.style.display = 'block';
    progressText.style.display = 'block';
    progressFill.style.width = '0%';

    const COMPRESS_THRESHOLD = 40 * 1024 * 1024; // 40MB
    let uploadBlob = file;
    let mimeType = file.type || 'video/mp4';

    try {
        // Step 1: Compress if needed
        if (file.size > COMPRESS_THRESHOLD) {
            const originalMB = (file.size / (1024 * 1024)).toFixed(1);
            progressText.textContent = `Mengompres video (${originalMB} MB)...`;
            progressFill.classList.add('compressing');
            progressFill.style.width = '5%';

            try {
                uploadBlob = await compressVideo(file, (pct) => {
                    progressFill.style.width = Math.min(pct, 95) + '%';
                    progressText.textContent = `Mengompres video... ${pct}%`;
                });

                const compressedMB = (uploadBlob.size / (1024 * 1024)).toFixed(1);
                progressText.textContent = `Kompresi selesai: ${originalMB} MB → ${compressedMB} MB`;
                progressFill.style.width = '100%';
                mimeType = 'video/mp4';

                // Jika masih > 50MB setelah kompresi
                if (uploadBlob.size > 50 * 1024 * 1024) {
                    showToast('error', `File masih terlalu besar setelah kompresi (${compressedMB} MB). Maks 50MB.`);
                    btn.disabled = false;
                    btn.innerHTML = '<i class="ph ph-cloud-arrow-up"></i> Upload';
                    progressBar.style.display = 'none';
                    progressText.style.display = 'none';
                    progressFill.classList.remove('compressing');
                    return;
                }
            } catch (compErr) {
                console.error('Compression failed:', compErr);
                // Fallback: coba upload tanpa kompresi jika file < 50MB asli
                if (file.size <= 50 * 1024 * 1024) {
                    showToast('info', 'Kompresi gagal, mencoba upload langsung...');
                    uploadBlob = file;
                } else {
                    showToast('error', 'File terlalu besar dan kompresi gagal. Coba video yang lebih pendek.');
                    btn.disabled = false;
                    btn.innerHTML = '<i class="ph ph-cloud-arrow-up"></i> Upload';
                    progressBar.style.display = 'none';
                    progressText.style.display = 'none';
                    progressFill.classList.remove('compressing');
                    return;
                }
            }

            progressFill.classList.remove('compressing');
        }

        // Step 2: Convert to base64
        progressText.textContent = 'Menyiapkan file untuk upload...';
        progressFill.style.width = '10%';
        const base64Data = await fileToBase64(uploadBlob);

        // Step 3: Build filename
        const now = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        const dateStr = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
        const ext = mimeType === 'video/mp4' ? '.mp4' : (file.name.includes('.') ? file.name.substring(file.name.lastIndexOf('.')) : '.mp4');
        const finalFileName = `${videoName.replace(/[^a-zA-Z0-9_\-\s]/g, '').replace(/\s+/g, '_')}_${dateStr}${ext}`;

        // Step 4: Upload to Apps Script
        progressText.textContent = 'Mengupload ke Google Drive...';
        progressFill.style.width = '30%';

        // Simulate progress during upload
        let uploadProgress = 30;
        const uploadTimer = setInterval(() => {
            uploadProgress = Math.min(uploadProgress + 2, 90);
            progressFill.style.width = uploadProgress + '%';
        }, 500);

        const payload = {
            action: 'uploadVideo',
            token: window.ieTokenPublic,
            fileName: finalFileName,
            fileData: base64Data,
            mimeType: mimeType
        };

        const res = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(payload)
        });

        clearInterval(uploadTimer);
        const data = await res.json();

        if (data.status === 'success') {
            progressFill.style.width = '100%';
            progressText.textContent = '✅ Upload selesai!';
            progressText.style.color = 'var(--neon-green)';

            // Show result
            resultBox.style.display = 'block';
            document.getElementById('uploadResultLink').value = data.link;

            // Add to session history
            addToUploadHistory(videoName, data.link);

            showToast('success', 'Video berhasil diupload! Copy link lalu paste ke form IE.');

            // Reset inputs for next upload
            nameInput.value = '';
            document.getElementById('videoUploadFile').value = '';
            document.getElementById('videoFileInfo').textContent = 'Belum ada file dipilih';
            document.getElementById('videoFileInfo').style.color = '';
            _selectedVideoFile = null;

        } else {
            showToast('error', 'Upload gagal: ' + (data.message || 'Unknown error'));
            progressText.textContent = '❌ Upload gagal';
            progressText.style.color = 'var(--neon-red)';
        }

    } catch (err) {
        console.error('Upload error:', err);
        showToast('error', 'Koneksi terganggu saat upload.');
        progressText.textContent = '❌ Error koneksi';
        progressText.style.color = 'var(--neon-red)';
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-cloud-arrow-up"></i> Upload';

        // Hide progress after a delay
        setTimeout(() => {
            progressBar.style.display = 'none';
            progressText.style.display = 'none';
            progressText.style.color = '';
            progressFill.style.width = '0%';
        }, 3000);
    }
};

// Copy uploaded video link
window.copyUploadedVideoLink = function () {
    const linkInput = document.getElementById('uploadResultLink');
    if (!linkInput || !linkInput.value) return;

    navigator.clipboard.writeText(linkInput.value).then(() => {
        const btn = document.getElementById('btnCopyUploadLink');
        btn.innerHTML = '<i class="ph ph-check"></i> Copied!';
        btn.classList.add('copied');
        showToast('success', 'Link berhasil di-copy!');
        setTimeout(() => {
            btn.innerHTML = '<i class="ph ph-copy"></i> Copy Link';
            btn.classList.remove('copied');
        }, 2000);
    }).catch(() => {
        // Fallback
        linkInput.select();
        document.execCommand('copy');
        showToast('success', 'Link berhasil di-copy!');
    });
};

// Preview uploaded video
window.previewUploadedVideo = function () {
    const linkInput = document.getElementById('uploadResultLink');
    if (!linkInput || !linkInput.value) return;
    openVideoPlayer(linkInput.value, 'Preview Upload');
};

// Upload History (session-based)
function addToUploadHistory(name, link) {
    let history = JSON.parse(sessionStorage.getItem('videoUploadHistory') || '[]');
    history.unshift({ name, link, time: new Date().toLocaleTimeString('id-ID') });
    if (history.length > 5) history = history.slice(0, 5);
    sessionStorage.setItem('videoUploadHistory', JSON.stringify(history));
    renderUploadHistory();
}

function renderUploadHistory() {
    const area = document.getElementById('uploadHistoryArea');
    const list = document.getElementById('uploadHistoryList');
    if (!area || !list) return;

    const history = JSON.parse(sessionStorage.getItem('videoUploadHistory') || '[]');
    if (history.length === 0) {
        area.style.display = 'none';
        return;
    }

    area.style.display = 'block';
    list.innerHTML = history.map(h =>
        `<div class="history-item">
            <span class="h-name" title="${h.link}">🎬 ${h.name} <span style="color:var(--text-dim);font-size:0.6rem;">(${h.time})</span></span>
            <button class="h-copy" onclick="navigator.clipboard.writeText('${h.link}');showToast('success','Link di-copy!')">📋 Copy</button>
        </div>`
    ).join('');
}

// Render history on page load (if any from session)
document.addEventListener('DOMContentLoaded', () => {
    renderUploadHistory();
});