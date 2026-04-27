// ==========================================
// WHATSAPP INTEGRATION (Opsi 1 & Opsi 3)
// ==========================================

window.shareTextWA = function (lineFilter = 'ALL') {
    if (!allData || !allData.templateInLine) return showToast('error', 'Data belum dimuat secara penuh.');

    var totalSmv = 0, totalAct = 0, totalSav = 0;

    var validRows = allData.templateInLine.filter(function (r) {
        return (r.style || '').trim() !== '' || (r.proses || '').trim() !== '';
    });

    // Fill missing Line/Style (Forward Fill) for accurate filtering
    var currentLineFill = null;
    var currentStyleFill = null;
    var filledRows = validRows.map(function(r) {
        if ((r.line || '').trim() !== '') currentLineFill = (r.line || '').trim();
        if ((r.style || '').trim() !== '') currentStyleFill = (r.style || '').trim();
        return Object.assign({}, r, { effectiveLine: currentLineFill, effectiveStyle: currentStyleFill });
    });

    if (lineFilter !== 'ALL') {
        filledRows = filledRows.filter(r => (r.effectiveLine || '').toUpperCase() === lineFilter.toUpperCase());
        if (filledRows.length === 0) return showToast('error', 'Tidak ada data untuk ' + lineFilter);
    }

    var dt = new Date();
    var jam = dt.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
    var tgl = dt.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

    var sep = '====================';
    var subSep = '--------------------';

    var waText = '*[ TEMPLATE IN LINE ]*\n';
    if (lineFilter !== 'ALL') waText = '*[ LAPORAN ' + lineFilter.toUpperCase() + ' ]*\n';
    waText += 'Tanggal : *' + tgl + '*\n';
    waText += 'Pukul   : *' + jam + ' WIB*\n';
    waText += sep + '\n\n';
    waText += '*RINCIAN PER LINE:*\n';

    var currentLine = null;
    var currentStyle = null;
    var isFirst = true;

    filledRows.forEach(function (r) {
        var rowLine = r.effectiveLine || '-';
        var rowStyle = r.effectiveStyle || '-';
        
        // Cek apakah Line atau Style berubah dari baris sebelumnya
        if (rowLine !== currentLine || rowStyle !== currentStyle) {
            currentLine = rowLine;
            currentStyle = rowStyle;

            if (!isFirst) waText += '\n';
            waText += '>> *' + currentLine.toUpperCase() + '*  (Style: ' + currentStyle + ')\n';
            waText += subSep + '\n';
            isFirst = false;
        }

        // Konversi koma ke titik jika ada, agar parseFloat membaca desimal dengan benar
        var smvNum = parseFloat(String(r.smv || '0').replace(',', '.')) || 0;
        var actNum = parseFloat(String(r.actual || '0').replace(',', '.')) || 0;
        var savNum = parseFloat(String(r.saving || '0').replace(',', '.')) || 0;

        totalSmv += smvNum;
        totalAct += actNum;
        totalSav += savNum;

        var status = '[' + (r.status || '-') + ']';
        var pName = r.proses || '-';
        var rateP = r.rate ? r.rate + '%' : '0%';

        waText += '  ' + status + ' *' + pName + '*\n';
        waText += '    SMV: *' + smvNum.toFixed(1) + '* | Act: *' + actNum.toFixed(1) + '* | Sav: *' + savNum.toFixed(1) + '* | Rate: ' + rateP + '\n';
    });

    waText += '\n' + sep + '\n';
    waText += '*SUMMARY TOTAL:*\n';

    var rate = 0;
    if (totalSmv > 0 && totalAct > 0) rate = Math.round((totalSmv / totalAct) * 100);

    var efStatus = rate >= 100 ? '[TERCAPAI]' : '[BELUM]';

    waText += '  Total SMV    : *' + totalSmv.toFixed(2) + '* Mnt\n';
    waText += '  Total Actual : *' + totalAct.toFixed(2) + '* Mnt\n';
    waText += '  Total Saving : *' + totalSav.toFixed(2) + '* Mnt\n';
    waText += '  Efisiensi    : *' + rate + '%* ' + efStatus + '\n\n';
    waText += '_Dikirim oleh Agentic AI Dept Template_\n';

    var waUrl = 'https://wa.me/?text=' + encodeURIComponent(waText);
    window.open(waUrl, '_blank');
};

window.shareImageWA = async function () {
    var target = document.getElementById('templateCardArea');
    if (!target) return showToast('error', 'Area tabel template tidak ditemukan.');

    var btnText = document.querySelector('button[onclick="shareImageWA()"]');
    if (btnText) btnText.innerHTML = '<i class="ph ph-spinner spin"></i> Memfoto...';

    try {
        var canvas = await html2canvas(target, {
            backgroundColor: '#0a0e17',
            scale: 2
        });

        canvas.toBlob(function (blob) {
            if (!blob) { showToast('error', 'Blob Kosong.'); return; }

            var item = new ClipboardItem({ 'image/png': blob });
            navigator.clipboard.write([item]).then(function () {
                showToast('success', 'Tabel tersalin ke Clipboard! Buka WA Web lalu tekan Ctrl+V.', 4000);
            }).catch(function () {
                showToast('error', 'Clipboard ditolak oleh browser.', 3000);
            });
        }, 'image/png');
    } catch (err) {
        showToast('error', 'Screenshot Gagal: ' + err.message);
    } finally {
        if (btnText) btnText.innerHTML = '<i class="ph ph-camera"></i> Copy as Image';
    }
};
