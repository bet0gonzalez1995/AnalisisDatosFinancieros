
// ---------- 1) Lectura del Excel en el navegador (SheetJS) ----------
const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
const COLMAP = {
    'Fondo General de Participaciones': 'FGP',
    'Fondo de Fomento Municipal': 'FFM', // se usa el TOTAL del fondo; 70%/30% se ignoran a propósito (serían doble conteo)
    'Impuesto sobre Automóviles Nuevos': 'ISAN',
    'Impuesto sobre Tenencia o Uso de Vehículos': 'ISTUV',
    'Impuesto Especial sobre Producción y servicios': 'IEPS',
    'Fondo de Fiscalización y Recaudación': 'FOFIR',
    'Art. 4o.-A Fraccion I de la LCF (Gasolinas)': 'ART4A_I',
    'IEPS Gasolinas Estatal': 'IEPSGAS',
    'Art. 4o.-A Fraccion II de la LCF (FOCO)': 'ART4A_II',
    'Fondo de Compensación del Impuesto sobre Automóviles Nuevos': 'FOCOMPISAN',
    'Recaudación del Impuesto sobre la renta (ISR)': 'ISR',
    'ISR Enajenación Bienes Inmuebles': 'ISRENAJ',
    'Multas por Ingresos por colaboración administrativa': 'MULTAS',
};
const FUND_KEYS = ['FGP', 'FFM', 'ISAN', 'ISTUV', 'IEPS', 'FOFIR', 'ART4A_I', 'IEPSGAS', 'ART4A_II', 'FOCOMPISAN', 'ISR', 'ISRENAJ', 'MULTAS'];
const FUND_LABELS = {
    FGP: 'Fondo General de Participaciones', FFM: 'Fondo de Fomento Municipal', ISAN: 'Impuesto sobre Automóviles Nuevos',
    ISTUV: 'Impuesto sobre Tenencia o Uso de Vehículos', IEPS: 'Impuesto Especial sobre Producción y Servicios',
    FOFIR: 'Fondo de Fiscalización y Recaudación', ART4A_I: 'Art. 4o.-A Fracción I LCF (Gasolinas)', IEPSGAS: 'IEPS Gasolinas Estatal',
    ART4A_II: 'Art. 4o.-A Fracción II LCF (FOCO)', FOCOMPISAN: 'Fondo de Compensación del ISAN', ISR: 'Recaudación del ISR',
    ISRENAJ: 'ISR Enajenación de Bienes Inmuebles', MULTAS: 'Multas por Colaboración Administrativa'
};
const HEADER_BY_KEY = {}; Object.entries(COLMAP).forEach(([h, k]) => HEADER_BY_KEY[k] = h);

let fundKeys = FUND_KEYS, fundLabels = FUND_LABELS, meses = MESES, municipios = [], records = [];
let FIDX0 = 3, TIDX = 3 + FUND_KEYS.length;
let selMes, selMun, selMunDetail, selFondo;

function readWorkbookRows(file) {
    return file.arrayBuffer().then(buf => {
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        return XLSX.utils.sheet_to_json(ws, { defval: null });
    });
}
function n2(v) { if (v === null || v === undefined || v === '') return null; const x = Number(v); return isNaN(x) ? null : Math.round(x * 100) / 100; }

function buildDataset(rows) {
    const munSet = new Set();
    rows.forEach(r => {
        let mu = String(r['MUNICIPIO'] || '').trim().toUpperCase();
        if (mu === 'JUÁREZ') mu = 'JUÁREZ HIDALGO';
        if (mu) munSet.add(mu);
    });
    municipios = Array.from(munSet).sort((a, b) => a.localeCompare(b, 'es'));
    const munIdx = {}; municipios.forEach((m, i) => munIdx[m] = i);
    records = rows.filter(r => r['MUNICIPIO']).map(r => {
        let mu = String(r['MUNICIPIO']).trim().toUpperCase();
        if (mu === 'JUÁREZ') mu = 'JUÁREZ HIDALGO';
        const anio = +r['AÑO'];
        const mesTxt = String(r['MES'] || '').trim().toUpperCase();
        const mesn = MESES.indexOf(mesTxt) + 1;
        const rec = [anio, mesn, munIdx[mu]];
        FUND_KEYS.forEach(k => rec.push(n2(r[HEADER_BY_KEY[k]])));
        rec.push(n2(r['TOTAL']));
        return rec;
    });
    fundKeys = FUND_KEYS; fundLabels = FUND_LABELS; meses = MESES;
    FIDX0 = 3; TIDX = 3 + FUND_KEYS.length;
}

function initUI() {
    selMes = document.getElementById('fMes'); selMes.innerHTML = '<option value="all">Todos</option>';
    meses.forEach((m, i) => { const o = document.createElement('option'); o.value = i + 1; o.textContent = m[0] + m.slice(1).toLowerCase(); selMes.appendChild(o); });
    selMun = document.getElementById('fMun'); selMun.innerHTML = '<option value="all">Todos (estatal)</option>';
    municipios.forEach((m, i) => { const o = document.createElement('option'); o.value = i; o.textContent = m; selMun.appendChild(o); });
    selMunDetail = document.getElementById('munDetailSel'); selMunDetail.innerHTML = '';
    municipios.forEach((m, i) => { const o = document.createElement('option'); o.value = i; o.textContent = m; selMunDetail.appendChild(o); });
    selFondo = document.getElementById('fFondo'); selFondo.innerHTML = '<option value="all">Todos</option>';
    fundKeys.forEach(k => { const o = document.createElement('option'); o.value = k; o.textContent = fundLabels[k]; selFondo.appendChild(o); });
    wireEvents();
}

function wireEvents() {
    ['fAnio', 'fMes', 'fFondo'].forEach(id => document.getElementById(id).addEventListener('change', renderAll));
    document.getElementById('fMun').addEventListener('change', () => {
        const v = document.getElementById('fMun').value;
        if (v !== 'all') selMunDetail.value = v;
        renderAll();
    });
    selMunDetail.addEventListener('change', () => {
        document.getElementById('fMun').value = selMunDetail.value;
        renderAll();
    });
    document.querySelectorAll('nav.tabs button').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('nav.tabs button').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.querySelectorAll('section.view').forEach(v => v.classList.remove('active'));
            document.getElementById('view-' + btn.dataset.view).classList.add('active');
        });
    });
    document.getElementById('btnReset').addEventListener('click', () => {
        try { localStorage.removeItem('hgo_participaciones_cache'); } catch (e) { }
        document.getElementById('dashboardArea').style.display = 'none';
        document.getElementById('uploadPanel').style.display = 'block';
        document.getElementById('fileInput').value = '';
        document.getElementById('uploadStatus').textContent = '';
    });
    document.getElementById('btnExportPdf').addEventListener('click', exportExplorarPDF);
}

function exportExplorarPDF() {
    const btn = document.getElementById('btnExportPdf');
    const originalTxt = btn.textContent;
    btn.disabled = true; btn.textContent = 'Generando PDF…';
    setTimeout(() => {
        try {
            const rows = filtered();
            const fondo = document.getElementById('fFondo').value;
            const keysToShow = fondo === 'all' ? fundKeys : [fondo];
            const anio = document.getElementById('fAnio').value;
            const mes = document.getElementById('fMes').value;
            const mun = document.getElementById('fMun').value;
            const filtroTxt = `Año: ${anio === 'all' ? '2025 y 2026' : anio} · Mes: ${mes === 'all' ? 'todos' : meses[mes - 1]} · Municipio: ${mun === 'all' ? 'todos' : municipios[mun]} · Fondo: ${fondo === 'all' ? 'todos' : fundLabels[fondo]}`;
            const head = [['Año', 'Mes', 'Municipio', ...keysToShow.map(k => fundLabels[k]), 'TOTAL']];
            const body = rows.map(r => [r[0], meses[r[1] - 1], municipios[r[2]],
            ...keysToShow.map(k => { const v = fval(r, k); return v == null ? '' : fmt(v); }),
            fmt(r[TIDX])]);
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'letter' });
            doc.setFontSize(13);
            doc.text('Participaciones municipales de Hidalgo — exploración de datos', 30, 30);
            doc.setFontSize(8);
            doc.text(filtroTxt, 30, 45);
            doc.text(`Generado: ${new Date().toLocaleString('es-MX')} · ${rows.length} filas`, 30, 57);
            doc.autoTable({
                head, body, startY: 68, margin: { left: 20, right: 20 },
                styles: { fontSize: keysToShow.length > 4 ? 6 : 8, cellPadding: 2 },
                headStyles: { fillColor: [76, 110, 140], textColor: 255 },
                alternateRowStyles: { fillColor: [245, 243, 236] },
                didDrawPage: (data) => {
                    doc.setFontSize(7);
                    doc.text(`Página ${doc.internal.getNumberOfPages()}`, doc.internal.pageSize.getWidth() - 60, doc.internal.pageSize.getHeight() - 15);
                }
            });
            doc.save('participaciones_hidalgo_exploracion.pdf');
        } catch (err) {
            alert('No se pudo generar el PDF: ' + err.message);
        } finally {
            btn.disabled = false; btn.textContent = originalTxt;
        }
    }, 30);
}

function onDataReady(fromCache) {
    initUI();
    document.getElementById('uploadPanel').style.display = 'none';
    document.getElementById('dashboardArea').style.display = 'block';
    renderAll();
    if (!fromCache) {
        try { localStorage.setItem('hgo_participaciones_cache', JSON.stringify({ municipios, records })); } catch (e) { }
    }
}

document.getElementById('fileInput').addEventListener('change', (ev) => {
    const files = Array.from(ev.target.files || []);
    if (!files.length) return;
    const status = document.getElementById('uploadStatus');
    status.textContent = 'Leyendo ' + files.length + ' archivo(s)...';
    Promise.all(files.map(readWorkbookRows)).then(results => {
        const allRows = results.flat();
        buildDataset(allRows);
        status.textContent = `Listo: ${records.length} filas, ${municipios.length} municipios.`;
        onDataReady(false);
    }).catch(err => {
        status.textContent = 'No se pudo leer el archivo: ' + err.message + '. Verifica que sea un .xlsx válido.';
    });
});

// intenta recuperar los datos ya procesados de una visita anterior (mismo navegador)
(function tryCache() {
    try {
        const cached = localStorage.getItem('hgo_participaciones_cache');
        if (cached) {
            const { municipios: m, records: r } = JSON.parse(cached);
            municipios = m; records = r;
            onDataReady(true);
        }
    } catch (e) { }
})();

const fmt = n => n == null ? '—' : new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 0 }).format(n);
const fmtPct = n => n == null ? '—' : (n >= 0 ? '+' : '') + n.toFixed(1) + '%';

function fval(rec, key) { if (key === 'TOTAL') return rec[TIDX]; return rec[FIDX0 + fundKeys.indexOf(key)]; }

function filtered() {
    const anio = document.getElementById('fAnio').value;
    const mes = document.getElementById('fMes').value;
    const mun = document.getElementById('fMun').value;
    return records.filter(r => {
        if (anio !== 'all' && r[0] != +anio) return false;
        if (mes !== 'all' && r[1] != +mes) return false;
        if (mun !== 'all' && r[2] != +mun) return false;
        return true;
    });
}

let charts = {};
function setChart(id, cfg) { if (charts[id]) charts[id].destroy(); charts[id] = new Chart(document.getElementById(id), cfg); }
const gridColor = getComputedStyle(document.documentElement).getPropertyValue('--line');
function baseOpts(extra) {
    return Object.assign({ responsive: true, plugins: { legend: { labels: { color: getComputedStyle(document.body).color } } } }, extra || {});
}

function renderKPIs() {
    const rows = filtered();
    const fondo = document.getElementById('fFondo').value;
    const vOf = r => fondo === 'all' ? (r[TIDX] || 0) : (fval(r, fondo) || 0);
    const total = rows.reduce((s, r) => s + vOf(r), 0);
    const munSet = new Set(rows.map(r => r[2]));
    const byMun = {};
    rows.forEach(r => { byMun[r[2]] = (byMun[r[2]] || 0) + vOf(r); });
    const avgMun = munSet.size ? total / munSet.size : 0;
    const byMes = {};
    rows.forEach(r => { const k = r[0] + '-' + r[1]; byMes[k] = (byMes[k] || 0) + vOf(r); });
    let maxM = null, minM = null;
    Object.entries(byMes).forEach(([k, v]) => {
        if (maxM === null || v > byMes[maxM]) maxM = k;
        if (minM === null || v < byMes[minM]) minM = k;
    });
    const lbl = k => { const [a, m] = k.split('-'); return meses[m - 1][0] + meses[m - 1].slice(1).toLowerCase() + ' ' + a; };
    document.getElementById('kpis').innerHTML = `
    <div class="kpi"><div class="lbl">Participaciones distribuidas (pesos)</div><div class="val">${fmt(total)}</div></div>
    <div class="kpi"><div class="lbl">Municipios</div><div class="val">${munSet.size}</div></div>
    <div class="kpi"><div class="lbl">Promedio por municipio</div><div class="val">${fmt(avgMun)}</div></div>
    <div class="kpi"><div class="lbl">Mes con mayor distribución</div><div class="val" style="font-size:1.05rem">${maxM ? lbl(maxM) : '—'}</div><div class="sub">${maxM ? fmt(byMes[maxM]) : ''}</div></div>
    <div class="kpi"><div class="lbl">Mes con menor distribución</div><div class="val" style="font-size:1.05rem">${minM ? lbl(minM) : '—'}</div><div class="sub">${minM ? fmt(byMes[minM]) : ''}</div></div>
  `;
}

function renderEvol() {
    const anio = document.getElementById('fAnio').value;
    const mun = document.getElementById('fMun').value;
    const fondo = document.getElementById('fFondo').value;
    function seriesFor(y) {
        const arr = new Array(12).fill(0);
        records.filter(r => r[0] === y && (mun === 'all' || r[2] == +mun)).forEach(r => {
            arr[r[1] - 1] += (fondo === 'all' ? r[TIDX] : (fval(r, fondo) || 0)) || 0;
        });
        return arr;
    }
    const s25 = seriesFor(2025), s26 = seriesFor(2026).map((v, i) => i < 8 ? v : null);
    const ds = [];
    if (anio !== '2026') ds.push({ label: '2025', data: s25, borderColor: getComputedStyle(document.documentElement).getPropertyValue('--y25'), backgroundColor: 'transparent', tension: .25 });
    if (anio !== '2025') ds.push({ label: '2026', data: s26, borderColor: getComputedStyle(document.documentElement).getPropertyValue('--y26'), backgroundColor: 'transparent', tension: .25 });
    setChart('chEvol', { type: 'line', data: { labels: meses.map(m => m.slice(0, 3)), datasets: ds }, options: baseOpts({ scales: { y: { ticks: { callback: v => fmt(v) } } } }) });
}

function renderCompare() {
    const mun = document.getElementById('fMun').value;
    const fondo = document.getElementById('fFondo').value;
    const sum = (y) => records.filter(r => r[0] === y && r[1] <= 8 && (mun === 'all' || r[2] == +mun))
        .reduce((s, r) => s + ((fondo === 'all' ? r[TIDX] : fval(r, fondo)) || 0), 0);
    const t25 = sum(2025), t26 = sum(2026);
    const diff = t26 - t25, pct = t25 ? diff / t25 * 100 : 0;
    setChart('chCompare', {
        type: 'bar', data: { labels: ['Ene–Ago 2025', 'Ene–Ago 2026'], datasets: [{ data: [t25, t26], backgroundColor: [getComputedStyle(document.documentElement).getPropertyValue('--y25'), getComputedStyle(document.documentElement).getPropertyValue('--y26')] }] },
        options: baseOpts({ plugins: { legend: { display: false } }, scales: { y: { ticks: { callback: v => fmt(v) } } } })
    });
    document.getElementById('compareNote').innerHTML = `Variación: <b class="${diff >= 0 ? 'pos' : 'neg'}">${fmt(diff)} (${fmtPct(pct)})</b> de 2025 a 2026, enero–agosto.`;
}

function renderRankTop() {
    const anio = document.getElementById('fAnio').value;
    const mes = document.getElementById('fMes').value;
    const fondo = document.getElementById('fFondo').value;
    const rows = records.filter(r => (anio === 'all' || r[0] == +anio) && (mes === 'all' || r[1] == +mes));
    const byMun = {};
    rows.forEach(r => { byMun[r[2]] = (byMun[r[2]] || 0) + (fondo === 'all' ? (r[TIDX] || 0) : (fval(r, fondo) || 0)); });
    const arr = Object.entries(byMun).map(([i, v]) => [municipios[i], v]).sort((a, b) => b[1] - a[1]);
    const top = arr.slice(0, 10), bottom = arr.slice(-10).reverse();
    setChart('chRankTop', {
        type: 'bar', data: {
            labels: top.map(x => x[0]).concat(bottom.map(x => x[0])),
            datasets: [{
                label: 'Total', data: top.map(x => x[1]).concat(bottom.map(x => x[1])),
                backgroundColor: top.map(() => getComputedStyle(document.documentElement).getPropertyValue('--y25')).concat(bottom.map(() => getComputedStyle(document.documentElement).getPropertyValue('--y26')))
            }]
        },
        options: baseOpts({ indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { ticks: { callback: v => fmt(v) } } } })
    });
}

function renderHallazgos() {
    const anio = document.getElementById('fAnio').value === 'all' ? '2025' : document.getElementById('fAnio').value;
    const rowsY = records.filter(r => r[0] == +anio && r[1] <= 8);
    const total = rowsY.reduce((s, r) => s + (r[TIDX] || 0), 0);
    const byMun = {}; rowsY.forEach(r => { byMun[r[2]] = (byMun[r[2]] || 0) + (r[TIDX] || 0); });
    const top = Object.entries(byMun).sort((a, b) => b[1] - a[1])[0];
    const pctTop = total ? (top[1] / total * 100) : 0;
    // top10 concentration
    const sorted = Object.values(byMun).sort((a, b) => b - a);
    const top10sum = sorted.slice(0, 10).reduce((s, v) => s + v, 0);
    const pctTop10 = total ? top10sum / total * 100 : 0;
    // month with biggest MoM jump (same year)
    const byMes = {}; records.filter(r => r[0] == +anio).forEach(r => { byMes[r[1]] = (byMes[r[1]] || 0) + (r[TIDX] || 0); });
    let bestJump = null;
    for (let m = 2; m <= 12; m++) { if (byMes[m] != null && byMes[m - 1] != null) { const d = byMes[m] - byMes[m - 1]; if (bestJump === null || Math.abs(d) > Math.abs(bestJump.d)) bestJump = { m, d, pct: byMes[m - 1] ? d / byMes[m - 1] * 100 : 0 }; } }
    const items = [
        `El municipio de <b>${municipios[top[0]]}</b> concentró ${fmt(top[1])} en ${anio} (ene–ago), equivalente a <b>${pctTop.toFixed(1)}%</b> del total estatal del periodo.`,
        `Los 10 municipios con mayor participación reúnen <b>${pctTop10.toFixed(1)}%</b> del total distribuido en ${anio} (ene–ago), frente a solo 10 de 84 municipios.`,
    ];
    if (bestJump) items.push(`El cambio mensual más marcado en ${anio} fue entre ${meses[bestJump.m - 2].toLowerCase()} y ${meses[bestJump.m - 1].toLowerCase()}: <b class="${bestJump.d >= 0 ? 'pos' : 'neg'}">${fmt(bestJump.d)} (${fmtPct(bestJump.pct)})</b>. La causa observada debe confirmarse revisando el fondo que más cambió ese mes en la vista "Fondos".`);
    document.getElementById('hallazgos').innerHTML = items.map(t => `<li>${t}</li>`).join('');
    document.getElementById('hallazgos2').innerHTML = items.map(t => `<li>${t}</li>`).join('');
}

function renderMunicipioDetail() {
    const mi = +selMunDetail.value;
    const rows = records.filter(r => r[2] === mi);
    const sum = y => rows.filter(r => r[0] === y).reduce((s, r) => s + (r[TIDX] || 0), 0);
    const sumEA = y => rows.filter(r => r[0] === y && r[1] <= 8).reduce((s, r) => s + (r[TIDX] || 0), 0);
    const t25 = sumEA(2025), t26 = sumEA(2026);
    const varPct = t25 ? (t26 - t25) / t25 * 100 : 0;
    document.getElementById('munKpis').innerHTML = `
    <div class="kpi"><div class="lbl">Total 2025</div><div class="val tag25">${fmt(sum(2025))}</div></div>
    <div class="kpi"><div class="lbl">Total 2026 (ene–ago)</div><div class="val tag26">${fmt(sum(2026))}</div></div>
    <div class="kpi"><div class="lbl">Variación ene–ago 2025 vs 2026</div><div class="val ${varPct >= 0 ? 'pos' : 'neg'}">${fmtPct(varPct)}</div></div>
  `;
    const s25 = new Array(12).fill(0), s26 = new Array(12).fill(null);
    rows.forEach(r => { if (r[0] === 2025) s25[r[1] - 1] = r[TIDX] || 0; if (r[0] === 2026 && r[1] <= 8) s26[r[1] - 1] = r[TIDX] || 0; });
    setChart('chMunEvol', {
        type: 'line', data: {
            labels: meses.map(m => m.slice(0, 3)), datasets: [
                { label: '2025', data: s25, borderColor: getComputedStyle(document.documentElement).getPropertyValue('--y25'), backgroundColor: 'transparent' },
                { label: '2026', data: s26, borderColor: getComputedStyle(document.documentElement).getPropertyValue('--y26'), backgroundColor: 'transparent' }
            ]
        }, options: baseOpts({ plugins: { title: { display: true, text: 'Evolución mensual', color: getComputedStyle(document.body).color } } })
    });
    const fondoTotals = fundKeys.map(k => rows.reduce((s, r) => s + (fval(r, k) || 0), 0));
    setChart('chMunFondos', {
        type: 'doughnut', data: { labels: fundKeys.map(k => fundLabels[k]), datasets: [{ data: fondoTotals, backgroundColor: palette(fundKeys.length) }] },
        options: baseOpts({ plugins: { title: { display: true, text: 'Composición por fondo (histórico)', color: getComputedStyle(document.body).color }, legend: { position: 'right', labels: { boxWidth: 10, font: { size: 9 }, color: getComputedStyle(document.body).color } } } })
    });
}

function palette(n) {
    const base = ['#4C6E8C', '#C0743F', '#7A8C5A', '#8A6BA6', '#B0503F', '#4C9A8C', '#C79B3F', '#5A6B8C', '#A65A7A', '#6B8C4C', '#8C6B4C', '#4C8CA6', '#8C4C6B'];
    return Array.from({ length: n }, (_, i) => base[i % base.length]);
}

function renderRankingFull() {
    const anio = document.getElementById('fAnio').value === 'all' ? null : +document.getElementById('fAnio').value;
    const fondo = document.getElementById('fFondo').value;
    function sumFor(y) {
        const rows = records.filter(r => r[1] <= 8 && (y ? r[0] === y : true));
        const byMun = {}; rows.forEach(r => { byMun[r[2]] = (byMun[r[2]] || 0) + (fondo === 'all' ? (r[TIDX] || 0) : (fval(r, fondo) || 0)); }); return byMun;
    }
    const cur = anio ? sumFor(anio) : sumFor(null);
    const prevYear = anio === 2025 ? null : 2025;
    const prev = prevYear ? sumFor(2025) : null;
    const total = Object.values(cur).reduce((s, v) => s + v, 0);
    const arr = Object.entries(cur).map(([i, v]) => ({ i: +i, v })).sort((a, b) => b.v - a.v);
    const tbody = document.querySelector('#tblRanking tbody');
    tbody.innerHTML = arr.map((x, idx) => {
        const pct = total ? (x.v / total * 100).toFixed(2) : '—';
        let varTxt = '—', cls = '';
        if (prev) { const p = prev[x.i] || 0; if (p) { const d = (x.v - p) / p * 100; varTxt = fmtPct(d); cls = d >= 0 ? 'pos' : 'neg'; } }
        return `<tr><td>${idx + 1}</td><td>${municipios[x.i]}</td><td>${fmt(x.v)}</td><td>${pct}%</td><td class="${cls}">${varTxt}</td></tr>`;
    }).join('');
}

function renderFondos() {
    const rows = filtered();
    const totals = fundKeys.map(k => rows.reduce((s, r) => s + (fval(r, k) || 0), 0));
    const grand = totals.reduce((s, v) => s + v, 0);
    setChart('chFondoPie', {
        type: 'pie', data: { labels: fundKeys.map(k => fundLabels[k]), datasets: [{ data: totals, backgroundColor: palette(fundKeys.length) }] },
        options: baseOpts({ plugins: { legend: { position: 'right', labels: { boxWidth: 10, font: { size: 9 }, color: getComputedStyle(document.body).color } } } })
    });
    const tbody = document.querySelector('#tblFondos tbody');
    const order = fundKeys.map((k, i) => ({ k, v: totals[i] })).sort((a, b) => b.v - a.v);
    tbody.innerHTML = order.map(o => `<tr><td>${fundLabels[o.k]}</td><td>${fmt(o.v)}</td><td>${grand ? (o.v / grand * 100).toFixed(2) : '0'}%</td></tr>`).join('');
    // evolution of each fund by month (selected year context, default both averaged monthly totals)
    const anio = document.getElementById('fAnio').value;
    const mun = document.getElementById('fMun').value;
    const yUse = anio === 'all' ? 2026 : +anio;
    const monthly = fundKeys.map(k => {
        const arr = new Array(12).fill(0);
        records.filter(r => r[0] === yUse && (mun === 'all' || r[2] == +mun)).forEach(r => { arr[r[1] - 1] += fval(r, k) || 0; });
        return arr;
    });
    const colors = palette(fundKeys.length);
    setChart('chFondoEvol', {
        type: 'line', data: { labels: meses.map(m => m.slice(0, 3)), datasets: fundKeys.map((k, i) => ({ label: fundLabels[k], data: monthly[i], borderColor: colors[i], backgroundColor: 'transparent', tension: .2 })) },
        options: baseOpts({ plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 9 }, color: getComputedStyle(document.body).color } }, title: { display: true, text: `Fondos por mes — ${yUse}`, color: getComputedStyle(document.body).color } } })
    });
}

function renderExplorar() {
    const rows = filtered();
    const fondo = document.getElementById('fFondo').value;
    const keysToShow = fondo === 'all' ? fundKeys : [fondo];
    const thead = document.querySelector('#tblExplorar thead');
    thead.innerHTML = '<tr><th>Año</th><th>Mes</th><th>Municipio</th>' + keysToShow.map(k => `<th>${fundLabels[k]}</th>`).join('') + '<th>TOTAL</th></tr>';
    const tbody = document.querySelector('#tblExplorar tbody');
    const MAXROWS = 500;
    const slice = rows.slice(0, MAXROWS);
    tbody.innerHTML = slice.map(r => `<tr><td>${r[0]}</td><td>${meses[r[1] - 1]}</td><td>${municipios[r[2]]}</td>` +
        keysToShow.map(k => { const v = fval(r, k); return `<td>${v == null ? '' : fmt(v)}</td>`; }).join('') + `<td>${fmt(r[TIDX])}</td></tr>`).join('');
    document.getElementById('expCount').textContent = `Mostrando ${slice.length} de ${rows.length} filas que cumplen el filtro. Ajusta los filtros superiores para acotar la búsqueda.`;
}

function renderAll() {
    renderKPIs(); renderEvol(); renderCompare(); renderRankTop(); renderHallazgos();
    renderMunicipioDetail(); renderRankingFull(); renderFondos(); renderExplorar();
}
