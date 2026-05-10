import { db, auth, collection, getDocs, serverTimestamp,
         GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged }
  from './firebase.js';

// ── UID autorizzato — sostituisci dopo il primo login ──
const AUTHORIZED_UID = 'G4iFBuAYLdXc6tVYsAtluErdpsY2';

let rsvpData = [];

/* AUTH */
onAuthStateChanged(auth, user => {
  if (user && (user.uid === AUTHORIZED_UID || AUTHORIZED_UID === 'G4iFBuAYLdXc6tVYsAtluErdpsY2')) {
    // Mostra UID in console per copiarlo la prima volta
    console.log('Il tuo UID è:', user.uid);
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app').style.display = 'block';
    loadRsvp();
  } else if (user) {
    // Utente loggato ma non autorizzato
    document.getElementById('login-error').style.display = 'block';
    signOut(auth);
  }
});

document.getElementById('login-btn').addEventListener('click', async () => {
  const provider = new GoogleAuthProvider();
  try { await signInWithPopup(auth, provider); }
  catch(e) { console.error(e); }
});

document.getElementById('logout-btn').addEventListener('click', async () => {
    await signOut(auth);
    document.getElementById('app').style.display = 'none';
    document.getElementById('login-screen').style.display = 'flex';
    document.getElementById('login-error').style.display = 'none';
  });

/* LOAD DATA */
async function loadRsvp() {
  const snap = await getDocs(collection(db, 'rsvp'));
  rsvpData = snap.docs.map(d => ({ id: d.id, ...d.data() }));
  renderTable();
}

/* RENDER */
function renderTable() {
  const tbody = document.getElementById('rsvp-tbody');
  tbody.innerHTML = '';
  let totAdulti = 0, totBambini = 0, totIntolleranze = 0;

  if (rsvpData.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="loading">Nessuna risposta ancora.</td></tr>';
    updateStats(0, 0, 0, 0, 0);
    return;
  }

  rsvpData.forEach((r, idx) => {
    const persone = r.persone || [];
    const hasInt  = persone.some(p => p.intolleranze?.length > 0);
    if (hasInt) totIntolleranze++;
    totAdulti  += r.adulti  || 0;
    totBambini += r.bambini || 0;

    const tr = document.createElement('tr');
    if (hasInt) tr.classList.add('has-intolleranze');

    let personeHTML = persone.length > 0
      ? '<ul class="persone-list">' + persone.map(p => {
          const tipoCls  = p.tipo === 'Bambino' ? 'tipo-bambino' : 'tipo-adulto';
          const tags     = (p.intolleranze||[]).length > 0
            ? p.intolleranze.map(i => `<span class="tag-intolleranza">${i}</span>`).join('')
            : '<span class="no-intolleranze">nessuna</span>';
          return `<li>
            <span class="persona-nome-admin">${p.nome}</span>
            <span class="persona-tipo-badge ${tipoCls}">${p.tipo}</span>
            <div class="intolleranze-tags">${tags}</div>
          </li>`;
        }).join('') + '</ul>'
      : '<span class="no-intolleranze">—</span>';

    const data = r.timestamp?.toDate
      ? r.timestamp.toDate().toLocaleDateString('it-IT')
      : '—';

    tr.innerHTML = `
      <td class="td-muted">${idx + 1}</td>
      <td><strong>${r.nome}</strong></td>
      <td>${r.cognome}</td>
      <td class="td-muted">${r.email || '—'}</td>
      <td><span class="badge ${r.presenza === 'viene' ? 'badge-si' : 'badge-no'}">
        ${r.presenza === 'viene' ? '✓ Sì' : '✗ No'}
      </span></td>
      <td>${r.adulti  || '—'}</td>
      <td>${r.bambini || '—'}</td>
      <td>${personeHTML}</td>
      <td class="td-muted">${data}</td>`;
    tbody.appendChild(tr);
  });

  document.getElementById('foot-adulti').textContent  = totAdulti;
  document.getElementById('foot-bambini').textContent = totBambini;
  document.getElementById('foot-intolleranze').textContent =
    `${totIntolleranze} con esigenze`;

  const nonVengono = rsvpData.filter(r => r.presenza === 'non-viene').length;
  updateStats(rsvpData.length, totAdulti, totBambini, totIntolleranze, nonVengono);
  document.getElementById('last-update').textContent =
    'Aggiornato: ' + new Date().toLocaleString('it-IT');
}

function updateStats(rsvp, adulti, bambini, intolleranze, declinati) {
  document.getElementById('stat-rsvp').textContent         = rsvp;
  document.getElementById('stat-presenti').textContent     = adulti;
  document.getElementById('stat-bambini').textContent      = bambini;
  document.getElementById('stat-totale').textContent       = adulti + bambini;
  document.getElementById('stat-intolleranze').textContent = intolleranze;
  document.getElementById('stat-declinato').textContent    = declinati;
}

/* EXPORT EXCEL con ExcelJS + palette matrimonio */
window.exportExcel = async function() {
  const ANTR         = '3D3530';
  const ANTR_MID     = '5C504A';
  const PEWTER       = 'DFDFE3';
  const PEWTER_LIGHT = 'F5F5F7';
  const SEASHELL     = 'F1A17E';
  const SEASHELL_LT  = 'FDE8DC';
  const OLIVE        = 'A3AA5A';
  const OLIVE_LT     = 'EEF0DC';
  const NUDE         = 'E7B46A';
  const NUDE_LT      = 'FBF0DC';
  const WHITE        = 'FFFFFF';
  const TEXT_GRAY    = '999999';
  const TEXT_DARK    = '2C2420';

  const totVengono    = rsvpData.filter(r => r.presenza === 'viene').length;
  const totNonVengono = rsvpData.filter(r => r.presenza === 'non-viene').length;
  const totAdulti     = rsvpData.reduce((s,r) => s+(r.adulti||0), 0);
  const totBambini    = rsvpData.reduce((s,r) => s+(r.bambini||0), 0);
  const totInt        = rsvpData.filter(r => (r.persone||[]).some(p => p.intolleranze?.length > 0)).length;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Matrimonio A&M';
  wb.created = new Date();

  const fillS  = hex => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF'+hex } });
  const fontW  = (size=10, color=TEXT_DARK, bold=false, italic=false) => ({
    name: 'Calibri', size, bold, italic, color: { argb: 'FF'+color }
  });
  const alignC = { horizontal: 'center', vertical: 'middle', wrapText: false };
  const alignL = (wrap=false) => ({ horizontal: 'left', vertical: 'middle', wrapText: wrap });
  const borderThin = {
    top:    { style:'thin', color:{ argb:'FFD8D0C8' }},
    bottom: { style:'thin', color:{ argb:'FFD8D0C8' }},
    left:   { style:'thin', color:{ argb:'FFD8D0C8' }},
    right:  { style:'thin', color:{ argb:'FFD8D0C8' }},
  };

  function setMerged(ws, range, value, bg, fontObj, align) {
    ws.mergeCells(range);
    const cell    = ws.getCell(range.split(':')[0]);
    cell.value    = value;
    cell.fill     = fillS(bg);
    cell.font     = fontObj;
    cell.alignment = align;
  }

  // ══════════════════════
  // FOGLIO 1 — RSVP
  // ══════════════════════
  const ws1 = wb.addWorksheet('RSVP', { views: [{ showGridLines: false }] });
  ws1.columns = [
    {width:4},{width:14},{width:14},{width:26},
    {width:10},{width:8},{width:10},{width:40},{width:12}
  ];

  // Titolo
  setMerged(ws1,'A1:I1','Augusto & Mariapia · Riepilogo RSVP',
    ANTR, fontW(16,WHITE,true,true), alignC);
  ws1.getRow(1).height = 38;

  setMerged(ws1,'A2:I2',
    `10 Ottobre 2026  ·  Chiesa di San Michele  ·  Villa Balke  ·  Aggiornato: ${new Date().toLocaleDateString('it-IT')}`,
    ANTR_MID, fontW(9,PEWTER), alignC);
  ws1.getRow(2).height = 20;

  ws1.getRow(3).height = 8;
  for(let c=1;c<=9;c++) ws1.getCell(3,c).fill = fillS(PEWTER);

  setMerged(ws1,'A4:I4','  PANORAMICA', ANTR, fontW(9,PEWTER,true), alignL());
  ws1.getRow(4).height = 18;

  // Card row 1 — labels
  [['A5:C5','ADULTI'],['D5:F5','BAMBINI'],['G5:I5','CON ESIGENZE ALIMENTARI']].forEach(([rng,label],i) => {
    const bg = i < 2 ? PEWTER_LIGHT : NUDE_LT;
    setMerged(ws1, rng, label, bg, fontW(8,ANTR,true), alignC);
  });
  ws1.getRow(5).height = 16;

  // Card row 1 — valori
  [['A6:C6',totAdulti,PEWTER_LIGHT,ANTR],
   ['D6:F6',totBambini,PEWTER_LIGHT,ANTR],
   ['G6:I6',totInt,NUDE_LT,NUDE]
  ].forEach(([rng,val,bg,fc]) => setMerged(ws1,rng,val,bg,fontW(28,fc,true),alignC));
  ws1.getRow(6).height = 44;

  // Card row 2 — labels
  ws1.mergeCells('A7:F7');
  Object.assign(ws1.getCell('A7'), {
    value:'TOTALE PARTECIPANTI', fill:fillS(OLIVE_LT),
    font:fontW(8,OLIVE,true), alignment:alignC
  });
  setMerged(ws1,'G7:I7','NON PARTECIPANO', SEASHELL_LT, fontW(8,SEASHELL,true), alignC);
  ws1.getRow(7).height = 16;

  // Card row 2 — valori
  ws1.mergeCells('A8:F8');
  Object.assign(ws1.getCell('A8'), {
    value: totAdulti+totBambini, fill:fillS(OLIVE_LT),
    font:fontW(28,OLIVE,true), alignment:alignC
  });
  setMerged(ws1,'G8:I8', totNonVengono, SEASHELL_LT, fontW(28,SEASHELL,true), alignC);
  ws1.getRow(8).height = 44;

  // Separatore
  ws1.getRow(9).height = 8;
  for(let c=1;c<=9;c++) ws1.getCell(9,c).fill = fillS(PEWTER);

  setMerged(ws1,'A10:I10','  ELENCO INVITATI', ANTR, fontW(9,PEWTER,true), alignL());
  ws1.getRow(10).height = 18;

  // Intestazioni
  const headers = ['#','Nome','Cognome','Email','Presenza','Adulti','Bambini','Esigenze alimentari','Data RSVP'];
  headers.forEach((h,i) => {
    const cell = ws1.getCell(11,i+1);
    cell.value=h; cell.fill=fillS(ANTR);
    cell.font=fontW(10,WHITE,true); cell.alignment=alignC; cell.border=borderThin;
  });
  ws1.getRow(11).height = 22;

  // Righe dati
  rsvpData.forEach((r, idx) => {
    const rowNum  = idx+12;
    const viene   = r.presenza === 'viene';
    const persone = r.persone || [];
    const hasInt  = persone.some(p => p.intolleranze?.length > 0);

    const esigenzeList = persone
      .filter(p => p.intolleranze?.length > 0)
      .map(p => `${p.nome} (${p.tipo}): ${p.intolleranze.join(', ')}`);
    const esigenzeStr = esigenzeList.join(' | ') || 'Nessuna';
    const dataStr = r.timestamp?.toDate ? r.timestamp.toDate().toLocaleDateString('it-IT') : '—';

    let bg;
    if (!viene)      bg = SEASHELL_LT;
    else if (hasInt) bg = NUDE_LT;
    else             bg = idx%2===0 ? WHITE : PEWTER_LIGHT;

    const values = [
      idx+1, r.nome, r.cognome, r.email||'',
      viene ? 'Sì' : 'No',
      viene ? (r.adulti||0) : '—',
      viene ? (r.bambini||0) : '—',
      esigenzeStr, dataStr
    ];

    values.forEach((val, ci) => {
      const cell = ws1.getCell(rowNum, ci+1);
      cell.value     = val;
      cell.fill      = fillS(bg);
      cell.border    = borderThin;
      cell.alignment = [1,2,7].includes(ci) ? alignL(true) : alignC;
      if (ci===4)       cell.font = fontW(10, viene ? OLIVE : SEASHELL, true);
      else if (ci===0||ci===8) cell.font = fontW(9, TEXT_GRAY);
      else              cell.font = fontW(10, TEXT_DARK);
    });

    ws1.getRow(rowNum).height = Math.max(15*Math.max(1,persone.length), 18);
  });

  // Totali
  const totRow = rsvpData.length+12;
  ws1.mergeCells(`A${totRow}:E${totRow}`);
  Object.assign(ws1.getCell(`A${totRow}`), {
    value:'TOTALE', fill:fillS(ANTR), font:fontW(10,WHITE,true), alignment:alignC
  });
  [[6,totAdulti],[7,totBambini],[8,`${totInt} con esigenze`]].forEach(([col,val]) => {
    const c = ws1.getCell(totRow,col);
    c.value=val; c.fill=fillS(ANTR); c.font=fontW(10,PEWTER,true); c.alignment=alignC;
  });
  ws1.getCell(totRow,9).fill = fillS(ANTR);
  ws1.getRow(totRow).height = 22;

  // Legenda
  const lr = totRow+2;
  setMerged(ws1,`A${lr}:I${lr}`,'LEGENDA', PEWTER, fontW(9,TEXT_DARK,true), alignL());
  ws1.getRow(lr).height = 16;
  [
    [NUDE_LT,     NUDE,     '  Con esigenze alimentari'],
    [SEASHELL_LT, SEASHELL, '  Non partecipa'],
    [WHITE,       TEXT_DARK,'  Confermato senza esigenze'],
    [PEWTER_LIGHT,TEXT_DARK,'  Confermato senza esigenze (riga alternata)'],
  ].forEach(([bg,fc,label],i) => {
    const r = lr+1+i;
    setMerged(ws1,`A${r}:I${r}`,label,bg,fontW(9,fc),alignL());
    ws1.getRow(r).height = 16;
  });

  // ══════════════════════
  // FOGLIO 2 — DETTAGLIO PERSONE
  // ══════════════════════
  const ws2 = wb.addWorksheet('Dettaglio persone', { views: [{ showGridLines: false }] });
  ws2.columns = [{width:8},{width:22},{width:16},{width:10},{width:12},{width:13},{width:10},{width:32}];

  setMerged(ws2,'A1:H1','Dettaglio Partecipanti · Esigenze Alimentari',
    ANTR, fontW(14,WHITE,true,true), alignC);
  ws2.getRow(1).height = 32;

  setMerged(ws2,'A2:H2',"Una riga per ogni partecipante — ✓ indica la presenza dell'esigenza",
    ANTR_MID, fontW(9,PEWTER), alignC);
  ws2.getRow(2).height = 18;

  ws2.getRow(3).height = 8;
  for(let c=1;c<=8;c++) ws2.getCell(3,c).fill = fillS(PEWTER);

  const detHeaders = ['RSVP #','Referente','Nome','Tipo','Celiachia','Vegetariano','Vegano','Altro'];
  detHeaders.forEach((h,i) => {
    const cell = ws2.getCell(4,i+1);
    cell.value=h; cell.fill=fillS(ANTR);
    cell.font=fontW(10,WHITE,true); cell.alignment=alignC; cell.border=borderThin;
  });
  ws2.getRow(4).height = 22;

  const groupColors = [WHITE, PEWTER_LIGHT];
  let dr = 5;

  rsvpData.forEach((r, idx) => {
    const persone = r.persone || [];
    if (!persone.length) return;
    const groupBg = groupColors[idx%2];

    persone.forEach((p, pIdx) => {
      const isBambino = p.tipo === 'Bambino';
      const ints      = p.intolleranze || [];
      const hasI      = ints.length > 0;
      const bg        = hasI ? NUDE_LT : groupBg;
      const isFirst   = pIdx === 0;

      const brdTop = isFirst
        ? { style:'medium', color:{ argb:'FF'+ANTR }}
        : { style:'thin',   color:{ argb:'FFD8D0C8' }};
      const brdStd   = { style:'thin', color:{ argb:'FFD8D0C8' }};
      const grpBorder = { left:brdStd, right:brdStd, top:brdTop, bottom:brdStd };

      const altroVal = ints.filter(x=>x.startsWith('Altro')).join(', ') || '';
      const vals = [
        isFirst ? idx+1 : '',
        isFirst ? `${r.nome} ${r.cognome}` : '',
        p.nome, p.tipo,
        ints.includes('Celiachia')   ? '✓' : '',
        ints.includes('Vegetariano') ? '✓' : '',
        ints.includes('Vegano')      ? '✓' : '',
        altroVal
      ];

      vals.forEach((val, ci) => {
        const cell = ws2.getCell(dr, ci+1);
        cell.value     = val;
        cell.fill      = fillS(bg);
        cell.border    = grpBorder;
        cell.alignment = [1,2,7].includes(ci) ? alignL() : alignC;

        if (ci===0)        cell.font = fontW(9, TEXT_GRAY, true);
        else if (ci===1)   cell.font = isFirst ? fontW(9,ANTR,true) : fontW(9,TEXT_GRAY);
        else if (ci===3)   cell.font = fontW(9, isBambino ? SEASHELL : OLIVE, true);
        else if ([4,5,6].includes(ci) && val==='✓') cell.font = fontW(11, NUDE, true);
        else if (ci===7 && val) cell.font = fontW(9, SEASHELL, true);
        else               cell.font = fontW(10, TEXT_DARK);
      });

      ws2.getRow(dr).height = 18;
      dr++;
    });
  });

  // Download
  const buffer = await wb.xlsx.writeBuffer();
  const blob   = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download='rsvp-augusto-mariapia.xlsx';
  a.click();
  URL.revokeObjectURL(url);
};

  const alignL = (wrap=false) => ({ horizontal: 'left', vertical: 'middle', wrapText: wrap });
  const borderThin = {
    top:    { style:'thin', color:{ argb:'FFDDDDDD' }},
    bottom: { style:'thin', color:{ argb:'FFDDDDDD' }},
    left:   { style:'thin', color:{ argb:'FFDDDDDD' }},
    right:  { style:'thin', color:{ argb:'FFDDDDDD' }},
  };

  function styleRow(ws, rowNum, bg, fontObj, align, border=false) {
    const row = ws.getRow(rowNum);
    row.eachCell({ includeEmpty: true }, cell => {
      cell.fill      = fillSolid(bg);
      cell.font      = fontObj;
      cell.alignment = align;
      if (border) cell.border = borderThin;
    });
  }

  function setMerged(ws, range, value, bg, fontObj, align) {
    ws.mergeCells(range);
    const cell    = ws.getCell(range.split(':')[0]);
    cell.value    = value;
    cell.fill     = fillSolid(bg);
    cell.font     = fontObj;
    cell.alignment = align;
  }

  // ══════════════════════════════
  // FOGLIO 1 — RSVP
  // ══════════════════════════════
  const ws1 = wb.addWorksheet('RSVP', { views: [{ showGridLines: false }] });
  ws1.columns = [
    { width: 4  }, { width: 14 }, { width: 14 }, { width: 26 },
    { width: 10 }, { width: 8  }, { width: 10 }, { width: 38 }, { width: 12 }
  ];

  // Titolo
  setMerged(ws1,'A1:I1','Mariapia & Augusto · Riepilogo RSVP',
    GREEN_DEEP, fontW(16,WHITE,true,true), alignC);
  ws1.getRow(1).height = 38;

  setMerged(ws1,'A2:I2',
    `10 Ottobre 2026  ·  Chiesa di San Michele  ·  Villa Balke  ·  Aggiornato: ${new Date().toLocaleDateString('it-IT')}`,
    GREEN_MID, fontW(9,IVORY), alignC);
  ws1.getRow(2).height = 20;

  // Separatore
  ws1.getRow(3).height = 10;
  for(let c=1;c<=9;c++) ws1.getCell(3,c).fill = fillSolid(IVORY_DARK);

  // Label statistiche
  setMerged(ws1,'A4:I4','  PANORAMICA', BARK, fontW(9,COPPER_LIGHT,true), alignL());
  ws1.getRow(4).height = 18;

  // Card row 1 — labels
  const cardsTop = [
    ['A5:C5','RISPOSTE RICEVUTE',  GREEN_DEEP,   COPPER_LIGHT],
    ['D5:F5','ADULTI CONFERMATI',  GREEN_MID,    IVORY],
    ['G5:I5','BAMBINI CONFERMATI', COPPER,       IVORY],
  ];
  cardsTop.forEach(([rng,label,bg,fc]) =>
    setMerged(ws1,rng,label,bg,fontW(8,fc,true),alignC));
  ws1.getRow(5).height = 16;

  // Card row 1 — valori
  const cardsTopVal = [
    ['A6:C6', rsvpData.length, GREEN_DEEP],
    ['D6:F6', totAdulti,       GREEN_MID],
    ['G6:I6', totBambini,      COPPER],
  ];
  cardsTopVal.forEach(([rng,val,bg]) =>
    setMerged(ws1,rng,val,bg,fontW(28,WHITE,true),alignC));
  ws1.getRow(6).height = 44;

  // Card row 2 — labels
  const cardsBot = [
    ['A7:C7','TOTALE PARTECIPANTI',      BARK,         COPPER_LIGHT],
    ['D7:F7','CON ESIGENZE ALIMENTARI',  COPPER_LIGHT, BARK],
    ['G7:I7','NON PARTECIPANO',          RED,          IVORY],
  ];
  cardsBot.forEach(([rng,label,bg,fc]) =>
    setMerged(ws1,rng,label,bg,fontW(8,fc,true),alignC));
  ws1.getRow(7).height = 16;

  // Card row 2 — valori
  const cardsBotVal = [
    ['A8:C8', totAdulti+totBambini, BARK,         WHITE],
    ['D8:F8', totInt,               COPPER_LIGHT, BARK],
    ['G8:I8', totNonVengono,        RED,          WHITE],
  ];
  cardsBotVal.forEach(([rng,val,bg,fc]) =>
    setMerged(ws1,rng,val,bg,fontW(28,fc||WHITE,true),alignC));
  ws1.getRow(8).height = 44;

  // Separatore
  ws1.getRow(9).height = 10;
  for(let c=1;c<=9;c++) ws1.getCell(9,c).fill = fillSolid(IVORY_DARK);

  // Titoletto tabella
  setMerged(ws1,'A10:I10','  ELENCO INVITATI', BARK, fontW(9,COPPER_LIGHT,true), alignL());
  ws1.getRow(10).height = 18;

  // Intestazioni
  const headers = ['#','Nome','Cognome','Email','Presenza','Adulti','Bambini','Esigenze alimentari','Data RSVP'];
  headers.forEach((h, i) => {
    const cell = ws1.getCell(11, i+1);
    cell.value     = h;
    cell.fill      = fillSolid(GREEN_DEEP);
    cell.font      = fontW(10, COPPER_LIGHT, true);
    cell.alignment = alignC;
    cell.border    = borderThin;
  });
  ws1.getRow(11).height = 22;

  // Righe dati
  rsvpData.forEach((r, idx) => {
    const rowNum  = idx + 12;
    const viene   = r.presenza === 'viene';
    const persone = r.persone || [];
    const hasInt  = persone.some(p => p.intolleranze?.length > 0);

    const esigenzeList = persone
      .filter(p => p.intolleranze?.length > 0)
      .map(p => `${p.nome} (${p.tipo}): ${p.intolleranze.join(', ')}`);
    const esigenzeStr = esigenzeList.join(' | ') || 'Nessuna';

    const dataStr = r.timestamp?.toDate
      ? r.timestamp.toDate().toLocaleDateString('it-IT') : '—';

    let bg;
    if (!viene)       bg = RED_PALE;
    else if (hasInt)  bg = AMBER_PALE;
    else              bg = idx%2===0 ? WHITE.slice(2) : IVORY;

    const values = [
      idx+1, r.nome, r.cognome, r.email||'',
      viene ? 'Sì' : 'No',
      viene ? (r.adulti||0) : '—',
      viene ? (r.bambini||0) : '—',
      esigenzeStr, dataStr
    ];

    values.forEach((val, ci) => {
      const cell = ws1.getCell(rowNum, ci+1);
      cell.value     = val;
      cell.fill      = fillSolid(bg);
      cell.border    = borderThin;
      cell.alignment = [2,3,8].includes(ci) ? alignL(true) : alignC;

      if (ci === 4) {
        cell.font = fontW(10, viene ? GREEN_DEEP : RED, true);
      } else if (ci === 0 || ci === 8) {
        cell.font = fontW(9, '888888');
      } else {
        cell.font = fontW(10, BARK);
      }
    });

    ws1.getRow(rowNum).height = Math.max(15 * Math.max(1, persone.length), 18);
  });

  // Riga totali
  const totRow = rsvpData.length + 12;
  ws1.mergeCells(`A${totRow}:E${totRow}`);
  const totCell = ws1.getCell(`A${totRow}`);
  totCell.value = 'TOTALE'; totCell.fill = fillSolid(GREEN_DEEP);
  totCell.font = fontW(10,WHITE,true); totCell.alignment = alignC;

  [[6,totAdulti],[7,totBambini],[8,`${totInt} con esigenze`]].forEach(([col,val]) => {
    const c = ws1.getCell(totRow, col);
    c.value = val; c.fill = fillSolid(GREEN_DEEP);
    c.font = fontW(10,COPPER_LIGHT,true); c.alignment = alignC;
  });
  ws1.getCell(totRow,9).fill = fillSolid(GREEN_DEEP);
  ws1.getRow(totRow).height = 22;

  // Legenda
  const lr = totRow + 2;
  setMerged(ws1,`A${lr}:I${lr}`,'LEGENDA',IVORY_DARK,fontW(9,BARK,true),alignL());
  ws1.getRow(lr).height = 16;
  [
    [AMBER_PALE, COPPER, '  Con esigenze alimentari'],
    [RED_PALE,   RED,    '  Non partecipa'],
    [IVORY,      BARK,   '  Confermato senza esigenze'],
  ].forEach(([bg,fc,label], i) => {
    const r = lr+1+i;
    setMerged(ws1,`A${r}:I${r}`,label,bg,fontW(9,fc),alignL());
    ws1.getRow(r).height = 16;
  });

  // ══════════════════════════════
  // FOGLIO 2 — DETTAGLIO PERSONE
  // ══════════════════════════════
  const ws2 = wb.addWorksheet('Dettaglio persone', { views: [{ showGridLines: false }] });
  ws2.columns = [
    {width:8},{width:20},{width:16},{width:10},
    {width:12},{width:13},{width:10},{width:32}
  ];

  setMerged(ws2,'A1:H1','Dettaglio Partecipanti · Esigenze Alimentari',
    GREEN_DEEP, fontW(14,WHITE,true,true), alignC);
  ws2.getRow(1).height = 32;

  setMerged(ws2,'A2:H2',"Una riga per ogni partecipante — ✓ indica la presenza dell'esigenza",
    GREEN_MID, fontW(9,IVORY), alignC);
  ws2.getRow(2).height = 18;

  ws2.getRow(3).height = 8;
  for(let c=1;c<=8;c++) ws2.getCell(3,c).fill = fillSolid(IVORY_DARK);

  const det_headers = ['RSVP #','Referente','Nome','Tipo','Celiachia','Vegetariano','Vegano','Altro'];
  det_headers.forEach((h,i) => {
    const cell = ws2.getCell(4,i+1);
    cell.value=h; cell.fill=fillSolid(BARK);
    cell.font=fontW(10,COPPER_LIGHT,true); cell.alignment=alignC; cell.border=borderThin;
  });
  ws2.getRow(4).height = 22;

  let dr = 5;
  rsvpData.forEach((r, idx) => {
    const persone = r.persone || [];
    if (!persone.length) return;
    persone.forEach(p => {
      const isBambino = p.tipo === 'Bambino';
      const ints      = p.intolleranze || [];
      const hasI      = ints.length > 0;
      const bg        = hasI ? AMBER_PALE : (isBambino ? IVORY : WHITE.slice(2));

      const altroVal = ints.filter(x=>x.startsWith('Altro')).join(', ') || '';
      const vals = [
        idx+1,
        `${r.nome} ${r.cognome}`,
        p.nome, p.tipo,
        ints.includes('Celiachia')   ? '✓' : '',
        ints.includes('Vegetariano') ? '✓' : '',
        ints.includes('Vegano')      ? '✓' : '',
        altroVal
      ];

      vals.forEach((val, ci) => {
        const cell = ws2.getCell(dr, ci+1);
        cell.value     = val;
        cell.fill      = fillSolid(bg);
        cell.border    = borderThin;
        cell.alignment = [1,2,7].includes(ci) ? alignL() : alignC;

        if (ci === 3) {
          cell.font = fontW(9, isBambino ? COPPER : GREEN_DEEP, true);
        } else if ([4,5,6].includes(ci) && val==='✓') {
          cell.font = fontW(12, COPPER, true);
        } else {
          cell.font = fontW(10, BARK);
        }
      });
      ws2.getRow(dr).height = 18;
      dr++;
    });
  });

  // ── Download ──
  const buffer = await wb.xlsx.writeBuffer();
  const blob   = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'rsvp-matrimonio-mariapia-augusto.xlsx';
  a.click();
  URL.revokeObjectURL(url);
};