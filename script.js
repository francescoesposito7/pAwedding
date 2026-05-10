/* ── COUNTDOWN ── */
function updateCountdown() {
  const target = new Date('2026-10-10T11:00:00');
  const now = new Date();
  const diff = target - now;
  if (diff <= 0) {
    ['cd-giorni','cd-ore','cd-minuti','cd-secondi'].forEach(id => {
      document.getElementById(id).textContent = '0';
    });
    return;
  }
  document.getElementById('cd-giorni').textContent  = String(Math.floor(diff / 86400000));
  document.getElementById('cd-ore').textContent     = String(Math.floor((diff % 86400000) / 3600000)).padStart(2,'0');
  document.getElementById('cd-minuti').textContent  = String(Math.floor((diff % 3600000) / 60000)).padStart(2,'0');
  document.getElementById('cd-secondi').textContent = String(Math.floor((diff % 60000) / 1000)).padStart(2,'0');
}
updateCountdown();
setInterval(updateCountdown, 1000);

/* ── SCROLL REVEAL ── */
const revealObserver = new IntersectionObserver(entries => {
  entries.forEach((e, i) => {
    if (e.isIntersecting) {
      setTimeout(() => e.target.classList.add('visible'), i * 80);
      revealObserver.unobserve(e.target);
    }
  });
}, { threshold: 0.12 });
document.querySelectorAll('.reveal').forEach(el => revealObserver.observe(el));

/* ── COPY IBAN ── */
function copyIban() {
  const iban = document.getElementById('iban-text').textContent;
  navigator.clipboard.writeText(iban.replace(/\s/g, '')).then(() => {
    const btn = document.querySelector('.copy-btn');
    btn.textContent = 'Copiato ✓';
    setTimeout(() => btn.textContent = 'Copia', 2000);
  });
}

/* ── RSVP ── */
const btnViene    = document.getElementById('btn-viene');
const btnNonViene = document.getElementById('btn-non-viene');
const inputViene  = document.getElementById('input-viene');
const numWrap     = document.getElementById('presenze-num-wrap');
const personeSection = document.getElementById('persone-section');

const adultiInput  = document.getElementById('rsvp-adulti');
const bambiniInput = document.getElementById('rsvp-bambini');

const INTOLLERANZE_LABELS = ['Celiachia', 'Vegetariano', 'Vegano', 'Altro'];

btnViene.addEventListener('click', () => {
  btnViene.classList.add('active');
  btnNonViene.classList.remove('active');
  inputViene.value = 'viene';
  numWrap.classList.add('visible');
  buildPersoneRows();
});

btnNonViene.addEventListener('click', () => {
  btnNonViene.classList.add('active');
  btnViene.classList.remove('active');
  inputViene.value = 'non-viene';
  numWrap.classList.remove('visible');
  personeSection.classList.remove('visible');
  personeSection.querySelector('#persone-rows').innerHTML = '';
});

adultiInput.addEventListener('input', buildPersoneRows);
bambiniInput.addEventListener('input', buildPersoneRows);

function buildPersoneRows() {
  const adulti  = Math.max(0, parseInt(adultiInput.value)  || 0);
  const bambini = Math.max(0, parseInt(bambiniInput.value) || 0);
  const totale  = adulti + bambini;

  const container = document.getElementById('persone-rows');
  container.innerHTML = '';

  if (totale === 0) {
    personeSection.classList.remove('visible');
    return;
  }
  personeSection.classList.add('visible');

  for (let i = 0; i < totale; i++) {
    const tipo  = i < adulti ? 'Adulto' : 'Bambino';
    const num   = i < adulti ? i + 1 : i - adulti + 1;
    const label = `${tipo} ${num}`;
    const key   = `persona-${i}`;

    const row = document.createElement('div');
    row.className = 'persona-row';
    row.innerHTML = `
      <div class="persona-header">
        <span class="persona-label">${label}</span>
        <input type="text" class="form-input persona-nome"
          placeholder="Nome (opzionale)"
          id="nome-${key}" autocomplete="off">
      </div>
      <div class="persona-checks">
        ${INTOLLERANZE_LABELS.map(l => `
          <label class="persona-check-label" data-persona="${key}" data-label="${l}">
            <input type="checkbox"
              id="chk-${key}-${l.toLowerCase()}"
              class="persona-checkbox"
              data-persona="${key}"
              data-label="${l}">
            ${l}
          </label>
        `).join('')}
      </div>
      <div class="altro-persona-wrap" id="altro-${key}" style="display:none;">
        <input type="text" class="form-input"
          id="altro-detail-${key}"
          placeholder="Specifica l'intolleranza…">
      </div>
    `;
    container.appendChild(row);

    // Mostra campo testo se "Altro" selezionato
    const chkAltro = row.querySelector(`#chk-${key}-altro`);
    chkAltro.addEventListener('change', () => {
      document.getElementById(`altro-${key}`).style.display =
        chkAltro.checked ? 'block' : 'none';
    });
  }
}

/* ── SUBMIT ── */
function submitRsvp(e) {
  e.preventDefault();

  if (!inputViene.value) {
    alert('Indica se parteciperai o meno.');
    return;
  }

  const nome     = document.getElementById('rsvp-nome').value.trim();
  const cognome  = document.getElementById('rsvp-cognome').value.trim();
  const email    = document.getElementById('rsvp-email').value.trim();
  const presenza = inputViene.value;
  const adulti   = presenza === 'viene' ? parseInt(adultiInput.value)  || 0 : 0;
  const bambini  = presenza === 'viene' ? parseInt(bambiniInput.value) || 0 : 0;

  const persone = [];
  if (presenza === 'viene') {
    const rows = document.querySelectorAll('.persona-row');
    rows.forEach((row, i) => {
      const tipo   = i < adulti ? 'Adulto' : 'Bambino';
      const num    = i < adulti ? i + 1 : i - adulti + 1;
      const key    = `persona-${i}`;
      const nome_p = document.getElementById(`nome-${key}`).value.trim();
      const checks = [];
      INTOLLERANZE_LABELS.forEach(l => {
        const chk = document.getElementById(`chk-${key}-${l.toLowerCase()}`);
        if (chk && chk.checked) {
          if (l === 'Altro') {
            const det = document.getElementById(`altro-detail-${key}`).value.trim();
            checks.push(det ? `Altro: ${det}` : 'Altro');
          } else {
            checks.push(l);
          }
        }
      });
      persone.push({
        tipo, num,
        nome: nome_p || `${tipo} ${num}`,
        intolleranze: checks
      });
    });
  }

  // TODO: inviare a Firebase Firestore
  const payload = { nome, cognome, email, presenza, adulti, bambini, persone };
  console.log('RSVP payload:', JSON.stringify(payload, null, 2));

  // UI feedback
  document.getElementById('rsvp-form').style.display = 'none';
  const success = document.getElementById('form-success');
  success.style.display = 'block';
  setTimeout(() => success.classList.add('visible'), 50);
}
