/* ============================================================
   E-IPL AUCTION — app.js
   Pure vanilla JS. No frameworks. localStorage persistence.
   ============================================================ */

// ── State ─────────────────────────────────────────────────────
let owners    = [];   // { name, team, purse, totalPurse, slots, bought: [] }
let allPlayers = [];  // parsed from uploaded file
let setQueue  = [];   // shuffled players for current set
let currentIdx = -1;  // index in setQueue
let currentPlayer = null;

// ── Helpers ───────────────────────────────────────────────────
const $  = id => document.getElementById(id);
const show = id => $(id).classList.remove('hidden');
const hide = id => $(id).classList.add('hidden');

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function saveState() {
  localStorage.setItem('eipl_owners', JSON.stringify(owners));
}

function loadState() {
  const saved = localStorage.getItem('eipl_owners');
  if (saved) owners = JSON.parse(saved);
}

// ── Particle canvas ───────────────────────────────────────────
(function initParticles() {
  const canvas = $('particles-canvas');
  const ctx = canvas.getContext('2d');
  let W, H, particles = [];

  function resize() {
    W = canvas.width  = window.innerWidth;
    H = canvas.height = window.innerHeight;
  }
  resize();
  window.addEventListener('resize', resize);

  for (let i = 0; i < 70; i++) {
    particles.push({
      x: Math.random() * window.innerWidth,
      y: Math.random() * window.innerHeight,
      r: Math.random() * 1.5 + 0.3,
      dx: (Math.random() - 0.5) * 0.35,
      dy: (Math.random() - 0.5) * 0.35,
      alpha: Math.random() * 0.5 + 0.1
    });
  }

  function draw() {
    ctx.clearRect(0, 0, W, H);
    particles.forEach(p => {
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(245,197,24,${p.alpha})`;
      ctx.fill();
      p.x += p.dx; p.y += p.dy;
      if (p.x < 0 || p.x > W) p.dx *= -1;
      if (p.y < 0 || p.y > H) p.dy *= -1;
    });
    requestAnimationFrame(draw);
  }
  draw();
})();

// ── PAGE 1: Add Owner ─────────────────────────────────────────
function addOwner() {
  const name  = $('inp-name').value.trim();
  const team  = $('inp-team').value.trim();
  const purse = parseFloat($('inp-purse').value);
  const slots = parseInt($('inp-slots').value);

  if (!name || !team || isNaN(purse) || purse <= 0 || isNaN(slots) || slots <= 0) {
    showSetupMsg('Please fill in all fields correctly.', 'error');
    return;
  }

  owners.push({ name, team, purse, totalPurse: purse, slots, totalSlots: slots, bought: [] });
  saveState();

  // clear form
  $('inp-name').value = '';
  $('inp-team').value = '';
  $('inp-purse').value = '';
  $('inp-slots').value = '';
  $('inp-name').focus();

  showSetupMsg('', '');
  renderOwnerCards();
}

function removeOwner(idx) {
  owners.splice(idx, 1);
  saveState();
  renderOwnerCards();
}

function renderOwnerCards() {
  const grid = $('owners-grid');
  grid.innerHTML = '';

  owners.forEach((o, i) => {
    const card = document.createElement('div');
    card.className = 'owner-card';
    card.innerHTML = `
      <button class="btn-remove-owner" onclick="removeOwner(${i})" title="Remove">✕</button>
      <div class="owner-card-name">${o.name}</div>
      <div class="owner-card-team">${o.team}</div>
      <div class="owner-card-meta">
        <div class="owner-meta-chip">💰 <span>₹${o.purse} Cr</span></div>
        <div class="owner-meta-chip">👤 <span>${o.slots}</span> slots</div>
      </div>
    `;
    grid.appendChild(card);
  });

  $('btn-next').disabled = owners.length < 2;
  if (owners.length >= 2) {
    showSetupMsg(`${owners.length} teams ready. Proceed when everyone's registered.`, 'ok');
  } else if (owners.length === 1) {
    showSetupMsg('Add at least one more team.', 'warn');
  } else {
    showSetupMsg('', '');
  }
}

function showSetupMsg(msg, type) {
  const el = $('setup-msg');
  el.textContent = msg;
  el.style.color = type === 'error' ? '#e63946'
                 : type === 'ok'    ? '#00c9a7'
                 : type === 'warn'  ? '#ff6b35'
                 : '#8a93a8';
}

// ── Page navigation ───────────────────────────────────────────
function goToAuction() {
  if (owners.length < 2) return;
  $('page-setup').classList.remove('active');
  $('page-auction').classList.add('active');
  renderScoreboard();
  buildTeamButtons();
}

function goBack() {
  $('page-auction').classList.remove('active');
  $('page-setup').classList.add('active');
  renderOwnerCards();
}

// ── FILE UPLOAD (Excel / CSV) ──────────────────────────────────
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  $('file-label-text').textContent = '📄 ' + file.name;
  $('file-status').textContent = 'Reading file…';

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data  = new Uint8Array(e.target.result);
      const wb    = XLSX.read(data, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows  = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      if (!rows.length) {
        $('file-status').textContent = '⚠ No data found in file.';
        return;
      }

      // Normalise column names — strips spaces, underscores, dashes, brackets, units, case-insensitive
      const norm = s => String(s).toLowerCase().replace(/[\s_\-\(\)]/g, '').replace(/cr$/, '');

      allPlayers = rows.map(r => {
        const keys = Object.keys(r);
        const get  = (...names) => {
          for (const n of names) {
            const k = keys.find(k => norm(k) === norm(n));
            if (k !== undefined && r[k] !== '') return r[k];
          }
          return '';
        };

        const rawPrice = get(
          'price', 'baseprice', 'base price', 'base_price',
          'baserate', 'base rate', 'amount', 'value', 'cost'
        );

        return {
          setNo:      Number(get('setno','setnumber','set no','set number','set')) || 0,
          playerName: String(get('playername','player name','player_name','name','player')).trim(),
          price:      parseFloat(String(rawPrice).replace(/[^\d.]/g, '')) || 0,
          isOverseas: Number(get('isoverseas','is overseas','is_overseas','overseas')) || 0,
          sold:       false
        };
      }).filter(p => p.playerName);

      // Debug: log first row keys so we can diagnose if still broken
      if (rows.length) console.log('Excel columns detected:', Object.keys(rows[0]));

      $('file-status').textContent = `✅ Loaded ${allPlayers.length} players from "${file.name}"`;
    } catch (err) {
      $('file-status').textContent = '❌ Error reading file: ' + err.message;
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── START SET ─────────────────────────────────────────────────
function startSet() {
  const setNo = parseInt($('inp-set').value);
  if (isNaN(setNo) || setNo < 1) {
    alert('Enter a valid set number.'); return;
  }
  if (!allPlayers.length) {
    alert('Please upload a player file first.'); return;
  }

  const setPlayers = allPlayers.filter(p => p.setNo === setNo && !p.sold);
  if (!setPlayers.length) {
    alert(`No unsold players found in Set ${setNo}.`); return;
  }

  setQueue  = shuffle(setPlayers);
  currentIdx = -1;

  hide('set-empty-msg');
  show('player-card-wrap');
  $('card-set-badge').textContent = `SET ${setNo}`;

  nextPlayer();
}

// ── NEXT PLAYER ───────────────────────────────────────────────
function nextPlayer() {
  currentIdx++;
  if (currentIdx >= setQueue.length) {
    showSetComplete();
    return;
  }
  currentPlayer = setQueue[currentIdx];
  renderPlayerCard(currentPlayer);
  $('inp-bid').value = '';
}

function renderPlayerCard(p) {
  hide('set-empty-msg');

  // Initials
  const parts    = p.playerName.split(' ');
  const initials = parts.length >= 2
    ? (parts[0][0] + parts[parts.length-1][0]).toUpperCase()
    : p.playerName.slice(0,2).toUpperCase();

  $('card-initials').textContent = initials;
  $('card-name').textContent     = p.playerName;
  $('card-price').textContent    = `₹${p.price} Cr`;
  $('card-type').textContent     = p.isOverseas ? 'Overseas 🌍' : 'Local 🇮🇳';

  const ovBadge = $('card-overseas-badge');
  if (p.isOverseas) {
    ovBadge.textContent = '🌍 OVERSEAS';
    ovBadge.classList.remove('hidden');
  } else {
    ovBadge.classList.add('hidden');
  }

  // animate card
  const card = $('player-card');
  card.style.animation = 'none';
  void card.offsetWidth;
  card.style.animation = 'fadeIn 0.35s ease';
}

function showSetComplete() {
  currentPlayer = null;
  // hide player card content visually but show empty msg
  show('set-empty-msg');
  $('card-name').textContent    = '—';
  $('card-initials').textContent = '?';
  $('card-price').textContent   = '₹—';
  $('card-type').textContent    = '—';
  $('card-overseas-badge').classList.add('hidden');
  $('inp-bid').value = '';
}

// ── TEAM BUTTONS ──────────────────────────────────────────────
function buildTeamButtons() {
  const wrap = $('team-buttons');
  wrap.innerHTML = '';

  owners.forEach((o, i) => {
    const btn = document.createElement('button');
    btn.className = 'btn-sell-team';
    btn.id = `team-btn-${i}`;
    btn.innerHTML = `
      ${o.team}
      <small class="team-purse-left">₹${o.purse.toFixed(1)} Cr · ${o.slots} slots</small>
    `;
    btn.onclick = () => sellTo(i);
    wrap.appendChild(btn);
  });
}

function refreshTeamButtons() {
  owners.forEach((o, i) => {
    const btn = $(`team-btn-${i}`);
    if (!btn) return;
    btn.innerHTML = `
      ${o.team}
      <small class="team-purse-left">₹${o.purse.toFixed(1)} Cr · ${o.slots} slots</small>
    `;
    btn.disabled = o.purse <= 0 || o.slots <= 0;
    if (btn.disabled) btn.style.opacity = '0.4';
  });
}

// ── SELL PLAYER ───────────────────────────────────────────────
function sellTo(ownerIdx) {
  if (!currentPlayer) {
    alert('No player is currently up for auction.'); return;
  }

  const bidVal = parseFloat($('inp-bid').value);
  if (isNaN(bidVal) || bidVal < 0) {
    alert('Enter a valid winning bid amount.'); return;
  }

  const owner = owners[ownerIdx];

  if (bidVal > owner.purse) {
    alert(`${owner.team} doesn't have enough purse! Available: ₹${owner.purse.toFixed(1)} Cr`);
    return;
  }
  if (owner.slots <= 0) {
    alert(`${owner.team} has no player slots remaining.`); return;
  }

  // Record purchase
  owner.purse  = parseFloat((owner.purse - bidVal).toFixed(2));
  owner.slots -= 1;
  owner.bought.push({ name: currentPlayer.playerName, price: bidVal, isOverseas: currentPlayer.isOverseas });

  // Mark player as sold in master list
  const masterIdx = allPlayers.findIndex(p => p.playerName === currentPlayer.playerName && p.setNo === currentPlayer.setNo);
  if (masterIdx !== -1) allPlayers[masterIdx].sold = true;

  saveState();
  renderScoreboard();
  refreshTeamButtons();
  showSoldToast(currentPlayer.playerName, bidVal, owner.team);

  currentPlayer = null;
  $('inp-bid').value = '';
}

// ── MARK UNSOLD ───────────────────────────────────────────────
function markUnsold() {
  if (!currentPlayer) { alert('No active player.'); return; }

  const masterIdx = allPlayers.findIndex(p => p.playerName === currentPlayer.playerName && p.setNo === currentPlayer.setNo);
  if (masterIdx !== -1) allPlayers[masterIdx].sold = true; // mark so it's skipped

  addUnsoldEntry(currentPlayer);
  currentPlayer = null;
  $('inp-bid').value = '';
  nextPlayer();
}

function addUnsoldEntry(p) {
  const list = $('unsold-list');
  const note = list.querySelector('.empty-note');
  if (note) note.remove();

  const item = document.createElement('div');
  item.className = 'unsold-item';
  item.innerHTML = `<span>${p.playerName}</span><span class="u-price">₹${p.price} Cr</span>`;
  list.appendChild(item);
}

// ── SCOREBOARD ────────────────────────────────────────────────
function renderScoreboard() {
  const list = $('scoreboard-list');
  list.innerHTML = '';

  owners.forEach((o, i) => {
    const spentPct = Math.round(((o.totalPurse - o.purse) / o.totalPurse) * 100);
    const boughtCount = o.bought.length;
    const overseasCount = o.bought.filter(b => b.isOverseas).length;

    const card = document.createElement('div');
    card.className = 'score-card';
    card.innerHTML = `
      <div class="score-card-header">
        <div>
          <div class="score-team-name">${o.team}</div>
          <div class="score-owner-name">${o.name}</div>
        </div>
        <div class="score-purse">
          ₹${o.purse.toFixed(1)} Cr
          <small>remaining</small>
        </div>
      </div>
      <div class="score-bar-wrap">
        <div class="score-bar" style="width:${spentPct}%"></div>
      </div>
      <div class="score-players-count">
        <span>${boughtCount} player${boughtCount !== 1 ? 's' : ''} bought</span>
        <span>${o.slots} slot${o.slots !== 1 ? 's' : ''} left</span>
      </div>
      <div class="score-overseas-count">
        🌍 <span>${overseasCount}</span> overseas
      </div>
      ${boughtCount > 0 ? `
        <button class="btn-toggle-players" onclick="toggleBought(${i})">▾ View Squad</button>
        <div class="score-players-list" id="bought-list-${i}">
          ${o.bought.map(b => `
            <div class="bought-player-row">
              <span class="bp-name">${b.isOverseas ? '🌍 ' : ''}${b.name}</span>
              <span class="bp-price">₹${b.price} Cr</span>
            </div>
          `).join('')}
        </div>
      ` : ''}
    `;
    list.appendChild(card);
  });
}

function toggleBought(idx) {
  const el = $(`bought-list-${idx}`);
  if (!el) return;
  el.classList.toggle('open');
  const btn = el.previousElementSibling;
  btn.textContent = el.classList.contains('open') ? '▴ Hide Squad' : '▾ View Squad';
}

// ── SOLD TOAST ────────────────────────────────────────────────
function showSoldToast(playerName, bid, teamName) {
  const toast = $('sold-toast');
  $('toast-detail').textContent = `${playerName}  →  ${teamName}  @  ₹${bid} Cr`;
  toast.classList.remove('hidden');
  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.classList.add('hidden'), 350);
  }, 2800);
}

// ── INIT ──────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  loadState();
  renderOwnerCards();

  // Allow Enter key in setup form
  ['inp-name','inp-team','inp-purse','inp-slots'].forEach(id => {
    $(id)?.addEventListener('keydown', e => { if (e.key === 'Enter') addOwner(); });
  });
});

// ── END AUCTION — PDF EXPORT ───────────────────────────────────
function endAuction() {
  if (!owners.length) { alert('No teams found.'); return; }

  if (!confirm('End the auction and download the squads as PDF?')) return;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  const PAGE_W  = 210;
  const PAGE_H  = 297;
  const MARGIN  = 14;
  const COL_W   = PAGE_W - MARGIN * 2;

  // ── colour helpers ──
  const hexToRgb = hex => {
    const r = parseInt(hex.slice(1,3),16);
    const g = parseInt(hex.slice(3,5),16);
    const b = parseInt(hex.slice(5,7),16);
    return [r,g,b];
  };
  const GOLD   = hexToRgb('#f5c518');
  const ORANGE = hexToRgb('#ff6b35');
  const DARK   = hexToRgb('#080c14');
  const CARD   = hexToRgb('#0d1526');
  const WHITE  = [255,255,255];
  const DIM    = hexToRgb('#8a93a8');
  const TEAL   = hexToRgb('#00c9a7');
  const RED    = hexToRgb('#e63946');

  let y = 0;

  function newPage() {
    doc.addPage();
    y = 0;
    drawPageBg();
  }

  function drawPageBg() {
    doc.setFillColor(...DARK);
    doc.rect(0, 0, PAGE_W, PAGE_H, 'F');
  }

  function checkY(needed) {
    if (y + needed > PAGE_H - 10) newPage();
  }

  // ── COVER PAGE ──────────────────────────────────────────────
  drawPageBg();

  // gold top bar
  doc.setFillColor(...GOLD);
  doc.rect(0, 0, PAGE_W, 3, 'F');

  // title
  y = 40;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(32);
  doc.setTextColor(...GOLD);
  doc.text('E-IPL AUCTION', PAGE_W / 2, y, { align: 'center' });

  y += 10;
  doc.setFontSize(13);
  doc.setTextColor(...DIM);
  doc.text('OFFICIAL SQUAD REPORT', PAGE_W / 2, y, { align: 'center' });

  y += 20;
  // decorative line
  doc.setDrawColor(...GOLD);
  doc.setLineWidth(0.5);
  doc.line(MARGIN + 20, y, PAGE_W - MARGIN - 20, y);

  y += 14;
  doc.setFontSize(11);
  doc.setTextColor(...WHITE);
  doc.text(`Total Teams: ${owners.length}`, PAGE_W / 2, y, { align: 'center' });

  y += 8;
  const totalSold = owners.reduce((s, o) => s + o.bought.length, 0);
  const totalSpent = owners.reduce((s, o) => s + (o.totalPurse - o.purse), 0);
  doc.text(`Total Players Sold: ${totalSold}`, PAGE_W / 2, y, { align: 'center' });

  y += 8;
  doc.text(`Total Amount Spent: Rs. ${totalSpent.toFixed(1)} Cr`, PAGE_W / 2, y, { align: 'center' });

  y += 8;
  const now = new Date();
  doc.setTextColor(...DIM);
  doc.setFontSize(9);
  doc.text(`Generated: ${now.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' })}  ${now.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit' })}`, PAGE_W / 2, y, { align: 'center' });

  // decorative cricket text
  y += 30;
  doc.setFontSize(22);
  doc.setTextColor(...GOLD);
  doc.text('[ E-IPL AUCTION ]', PAGE_W / 2, y, { align: 'center' });

  // gold bottom bar
  doc.setFillColor(...GOLD);
  doc.rect(0, PAGE_H - 3, PAGE_W, 3, 'F');

  // ── TEAM PAGES ──────────────────────────────────────────────
  owners.forEach((owner, oi) => {
    newPage();

    // top accent bar (gradient simulation: gold left to orange right)
    doc.setFillColor(...GOLD);
    doc.rect(0, 0, PAGE_W / 2, 2, 'F');
    doc.setFillColor(...ORANGE);
    doc.rect(PAGE_W / 2, 0, PAGE_W / 2, 2, 'F');

    y = 16;

    // Team name header box
    doc.setFillColor(...CARD);
    doc.roundedRect(MARGIN, y, COL_W, 28, 3, 3, 'F');
    doc.setDrawColor(...GOLD);
    doc.setLineWidth(0.4);
    doc.roundedRect(MARGIN, y, COL_W, 28, 3, 3, 'S');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(18);
    doc.setTextColor(...GOLD);
    doc.text(owner.team.toUpperCase(), MARGIN + 8, y + 11);

    doc.setFontSize(9);
    doc.setTextColor(...DIM);
    doc.text(`Owner: ${owner.name}`, MARGIN + 8, y + 19);

    // summary chips
    const spent = owner.totalPurse - owner.purse;
    const overseasCt = owner.bought.filter(b => b.isOverseas).length;

    const chips = [
      { label: 'PURSE USED',   val: `Rs. ${spent.toFixed(1)} Cr`,      col: GOLD },
      { label: 'REMAINING',    val: `Rs. ${owner.purse.toFixed(1)} Cr`, col: TEAL },
      { label: 'PLAYERS',      val: `${owner.bought.length}`,      col: WHITE },
      { label: 'OVERSEAS',     val: `${overseasCt}`,               col: ORANGE },
    ];

    y += 36;
    const chipW = (COL_W - 9) / 4;
    chips.forEach((c, ci) => {
      const cx = MARGIN + ci * (chipW + 3);
      doc.setFillColor(...CARD);
      doc.roundedRect(cx, y, chipW, 18, 2, 2, 'F');
      doc.setDrawColor(50, 60, 80);
      doc.setLineWidth(0.3);
      doc.roundedRect(cx, y, chipW, 18, 2, 2, 'S');

      doc.setFontSize(6.5);
      doc.setTextColor(...DIM);
      doc.text(c.label, cx + chipW / 2, y + 6, { align: 'center' });

      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(...c.col);
      doc.text(c.val, cx + chipW / 2, y + 14, { align: 'center' });
    });

    y += 26;

    // Table header
    doc.setFillColor(20, 32, 58);
    doc.roundedRect(MARGIN, y, COL_W, 9, 1, 1, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor(...GOLD);
    doc.text('#',              MARGIN + 4,          y + 6);
    doc.text('PLAYER NAME',   MARGIN + 12,          y + 6);
    doc.text('TYPE',          MARGIN + COL_W - 46,  y + 6);
    doc.text('BID PRICE',     MARGIN + COL_W - 18,  y + 6, { align: 'right' });

    y += 11;

    if (!owner.bought.length) {
      doc.setFont('helvetica', 'italic');
      doc.setFontSize(9);
      doc.setTextColor(...DIM);
      doc.text('No players purchased.', MARGIN + 8, y + 6);
      y += 12;
    } else {
      owner.bought.forEach((p, pi) => {
        checkY(10);

        // alternating row bg
        if (pi % 2 === 0) {
          doc.setFillColor(16, 26, 46);
          doc.rect(MARGIN, y, COL_W, 9, 'F');
        }

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8.5);
        doc.setTextColor(...WHITE);
        doc.text(`${pi + 1}`, MARGIN + 4, y + 6);

        // overseas flag before name
        const nameText = (p.isOverseas ? '* ' : '') + p.name;
        doc.text(nameText, MARGIN + 12, y + 6);

        doc.setFontSize(7.5);
        const typeLabel = p.isOverseas ? 'Overseas' : 'Local';
        const typeColor = p.isOverseas ? TEAL : DIM;
        doc.setTextColor(...typeColor);
        doc.text(typeLabel, MARGIN + COL_W - 46, y + 6);

        doc.setFontSize(8.5);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(...GOLD);
        doc.text(`Rs. ${p.price} Cr`, MARGIN + COL_W - 18, y + 6, { align: 'right' });

        y += 9;
      });
    }

    // total spent row
    y += 2;
    doc.setDrawColor(...GOLD);
    doc.setLineWidth(0.3);
    doc.line(MARGIN, y, MARGIN + COL_W, y);
    y += 5;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.5);
    doc.setTextColor(...GOLD);
    doc.text('TOTAL SPENT', MARGIN + 4, y + 4);
    doc.text(`Rs. ${spent.toFixed(1)} Cr`, MARGIN + COL_W - 18, y + 4, { align: 'right' });

    // bottom bar
    doc.setFillColor(...GOLD);
    doc.rect(0, PAGE_H - 2, PAGE_W, 2, 'F');

    // page number
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(...DIM);
    doc.text(`${owner.team} — E-IPL Auction`, PAGE_W / 2, PAGE_H - 5, { align: 'center' });
  });

  doc.save('E-IPL-Auction-Squads.pdf');
}
