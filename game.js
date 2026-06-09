const board = document.getElementById('board');
const ctx = board.getContext('2d');
const nextCanvas = document.getElementById('next');
const nctx = nextCanvas.getContext('2d');
const overlay = document.getElementById('overlay');

const COLS = 10;
const ROWS = 20;
const CELL = 30; // board is 300x600

// Tetromino shapes and colors
const PIECES = [
  { shape: [[1,1,1,1]],                         color: '#00e5ff' }, // I
  { shape: [[1,1],[1,1]],                        color: '#ffd600' }, // O
  { shape: [[0,1,0],[1,1,1]],                    color: '#aa00ff' }, // T
  { shape: [[0,1,1],[1,1,0]],                    color: '#76ff03' }, // S
  { shape: [[1,1,0],[0,1,1]],                    color: '#ff1744' }, // Z
  { shape: [[1,0],[1,0],[1,1]],                  color: '#ff6d00' }, // L
  { shape: [[0,1],[0,1],[1,1]],                  color: '#2979ff' }, // J
];

const SCORE_TABLE = [0, 100, 300, 500, 800]; // 0-4 lines

let grid, piece, nextPiece, score, level, lines, dropTimer, dropInterval, running, gameOver;

function newGrid() {
  return Array.from({ length: ROWS }, () => Array(COLS).fill(null));
}

function randomPiece() {
  const p = PIECES[Math.floor(Math.random() * PIECES.length)];
  return {
    shape: p.shape.map(r => [...r]),
    color: p.color,
    x: Math.floor(COLS / 2) - Math.ceil(p.shape[0].length / 2),
    y: 0,
  };
}

function rotate(shape) {
  const rows = shape.length, cols = shape[0].length;
  return Array.from({ length: cols }, (_, c) =>
    Array.from({ length: rows }, (_, r) => shape[rows - 1 - r][c])
  );
}

function valid(p, dx = 0, dy = 0, shape = p.shape) {
  return shape.every((row, r) =>
    row.every((v, c) => {
      if (!v) return true;
      const nx = p.x + c + dx, ny = p.y + r + dy;
      return nx >= 0 && nx < COLS && ny < ROWS && (ny < 0 || !grid[ny][nx]);
    })
  );
}

function place() {
  piece.shape.forEach((row, r) =>
    row.forEach((v, c) => {
      if (v) grid[piece.y + r][piece.x + c] = piece.color;
    })
  );
  clearLines();
  piece = nextPiece;
  nextPiece = randomPiece();
  drawNext();
  if (!valid(piece)) {
    endGame();
  }
}

function clearLines() {
  let cleared = 0;
  for (let r = ROWS - 1; r >= 0; r--) {
    if (grid[r].every(c => c)) {
      grid.splice(r, 1);
      grid.unshift(Array(COLS).fill(null));
      cleared++;
      r++; // re-check same row
    }
  }
  if (cleared) {
    score += SCORE_TABLE[cleared] * level;
    lines += cleared;
    level = Math.floor(lines / 10) + 1;
    dropInterval = Math.max(100, 1000 - (level - 1) * 100);
    updateUI();
  }
}

function moveDown() {
  if (valid(piece, 0, 1)) {
    piece.y++;
  } else {
    place();
  }
}

function hardDrop() {
  while (valid(piece, 0, 1)) piece.y++;
  place();
}

function moveLeft()  { if (valid(piece, -1, 0)) piece.x--; }
function moveRight() { if (valid(piece,  1, 0)) piece.x++; }
function rotatePiece() {
  const r = rotate(piece.shape);
  // Wall-kick: try center, then nudge left/right
  for (const dx of [0, -1, 1, -2, 2]) {
    if (valid(piece, dx, 0, r)) { piece.shape = r; piece.x += dx; return; }
  }
}

function updateUI() {
  document.getElementById('score').textContent = score;
  document.getElementById('level').textContent = level;
  document.getElementById('lines').textContent = lines;
}

function startGame() {
  grid = newGrid();
  score = 0; level = 1; lines = 0;
  dropInterval = 900;
  dropTimer = 0;
  gameOver = false;
  running = true;
  piece = randomPiece();
  nextPiece = randomPiece();
  overlay.style.display = 'none';
  updateUI();
  drawNext();
  lastTime = performance.now();
  requestAnimationFrame(loop);
}

function endGame() {
  running = false;
  gameOver = true;
  overlay.innerHTML = '<h2>ゲームオーバー</h2><p>スペース / タップで もう一度</p>';
  overlay.style.display = 'flex';
}

// ─── Drawing ─────────────────────────────────────────────────────────────────

function drawCell(context, x, y, color, size = CELL) {
  const pad = 1;
  context.fillStyle = color;
  context.fillRect(x * size + pad, y * size + pad, size - pad * 2, size - pad * 2);
  // highlight
  context.fillStyle = 'rgba(255,255,255,0.25)';
  context.fillRect(x * size + pad, y * size + pad, size - pad * 2, 5);
  context.fillStyle = 'rgba(0,0,0,0.2)';
  context.fillRect(x * size + pad, y * size + size - pad - 5, size - pad * 2, 5);
}

function ghostY() {
  let dy = 0;
  while (valid(piece, 0, dy + 1)) dy++;
  return piece.y + dy;
}

function drawBoard() {
  ctx.fillStyle = '#0d0d1a';
  ctx.fillRect(0, 0, board.width, board.height);

  // grid lines
  ctx.strokeStyle = 'rgba(255,255,255,0.04)';
  ctx.lineWidth = 1;
  for (let r = 0; r <= ROWS; r++) {
    ctx.beginPath(); ctx.moveTo(0, r * CELL); ctx.lineTo(COLS * CELL, r * CELL); ctx.stroke();
  }
  for (let c = 0; c <= COLS; c++) {
    ctx.beginPath(); ctx.moveTo(c * CELL, 0); ctx.lineTo(c * CELL, ROWS * CELL); ctx.stroke();
  }

  // placed cells
  grid.forEach((row, r) =>
    row.forEach((color, c) => { if (color) drawCell(ctx, c, r, color); })
  );

  if (!running) return;

  // ghost piece
  const gy = ghostY();
  if (gy !== piece.y) {
    piece.shape.forEach((row, r) =>
      row.forEach((v, c) => {
        if (v) {
          ctx.fillStyle = 'rgba(255,255,255,0.12)';
          ctx.fillRect((piece.x + c) * CELL + 1, (gy + r) * CELL + 1, CELL - 2, CELL - 2);
        }
      })
    );
  }

  // active piece
  piece.shape.forEach((row, r) =>
    row.forEach((v, c) => { if (v) drawCell(ctx, piece.x + c, piece.y + r, piece.color); })
  );
}

function drawNext() {
  nctx.fillStyle = '#0d0d1a';
  nctx.fillRect(0, 0, nextCanvas.width, nextCanvas.height);
  if (!nextPiece) return;
  const s = nextPiece.shape;
  const size = 22;
  const ox = Math.floor((nextCanvas.width  - s[0].length * size) / 2);
  const oy = Math.floor((nextCanvas.height - s.length    * size) / 2);
  s.forEach((row, r) =>
    row.forEach((v, c) => {
      if (!v) return;
      nctx.fillStyle = nextPiece.color;
      nctx.fillRect(ox + c * size + 1, oy + r * size + 1, size - 2, size - 2);
      nctx.fillStyle = 'rgba(255,255,255,0.25)';
      nctx.fillRect(ox + c * size + 1, oy + r * size + 1, size - 2, 4);
    })
  );
}

// ─── Game Loop ────────────────────────────────────────────────────────────────

let lastTime = 0;

function loop(ts) {
  if (!running) return;
  const dt = ts - lastTime;
  lastTime = ts;
  dropTimer += dt;
  if (dropTimer >= dropInterval) {
    dropTimer = 0;
    moveDown();
  }
  drawBoard();
  requestAnimationFrame(loop);
}

// ─── Keyboard ────────────────────────────────────────────────────────────────

let softDropActive = false;
let softDropTimer = null;

function startSoftDrop() {
  if (softDropActive) return;
  softDropActive = true;
  dropInterval = 80;
}
function stopSoftDrop() {
  if (!softDropActive) return;
  softDropActive = false;
  dropInterval = Math.max(100, 1000 - (level - 1) * 100);
}

window.addEventListener('keydown', e => {
  if (e.code === 'Space') {
    if (!running) { startGame(); return; }
    hardDrop();
  }
  if (!running) return;
  if (e.code === 'ArrowLeft')  { moveLeft();    e.preventDefault(); }
  if (e.code === 'ArrowRight') { moveRight();   e.preventDefault(); }
  if (e.code === 'ArrowUp')    { rotatePiece(); e.preventDefault(); }
  if (e.code === 'ArrowDown')  { startSoftDrop(); e.preventDefault(); }
  if (e.code === 'KeyR')       { if (gameOver) startGame(); }
});
window.addEventListener('keyup', e => {
  if (e.code === 'ArrowDown') stopSoftDrop();
});

// ─── Touch buttons ────────────────────────────────────────────────────────────

function touchBtn(id, onDown, onUp) {
  const el = document.getElementById(id);
  if (!el) return;
  const press = e => { e.preventDefault(); el.classList.add('pressed'); onDown && onDown(); };
  const release = e => { e.preventDefault(); el.classList.remove('pressed'); onUp && onUp(); };
  el.addEventListener('touchstart', press, { passive: false });
  el.addEventListener('touchend',   release, { passive: false });
  el.addEventListener('touchcancel',release, { passive: false });
  el.addEventListener('mousedown',  press);
  el.addEventListener('mouseup',    release);
  el.addEventListener('mouseleave', release);
}

touchBtn('btn-left',   () => running && moveLeft());
touchBtn('btn-right',  () => running && moveRight());
touchBtn('btn-rotate', () => running && rotatePiece());
touchBtn('btn-down',   () => running && startSoftDrop(), () => stopSoftDrop());
touchBtn('btn-drop',   () => { if (!running) { startGame(); return; } hardDrop(); });

// tap overlay to start
overlay.addEventListener('click', () => { if (!running) startGame(); });

// initial draw
drawBoard();
