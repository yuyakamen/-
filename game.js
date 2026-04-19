const canvas = document.getElementById('gameCanvas');
const ctx = canvas.getContext('2d');
const W = canvas.width;
const H = canvas.height;

const GRAVITY = 0.55;
const GROUND_Y = H - 60;
const TILE = 40;
const PLAYER_W = 32;
const PLAYER_H = 40;
const ENEMY_W = 36;
const ENEMY_H = 32;
const COIN_R = 10;

let score, coins, lives, cameraX, frameCount, gameOver, win;

function initGame() {
  score = 0;
  coins = 0;
  lives = 3;
  cameraX = 0;
  frameCount = 0;
  gameOver = false;
  win = false;
  document.getElementById('message').textContent = '';
  buildLevel();
  spawnPlayer();
  updateUI();
}

let platforms, enemies, coinList, goal;

function buildLevel() {
  platforms = [];
  enemies = [];
  coinList = [];

  for (let x = 0; x < 4800; x += TILE) {
    platforms.push({ x, y: GROUND_Y, w: TILE, h: 60 });
  }

  const layout = [
    [200, GROUND_Y - 120, 4],
    [420, GROUND_Y - 200, 3],
    [640, GROUND_Y - 140, 3],
    [840, GROUND_Y - 220, 4],
    [1100, GROUND_Y - 160, 3],
    [1300, GROUND_Y - 240, 4],
    [1560, GROUND_Y - 180, 3],
    [1760, GROUND_Y - 120, 5],
    [2000, GROUND_Y - 200, 3],
    [2200, GROUND_Y - 280, 3],
    [2420, GROUND_Y - 160, 4],
    [2700, GROUND_Y - 220, 3],
    [2940, GROUND_Y - 140, 5],
    [3200, GROUND_Y - 260, 3],
    [3440, GROUND_Y - 180, 4],
    [3700, GROUND_Y - 240, 3],
    [3900, GROUND_Y - 140, 4],
    [4100, GROUND_Y - 200, 3],
    [4300, GROUND_Y - 300, 3],
  ];
  layout.forEach(([px, py, tiles]) => {
    for (let i = 0; i < tiles; i++) {
      platforms.push({ x: px + i * TILE, y: py, w: TILE, h: TILE, top: true });
    }
    const cx = px + Math.floor(tiles / 2) * TILE + TILE / 2;
    coinList.push({ x: cx, y: py - 30, r: COIN_R, collected: false });
  });

  [150, 300, 500, 700, 950, 1200, 1450, 1680, 2100, 2350, 2600, 2850, 3050, 3350, 3600, 3800, 4050].forEach(cx => {
    coinList.push({ x: cx, y: GROUND_Y - 30, r: COIN_R, collected: false });
  });

  [350, 650, 950, 1350, 1750, 2050, 2450, 2800, 3100, 3500, 3900, 4200].forEach(ex => {
    enemies.push({ x: ex, y: GROUND_Y - ENEMY_H, w: ENEMY_W, h: ENEMY_H, vx: -1.2, vy: 0, alive: true });
  });

  goal = { x: 4550, y: GROUND_Y - 180, w: 20, h: 180 };
}

let player;

function spawnPlayer() {
  player = {
    x: 80, y: GROUND_Y - PLAYER_H,
    w: PLAYER_W, h: PLAYER_H,
    vx: 0, vy: 0,
    onGround: false,
    dead: false,
    respawnTimer: 0,
    facing: 1,
    walkFrame: 0,
  };
}

const keys = {};

window.addEventListener('keydown', e => {
  keys[e.code] = true;
  if (e.code === 'KeyR') initGame();
  if (['Space', 'ArrowUp', 'ArrowLeft', 'ArrowRight'].includes(e.code)) e.preventDefault();
});
window.addEventListener('keyup', e => { keys[e.code] = false; });

function setupTouchBtn(id, keyCode) {
  const el = document.getElementById(id);
  if (!el) return;
  const press = e => {
    e.preventDefault();
    keys[keyCode] = true;
    el.classList.add('pressed');
  };
  const release = e => {
    e.preventDefault();
    keys[keyCode] = false;
    el.classList.remove('pressed');
    if (keyCode === 'Space' && (gameOver || win)) initGame();
  };
  el.addEventListener('touchstart', press, { passive: false });
  el.addEventListener('touchend', release, { passive: false });
  el.addEventListener('touchcancel', release, { passive: false });
  el.addEventListener('mousedown', press);
  el.addEventListener('mouseup', release);
  el.addEventListener('mouseleave', release);
}

setupTouchBtn('btn-left', 'ArrowLeft');
setupTouchBtn('btn-right', 'ArrowRight');
setupTouchBtn('btn-jump', 'Space');

function rectsOverlap(a, b) {
  return a.x < b.x + b.w && a.x + a.w > b.x &&
         a.y < b.y + b.h && a.y + a.h > b.y;
}

function resolvePlayerPlatform(p, plat) {
  const prevBottom = p.y + p.h - p.vy;
  const curBottom = p.y + p.h;
  if (p.vy >= 0 && prevBottom <= plat.y + 2 && curBottom >= plat.y) {
    if (p.x + p.w > plat.x + 4 && p.x < plat.x + plat.w - 4) {
      p.y = plat.y - p.h;
      p.vy = 0;
      p.onGround = true;
      return;
    }
  }
  if (p.vy < 0) {
    if (p.y - p.vy >= plat.y + plat.h - 2) p.vy = 0;
  }
  if (p.vy >= 0) {
    const px = p.x + p.vx;
    if (px + p.w > plat.x && px < plat.x + plat.w &&
        p.y + p.h > plat.y + 4 && p.y < plat.y + plat.h) {
      p.vx = 0;
    }
  }
}

function update() {
  if (gameOver || win) return;
  frameCount++;

  if (player.dead) {
    player.respawnTimer--;
    if (player.respawnTimer <= 0) {
      if (lives <= 0) {
        gameOver = true;
        document.getElementById('message').textContent = 'ゲームオーバー！ もう一度？';
        return;
      }
      spawnPlayer();
    }
    return;
  }

  const SPEED = 4.5;
  const JUMP = -13;
  player.onGround = false;

  if (keys['ArrowLeft']) { player.vx = -SPEED; player.facing = -1; }
  else if (keys['ArrowRight']) { player.vx = SPEED; player.facing = 1; }
  else { player.vx *= 0.8; if (Math.abs(player.vx) < 0.2) player.vx = 0; }

  if ((keys['Space'] || keys['ArrowUp']) && player.onGround) {
    player.vy = JUMP;
  }

  player.vy += GRAVITY;
  player.x += player.vx;
  player.y += player.vy;

  if (player.x < 0) { player.x = 0; player.vx = 0; }

  platforms.forEach(plat => {
    if (rectsOverlap(player, plat)) resolvePlayerPlatform(player, plat);
  });

  if (player.y > H + 50) { playerDie(); return; }

  if (player.onGround && Math.abs(player.vx) > 0.3) player.walkFrame += 0.2;

  cameraX = Math.max(0, player.x - W / 3);

  coinList.forEach(c => {
    if (c.collected) return;
    const dx = player.x + player.w / 2 - c.x;
    const dy = player.y + player.h / 2 - c.y;
    if (Math.hypot(dx, dy) < player.w / 2 + c.r) {
      c.collected = true;
      coins++;
      score += 100;
      updateUI();
    }
  });

  enemies.forEach(e => {
    if (!e.alive) return;
    e.vy += GRAVITY;
    e.x += e.vx;
    e.y += e.vy;

    e.onGround = false;
    platforms.forEach(plat => {
      if (rectsOverlap(e, plat)) {
        const prevBottom = e.y + e.h - e.vy;
        if (prevBottom <= plat.y + 2 && e.vy >= 0) {
          e.y = plat.y - e.h;
          e.vy = 0;
          e.onGround = true;
        }
      }
    });

    if (e.onGround) {
      const frontX = e.vx > 0 ? e.x + e.w : e.x;
      const gc = { x: frontX - 2, y: e.y + e.h + 2, w: 4, h: 4 };
      if (!platforms.some(p => rectsOverlap(gc, p))) e.vx *= -1;
    }
    if (platforms.some(p => rectsOverlap(e, p) && !(e.y + e.h <= p.y + 4))) e.vx *= -1;
    if (e.y > H + 100) { e.alive = false; return; }

    if (!rectsOverlap(player, e)) return;
    if (player.y + player.h - player.vy <= e.y + 8 && player.vy > 0) {
      e.alive = false;
      player.vy = -9;
      score += 200;
      updateUI();
    } else {
      playerDie();
    }
  });

  if (rectsOverlap(player, goal)) {
    win = true;
    score += 1000;
    updateUI();
    document.getElementById('message').textContent = 'クリア！ おめでとう！ 🎉';
  }
}

function playerDie() {
  player.dead = true;
  player.respawnTimer = 90;
  lives--;
  updateUI();
  if (lives <= 0) document.getElementById('message').textContent = 'ゲームオーバー！ もう一度？';
}

function updateUI() {
  document.getElementById('score').textContent = score;
  document.getElementById('coins').textContent = coins;
  document.getElementById('lives').textContent = Math.max(0, lives);
}

function draw() {
  const sky = ctx.createLinearGradient(0, 0, 0, H);
  sky.addColorStop(0, '#64b5f6');
  sky.addColorStop(1, '#bbdefb');
  ctx.fillStyle = sky;
  ctx.fillRect(0, 0, W, H);

  drawClouds();

  ctx.save();
  ctx.translate(-cameraX, 0);
  platforms.forEach(drawPlatform);
  coinList.forEach(c => { if (!c.collected) drawCoin(c); });
  drawGoal();
  enemies.forEach(e => { if (e.alive) drawEnemy(e); });
  if (!player.dead) drawPlayer();
  ctx.restore();
}

function drawClouds() {
  ctx.fillStyle = 'rgba(255,255,255,0.9)';
  [100, 300, 550, 720, 200, 450].forEach((cx, i) => {
    const ox = ((cx - cameraX * 0.3) % (W + 200) + W + 200) % (W + 200) - 100;
    const cy = 60 + (i % 3) * 30;
    ctx.beginPath();
    ctx.arc(ox, cy, 25, 0, Math.PI * 2);
    ctx.arc(ox + 30, cy - 10, 30, 0, Math.PI * 2);
    ctx.arc(ox + 60, cy, 25, 0, Math.PI * 2);
    ctx.fill();
  });
}

function drawPlatform(p) {
  if (p.top) {
    ctx.fillStyle = '#4caf50';
    ctx.fillRect(p.x, p.y, p.w, 12);
    ctx.fillStyle = '#388e3c';
    ctx.fillRect(p.x, p.y + 12, p.w, p.h - 12);
    ctx.strokeStyle = '#2e7d32';
  } else {
    ctx.fillStyle = '#8d6e63';
    ctx.fillRect(p.x, p.y, p.w, p.h);
    ctx.fillStyle = '#5d4037';
    ctx.fillRect(p.x, p.y + 12, p.w, p.h - 12);
    ctx.strokeStyle = '#4e342e';
  }
  ctx.lineWidth = 1;
  ctx.strokeRect(p.x, p.y, p.w, p.h);
}

function drawCoin(c) {
  const bob = Math.sin(frameCount * 0.1 + c.x * 0.05) * 3;
  ctx.fillStyle = '#ffd600';
  ctx.strokeStyle = '#f9a825';
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.arc(c.x, c.y + bob, c.r, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  ctx.fillStyle = '#fff9c4';
  ctx.beginPath();
  ctx.arc(c.x - 3, c.y + bob - 3, 3, 0, Math.PI * 2);
  ctx.fill();
}

function drawGoal() {
  ctx.fillStyle = '#bdbdbd';
  ctx.fillRect(goal.x, goal.y, 6, goal.h);
  ctx.fillStyle = '#e53935';
  ctx.beginPath();
  ctx.moveTo(goal.x + 6, goal.y);
  ctx.lineTo(goal.x + 46, goal.y + 20);
  ctx.lineTo(goal.x + 6, goal.y + 40);
  ctx.fill();
  ctx.fillStyle = '#757575';
  ctx.fillRect(goal.x - 10, goal.y + goal.h - 8, 26, 8);
}

function drawPlayer() {
  const px = player.x;
  const py = player.y;

  ctx.save();
  if (player.facing === -1) {
    ctx.translate(px + player.w, py);
    ctx.scale(-1, 1);
    ctx.translate(0, -py);
  }

  ctx.fillStyle = '#e53935';
  ctx.fillRect(px + 4, py + 16, 24, 18);
  ctx.fillStyle = '#1565c0';
  ctx.fillRect(px + 2, py + 22, 28, 12);
  ctx.fillStyle = '#ffcc80';
  ctx.fillRect(px + 6, py + 2, 20, 16);
  ctx.fillStyle = '#e53935';
  ctx.fillRect(px + 4, py, 24, 8);
  ctx.fillRect(px + 2, py + 6, 28, 4);
  ctx.fillStyle = '#212121';
  ctx.fillRect(px + 18, py + 7, 4, 4);
  ctx.fillStyle = '#4e342e';
  ctx.fillRect(px + 12, py + 14, 12, 3);
  ctx.fillStyle = '#1565c0';
  const leg = player.onGround ? Math.sin(player.walkFrame) * 4 : 0;
  ctx.fillRect(px + 5, py + 32, 10, 10 + leg);
  ctx.fillRect(px + 17, py + 32, 10, 10 - leg);
  ctx.fillStyle = '#4e342e';
  ctx.fillRect(px + 3, py + 38 + leg, 13, 4);
  ctx.fillRect(px + 15, py + 38 - leg, 13, 4);

  ctx.restore();
}

function drawEnemy(e) {
  const wobble = Math.sin(frameCount * 0.15 + e.x * 0.1) * 2;
  const ex = e.x, ey = e.y + wobble;

  ctx.fillStyle = '#6d4c41';
  ctx.beginPath();
  ctx.ellipse(ex + e.w / 2, ey + e.h * 0.7, e.w / 2, e.h * 0.35, 0, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = '#8b0000';
  ctx.beginPath();
  ctx.ellipse(ex + e.w / 2, ey + e.h * 0.4, e.w / 2 + 4, e.h * 0.45, 0, Math.PI, 0);
  ctx.fill();
  ctx.fillStyle = '#ffccbc';
  ctx.beginPath();
  ctx.arc(ex + e.w / 2 - 8, ey + e.h * 0.25, 4, 0, Math.PI * 2);
  ctx.arc(ex + e.w / 2 + 8, ey + e.h * 0.2, 3, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = '#fff';
  ctx.fillRect(ex + 7, ey + e.h * 0.55, 7, 7);
  ctx.fillRect(ex + e.w - 14, ey + e.h * 0.55, 7, 7);
  ctx.fillStyle = '#000';
  ctx.fillRect(ex + 8, ey + e.h * 0.58, 4, 4);
  ctx.fillRect(ex + e.w - 13, ey + e.h * 0.58, 4, 4);
  ctx.save();
  ctx.translate(ex + 10, ey + e.h * 0.52); ctx.rotate(0.3);
  ctx.fillRect(-4, -2, 8, 2);
  ctx.restore();
  ctx.save();
  ctx.translate(ex + e.w - 10, ey + e.h * 0.52); ctx.rotate(-0.3);
  ctx.fillRect(-4, -2, 8, 2);
  ctx.restore();
}

function drawHUD() {
  if (gameOver || win) {
    ctx.fillStyle = 'rgba(0,0,0,0.5)';
    ctx.fillRect(0, 0, W, H);
    ctx.textAlign = 'center';
    if (win) {
      ctx.fillStyle = '#ffd600';
      ctx.font = 'bold 52px Courier New';
      ctx.fillText('クリア！', W / 2, H / 2 - 30);
      ctx.fillStyle = '#fff';
      ctx.font = '24px Courier New';
      ctx.fillText('スコア: ' + score, W / 2, H / 2 + 20);
      ctx.fillText('ジャンプボタンで再開', W / 2, H / 2 + 60);
    } else {
      ctx.fillStyle = '#fff';
      ctx.font = 'bold 44px Courier New';
      ctx.fillText('ゲームオーバー', W / 2, H / 2 - 20);
      ctx.font = '22px Courier New';
      ctx.fillText('ジャンプボタンで再開', W / 2, H / 2 + 30);
    }
    ctx.textAlign = 'left';
  }
}

function loop() {
  update();
  draw();
  drawHUD();
  requestAnimationFrame(loop);
}

initGame();
loop();
