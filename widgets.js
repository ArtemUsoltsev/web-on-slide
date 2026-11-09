// Інструменти тренера, що працюють без інтернету і без сторонніх сайтів:
// таймер, випадковий вибір учасника, рахунок команд.
// Налаштування (тривалість, списки) зберігаються у файлі; поточний стан (час, бали, хто вже вибраний)
// живе лише в пам'яті й починається заново з кожним показом.
(function () {
  "use strict";
  function L(uk, en) { return window.WOS_I18N ? window.WOS_I18N.L(uk, en) : uk; }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function btn(cls, text, onClick) {
    var b = el("button", "btn " + cls, text);
    b.type = "button";
    b.addEventListener("click", onClick);
    return b;
  }
  function lines(text, max) {
    var out = [], arr = String(text || "").split(/\r?\n/);
    for (var i = 0; i < arr.length && out.length < max; i++) {
      var t = arr[i].trim();
      if (t) out.push(t.slice(0, 60));
    }
    return out;
  }

  // Короткий сигнал наприкінці таймера. Звук можливий, бо таймер запускають кліком.
  var audioCtx = null;
  function beep() {
    try {
      var AC = window.AudioContext || window.webkitAudioContext;
      if (!AC) return;
      audioCtx = audioCtx || new AC();
      [0, 0.35, 0.7].forEach(function (t) {
        var o = audioCtx.createOscillator(), g = audioCtx.createGain();
        o.frequency.value = 880;
        g.gain.setValueAtTime(0.0001, audioCtx.currentTime + t);
        g.gain.exponentialRampToValueAtTime(0.3, audioCtx.currentTime + t + 0.02);
        g.gain.exponentialRampToValueAtTime(0.0001, audioCtx.currentTime + t + 0.25);
        o.connect(g); g.connect(audioCtx.destination);
        o.start(audioCtx.currentTime + t); o.stop(audioCtx.currentTime + t + 0.3);
      });
    } catch (e) { /* без звуку */ }
  }
  function unlockAudio() {
    try {
      var AC = window.AudioContext || window.webkitAudioContext;
      if (AC && !audioCtx) audioCtx = new AC();
      if (audioCtx && audioCtx.state === "suspended") audioCtx.resume();
    } catch (e) { /* без звуку */ }
  }

  function fmt(sec) {
    sec = Math.max(0, Math.ceil(sec));
    var m = Math.floor(sec / 60), s = sec % 60;
    return m + ":" + (s < 10 ? "0" : "") + s;
  }

  // ---------- Таймер ----------
  function timer(root, w) {
    var total = Math.round(Math.min(Math.max(parseFloat(w.minutes) || 5, 0.25), 180) * 60);
    var left = total, running = false, endAt = 0, tick = null;
    root.classList.add("w-timer");
    var title = w.title ? el("div", "w-title", w.title) : null;
    var face = el("div", "w-time", fmt(left));
    face.setAttribute("role", "timer");
    var startB = btn("primary", L("Старт", "Start"), toggle);
    var resetB = btn("", L("Скинути", "Reset"), reset);
    var plusB = btn("", "+1 " + L("хв", "min"), function () { left += 60; if (running) endAt += 60000; draw(); });
    var row = el("div", "w-actions");
    row.appendChild(startB); row.appendChild(plusB); row.appendChild(resetB);
    if (title) root.appendChild(title);
    root.appendChild(face); root.appendChild(row);

    function draw() {
      face.textContent = fmt(left);
      root.classList.toggle("w-warn", left > 0 && left <= 10);
      root.classList.toggle("w-done", left <= 0);
      startB.textContent = running ? L("Пауза", "Pause") : (left < total && left > 0 ? L("Далі", "Resume") : L("Старт", "Start"));
    }
    function loop() {
      left = (endAt - Date.now()) / 1000;
      if (left <= 0) { left = 0; running = false; clearInterval(tick); tick = null; beep(); }
      draw();
    }
    function toggle() {
      unlockAudio();
      if (running) {
        left = Math.max(0, (endAt - Date.now()) / 1000);
        running = false; clearInterval(tick); tick = null;
        draw();
        return;
      }
      if (left <= 0) left = total;
      running = true; endAt = Date.now() + left * 1000;
      tick = setInterval(loop, 200);
      draw();
    }
    function reset() { running = false; if (tick) clearInterval(tick); tick = null; left = total; draw(); }
    draw();
    return function destroy() { if (tick) clearInterval(tick); };
  }

  // ---------- Випадковий вибір ----------
  function picker(root, w) {
    var all = lines(w.names, 200), pool = all.slice(), spin = null;
    root.classList.add("w-picker");
    if (w.title) root.appendChild(el("div", "w-title", w.title));
    var face = el("div", "w-pick", all.length ? "?" : L("Додайте імена у налаштуваннях", "Add names in the settings"));
    face.setAttribute("aria-live", "polite");
    var info = el("div", "w-note", "");
    var go = btn("primary big", L("Вибрати", "Pick"), pick);
    var again = btn("", L("Почати знову", "Start over"), function () { pool = all.slice(); face.textContent = "?"; draw(); });
    var row = el("div", "w-actions");
    row.appendChild(go); row.appendChild(again);
    root.appendChild(face); root.appendChild(row); root.appendChild(info);

    function draw() {
      go.disabled = !pool.length || !!spin;
      info.textContent = all.length ? (w.repeat ? L("Учасників: ", "Names: ") + all.length
        : L("Залишилось ", "Left: ") + pool.length + L(" з ", " of ") + all.length) : "";
    }
    function pick() {
      if (!pool.length || spin) return;
      var steps = 14, n = 0;
      root.classList.remove("w-picked");
      spin = setInterval(function () {
        face.textContent = all[Math.floor(Math.random() * all.length)];
        if (++n >= steps) {
          clearInterval(spin); spin = null;
          var i = Math.floor(Math.random() * pool.length);
          face.textContent = pool[i];
          if (!w.repeat) pool.splice(i, 1);
          root.classList.add("w-picked");
          draw();
        }
      }, 70);
      draw();
    }
    draw();
    return function destroy() { if (spin) clearInterval(spin); };
  }

  // ---------- Рахунок команд ----------
  function score(root, w) {
    var teams = lines(w.teams, 6);
    if (teams.length < 2) teams = [L("Команда 1", "Team 1"), L("Команда 2", "Team 2")];
    var step = Math.max(1, Math.min(100, parseInt(w.step, 10) || 1));
    var pts = teams.map(function () { return 0; });
    root.classList.add("w-score");
    if (w.title) root.appendChild(el("div", "w-title", w.title));
    var grid = el("div", "w-teams");
    grid.style.setProperty("--n", teams.length);
    var nums = [];
    teams.forEach(function (name, i) {
      var card = el("div", "w-team");
      card.appendChild(el("div", "w-team-name", name));
      var num = el("div", "w-team-pts", "0");
      num.setAttribute("aria-live", "polite");
      nums.push(num);
      card.appendChild(num);
      var r = el("div", "w-actions");
      r.appendChild(btn("", "−" + step, function () { pts[i] -= step; draw(); }));
      r.appendChild(btn("primary", "+" + step, function () { pts[i] += step; draw(); }));
      card.appendChild(r);
      grid.appendChild(card);
    });
    root.appendChild(grid);
    var reset = btn("ghost", L("Обнулити", "Reset"), function () { pts = pts.map(function () { return 0; }); draw(); });
    root.appendChild(reset);
    function draw() {
      var max = Math.max.apply(null, pts);
      nums.forEach(function (n, i) {
        n.textContent = pts[i];
        n.parentNode.classList.toggle("w-lead", pts[i] === max && max > 0);
      });
    }
    return function destroy() {};
  }

  var KINDS = { timer: timer, picker: picker, score: score };
  var destroyFn = null, mountedKey = null;

  window.WOS_WIDGETS = {
    kinds: ["timer", "picker", "score"],
    // Монтує віджет у контейнер; повторний виклик з тими ж налаштуваннями зберігає стан.
    mount: function (container, kind, w) {
      var key = kind + "|" + JSON.stringify(w || {});
      if (key === mountedKey) return;
      this.unmount(container);
      var root = el("div", "w-root");
      container.appendChild(root);
      destroyFn = (KINDS[kind] || timer)(root, w || {});
      mountedKey = key;
    },
    unmount: function (container) {
      if (destroyFn) destroyFn();
      destroyFn = null; mountedKey = null;
      while (container.firstChild) container.removeChild(container.firstChild);
    }
  };
})();
