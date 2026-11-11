// Інструменти тренера, що працюють без інтернету і без сторонніх сайтів.
// Налаштування (тривалість, списки) зберігаються у файлі; поточний стан (час, бали, хто вже вибраний)
// живе в пам'яті й починається заново з кожним показом. Виняток: спільний рахунок команд
// (w.shared) живе в localStorage цього комп'ютера 12 годин після останньої зміни
// і синхронізується між блоками через BroadcastChannel.
//
// Кожен віджет повертає { destroy, key(e) -> true, якщо клавішу оброблено, results() -> текст для звіту }.
(function () {
  "use strict";
  function L(uk, en) { return window.WOS_I18N ? window.WOS_I18N.L(uk, en) : uk; }

  // ---------- Допоміжне ----------
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
  function row() {
    var r = el("div", "w-actions");
    for (var i = 0; i < arguments.length; i++) if (arguments[i]) r.appendChild(arguments[i]);
    return r;
  }
  function lines(text, max) {
    var out = [], arr = String(text || "").split(/\r?\n/);
    for (var i = 0; i < arr.length && out.length < max; i++) {
      var t = arr[i].trim();
      if (t) out.push(t.slice(0, 120));
    }
    return out;
  }
  function shuffle(a) {
    a = a.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1)), t = a[i];
      a[i] = a[j]; a[j] = t;
    }
    return a;
  }
  function title(root, w, fallback) {
    var t = w.title || fallback;
    if (t) root.appendChild(el("div", "w-title", t));
  }
  function fmt(sec) {
    sec = Math.max(0, Math.ceil(sec));
    var h = Math.floor(sec / 3600), m = Math.floor(sec % 3600 / 60), s = sec % 60;
    return (h ? h + ":" + (m < 10 ? "0" : "") : "") + m + ":" + (s < 10 ? "0" : "") + s;
  }
  function isPrimaryKey(e) { return e.key === " " || e.key === "Enter"; }

  // Кнопка «Скопіювати результати»: буфер обміну, а якщо не вийшло, показуємо текст для ручного копіювання.
  function copyButton(root, getText) {
    var note = el("div", "w-note w-copied");
    var b = btn("ghost", L("Скопіювати результати", "Copy results"), function () {
      var text = getText();
      function fallback() { note.textContent = text; note.style.userSelect = "all"; }
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () {
          b.textContent = L("Скопійовано", "Copied");
          setTimeout(function () { b.textContent = L("Скопіювати результати", "Copy results"); }, 1500);
        }, fallback);
      } else fallback();
    });
    var r = row(b); r.classList.add("w-tools");
    root.appendChild(r);
    root.appendChild(note);
  }

  // Короткий сигнал. Звук можливий, бо віджети запускають кліком.
  var audioCtx = null;
  function unlockAudio() {
    try {
      var AC = window.AudioContext || window.webkitAudioContext;
      if (AC && !audioCtx) audioCtx = new AC();
      if (audioCtx && audioCtx.state === "suspended") audioCtx.resume();
    } catch (e) { /* без звуку */ }
  }
  function beep(times) {
    try {
      unlockAudio();
      if (!audioCtx) return;
      for (var k = 0; k < (times || 3); k++) {
        var t = k * 0.35, o = audioCtx.createOscillator(), g = audioCtx.createGain();
        o.frequency.value = 880;
        g.gain.setValueAtTime(0.0001, audioCtx.currentTime + t);
        g.gain.exponentialRampToValueAtTime(0.3, audioCtx.currentTime + t + 0.02);
        g.gain.exponentialRampToValueAtTime(0.0001, audioCtx.currentTime + t + 0.25);
        o.connect(g); g.connect(audioCtx.destination);
        o.start(audioCtx.currentTime + t); o.stop(audioCtx.currentTime + t + 0.3);
      }
    } catch (e) { /* без звуку */ }
  }

  // Зворотний відлік: { start, pause, toggle, reset(total), add(sec), left(), running() }
  function countdown(total, onTick, onDone) {
    var left = total, running = false, endAt = 0, tick = null;
    function stop() { running = false; if (tick) clearInterval(tick); tick = null; }
    function loop() {
      left = (endAt - Date.now()) / 1000;
      if (left <= 0) { left = 0; stop(); onTick(); onDone(); return; }
      onTick();
    }
    var api = {
      start: function () {
        unlockAudio();
        if (running) return;
        if (left <= 0) left = total;
        running = true; endAt = Date.now() + left * 1000;
        tick = setInterval(loop, 200);
        onTick();
      },
      pause: function () {
        if (!running) return;
        left = Math.max(0, (endAt - Date.now()) / 1000);
        stop(); onTick();
      },
      toggle: function () { if (running) api.pause(); else api.start(); },
      reset: function (t) { stop(); if (t != null) total = t; left = total; onTick(); },
      add: function (sec) { left += sec; if (running) endAt += sec * 1000; onTick(); },
      left: function () { return left; },
      total: function () { return total; },
      running: function () { return running; },
      destroy: stop
    };
    return api;
  }

  // ---------- Таймер ----------
  function timer(root, w) {
    var total = Math.round(Math.min(Math.max(parseFloat(w.minutes) || 5, 0.25), 180) * 60);
    root.classList.add("w-timer");
    title(root, w);
    var face = el("div", "w-time");
    face.setAttribute("role", "timer");
    var cd = countdown(total, draw, function () { beep(3); });
    var startB = btn("primary", "", cd.toggle);
    var plusB = btn("", "+1 " + L("хв", "min"), function () { cd.add(60); });
    var resetB = btn("", L("Скинути", "Reset"), function () { cd.reset(); });
    root.appendChild(face); root.appendChild(row(startB, plusB, resetB));
    function draw() {
      var left = cd.left();
      face.textContent = fmt(left);
      root.classList.toggle("w-warn", left > 0 && left <= 10);
      root.classList.toggle("w-done", left <= 0);
      startB.textContent = cd.running() ? L("Пауза", "Pause") : (left < cd.total() && left > 0 ? L("Далі", "Resume") : L("Старт", "Start"));
    }
    draw();
    return {
      destroy: cd.destroy,
      key: function (e) {
        if (isPrimaryKey(e)) { cd.toggle(); return true; }
        if (e.key === "+" || e.key === "=") { cd.add(60); return true; }
        if (e.key === "r" || e.key === "R" || e.key === "к" || e.key === "К") { cd.reset(); return true; }
        return false;
      }
    };
  }

  // ---------- Перерва до певного часу ----------
  function breakUntil(root, w) {
    var m = /^(\d{1,2}):(\d{2})$/.exec(w.until || "");
    var hh = m ? +m[1] : 14, mm = m ? +m[2] : 30;
    root.classList.add("w-timer", "w-break");
    title(root, w, L("Перерва", "Break"));
    var back = el("div", "w-sub", L("Повертаємось о ", "Back at ") + hh + ":" + (mm < 10 ? "0" : "") + mm);
    var face = el("div", "w-time");
    face.setAttribute("role", "timer");
    var now = el("div", "w-note");
    root.appendChild(back); root.appendChild(face); root.appendChild(now);
    var beeped = false;
    function target() {
      var d = new Date(); d.setHours(hh, mm, 0, 0);
      // Якщо час уже минув більш ніж на 6 годин, це, найімовірніше, наступний день.
      if (d.getTime() < Date.now() - 6 * 3600 * 1000) d.setDate(d.getDate() + 1);
      return d.getTime();
    }
    function draw() {
      var left = (target() - Date.now()) / 1000;
      var d = new Date();
      now.textContent = L("Зараз ", "Now ") + d.getHours() + ":" + (d.getMinutes() < 10 ? "0" : "") + d.getMinutes();
      if (left <= 0) {
        face.textContent = L("Час повертатися!", "Time to come back!");
        face.classList.add("w-time-text");
        root.classList.add("w-done"); root.classList.remove("w-warn");
        if (!beeped && left > -5) { beeped = true; beep(3); }
      } else {
        face.textContent = fmt(left);
        face.classList.remove("w-time-text");
        root.classList.remove("w-done");
        root.classList.toggle("w-warn", left <= 60);
      }
    }
    draw();
    var tick = setInterval(draw, 500);
    root.addEventListener("click", unlockAudio);
    return { destroy: function () { clearInterval(tick); }, key: function () { return false; } };
  }

  // ---------- Випадковий вибір (списком або колесом) ----------
  function picker(root, w) {
    var all = lines(w.names, 200), pool = all.slice(), busy = false, history = [];
    var wheel = w.style === "wheel" && all.length >= 2 && all.length <= 40;
    root.classList.add("w-picker");
    title(root, w);
    var face = el("div", "w-pick", all.length ? "?" : L("Додайте імена у налаштуваннях", "Add names in the settings"));
    face.setAttribute("aria-live", "polite");
    var canvas = null, angle = 0;
    if (wheel) {
      root.classList.add("w-wheel");
      var wrap = el("div", "w-wheel-wrap");
      canvas = el("canvas");
      canvas.width = canvas.height = 600;
      wrap.appendChild(canvas);
      wrap.appendChild(el("div", "w-wheel-pin"));
      root.appendChild(wrap);
    }
    var info = el("div", "w-note", "");
    var go = btn("primary big", wheel ? L("Крутити", "Spin") : L("Вибрати", "Pick"), pick);
    var again = btn("", L("Почати знову", "Start over"), function () { pool = all.slice(); history = []; face.textContent = "?"; root.classList.remove("w-picked"); drawWheel(); draw(); });
    root.appendChild(face); root.appendChild(row(go, again)); root.appendChild(info);
    copyButton(root, function () { return (w.title ? w.title + "\n" : "") + history.map(function (n, i) { return (i + 1) + ". " + n; }).join("\n"); });

    var COLORS = ["#FFD500", "#1d1f23", "#f4f5f7", "#ffb000", "#6b7079", "#fff1a8"];
    var INK = ["#111", "#fff", "#111", "#111", "#fff", "#111"];
    function drawWheel() {
      if (!canvas) return;
      var ctx = canvas.getContext("2d"), n = pool.length || all.length, list = pool.length ? pool : all, r = 290, c = 300;
      ctx.clearRect(0, 0, 600, 600);
      for (var i = 0; i < n; i++) {
        var a0 = angle + i * 2 * Math.PI / n, a1 = a0 + 2 * Math.PI / n, ci = i % COLORS.length;
        if (n % COLORS.length === 1 && i === n - 1) ci = 2; // сусідні сектори не однакові
        ctx.beginPath(); ctx.moveTo(c, c); ctx.arc(c, c, r, a0, a1); ctx.closePath();
        ctx.fillStyle = COLORS[ci]; ctx.fill();
        ctx.strokeStyle = "#dfe2e7"; ctx.lineWidth = 2; ctx.stroke();
        var mid = (a0 + a1) / 2, norm = ((mid % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
        // На лівій половині колеса повертаємо підпис на 180°, щоб він не був догори ногами.
        var flip = norm > Math.PI / 2 && norm < 3 * Math.PI / 2;
        ctx.save(); ctx.translate(c, c); ctx.rotate(flip ? mid + Math.PI : mid);
        ctx.fillStyle = INK[ci]; ctx.font = "600 " + (n > 20 ? 18 : 26) + "px Segoe UI, system-ui, sans-serif";
        ctx.textAlign = flip ? "left" : "right"; ctx.textBaseline = "middle";
        var label = list[i].length > 18 ? list[i].slice(0, 17) + "…" : list[i];
        ctx.fillText(label, flip ? -(r - 16) : r - 16, 0); ctx.restore();
      }
    }
    function draw() {
      go.disabled = !pool.length || busy;
      info.textContent = all.length ? (w.repeat ? L("Учасників: ", "Names: ") + all.length
        : L("Залишилось ", "Left: ") + pool.length + L(" з ", " of ") + all.length) : "";
    }
    function done(i) {
      var name = pool[i];
      face.textContent = name;
      history.push(name);
      if (!w.repeat) pool.splice(i, 1);
      root.classList.add("w-picked");
      busy = false;
      draw();
    }
    function pick() {
      if (!pool.length || busy) return;
      busy = true;
      root.classList.remove("w-picked");
      var i = Math.floor(Math.random() * pool.length);
      if (wheel) {
        // Крутимо так, щоб сектор i зупинився під покажчиком угорі (-90°).
        var n = pool.length, seg = 2 * Math.PI / n;
        var targetAngle = -Math.PI / 2 - (i + 0.5) * seg;
        var spins = 5 * 2 * Math.PI, start = angle;
        var diff = ((targetAngle - start) % (2 * Math.PI) + 2 * Math.PI) % (2 * Math.PI);
        var end = start + spins + diff;
        var t0 = Date.now(), dur = 3200;
        var reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
        if (reduce) { angle = end; drawWheel(); done(i); return; }
        (function anim() {
          var p = Math.min(1, (Date.now() - t0) / dur), e = 1 - Math.pow(1 - p, 3);
          angle = start + (end - start) * e;
          drawWheel();
          if (p < 1) requestAnimationFrame(anim); else done(i);
        })();
        face.textContent = "…";
      } else {
        var n2 = 0, spin = setInterval(function () {
          face.textContent = all[Math.floor(Math.random() * all.length)];
          if (++n2 >= 14) { clearInterval(spin); done(i); }
        }, 70);
      }
      draw();
    }
    drawWheel(); draw();
    return { destroy: function () { busy = false; }, key: function (e) { if (isPrimaryKey(e)) { pick(); return true; } return false; } };
  }

  // ---------- Рахунок команд (за бажанням спільний для всіх слайдів) ----------
  function score(root, w) {
    var teams = lines(w.teams, 6);
    if (teams.length < 2) teams = [L("Команда 1", "Team 1"), L("Команда 2", "Team 2")];
    var step = Math.max(1, Math.min(100, parseInt(w.step, 10) || 1));
    var storeKey = "webOnSlide.score:" + teams.join("|");
    var SHARED_TTL_MS = 12 * 3600 * 1000; // спільний рахунок забувається через 12 годин після останньої зміни
    var chan = null;
    var pts = teams.map(function () { return 0; });
    function validPts(a) {
      return Array.isArray(a) && a.length === teams.length && a.every(function (x) { return typeof x === "number" && isFinite(x); });
    }
    if (w.shared) {
      try {
        var saved = JSON.parse(localStorage.getItem(storeKey) || "null");
        // Старий формат (просто масив) не має часу: вважаємо його застарілим.
        if (saved && validPts(saved.pts) && typeof saved.t === "number" && Date.now() - saved.t < SHARED_TTL_MS) pts = saved.pts.slice();
        else localStorage.removeItem(storeKey);
      } catch (e) { /* немає сховища */ }
      try { chan = new BroadcastChannel("webOnSlide.score"); chan.onmessage = function (m) { if (m.data && m.data.key === storeKey && validPts(m.data.pts)) { pts = m.data.pts.slice(); draw(); } }; } catch (e2) { chan = null; }
    }
    root.classList.add("w-score");
    title(root, w);
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
      card.appendChild(row(btn("", "−" + step, function () { add(i, -step); }), btn("primary", "+" + step, function () { add(i, step); })));
      grid.appendChild(card);
    });
    root.appendChild(grid);
    var resetB = btn("ghost", L("Обнулити", "Reset"), function () { pts = pts.map(function () { return 0; }); changed(); });
    var resetRow = row(resetB); resetRow.classList.add("w-tools");
    root.appendChild(resetRow);
    if (w.shared) root.appendChild(el("div", "w-note", L("Спільний рахунок для всіх слайдів з цими командами. Скидається кнопкою «Обнулити» або сам через 12 годин.", "Shared score across slides with these teams. Reset it with Reset, or it clears itself after 12 hours.")));
    copyButton(root, function () { return (w.title ? w.title + "\n" : "") + teams.map(function (t, i) { return t + ": " + pts[i]; }).join("\n"); });
    function add(i, d) { pts[i] += d; changed(); }
    function changed() {
      if (w.shared) {
        try { localStorage.setItem(storeKey, JSON.stringify({ pts: pts, t: Date.now() })); } catch (e) { /* немає сховища */ }
        if (chan) chan.postMessage({ key: storeKey, pts: pts });
      }
      draw();
    }
    function draw() {
      var max = Math.max.apply(null, pts);
      nums.forEach(function (n, i) {
        n.textContent = pts[i];
        n.parentNode.classList.toggle("w-lead", pts[i] === max && max > 0);
      });
    }
    draw();
    return {
      destroy: function () { if (chan) chan.close(); },
      // 1–6 додають бали команді, Shift+1–6 віднімають.
      key: function (e) {
        var n = parseInt(e.code && /^Digit\d$/.test(e.code) ? e.code.slice(5) : e.key, 10);
        if (n >= 1 && n <= teams.length) { add(n - 1, e.shiftKey ? -step : step); return true; }
        return false;
      }
    };
  }

  // ---------- Розподіл на групи ----------
  function groups(root, w) {
    var names = lines(w.names, 300);
    var n = Math.max(1, parseInt(w.n, 10) || 4);
    root.classList.add("w-groups");
    title(root, w);
    var grid = el("div", "w-group-grid");
    root.appendChild(grid);
    var result = [];
    function split() {
      var mixed = shuffle(names);
      var count = w.by === "size" ? Math.max(1, Math.ceil(mixed.length / n)) : Math.min(n, Math.max(1, mixed.length));
      result = [];
      for (var g = 0; g < count; g++) result.push([]);
      mixed.forEach(function (name, i) { result[i % count].push(name); });
      draw();
    }
    function draw() {
      while (grid.firstChild) grid.removeChild(grid.firstChild);
      grid.style.setProperty("--n", Math.min(result.length, 6));
      result.forEach(function (members, i) {
        var col = el("div", "w-group");
        col.appendChild(el("div", "w-group-name", L("Група ", "Group ") + (i + 1)));
        var ul = el("ul");
        members.forEach(function (m) { ul.appendChild(el("li", "", m)); });
        col.appendChild(ul);
        grid.appendChild(col);
      });
      if (!names.length) grid.appendChild(el("div", "w-note", L("Додайте учасників у налаштуваннях", "Add participants in the settings")));
    }
    root.appendChild(row(btn("primary", L("Перемішати", "Shuffle"), split)));
    copyButton(root, function () { return result.map(function (m, i) { return L("Група ", "Group ") + (i + 1) + ": " + m.join(", "); }).join("\n"); });
    split();
    return { destroy: function () {}, key: function (e) { if (isPrimaryKey(e)) { split(); return true; } return false; } };
  }

  // ---------- Ручне голосування залу ----------
  function vote(root, w) {
    var opts = lines(w.options, 6);
    if (opts.length < 2) opts = [L("Так", "Yes"), L("Ні", "No")];
    var votes = opts.map(function () { return 0; });
    root.classList.add("w-vote");
    title(root, w);
    var list = el("div", "w-bars");
    var bars = [];
    opts.forEach(function (o, i) {
      var r = el("div", "w-bar-row");
      var name = el("div", "w-bar-name", (i + 1) + ". " + o);
      var track = el("div", "w-bar-track");
      var fill = el("div", "w-bar-fill");
      var val = el("div", "w-bar-val", "0");
      track.appendChild(fill);
      var minus = btn("small", "−", function () { if (votes[i] > 0) { votes[i]--; draw(); } });
      var plus = btn("primary small", "+", function () { votes[i]++; draw(); });
      r.appendChild(name); r.appendChild(track); r.appendChild(val); r.appendChild(minus); r.appendChild(plus);
      list.appendChild(r);
      bars.push({ fill: fill, val: val });
    });
    root.appendChild(list);
    var total = el("div", "w-note");
    root.appendChild(total);
    var voteReset = row(btn("ghost", L("Обнулити", "Reset"), function () { votes = votes.map(function () { return 0; }); draw(); }));
    voteReset.classList.add("w-tools");
    root.appendChild(voteReset);
    copyButton(root, function () {
      var sum = votes.reduce(function (a, b) { return a + b; }, 0) || 1;
      return (w.title ? w.title + "\n" : "") + opts.map(function (o, i) { return o + ": " + votes[i] + " (" + Math.round(votes[i] * 100 / sum) + "%)"; }).join("\n");
    });
    function draw() {
      var sum = votes.reduce(function (a, b) { return a + b; }, 0);
      var max = Math.max.apply(null, votes) || 1;
      bars.forEach(function (b, i) {
        b.fill.style.width = (votes[i] * 100 / max) + "%";
        b.val.textContent = votes[i] + (sum ? " · " + Math.round(votes[i] * 100 / sum) + "%" : "");
      });
      total.textContent = L("Усього голосів: ", "Total votes: ") + sum;
    }
    draw();
    return {
      destroy: function () {},
      key: function (e) {
        var n = parseInt(e.code && /^Digit\d$/.test(e.code) ? e.code.slice(5) : e.key, 10);
        if (n >= 1 && n <= opts.length) { if (e.shiftKey) { if (votes[n - 1] > 0) votes[n - 1]--; } else votes[n - 1]++; draw(); return true; }
        return false;
      }
    };
  }

  // ---------- Кубик, монетка, число ----------
  var PIPS = ["", "⚀", "⚁", "⚂", "⚃", "⚄", "⚅"];
  function dice(root, w) {
    var mode = w.dice === "coin" || w.dice === "number" ? w.dice : "dice";
    var count = Math.max(1, Math.min(6, parseInt(w.n, 10) || 1));
    var max = Math.max(2, Math.min(1000000, parseInt(w.n, 10) || 100));
    root.classList.add("w-dice");
    title(root, w);
    var face = el("div", "w-dice-face", "?");
    face.setAttribute("aria-live", "polite");
    var sum = el("div", "w-note");
    root.appendChild(face); root.appendChild(sum);
    var busy = false;
    function once() {
      if (mode === "coin") return Math.random() < 0.5 ? L("Орел", "Heads") : L("Решка", "Tails");
      if (mode === "number") return String(1 + Math.floor(Math.random() * max));
      var vals = [];
      for (var i = 0; i < count; i++) vals.push(1 + Math.floor(Math.random() * 6));
      return vals;
    }
    function show(v) {
      if (Array.isArray(v)) {
        face.textContent = v.map(function (x) { return PIPS[x]; }).join(" ");
        face.setAttribute("aria-label", v.join(", "));
        sum.textContent = v.length > 1 ? L("Сума: ", "Sum: ") + v.reduce(function (a, b) { return a + b; }, 0) : "";
      } else { face.textContent = v; face.removeAttribute("aria-label"); sum.textContent = ""; }
    }
    function roll() {
      if (busy) return;
      busy = true; root.classList.remove("w-picked");
      var k = 0, t = setInterval(function () {
        show(once());
        if (++k >= 10) { clearInterval(t); busy = false; root.classList.add("w-picked"); }
      }, 60);
    }
    var label = mode === "coin" ? L("Підкинути", "Flip") : mode === "number" ? L("Число від 1 до ", "Number 1 to ") + max : L("Кинути", "Roll");
    root.appendChild(row(btn("primary big", label, roll)));
    return { destroy: function () {}, key: function (e) { if (isPrimaryKey(e)) { roll(); return true; } return false; } };
  }

  // ---------- Квіз-картка ----------
  function quiz(root, w) {
    var raw = lines(w.options, 6);
    var opts = raw.map(function (o) { return { text: o.replace(/^\*\s*/, ""), ok: /^\*/.test(o) }; });
    var secs = Math.max(0, Math.min(600, parseInt(w.n, 10) || 0));
    root.classList.add("w-quiz");
    title(root, w, L("Питання", "Question"));
    var LETTERS = "ABCDEF";
    var grid = el("div", "w-quiz-grid");
    var items = opts.map(function (o, i) {
      var it = el("div", "w-quiz-opt");
      it.appendChild(el("span", "w-quiz-letter", LETTERS[i]));
      it.appendChild(el("span", "", o.text));
      grid.appendChild(it);
      return it;
    });
    root.appendChild(grid);
    var timeEl = secs ? el("div", "w-quiz-time") : null;
    var startB = null;
    function syncStart() { if (startB) startB.textContent = cd.running() ? L("Пауза", "Pause") : L("Старт часу", "Start timer"); }
    var cd = secs ? countdown(secs, function () { timeEl.textContent = fmt(cd.left()); timeEl.classList.toggle("w-warn-text", cd.left() <= 5); syncStart(); }, function () { beep(2); syncStart(); }) : null;
    if (timeEl) { root.appendChild(timeEl); timeEl.textContent = fmt(secs); }
    var revealed = false;
    function reveal() {
      revealed = !revealed;
      if (cd && revealed) { cd.pause(); syncStart(); }
      items.forEach(function (it, i) {
        it.classList.toggle("w-correct", revealed && opts[i].ok);
        it.classList.toggle("w-wrong", revealed && !opts[i].ok);
      });
      revB.textContent = revealed ? L("Сховати відповідь", "Hide answer") : L("Показати відповідь", "Show answer");
    }
    startB = cd ? btn("", L("Старт часу", "Start timer"), function () { cd.toggle(); syncStart(); }) : null;
    var revB = btn("primary", L("Показати відповідь", "Show answer"), reveal);
    root.appendChild(row(startB, revB));
    // Без позначеної відповіді кнопка нічого не покаже, а підказка потрібна лише тренерові: у показі їх не видно.
    if (!opts.some(function (o) { return o.ok; })) {
      revB.classList.add("w-edit-only");
      root.appendChild(el("div", "w-note w-edit-only", L("Позначте правильну відповідь зірочкою * у налаштуваннях", "Mark the correct answer with * in the settings")));
    }
    return {
      destroy: function () { if (cd) cd.destroy(); },
      key: function (e) {
        if (e.key === "Enter") { reveal(); return true; }
        if (e.key === " " && startB) { startB.click(); return true; }
        return false;
      }
    };
  }

  // ---------- Порядок денний з етапами ----------
  function agenda(root, w) {
    var steps = lines(w.steps, 20).map(function (l) {
      var bar = l.lastIndexOf("|");
      var mins = bar === -1 ? 5 : parseFloat(l.slice(bar + 1).replace(",", "."));
      return { name: (bar === -1 ? l : l.slice(0, bar)).trim(), sec: Math.round(Math.min(Math.max(mins || 5, 0.25), 180) * 60) };
    });
    if (!steps.length) steps = [{ name: L("Етап 1", "Step 1"), sec: 300 }];
    var cur = 0;
    root.classList.add("w-agenda");
    title(root, w);
    var layout = el("div", "w-agenda-layout");
    var main = el("div", "w-agenda-main");
    var stepName = el("div", "w-sub");
    var face = el("div", "w-time");
    face.setAttribute("role", "timer");
    main.appendChild(stepName); main.appendChild(face);
    var list = el("ol", "w-agenda-list");
    var lis = steps.map(function (s) {
      var li = el("li", "", s.name + " · " + fmt(s.sec));
      list.appendChild(li);
      return li;
    });
    layout.appendChild(main); layout.appendChild(list);
    root.appendChild(layout);
    var cd = countdown(steps[0].sec, draw, function () {
      beep(2);
      if (cur < steps.length - 1) { go(cur + 1); cd.start(); }
    });
    var startB = btn("primary", "", function () { cd.toggle(); });
    var prevB = btn("", "←", function () { if (cur > 0) go(cur - 1); });
    var nextB = btn("", "→", function () { if (cur < steps.length - 1) go(cur + 1); });
    prevB.setAttribute("aria-label", L("Попередній етап", "Previous step"));
    nextB.setAttribute("aria-label", L("Наступний етап", "Next step"));
    main.appendChild(row(prevB, startB, nextB));
    function go(i) { var was = cd.running(); cur = i; cd.reset(steps[i].sec); if (was) cd.start(); draw(); }
    function draw() {
      var left = cd.left();
      stepName.textContent = (cur + 1) + "/" + steps.length + " · " + steps[cur].name;
      face.textContent = fmt(left);
      root.classList.toggle("w-warn", left > 0 && left <= 10);
      root.classList.toggle("w-done", left <= 0 && cur === steps.length - 1);
      lis.forEach(function (li, i) { li.className = i < cur ? "w-past" : i === cur ? "w-now" : ""; });
      startB.textContent = cd.running() ? L("Пауза", "Pause") : L("Старт", "Start");
      prevB.disabled = cur === 0; nextB.disabled = cur === steps.length - 1;
    }
    draw();
    return {
      destroy: cd.destroy,
      key: function (e) {
        if (isPrimaryKey(e)) { cd.toggle(); return true; }
        if (e.key === "ArrowRight" && cur < steps.length - 1) { go(cur + 1); return true; }
        if (e.key === "ArrowLeft" && cur > 0) { go(cur - 1); return true; }
        return false;
      }
    };
  }

  // ---------- Колода карток ----------
  function cards(root, w) {
    var all = lines(w.cards, 200), pool = shuffle(all), drawn = [];
    root.classList.add("w-cards");
    title(root, w);
    var card = el("div", "w-card-face", all.length ? L("Натисніть «Витягнути картку»", "Click “Draw a card”") : L("Додайте картки в налаштуваннях", "Add cards in the settings"));
    card.setAttribute("aria-live", "polite");
    var info = el("div", "w-note");
    var drawB = btn("primary big", L("Витягнути картку", "Draw a card"), drawCard);
    var again = btn("", L("Перемішати знову", "Reshuffle"), function () { pool = shuffle(all); drawn = []; card.classList.remove("w-flipped"); card.textContent = "?"; update(); });
    root.appendChild(card); root.appendChild(row(drawB, again)); root.appendChild(info);
    function update() {
      drawB.disabled = !pool.length;
      info.textContent = all.length ? L("У колоді: ", "In deck: ") + pool.length + L(" з ", " of ") + all.length : "";
    }
    function drawCard() {
      if (!pool.length) return;
      var c = pool.pop();
      drawn.push(c);
      card.classList.remove("w-flipped");
      void card.offsetWidth; // перезапуск анімації
      card.textContent = c;
      card.classList.add("w-flipped");
      update();
    }
    update();
    return { destroy: function () {}, key: function (e) { if (isPrimaryKey(e)) { drawCard(); return true; } return false; } };
  }

  var KINDS = { timer: timer, "break": breakUntil, picker: picker, score: score, groups: groups, vote: vote, dice: dice, quiz: quiz, agenda: agenda, cards: cards };
  var active = null, mountedKey = null;

  // Гарячі клавіші працюють, коли фокус у блоці (після кліку по ньому). Не чіпаємо поля вводу.
  document.addEventListener("keydown", function (e) {
    if (!active || e.ctrlKey || e.altKey || e.metaKey) return;
    // Поки відкрита форма налаштувань, інструмент прихований і клавіші йому не належать.
    var setupForm = document.getElementById("setup");
    if (setupForm && !setupForm.classList.contains("hidden")) return;
    var box = document.getElementById("widget");
    if (box && box.classList.contains("hidden")) return;
    var t = e.target && e.target.tagName;
    if (t === "INPUT" || t === "TEXTAREA" || t === "SELECT") return;
    if ((e.key === " " || e.key === "Enter") && t === "BUTTON") return; // кнопка сама обробить клік
    if (active.key && active.key(e)) e.preventDefault();
  });

  window.WOS_WIDGETS = {
    kinds: Object.keys(KINDS),
    // Монтує віджет у контейнер; повторний виклик з тими ж налаштуваннями зберігає стан.
    mount: function (container, kind, w) {
      var key = kind + "|" + JSON.stringify(w || {});
      if (key === mountedKey) return;
      this.unmount(container);
      var root = el("div", "w-root");
      container.appendChild(root);
      active = (KINDS[kind] || timer)(root, w || {});
      mountedKey = key;
    },
    unmount: function (container) {
      if (active && active.destroy) active.destroy();
      active = null; mountedKey = null;
      while (container.firstChild) container.removeChild(container.firstChild);
    }
  };
})();
