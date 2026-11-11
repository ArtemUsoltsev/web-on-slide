// Кожна вставка content add-in має власні налаштування в .pptx,
// тож на кожному слайді може бути свій URL і свій режим.
(function () {
  "use strict";

  var KEY_URL = "url";         // адреса; може містити {код} (див. config.js)
  var KEY_MODE = "mode";       // "embed" (iframe) | "button" (кнопка, що відкриває вікно) | "qr" (лише QR-код)
  var KEY_SCALE = "scale";     // масштаб сторінки, 0.5–1 (1 = 100%)
  var KEY_SCROLL = "scroll";   // "show" = показувати прокрутку; інакше прихована
  var KEY_CODE = "code";       // код сесії, що підставляється замість {код}
  var KEY_QR = "qr";           // "corner" = QR-код адреси в куті блоку
  var KEY_REVEAL = "reveal";   // "1" = сайт з'являється лише після кліку по заставці
  var KEY_REVEAL_TEXT = "revealText"; // підпис на заставці
  var KEY_REFRESH = "refresh"; // автооновлення, секунди (0 = вимкнено)
  var KEY_ZOOM = "zoom";       // збільшення частини сторінки, 1–3
  var KEY_FX = "fx", KEY_FY = "fy"; // яку частину збільшувати: 0–100% по горизонталі й вертикалі
  var KEY_OFFLINE = "offline"; // власний текст, якщо сайт не відкрився
  var KEY_TABS = "tabs";       // JSON: [{t: назва, u: адреса}] — додаткові вкладки
  var KEY_MAIN_TITLE = "mainTitle"; // назва першої вкладки
  var KEY_BORDER = "border";   // "" | "gray" | "yellow" | "dark"
  var KEY_ROUNDED = "rounded"; // "1" = скруглені кути
  var KEY_CORNER_BG = "cornerBg"; // колір за скругленими кутами (під колір слайда)
  var MAX_TABS = 3;            // додаткових вкладок, крім основної
  var ZOOM_CHOICES = [1, 1.25, 1.5, 2, 3];
  var RECENT_KEY = "webOnSlide.recent"; // нещодавні адреси, лише на цьому комп'ютері
  var LOAD_TIMEOUT_MS = 10000; // через скільки показати картку «сторінка не відповідає»
  var REFRESH_CHOICES = [0, 10, 30, 60, 300];
  var KEY_AUTOPLAY = "autoplay"; // "1" = автозапуск відео YouTube/Vimeo без звуку
  var KEY_KIND = "kind";       // "site" або інструмент тренера з widgets.js (timer, picker, score, groups, vote, ...)
  var KEY_ROTATE = "rotate";   // карусель вкладок: секунди між перемиканням (0 = вимкнено)
  var ROTATE_CHOICES = [0, 10, 20, 30, 60];
  var TEMPLATES_KEY = "webOnSlide.templates"; // мої шаблони блоків, лише на цьому комп'ютері
  var KEY_WIDGET = "widget";   // JSON налаштувань інструмента тренера
  var KEY_TIMER = "timerCfg";  // JSON налаштувань «Таймера PRO» (timer/index.html) для цього слайда
  var TIMER_URL = "timer/index.html?embed=1&v=1.9.1";

  function L(uk, en) { return window.WOS_I18N ? window.WOS_I18N.L(uk, en) : uk; }
  function lang() { return window.WOS_I18N ? window.WOS_I18N.lang() : "uk"; }
  function reason(code) { return cfg.reason ? cfg.reason(code, lang()) : (cfg.reasonText[code] || "Ця адреса не дозволена."); }
  function defaultRevealText() { return L("Натисніть, щоб показати", "Click to show"); }

  var cfg = window.WEB_ON_SLIDE_CONFIG;

  function $(id) { return document.getElementById(id); }
  var frame = $("frame"), frameWrap = $("frameWrap"), launch = $("launch"), setup = $("setup"), bar = $("bar");
  var loadfail = $("loadfail"), fallback = $("fallback"), fatal = $("fatal"), loading = $("loading");
  var notice = $("notice");
  var cover = $("cover"), qrCard = $("qrCard"), qrCorner = $("qrCorner");
  var input = $("url"), codeInput = $("code"), codeRow = $("codeRow"), info = $("info"), err = $("err"), warn = $("warn");
  var modeSel = $("mode"), scaleSel = $("scale"), refreshSel = $("refresh"), allowScroll = $("allowScroll");
  var qrBox = $("qrCornerOn"), revealBox = $("reveal"), revealText = $("revealText"), more = $("more");
  var zoomSel = $("zoom"), fxIn = $("fx"), fyIn = $("fy"), offlineIn = $("offlineText"), tabsText = $("tabsText"), mainTitle = $("mainTitle");
  var kindSel = $("kind"), siteFields = $("siteFields"), widgetFields = $("widgetFields"), widgetBox = $("widget");
  var autoplayBox = $("autoplay");
  var timerWrap = $("timerWrap"), timerFrame = $("timerFrame");
  var timerSaveT = null;
  var formTimer = "";     // налаштування Таймера PRO, з якими відкрито форму (з блоку або з шаблону)
  var borderSel = $("border"), roundedBox = $("rounded"), cornerBg = $("cornerBg"), tabsBar = $("tabs"), frameZoom = $("frameZoom");

  var view = null;        // null (ще невідомо) | "edit" | "read" (показ слайдів)
  var current = null;     // збережене налаштування цієї вставки (див. readSettings)
  var editing = false;    // форма відкрита кнопкою «Змінити»; current при цьому зберігається
  var revealed = false;   // заставку вже прибрали кліком
  var activeTab = 0;      // 0 = основна адреса, 1..3 = додаткові вкладки
  var rotateTimer = null; // карусель вкладок
  var editingFromTemplate = false;
  var dialogSupported = false;
  var isWeb = false;
  var loadTimer = null;
  var refreshTimer = null;
  var frameLoaded = false;
  var qrDrawnFor = { corner: null, card: null };

  function hide() { for (var i = 0; i < arguments.length; i++) arguments[i].classList.add("hidden"); }
  function show() { for (var i = 0; i < arguments.length; i++) arguments[i].classList.remove("hidden"); }

  function cleanScale(v) {
    var n = parseFloat(v);
    return isFinite(n) && n >= 0.25 && n <= 1 ? n : 1;
  }
  function cleanRefresh(v) {
    var n = parseInt(v, 10);
    return REFRESH_CHOICES.indexOf(n) !== -1 ? n : 0;
  }
  function cleanRotate(v) {
    var n = parseInt(v, 10);
    return ROTATE_CHOICES.indexOf(n) !== -1 ? n : 0;
  }
  function cleanZoom(v) {
    var n = parseFloat(v);
    return ZOOM_CHOICES.indexOf(n) !== -1 ? n : 1;
  }
  function cleanPct(v, def) {
    var n = parseInt(v, 10);
    return isFinite(n) && n >= 0 && n <= 100 ? n : def;
  }
  function cleanColor(v) { return /^#[0-9a-f]{6}$/i.test(v || "") ? v : "#ffffff"; }
  function cleanTabs(v) {
    var arr;
    try { arr = typeof v === "string" ? JSON.parse(v) : v; } catch (e) { arr = null; }
    if (!Array.isArray(arr)) return [];
    var out = [];
    for (var i = 0; i < arr.length && out.length < MAX_TABS; i++) {
      if (arr[i] && typeof arr[i].u === "string") out.push({ t: String(arr[i].t || "").slice(0, 30), u: arr[i].u });
    }
    return out;
  }
  // Рядки «Назва | адреса» з форми → [{t, u}]; повертає { tabs, error }.
  function parseTabs(text, code) {
    var lines = String(text || "").split(/\r?\n/), tabs = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      var bar = line.lastIndexOf("|");
      var title = bar === -1 ? "" : line.slice(0, bar).trim();
      var url = cfg.normalize(bar === -1 ? line : line.slice(bar + 1));
      var res = cfg.check(cfg.fillCode(url, code || "x"));
      if (!res.ok) return { tabs: [], error: L("Вкладка «", "Tab “") + (title || url) + L("»: ", "”: ") + reason(res.reason) };
      if (tabs.length >= MAX_TABS) return { tabs: [], error: L("Можна додати не більше " + MAX_TABS + " вкладок.", "You can add up to " + MAX_TABS + " tabs.") };
      tabs.push({ t: title.slice(0, 30) || hostOf(url) || L("Сторінка ", "Page ") + (tabs.length + 2), u: url });
    }
    return { tabs: tabs, error: null };
  }
  function tabsToText(tabs) {
    return tabs.map(function (t) { return t.t + " | " + t.u; }).join("\n");
  }

  function cleanMode(v) {
    if (v === "button" && dialogSupported) return "button";
    if (v === "qr") return "qr";
    return "embed";
  }

  // Шаблон адреси активної вкладки.
  function activeTemplate(cur) {
    var i = activeTab > 0 && activeTab <= cur.tabs.length ? activeTab : 0;
    return i ? cur.tabs[i - 1].u : cur.url;
  }
  // Адреса, яку відкривають люди (у вікні, за QR-кодом): шаблон із підставленим кодом.
  function targetUrl(cur) { return cfg.normalize(cfg.fillCode(activeTemplate(cur), cur.code)); }
  // Адреса для блоку на слайді: та сама, але перетворена на embed-адресу, якщо сервіс відомий.
  function frameUrl(cur) {
    var t = targetUrl(cur);
    var e = cfg.toEmbed(t);
    var url = e.service && cfg.check(e.url).ok ? e.url : t;
    return cur.autoplay ? cfg.withAutoplay(url) : url;
  }
  function usesCode(cur) {
    if (cfg.hasCode(cur.url)) return true;
    for (var i = 0; i < cur.tabs.length; i++) if (cfg.hasCode(cur.tabs[i].u)) return true;
    return false;
  }
  function needsCode(cur) { return cur.kind === "site" && usesCode(cur) && !cur.code; }

  // Масштаб < 1: iframe більший за блок у 1/scale разів і зменшений transform-ом,
  // тож сайт розкладається як на більшому екрані й вміщується без прокрутки.
  function applyLook(cur) {
    var k = cleanScale(cur.scale);
    frame.style.width = (100 / k) + "%";
    frame.style.height = (100 / k) + "%";
    frame.style.transform = k === 1 ? "" : "scale(" + k + ")";
    // Кадр: збільшуємо обгортку навколо точки (fx%, fy%), тож видно лише цю частину сторінки.
    var z = cleanZoom(cur.zoom);
    frameZoom.style.transformOrigin = cur.fx + "% " + cur.fy + "%";
    frameZoom.style.transform = z === 1 ? "" : "scale(" + z + ")";
    frameWrap.className = (cur.border ? "b-" + cur.border : "") + (cur.rounded ? " rounded" : "");
    document.body.style.background = cur.rounded ? cur.cornerBg : "";
  }

  function hostOf(url) {
    try { return new URL(url).host; } catch (e) { return ""; }
  }
  // Коротка адреса для підпису під QR-кодом: без https:// і www.
  function shortUrl(url) {
    try {
      var u = new URL(url);
      var s = u.host.replace(/^www\./, "") + (u.pathname === "/" ? "" : u.pathname) + u.search;
      return s.length > 60 ? s.slice(0, 57) + "…" : s;
    } catch (e) { return url; }
  }

  // ---------- QR-код ----------
  // Малюється локально (vendor/qrcode.js), адреса нікуди не передається.
  var SVG_NS = "http://www.w3.org/2000/svg";
  function drawQr(container, text) {
    while (container.firstChild) container.removeChild(container.firstChild);
    if (typeof qrcode !== "function") return false;
    var qr;
    try {
      qrcode.stringToBytes = qrcode.stringToBytesFuncs["UTF-8"];
      qr = qrcode(0, "M");
      qr.addData(text, "Byte");
      qr.make();
    } catch (e) { return false; }
    var n = qr.getModuleCount(), q = 4, size = n + q * 2, d = "";
    for (var r = 0; r < n; r++) {
      for (var c = 0; c < n; c++) {
        if (qr.isDark(r, c)) d += "M" + (c + q) + " " + (r + q) + "h1v1h-1z";
      }
    }
    var svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("viewBox", "0 0 " + size + " " + size);
    svg.setAttribute("shape-rendering", "crispEdges");
    svg.setAttribute("role", "img");
    svg.setAttribute("aria-label", L("QR-код: ", "QR code: ") + text);
    var bg = document.createElementNS(SVG_NS, "rect");
    bg.setAttribute("width", size); bg.setAttribute("height", size); bg.setAttribute("fill", "#fff");
    var path = document.createElementNS(SVG_NS, "path");
    path.setAttribute("d", d); path.setAttribute("fill", "#000");
    svg.appendChild(bg); svg.appendChild(path);
    container.appendChild(svg);
    return true;
  }
  // Бібліотека QR (56 КБ) вантажиться лише тоді, коли блок справді показує QR-код.
  var QR_LIB = "vendor/qrcode.js?v=1.9.1";
  var qrLibState = typeof qrcode === "function" ? "ready" : "none"; // none | loading | ready | failed
  var qrWaiting = [];
  function withQrLib(cb) {
    if (qrLibState === "ready") { cb(); return; }
    if (qrLibState === "failed") return;
    qrWaiting.push(cb);
    if (qrLibState === "loading") return;
    qrLibState = "loading";
    var sc = document.createElement("script");
    sc.src = QR_LIB;
    sc.onload = function () {
      qrLibState = typeof qrcode === "function" ? "ready" : "failed";
      var list = qrWaiting; qrWaiting = [];
      if (qrLibState === "ready") list.forEach(function (f) { f(); });
    };
    sc.onerror = function () { qrLibState = "failed"; qrWaiting = []; };
    document.head.appendChild(sc);
  }
  var qrWanted = { corner: null, card: null }; // адреса, для якої QR має бути видно просто зараз
  function showQr(which, el, url) {
    qrWanted[which] = url;
    if (qrDrawnFor[which] === url) { show(el); return; }
    withQrLib(function () {
      // Поки бібліотека вантажилась, блок могли перемкнути: малюємо лише якщо QR досі потрібен.
      if (qrWanted[which] !== url) return;
      if (!drawQr(which === "card" ? $("qrBig") : el, url)) return;
      qrDrawnFor[which] = url;
      show(el);
    });
  }

  // ---------- Фатальні стани: без office.js або поза PowerPoint ----------
  function showFatal(text) {
    hide(frameWrap, launch, setup, bar, loadfail, fallback, loading, cover, qrCard, qrCorner, notice);
    $("fatalText").textContent = text;
    $("fatalUrl").textContent = location.href;
    show(fatal);
  }

  if (!cfg || typeof cfg.check !== "function" || typeof cfg.toEmbed !== "function") {
    showFatal("Не вдалося завантажити налаштування надбудови (config.js). Повідомте адміністратора. / Couldn't load the add-in settings (config.js). Contact your administrator.");
    return;
  }
  if (typeof Office === "undefined" || !Office.onReady) {
    showFatal("Не вдалося завантажити бібліотеку Office (appsforoffice.microsoft.com). Перевірте мережу або проксі. / Couldn't load the Office library. Check the network or proxy.");
    return;
  }

  // ---------- Білий список → CSP frame-src ----------
  // Якщо allowedOrigins заповнено, додаємо другий CSP лише з цими доменами. Браузер застосовує
  // обидва, тому список обмежує і переходи всередині iframe, а не лише першу адресу.
  function applyFrameCsp() {
    var allowed = Array.isArray(cfg.allowedOrigins) ? cfg.allowedOrigins : [];
    if (!allowed.length) return;
    var meta = document.createElement("meta");
    meta.httpEquiv = "Content-Security-Policy";
    // 'self' — для «Таймера PRO», що лежить на цьому ж сайті (адреси сайтів зі свого домену config.js не пропускає).
    meta.content = "frame-src 'self' " + allowed.join(" ") + ";";
    document.head.appendChild(meta);
  }

  // ---------- Відображення ----------
  function clearLoadTimer() {
    if (loadTimer) { clearTimeout(loadTimer); loadTimer = null; }
  }
  function stopRefresh() {
    if (refreshTimer) { clearInterval(refreshTimer); refreshTimer = null; }
  }
  // Автооновлення: тихо перезавантажує сторінку без індикатора, щоб табло не блимало заставкою.
  function startRefresh(cur) {
    stopRefresh();
    if (!cur.refresh) return;
    refreshTimer = setInterval(function () {
      if (document.hidden || !frame.getAttribute("src")) return;
      frame.src = frame.getAttribute("src");
    }, cur.refresh * 1000);
  }

  function unloadFrame() {
    clearLoadTimer();
    stopRefresh();
    if (rotateTimer) { clearInterval(rotateTimer); rotateTimer = null; }
    frame.removeAttribute("src");
  }

  function loadFrame(url) {
    clearLoadTimer();
    frameLoaded = false;
    hide(loadfail);
    $("loadingText").textContent = hostOf(url) ? L("Завантажую ", "Loading ") + hostOf(url) + "…" : L("Завантажую сторінку…", "Loading page…");
    show(loading);
    frame.title = hostOf(url) || L("Вбудована сторінка", "Embedded page");
    // scrolling="no" ховає смугу прокрутки вбудованої сторінки (Chromium/Edge WebView2).
    // Атрибут має стояти до завантаження, тому виставляємо його перед src.
    var noScroll = !current || current.scroll !== "show";
    if (noScroll) frame.setAttribute("scrolling", "no"); else frame.removeAttribute("scrolling");
    frame.src = url;
    // Без мережі не чекаємо 10 секунд: одразу показуємо картку (з власним текстом тренера, якщо є).
    loadTimer = setTimeout(function () {
      if (!frameLoaded) showLoadFail(url);
    }, navigator.onLine === false ? 300 : LOAD_TIMEOUT_MS);
  }

  function setEditLabel(text, title) {
    var b = $("editBtn");
    b.querySelector("span").textContent = text;
    b.title = title;
  }
  function defaultFailText(host) {
    return host
      ? L(host + " не відповідає. Сайт може забороняти показ у слайді, або немає інтернету.", host + " is not responding. The site may not allow showing it in a slide, or there is no internet.")
      : L("Сторінка не відповідає. Сайт може забороняти показ у слайді, або немає інтернету.", "The page is not responding. The site may not allow showing it in a slide, or there is no internet.");
  }
  function showLoadFail(url) {
    hide(loading);
    var own = current && current.offline;
    $("loadfailText").textContent = own || defaultFailText(hostOf(url));
    $("loadfailText").style.fontSize = own ? "clamp(16px, 5vmin, 40px)" : "";
    $("loadfailHost").textContent = "";
    // Підказка-виправлення лише для автора, не під час показу.
    if (view === "edit" && dialogSupported) show($("loadfailToButton")); else hide($("loadfailToButton"));
    show(loadfail);
  }

  frame.addEventListener("load", function () {
    // Подія приходить і для сторінки помилки браузера, але для X-Frame-Options часто не приходить,
    // тому таймер лишається головним сигналом.
    frameLoaded = true;
    clearLoadTimer();
    hide(loadfail, loading);
  });

  function setError(text, field) {
    err.textContent = text || "";
    input.removeAttribute("aria-invalid");
    codeInput.removeAttribute("aria-invalid");
    tabsText.removeAttribute("aria-invalid");
    ["wMinutes", "wNames", "wTime", "wLines", "wNum"].forEach(function (id) { $(id).removeAttribute("aria-invalid"); });
    if (text) (field || input).setAttribute("aria-invalid", "true");
  }

  // ---------- Таймер PRO ----------
  // Сторінка timer/index.html з цього ж сайту. Вона читає window.WOS_TIMER_HOST батьківської сторінки:
  // cfg — налаштування цього слайда, view — редагування чи показ, save — зберегти зміни у файл.
  function validKind(k) {
    if (k === "protimer") return true;
    return !!(window.WOS_WIDGETS && window.WOS_WIDGETS.kinds.indexOf(k) !== -1);
  }
  function unloadTimer() {
    if (timerFrame.getAttribute("src")) timerFrame.removeAttribute("src");
  }
  function showTimer() {
    window.WOS_TIMER_HOST.view = view;
    show(timerWrap);
    if (!timerFrame.getAttribute("src")) { timerFrame.src = TIMER_URL; return; }
    // Перемикання редагування ↔ показ без перезавантаження: таймер, що вже йде, не скидається.
    try { if (timerFrame.contentWindow.WOS_TIMER_VIEW) timerFrame.contentWindow.WOS_TIMER_VIEW(view); } catch (e) { /* ще вантажиться */ }
  }

  function render() {
    hide(frameWrap, launch, setup, bar, loadfail, fallback, loading, cover, qrCard, qrCorner, tabsBar, widgetBox, notice, timerWrap);
    qrWanted.corner = null; qrWanted.card = null;
    document.body.style.background = "";
    setEditLabel(L("Змінити", "Edit"), L("Змінити адресу та вигляд", "Change address and look"));
    // Службові написи й кнопки інструментів тренера видно лише в режимі редагування
    document.body.classList.toggle("editing", view === "edit");
    if (view === null) return; // не малюємо нічого, поки не знаємо режим: інакше форма блимає під час показу

    var showForm = !current || needsCode(current) || (editing && view === "edit");
    if (showForm) {
      unloadFrame();
      unloadTimer();
      if (view === "edit") {
        if (current && !needsCode(current)) show($("cancelBtn")); else hide($("cancelBtn"));
        show(setup);
        // Нова вставка: одразу ставимо курсор у поле адреси.
        if (!current) input.focus();
      }
      // Під час показу без налаштованого URL нічого не показуємо, а без коду сесії пояснюємо, що сталося,
      // щоб тренер не побачив перед залом порожній білий блок.
      if (view !== "edit" && current && needsCode(current)) {
        showNotice(L("Не задано код сесії. Вийдіть з показу й натисніть «Код» на цьому блоці.", "No session code is set. Leave the slide show and click Code on this block."));
      }
      return;
    }

    if (current.kind === "protimer") {
      unloadFrame();
      if (window.WOS_WIDGETS) window.WOS_WIDGETS.unmount(widgetBox);
      showTimer();
      if (view === "edit") {
        hide($("reloadBtn"), $("popBtn"), $("codeBtn"));
        // У таймера є власні налаштування, тож кнопка надбудови називається за тим, що вона відкриває
        setEditLabel(L("Тип блоку", "Block type"), L("Змінити тип блоку", "Change block type"));
        show(bar);
      }
      return;
    }
    unloadTimer();

    if (current.kind !== "site" && window.WOS_WIDGETS) {
      unloadFrame();
      window.WOS_WIDGETS.mount(widgetBox, current.kind, current.widget);
      show(widgetBox);
      if (view === "edit") {
        hide($("reloadBtn"), $("popBtn"), $("codeBtn"));
        show(bar);
      }
      return;
    }
    if (window.WOS_WIDGETS) window.WOS_WIDGETS.unmount(widgetBox);

    var target = targetUrl(current);
    var reloadable = false;
    // Код сесії підставляється вже після перевірки шаблону, тож перевіряємо остаточну адресу ще раз:
    // код у назві сайту (https://{код}.example.com) не повинен обходити білий список.
    var finalCheck = cfg.check(target);
    if (!finalCheck.ok) {
      unloadFrame();
      showNotice(reason(finalCheck.reason));
      if (view === "edit") { hide($("reloadBtn"), $("popBtn")); if (usesCode(current)) show($("codeBtn")); else hide($("codeBtn")); show(bar); }
      return;
    }
    if (current.mode === "button") {
      // Режим кнопки: iframe не використовується, тож «Оновити» лише викликало б хибне «не відповідає».
      unloadFrame();
      $("launchHost").textContent = hostOf(target);
      show(launch);
    } else if (current.mode === "qr") {
      unloadFrame();
      // Картка входу: якщо адреса містить {код}, показуємо адресу без коду і сам код великими цифрами.
      if (cfg.hasCode(activeTemplate(current)) && current.code) {
        $("qrUrl").textContent = shortUrl(cfg.normalize(cfg.fillCode(activeTemplate(current), ""))).replace(/[\/?=&#]+$/, "");
        $("qrCode").textContent = (L("Код ", "Code ") + current.code.replace(/(\d{4})(?=\d)/g, "$1 "));
      } else {
        $("qrUrl").textContent = shortUrl(target);
        $("qrCode").textContent = "";
      }
      showQr("card", qrCard, target);
    } else if (current.reveal && !revealed) {
      unloadFrame();
      $("coverBtn").textContent = current.revealText || defaultRevealText();
      show(cover);
    } else {
      var url = frameUrl(current);
      applyLook(current);
      reloadable = true;
      show(frameWrap);
      if (frame.getAttribute("src") !== url) loadFrame(url);
      else if (!frameLoaded && loadTimer) show(loading);
      if (!refreshTimer) startRefresh(current);
      if (current.tabs.length) {
        renderTabs();
        if (current.rotate && !rotateTimer) {
          rotateTimer = setInterval(function () {
            if (document.hidden || !current || !current.tabs.length) return;
            activeTab = (activeTab + 1) % (current.tabs.length + 1);
            render();
          }, current.rotate * 1000);
        }
      }
    }
    if (current.qrCorner && current.mode !== "qr") showQr("corner", qrCorner, target);

    if (view === "edit") {
      if (reloadable) show($("reloadBtn")); else hide($("reloadBtn"));
      if (dialogSupported) show($("popBtn"));
      if (usesCode(current)) show($("codeBtn")); else hide($("codeBtn"));
      show(bar);
    }
  }

  function showNotice(text) {
    $("noticeText").textContent = text;
    show(notice);
  }

  function renderTabs() {
    while (tabsBar.firstChild) tabsBar.removeChild(tabsBar.firstChild);
    var titles = [current.mainTitle || L("Сторінка", "Page")].concat(current.tabs.map(function (t) { return t.t; }));
    titles.forEach(function (title, i) {
      var b = document.createElement("button");
      b.type = "button";
      b.setAttribute("role", "tab");
      b.setAttribute("aria-selected", i === activeTab ? "true" : "false");
      b.textContent = title;
      b.addEventListener("click", function () {
        if (i === activeTab) return;
        activeTab = i;
        render();
      });
      tabsBar.appendChild(b);
    });
    show(tabsBar);
  }

  // ---------- Нещодавні адреси (лише на цьому комп'ютері, у файл не потрапляють) ----------
  function readRecent() {
    try { var a = JSON.parse(localStorage.getItem(RECENT_KEY) || "[]"); return Array.isArray(a) ? a.slice(0, 10) : []; }
    catch (e) { return []; }
  }
  function rememberRecent(url) {
    var list = readRecent().filter(function (u) { return u !== url; });
    list.unshift(url);
    try { localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0, 10))); } catch (e) { /* сховище недоступне */ }
    fillRecent();
  }
  function fillRecent() {
    var dl = $("recentUrls");
    while (dl.firstChild) dl.removeChild(dl.firstChild);
    readRecent().forEach(function (u) {
      var o = document.createElement("option");
      o.value = u;
      dl.appendChild(o);
    });
  }
  function fillPresets() {
    var list = Array.isArray(cfg.presets) ? cfg.presets : [];
    var box = $("presets");
    list.forEach(function (p) {
      if (!p || !p.name || !p.url) return;
      var b = document.createElement("button");
      b.type = "button";
      b.textContent = p.name;
      b.title = p.url;
      b.addEventListener("click", function () {
        input.value = p.url;
        updateForm();
        if (cfg.hasCode(p.url)) { codeInput.focus(); } else { input.focus(); }
      });
      box.appendChild(b);
    });
    if (box.querySelector("button")) show(box);
  }

  // ---------- Налаштування інструментів тренера у формі ----------
  // Для кожного інструмента: які рядки форми показати і як підписати спільні поля
  // (wLines — список, wSub — вибір, wNum — число).
  var WCFG = {
    protimer: { noTitle: true, rows: [],
      hint: ["Налаштуєте прямо в блоці після «Показати на слайді»: кнопка з повзунками відкриває налаштування. Під час показу видно лише циферблат.", "Set it up right in the block after Show on slide: the sliders button opens the settings. The slide show shows only the dial."] },
    timer: { title: ["Заголовок (необов'язково)", "Heading (optional)"], rows: ["wTimerRow"],
      hint: ["Клавіші після кліку по блоку: пробіл — старт/пауза, + — ще хвилина.", "Keys after clicking the block: Space start/pause, + adds a minute."] },
    "break": { title: ["Заголовок", "Heading"], titlePh: ["Перерва", "Break"], rows: ["wTimeRow"],
      hint: ["Відлік іде до вказаного часу за годинником комп'ютера, тож він правильний, хоч би коли відкрили слайд.", "Counts down to the clock time on this computer, so it's right whenever the slide is opened."] },
    picker: { title: ["Заголовок (необов'язково)", "Heading (optional)"], rows: ["wPickerRow"],
      hint: ["Клавіші: пробіл або Enter — вибрати. Колесо вміщує до 40 імен.", "Keys: Space or Enter picks. The wheel fits up to 40 names."] },
    score: { title: ["Заголовок (необов'язково)", "Heading (optional)"], rows: ["wScoreRow"],
      hint: ["Клавіші: 1–6 додають бали команді, Shift+цифра віднімає. Спільний рахунок живе на цьому комп'ютері 12 годин після останньої зміни (або до «Обнулити») і не потрапляє у файл.", "Keys: 1–6 add points to a team, Shift+digit subtracts. A shared score lives on this computer for 12 hours after the last change (or until Reset) and is not saved in the file."] },
    groups: { title: ["Заголовок (необов'язково)", "Heading (optional)"], lines: ["Учасники, по одному на рядок", "Participants, one per line"], linesKey: "names",
      ph: ["Олена\nАндрій\nМарта\nІгор", "Anna\nMark\nLucy\nTom"],
      sub: { key: "by", label: ["Ділити за", "Split by"], opts: [["count", "кількістю груп", "number of groups"], ["size", "розміром групи", "group size"]] },
      num: { key: "n", label: ["Скільки груп / людей у групі", "Groups / people per group"], def: 4, min: 1, max: 100 },
      hint: ["Клавіші: пробіл — перемішати ще раз.", "Keys: Space reshuffles."] },
    vote: { title: ["Питання", "Question"], lines: ["Варіанти, по одному на рядок (2–6)", "Options, one per line (2–6)"], linesKey: "options", min: 2,
      ph: ["Так\nНі\nВагаюся", "Yes\nNo\nNot sure"],
      hint: ["Тренер рахує підняті руки й натискає «+». Клавіші: 1–6 додають голос, Shift+цифра віднімає.", "Count raised hands and press +. Keys: 1–6 add a vote, Shift+digit subtracts."] },
    quiz: { title: ["Питання", "Question"], lines: ["Варіанти (2–6); правильний позначте * на початку", "Options (2–6); mark the correct one with * at the start"], linesKey: "options", min: 2,
      ph: ["Київ\n*Харків\nЛьвів", "Paris\n*Rome\nMadrid"],
      num: { key: "n", label: ["Час на відповідь, с (0 — без таймера)", "Answer time, s (0 = no timer)"], def: 30, min: 0, max: 600 },
      hint: ["Клавіші: Enter — показати відповідь, пробіл — таймер.", "Keys: Enter shows the answer, Space starts the timer."] },
    agenda: { title: ["Заголовок (необов'язково)", "Heading (optional)"], lines: ["Етапи: Назва | хвилини, по одному на рядок", "Stages: Name | minutes, one per line"], linesKey: "steps",
      ph: ["Обговорення в групах | 10\nПрезентації груп | 15\nПідсумки | 5", "Group discussion | 10\nGroup presentations | 15\nWrap-up | 5"],
      hint: ["Таймер сам переходить до наступного етапу. Клавіші: пробіл — старт/пауза, стрілки — етапи.", "Moves to the next stage by itself. Keys: Space start/pause, arrows switch stages."] },
    cards: { title: ["Заголовок (необов'язково)", "Heading (optional)"], lines: ["Картки, по одній на рядок", "Cards, one per line"], linesKey: "cards",
      ph: ["Клієнт незадоволений терміном\nКолега не виконав свою частину", "The client is unhappy with the deadline\nA colleague missed their part"],
      hint: ["Кейси, питання чи ролі. Клавіші: пробіл — витягнути картку.", "Cases, questions or roles. Keys: Space draws a card."] },
    dice: { title: ["Заголовок (необов'язково)", "Heading (optional)"],
      sub: { key: "dice", label: ["Що кидати", "What to roll"], opts: [["dice", "кубики", "dice"], ["coin", "монетку", "a coin"], ["number", "випадкове число", "a random number"]] },
      num: { key: "n", label: ["Кубиків (1–6) або до якого числа", "Dice (1–6) or the max number"], def: 1, min: 1, max: 1000000 },
      hint: ["Клавіші: пробіл — кинути.", "Keys: Space rolls."] }
  };
  var W_ROWS = ["wTimerRow", "wTimeRow", "wPickerRow", "wScoreRow", "wLinesRow", "wSubRow", "wNumRow"];
  function pair(p) { return p ? L(p[0], p[1]) : ""; }

  function fillWidgetForm(kind, w) {
    var c = WCFG[kind];
    if (!c) return;
    $("wTitle").value = w.title || "";
    $("wMinutes").value = w.minutes || 5;
    $("wTime").value = w.until || "14:30";
    $("wNames").value = w.names || "";
    $("wStyle").value = w.style === "wheel" ? "wheel" : "list";
    $("wRepeat").checked = !!w.repeat;
    $("wTeams").value = w.teams || "";
    $("wStep").value = w.step || 1;
    $("wShared").checked = !!w.shared;
    $("wLines").value = c.linesKey ? (w[c.linesKey] || "") : "";
    if (c.sub) setupSub(c, w[c.sub.key]);
    if (c.num) $("wNum").value = w[c.num.key] != null ? w[c.num.key] : c.num.def;
  }
  function setupSub(c, value) {
    var sel = $("wSub");
    while (sel.firstChild) sel.removeChild(sel.firstChild);
    c.sub.opts.forEach(function (o) {
      var opt = document.createElement("option");
      opt.value = o[0]; opt.textContent = L(o[1], o[2]);
      sel.appendChild(opt);
    });
    sel.value = value || c.sub.opts[0][0];
  }
  var lastKind = null;
  function updateWidgetForm(kind) {
    var c = WCFG[kind];
    if (!c) return;
    var need = (c.rows || []).slice();
    if (c.lines) need.push("wLinesRow");
    if (c.sub) need.push("wSubRow");
    if (c.num) need.push("wNumRow");
    W_ROWS.forEach(function (id) { if (need.indexOf(id) !== -1) show($(id)); else hide($(id)); });
    if (c.noTitle) hide($("wTitleLabel"), $("wTitle")); else show($("wTitleLabel"), $("wTitle"));
    $("wTitleLabel").textContent = pair(c.title);
    $("wTitle").setAttribute("placeholder", c.titlePh ? pair(c.titlePh) : "");
    if (c.lines) {
      $("wLinesLabel").textContent = pair(c.lines);
      $("wLines").setAttribute("placeholder", c.ph ? pair(c.ph).replace(/\\n/g, "\n") : "");
    }
    if (c.sub) {
      $("wSubLabel").textContent = pair(c.sub.label);
      if (lastKind !== kind) setupSub(c, $("wSub").value);
    }
    if (c.num) {
      $("wNumLabel").textContent = pair(c.num.label);
      if (kind === "dice") {
        var sub = $("wSub").value;
        $("wNumLabel").textContent = sub === "number" ? L("До якого числа", "Up to which number") : L("Скільки кубиків (1–6)", "How many dice (1–6)");
        if (sub === "coin") hide($("wNumRow")); else show($("wNumRow"));
      }
      $("wNum").min = c.num.min; $("wNum").max = c.num.max;
      if (lastKind !== kind && lastKind !== null && $("wNum").value === "") $("wNum").value = c.num.def;
    }
    $("wHint").textContent = pair(c.hint);
    lastKind = kind;
  }
  // Збирає налаштування інструмента з форми; повертає { w } або { error, field }.
  function collectWidget(kind) {
    var c = WCFG[kind], w = kind === "protimer" ? {} : { title: $("wTitle").value.trim().slice(0, 120) };
    if (kind === "timer") {
      var m = parseFloat($("wMinutes").value);
      if (!(m >= 0.25 && m <= 180)) return { error: L("Тривалість від 0,25 до 180 хвилин.", "Duration must be 0.25–180 minutes."), field: $("wMinutes") };
      w.minutes = m;
    } else if (kind === "break") {
      if (!/^\d{1,2}:\d{2}$/.test($("wTime").value)) return { error: L("Вкажіть час повернення.", "Enter the time to come back."), field: $("wTime") };
      w.until = $("wTime").value;
    } else if (kind === "picker") {
      w.names = $("wNames").value.slice(0, 5000);
      w.repeat = $("wRepeat").checked;
      w.style = $("wStyle").value;
      if (!w.names.trim()) return { error: L("Додайте хоча б одне ім'я.", "Add at least one name."), field: $("wNames") };
    } else if (kind === "score") {
      w.teams = $("wTeams").value.slice(0, 1000);
      w.step = Math.max(1, Math.min(100, parseInt($("wStep").value, 10) || 1));
      w.shared = $("wShared").checked;
    }
    if (c && c.linesKey) {
      var text = $("wLines").value.slice(0, 8000);
      var count = text.split(/\r?\n/).filter(function (l) { return l.trim(); }).length;
      if (count < (c.min || 1)) return { error: L("Додайте щонайменше " + (c.min || 1) + " рядки у список.", "Add at least " + (c.min || 1) + " lines to the list.").replace(" 1 рядки", " 1 рядок"), field: $("wLines") };
      w[c.linesKey] = text;
    }
    if (c && c.sub) w[c.sub.key] = $("wSub").value;
    if (c && c.num) {
      var n = parseFloat($("wNum").value);
      if (kind === "dice" && w.dice === "dice" && !(n >= 1 && n <= 6)) return { error: L("Кубиків може бути від 1 до 6.", "You can roll 1 to 6 dice."), field: $("wNum") };
      if (kind === "dice" && w.dice === "number" && !(n >= 2)) return { error: L("Число має бути щонайменше 2.", "The number must be at least 2."), field: $("wNum") };
      if (!(n >= c.num.min && n <= c.num.max)) return { error: pair(c.num.label) + L(": від ", ": from ") + c.num.min + L(" до ", " to ") + c.num.max + ".", field: $("wNum") };
      w[c.num.key] = n;
    }
    return { w: w };
  }

  // ---------- Мої шаблони блоків (лише на цьому комп'ютері) + шаблони організації з config.js ----------
  function readTemplates() {
    var own = [];
    try { own = JSON.parse(localStorage.getItem(TEMPLATES_KEY) || "[]"); } catch (e) { own = []; }
    if (!Array.isArray(own)) own = [];
    var org = Array.isArray(cfg.blockTemplates) ? cfg.blockTemplates : [];
    return { own: own.filter(function (t) { return t && t.name && t.data; }), org: org.filter(function (t) { return t && t.name && t.data; }) };
  }
  function writeTemplates(list) {
    try { localStorage.setItem(TEMPLATES_KEY, JSON.stringify(list.slice(0, 50))); return true; } catch (e) { return false; }
  }
  function fillTemplates() {
    var sel = $("tplSel"), t = readTemplates();
    while (sel.firstChild) sel.removeChild(sel.firstChild);
    var first = document.createElement("option");
    first.value = ""; first.textContent = L("— вибрати —", "— choose —");
    sel.appendChild(first);
    function add(list, prefix, group) {
      list.forEach(function (tpl, i) {
        var o = document.createElement("option");
        o.value = prefix + i; o.textContent = tpl.name;
        group.appendChild(o);
      });
    }
    if (t.org.length) { var g1 = document.createElement("optgroup"); g1.label = L("Організації", "Organization"); add(t.org, "o", g1); sel.appendChild(g1); }
    if (t.own.length) { var g2 = document.createElement("optgroup"); g2.label = L("Мої", "Mine"); add(t.own, "m", g2); sel.appendChild(g2); }
    if (t.org.length || t.own.length) show($("tplRow")); else hide($("tplRow"));
    $("tplDel").disabled = true;
  }
  function templateName(cur) {
    if (cur.kind !== "site") {
      var opt = kindSel.querySelector('option[value="' + cur.kind + '"]');
      return (opt ? opt.textContent : cur.kind) + (cur.widget && cur.widget.title ? ": " + cur.widget.title : "");
    }
    return hostOf(cfg.normalize(cfg.fillCode(cur.url, "x"))) + (cur.tabs.length ? " +" + cur.tabs.length : "") + (cur.mode !== "embed" ? " · " + cur.mode : "");
  }

  // ---------- Форма ----------
  function fillForm(cur) {
    kindSel.value = cur ? cur.kind : "site";
    formTimer = cur && cur.timer ? cur.timer : "";
    var w = cur && cur.widget ? cur.widget : {};
    lastKind = null;
    if (cur && cur.kind !== "site") fillWidgetForm(cur.kind, w);
    $("rotate").value = String(cur ? cur.rotate || 0 : 0);
    if (current && !editingFromTemplate) show($("tplSave")); else hide($("tplSave"));
    fillTemplates();
    autoplayBox.checked = cur ? cur.autoplay : false;
    input.value = cur ? cur.url : "";
    codeInput.value = cur ? cur.code : "";
    modeSel.value = cur ? cur.mode : "embed";
    scaleSel.value = String(cur ? cur.scale : 1);
    if (scaleSel.selectedIndex < 0) scaleSel.value = "1";
    refreshSel.value = String(cur ? cur.refresh : 0);
    allowScroll.checked = cur ? cur.scroll === "show" : false;
    qrBox.checked = cur ? cur.qrCorner : false;
    revealBox.checked = cur ? cur.reveal : false;
    revealText.value = cur ? cur.revealText : "";
    zoomSel.value = String(cur ? cur.zoom : 1);
    if (zoomSel.selectedIndex < 0) zoomSel.value = "1";
    fxIn.value = cur ? cur.fx : 50;
    fyIn.value = cur ? cur.fy : 0;
    offlineIn.value = cur ? cur.offline : "";
    tabsText.value = cur ? tabsToText(cur.tabs) : "";
    mainTitle.value = cur ? cur.mainTitle : "";
    borderSel.value = cur ? cur.border : "";
    roundedBox.checked = cur ? cur.rounded : false;
    cornerBg.value = cur ? cur.cornerBg : "#ffffff";
    // Розгортаємо «Додатково», якщо там щось змінено від типового.
    more.open = !!(cur && (cur.mode !== "embed" || cur.scale !== 1 || cur.scroll === "show" ||
      cur.qrCorner || cur.reveal || cur.refresh || cur.zoom !== 1 || cur.offline || cur.tabs.length || cur.border || cur.rounded || cur.autoplay || cur.rotate));
    warn.textContent = "";
    setError("");
    updateForm();
  }

  // Показує поле коду, підказку про перетворення адреси і вмикає лише доречні для режиму параметри.
  function updateForm() {
    var kind = kindSel.value;
    var site = kind === "site";
    if (site) { show(siteFields, more, $("urlHint")); hide(widgetFields); }
    else { hide(siteFields, more, $("urlHint")); show(widgetFields); }
    if (!site) updateWidgetForm(kind);
    var templ = cfg.normalize(input.value);
    var withCode = cfg.hasCode(templ) || cfg.hasCode(tabsText.value);
    if (withCode) show(codeRow); else hide(codeRow);
    if (tabsText.value.trim()) show($("mainTitleRow")); else hide($("mainTitleRow"));
    if (zoomSel.value !== "1" && modeSel.value === "embed") show($("zoomPos")); else hide($("zoomPos"));
    if (roundedBox.checked) show($("cornerBgRow")); else hide($("cornerBgRow"));

    var mode = modeSel.value;
    var embed = mode === "embed";
    scaleSel.disabled = !embed;
    refreshSel.disabled = !embed;
    allowScroll.disabled = !embed;
    revealBox.disabled = !embed;
    autoplayBox.disabled = !embed;
    zoomSel.disabled = !embed;
    tabsText.disabled = !embed;
    $("rotate").disabled = !embed || !tabsText.value.trim();
    qrBox.disabled = mode === "qr";
    if (embed && revealBox.checked) show($("revealRow")); else hide($("revealRow"));
    $("buttonHint").textContent = mode === "button"
      ? L("На слайді буде велика кнопка, що відкриває сайт у вікні. Для сайтів, які забороняють вбудовування.", "The slide shows a big button that opens the site in a window. For sites that forbid embedding.")
      : mode === "qr"
        ? L("На слайді буде лише великий QR-код і коротка адреса: учасники відкривають сайт на телефонах.", "The slide shows only a big QR code and a short address: people open the site on their phones.")
        : "";

    var res = cfg.check(cfg.fillCode(templ, withCode ? (codeInput.value || "x") : ""));
    warn.textContent = res.ok && res.warnSecret
      ? L("Схоже, адреса містить токен або ключ. Вона збережеться у файлі презентації відкритим текстом.", "The address seems to contain a token or key. It will be saved in the presentation file as plain text.")
      : "";
    var conv = res.ok && embed ? cfg.toEmbed(res.url) : { service: null };
    info.textContent = conv.service && cfg.check(conv.url).ok
      ? L("Адресу " + conv.service + " буде показано у форматі для вбудовування.", "The " + conv.service + " address will be shown in its embeddable form.")
      : "";
  }

  // ---------- Запасний варіант: діалог Office ----------
  function showFallback(url) {
    $("fallbackUrl").textContent = url;
    show(fallback);
  }

  function openInDialog(url) {
    if (!cfg.check(url).ok) return;
    if (!dialogSupported) { showFallback(url); return; }
    var dialogUrl = new URL("dialog.html", location.href);
    dialogUrl.searchParams.set("u", url);
    Office.context.ui.displayDialogAsync(
      dialogUrl.href,
      // У PowerPoint для вебу діалог = нове вікно браузера. Проміжне підтвердження від Office
      // дає «дію користувача», тож блокувальник спливних вікон його пропускає.
      { height: 85, width: 85, promptBeforeOpen: isWeb },
      function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) return;
        var code = res.error && res.error.code;
        if (code === 12007) return; // діалог уже відкритий
        if (code === 12009) return; // користувач сам відмовився відкривати
        // Інше (напр. заблоковано спливне вікно): показуємо адресу для копіювання.
        showFallback(url);
      }
    );
  }

  // ---------- Ініціалізація ----------
  Office.onReady(function (ready) {
    if (!ready || !ready.host) {
      showFatal("Ця сторінка працює лише всередині PowerPoint як надбудова «Веб на слайді». / This page only works inside PowerPoint as the “Web on Slide” add-in.");
      return;
    }

    var req = Office.context.requirements;
    dialogSupported = !!(req && req.isSetSupported("DialogApi", "1.1") && Office.context.ui && Office.context.ui.displayDialogAsync);
    isWeb = !!(Office.PlatformType && Office.context.platform === Office.PlatformType.OfficeOnline);
    if (!dialogSupported) {
      hide($("popBtn"));
      // Режим «кнопка» без діалогу не має сенсу: лишаємо вбудовування і QR.
      modeSel.querySelector('option[value="button"]').disabled = true;
    }

    applyFrameCsp();

    if (window.WOS_I18N) {
      window.WOS_I18N.init(cfg.language, Office.context.displayLanguage);
      if (lang() === "en") $("helpLink").setAttribute("href", "help-en.html");
    }

    var settings = Office.context.document.settings;
    var savedUrl = settings.get(KEY_URL) || "";
    var savedCode = settings.get(KEY_CODE) || "";
    var savedKind = settings.get(KEY_KIND);
    var kind = validKind(savedKind) ? savedKind : "site";
    var savedWidget = null;
    try { savedWidget = JSON.parse(settings.get(KEY_WIDGET) || "null"); } catch (e) { savedWidget = null; }
    var siteOk = !!savedUrl && cfg.check(cfg.fillCode(savedUrl, savedCode || "x")).ok;
    if (siteOk || kind !== "site") {
      current = {
        kind: kind,
        widget: savedWidget && typeof savedWidget === "object" ? savedWidget : {},
        timer: String(settings.get(KEY_TIMER) || ""),
        url: siteOk ? savedUrl : "",
        code: String(savedCode),
        autoplay: settings.get(KEY_AUTOPLAY) === "1",
        mode: cleanMode(settings.get(KEY_MODE)),
        scale: cleanScale(settings.get(KEY_SCALE)),
        scroll: settings.get(KEY_SCROLL) === "show" ? "show" : "hide",
        qrCorner: settings.get(KEY_QR) === "corner",
        reveal: settings.get(KEY_REVEAL) === "1",
        revealText: String(settings.get(KEY_REVEAL_TEXT) || ""),
        refresh: cleanRefresh(settings.get(KEY_REFRESH)),
        zoom: cleanZoom(settings.get(KEY_ZOOM)),
        fx: cleanPct(settings.get(KEY_FX), 50),
        fy: cleanPct(settings.get(KEY_FY), 0),
        offline: String(settings.get(KEY_OFFLINE) || ""),
        tabs: cleanTabs(settings.get(KEY_TABS)).filter(function (t) { return cfg.check(cfg.fillCode(t.u, savedCode || "x")).ok; }),
        mainTitle: String(settings.get(KEY_MAIN_TITLE) || ""),
        border: ["gray", "yellow", "dark"].indexOf(settings.get(KEY_BORDER)) !== -1 ? settings.get(KEY_BORDER) : "",
        rounded: settings.get(KEY_ROUNDED) === "1",
        cornerBg: cleanColor(settings.get(KEY_CORNER_BG)),
        rotate: cleanRotate(settings.get(KEY_ROTATE))
      };
      if (needsCode(current)) fillForm(current);
    } else {
      updateForm();
    }

    // Міст для «Таймера PRO». Під час показу зміни не пишемо у файл: бали чи ростер, змінені перед залом,
    // живуть до кінця показу, а збережене налаштування слайда лишається таким, як його зробив автор.
    window.WOS_TIMER_HOST = {
      cfg: current && current.kind === "protimer" ? current.timer : "",
      view: view,
      save: function (json) {
        if (view !== "edit" || !current || current.kind !== "protimer" || typeof json !== "string" || json.length > 200000) return;
        current.timer = json;
        window.WOS_TIMER_HOST.cfg = json;
        settings.set(KEY_TIMER, json);
        clearTimeout(timerSaveT);
        timerSaveT = setTimeout(function () { settings.saveAsync(function () { /* наступна зміна спробує ще раз */ }); }, 800);
      }
    };

    var activeViewSupported = !!(req && req.isSetSupported("ActiveView", "1.1"));
    if (activeViewSupported) {
      Office.context.document.getActiveViewAsync(function (r) {
        view = r.status === Office.AsyncResultStatus.Succeeded ? r.value : "edit";
        render();
      });
      Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, function (e) {
        // Кожен новий показ починається із заставки, навіть якщо в редакторі її вже відкрили.
        if (e.activeView !== view) {
          revealed = false;
          activeTab = 0;
          if (window.WOS_WIDGETS) window.WOS_WIDGETS.unmount(widgetBox);
          // Таймер PRO починає показ з налаштувань слайда, а не з того, що натиснули в редакторі.
          if (current && current.kind === "protimer") { window.WOS_TIMER_HOST.cfg = current.timer; unloadTimer(); }
        }
        view = e.activeView;
        render();
      });
    } else {
      view = "edit";
      render();
    }

    input.addEventListener("input", function () {
      if (err.textContent) setError("");
      updateForm();
    });
    codeInput.addEventListener("input", function () {
      if (err.textContent) setError("");
      updateForm();
    });
    modeSel.addEventListener("change", updateForm);
    kindSel.addEventListener("change", function () {
      setError("");
      if (kindSel.value !== "site" && lastKind !== kindSel.value) fillWidgetForm(kindSel.value, current && current.kind === kindSel.value ? current.widget : {});
      updateForm();
    });
    // Кубики й випадкове число ділять одне поле: при перемиканні ставимо доречне значення.
    $("wSub").addEventListener("change", function () {
      if (kindSel.value !== "dice") return;
      var n = parseFloat($("wNum").value);
      if ($("wSub").value === "number" && !(n > 6)) $("wNum").value = 100;
      if ($("wSub").value === "dice" && !(n >= 1 && n <= 6)) $("wNum").value = 1;
      updateWidgetForm("dice");
    });
    tabsText.addEventListener("input", function () { $("rotate").disabled = modeSel.value !== "embed" || !tabsText.value.trim(); });
    $("tplSel").addEventListener("change", function () {
      var v = $("tplSel").value, t = readTemplates();
      $("tplDel").disabled = v.charAt(0) !== "m";
      if (!v) return;
      var tpl = (v.charAt(0) === "o" ? t.org : t.own)[+v.slice(1)];
      if (!tpl) return;
      var d = tpl.data && typeof tpl.data === "object" ? tpl.data : {};
      // Шаблон заповнює форму; зберігається він лише після «Показати на слайді».
      var keep = $("tplSel").value;
      editingFromTemplate = true;
      fillForm({
        kind: validKind(d.kind) ? d.kind : "site",
        widget: d.widget && typeof d.widget === "object" ? d.widget : {}, timer: typeof d.timer === "string" ? d.timer : "",
        url: d.url || "", code: d.code || "",
        autoplay: !!d.autoplay, mode: cleanMode(d.mode), scale: cleanScale(d.scale), scroll: d.scroll === "show" ? "show" : "hide",
        qrCorner: !!d.qrCorner, reveal: !!d.reveal, revealText: d.revealText || "", refresh: cleanRefresh(d.refresh),
        zoom: cleanZoom(d.zoom), fx: cleanPct(d.fx, 50), fy: cleanPct(d.fy, 0), offline: d.offline || "", tabs: cleanTabs(d.tabs),
        mainTitle: d.mainTitle || "", border: d.border || "", rounded: !!d.rounded, cornerBg: cleanColor(d.cornerBg), rotate: cleanRotate(d.rotate)
      });
      editingFromTemplate = false;
      $("tplSel").value = keep;
      $("tplDel").disabled = keep.charAt(0) !== "m";
    });
    $("tplDel").addEventListener("click", function () {
      var v = $("tplSel").value;
      if (v.charAt(0) !== "m") return;
      var own = readTemplates().own;
      own.splice(+v.slice(1), 1);
      writeTemplates(own);
      fillTemplates();
    });
    $("tplSave").addEventListener("click", function () {
      if (!current) return;
      var own = readTemplates().own;
      var name = templateName(current);
      own = own.filter(function (t) { return t.name !== name; });
      own.unshift({ name: name, data: current });
      if (writeTemplates(own)) {
        fillTemplates();
        $("tplSave").textContent = L("Збережено як «", "Saved as “") + name + L("»", "”");
        setTimeout(function () { $("tplSave").textContent = L("Зберегти як шаблон", "Save as template"); }, 2500);
      }
    });
    revealBox.addEventListener("change", updateForm);
    zoomSel.addEventListener("change", updateForm);
    roundedBox.addEventListener("change", updateForm);
    tabsText.addEventListener("input", function () { if (err.textContent) setError(""); updateForm(); });
    fillPresets();
    fillRecent();

    function saveAll(next) {
      settings.set(KEY_KIND, next.kind);
      settings.set(KEY_TIMER, next.timer || "");
      settings.set(KEY_WIDGET, JSON.stringify(next.widget || {}));
      settings.set(KEY_URL, next.url);
      settings.set(KEY_CODE, next.code);
      settings.set(KEY_AUTOPLAY, next.autoplay ? "1" : "");
      settings.set(KEY_MODE, next.mode);
      settings.set(KEY_SCALE, next.scale);
      settings.set(KEY_SCROLL, next.scroll);
      settings.set(KEY_QR, next.qrCorner ? "corner" : "");
      settings.set(KEY_REVEAL, next.reveal ? "1" : "");
      settings.set(KEY_REVEAL_TEXT, next.revealText);
      settings.set(KEY_REFRESH, next.refresh);
      settings.set(KEY_ZOOM, next.zoom);
      settings.set(KEY_FX, next.fx);
      settings.set(KEY_FY, next.fy);
      settings.set(KEY_OFFLINE, next.offline);
      settings.set(KEY_TABS, JSON.stringify(next.tabs));
      settings.set(KEY_MAIN_TITLE, next.mainTitle);
      settings.set(KEY_BORDER, next.border);
      settings.set(KEY_ROUNDED, next.rounded ? "1" : "");
      settings.set(KEY_CORNER_BG, next.cornerBg);
      settings.set(KEY_ROTATE, next.rotate || 0);
      settings.saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          setError(L("Не вдалося зберегти налаштування у файлі: ", "Couldn't save the settings in the file: ") + (r.error && r.error.message ? r.error.message : L("невідома помилка", "unknown error")) + L(". Перевірте, чи файл не відкрито лише для читання.", ". Check that the file isn't read-only."));
          return;
        }
        current = next;
        window.WOS_TIMER_HOST.cfg = next.kind === "protimer" ? (next.timer || "") : "";
        unloadTimer();
        editing = false;
        revealed = false;
        activeTab = 0;
        if (next.kind === "site") rememberRecent(next.url);
        unloadFrame();
        render();
      });
    }

    setup.addEventListener("submit", function (e) {
      e.preventDefault();
      if (kindSel.value !== "site") {
        var base = current || {};
        var k = kindSel.value;
        var got = collectWidget(k);
        if (got.error) { setError(got.error, got.field); got.field.focus(); return; }
        var w = got.w;
        setError("");
        saveAll({
          kind: k, widget: w, timer: k === "protimer" ? (formTimer || base.timer || "") : "", url: base.url || "", code: base.code || "", autoplay: !!base.autoplay,
          mode: base.mode || "embed", scale: base.scale || 1, scroll: base.scroll || "hide", qrCorner: !!base.qrCorner,
          reveal: !!base.reveal, revealText: base.revealText || "", refresh: base.refresh || 0, zoom: base.zoom || 1,
          fx: base.fx == null ? 50 : base.fx, fy: base.fy || 0, offline: base.offline || "", tabs: base.tabs || [],
          mainTitle: base.mainTitle || "", border: base.border || "", rounded: !!base.rounded, cornerBg: base.cornerBg || "#ffffff",
          rotate: base.rotate || 0
        });
        return;
      }
      var templ = cfg.normalize(input.value);
      var withCode = cfg.hasCode(templ) || cfg.hasCode(tabsText.value);
      var code = withCode ? codeInput.value.replace(/\s+/g, "") : "";
      var embedMode = modeSel.value === "embed";
      var parsed = embedMode ? parseTabs(tabsText.value, code) : { tabs: [], error: null };
      var res = cfg.check(cfg.fillCode(templ, code || "x"));
      if (!res.ok) {
        setError(reason(res.reason));
        input.focus();
        return;
      }
      if (parsed.error) {
        setError(parsed.error, tabsText);
        more.open = true;
        tabsText.focus();
        return;
      }
      if (withCode && !code) {
        setError(L("Введіть код сесії.", "Enter the session code."), codeInput);
        codeInput.focus();
        return;
      }
      setError("");
      var next = {
        kind: "site",
        widget: {},
        timer: "",
        url: templ,
        code: code,
        autoplay: autoplayBox.checked && modeSel.value === "embed",
        mode: cleanMode(modeSel.value),
        scale: cleanScale(scaleSel.value),
        scroll: allowScroll.checked ? "show" : "hide",
        qrCorner: qrBox.checked && modeSel.value !== "qr",
        reveal: revealBox.checked && modeSel.value === "embed",
        revealText: revealText.value.trim(),
        refresh: cleanRefresh(refreshSel.value),
        zoom: embedMode ? cleanZoom(zoomSel.value) : 1,
        fx: cleanPct(fxIn.value, 50),
        fy: cleanPct(fyIn.value, 0),
        offline: offlineIn.value.trim().slice(0, 160),
        tabs: parsed.tabs,
        mainTitle: mainTitle.value.trim().slice(0, 30),
        border: borderSel.value,
        rounded: roundedBox.checked,
        cornerBg: cleanColor(cornerBg.value),
        rotate: parsed.tabs.length ? cleanRotate($("rotate").value) : 0
      };
      saveAll(next);
    });

    function openForm(focusEl) {
      fillForm(current);
      editing = true;
      render();
      focusEl.focus();
      if (focusEl.select) focusEl.select();
    }
    $("editBtn").addEventListener("click", function () { openForm(current && current.kind !== "site" ? kindSel : input); });
    $("codeBtn").addEventListener("click", function () { openForm(codeInput); });
    function cancelEdit() {
      if (!current || !editing || needsCode(current)) return;
      editing = false;
      setError("");
      render();
    }
    $("cancelBtn").addEventListener("click", cancelEdit);
    setup.addEventListener("keydown", function (e) {
      if (e.key === "Escape") { e.preventDefault(); cancelEdit(); }
    });
    $("coverBtn").addEventListener("click", function () { revealed = true; render(); });
    $("reloadBtn").addEventListener("click", function () { if (current) loadFrame(frameUrl(current)); });
    $("loadfailToButton").addEventListener("click", function () {
      if (!current) return;
      settings.set(KEY_MODE, "button");
      settings.saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return;
        current.mode = "button";
        render();
      });
    });
    $("loadfailRetry").addEventListener("click", function () { if (current) loadFrame(frameUrl(current)); });
    $("popBtn").addEventListener("click", function () { if (current) openInDialog(targetUrl(current)); });
    $("launchBtn").addEventListener("click", function () { if (current) openInDialog(targetUrl(current)); });
    $("loadfailOpen").addEventListener("click", function () { if (current) openInDialog(targetUrl(current)); });
    $("fallbackClose").addEventListener("click", function () { hide(fallback); });
    $("copyBtn").addEventListener("click", function () {
      var text = $("fallbackUrl").textContent;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () {
          $("copyBtn").textContent = L("Скопійовано", "Copied");
          setTimeout(function () { $("copyBtn").textContent = L("Скопіювати адресу", "Copy address"); }, 1500);
        }, copyFailed);
      } else {
        copyFailed();
      }
      function copyFailed() {
        // Адреса виділяється одним кліком (user-select: all), тож підказуємо скопіювати вручну.
        $("copyBtn").textContent = L("Не вдалося. Виділіть адресу й натисніть Ctrl+C", "Failed. Select the address and press Ctrl+C");
        setTimeout(function () { $("copyBtn").textContent = L("Скопіювати адресу", "Copy address"); }, 3000);
      }
    });
  });
})();
