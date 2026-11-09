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
  var KEY_CODE_SHARED = "codeShared"; // "1" = код береться зі спільного тегу презентації
  var KEY_KIND = "kind";       // "site" | "timer" | "picker" | "score"
  var KEY_WIDGET = "widget";   // JSON налаштувань інструмента тренера
  var SHARED_CODE_TAG = "WEBONSLIDE_CODE"; // тег презентації зі спільним кодом сесії

  function L(uk, en) { return window.WOS_I18N ? window.WOS_I18N.L(uk, en) : uk; }
  function lang() { return window.WOS_I18N ? window.WOS_I18N.lang() : "uk"; }
  function reason(code) { return cfg.reason ? cfg.reason(code, lang()) : (cfg.reasonText[code] || "Ця адреса не дозволена."); }
  function defaultRevealText() { return L("Натисніть, щоб показати", "Click to show"); }

  var cfg = window.WEB_ON_SLIDE_CONFIG;

  function $(id) { return document.getElementById(id); }
  var frame = $("frame"), frameWrap = $("frameWrap"), launch = $("launch"), setup = $("setup"), bar = $("bar");
  var loadfail = $("loadfail"), fallback = $("fallback"), fatal = $("fatal"), loading = $("loading");
  var cover = $("cover"), qrCard = $("qrCard"), qrCorner = $("qrCorner");
  var input = $("url"), codeInput = $("code"), codeRow = $("codeRow"), info = $("info"), err = $("err"), warn = $("warn");
  var modeSel = $("mode"), scaleSel = $("scale"), refreshSel = $("refresh"), allowScroll = $("allowScroll");
  var qrBox = $("qrCornerOn"), revealBox = $("reveal"), revealText = $("revealText"), more = $("more");
  var zoomSel = $("zoom"), fxIn = $("fx"), fyIn = $("fy"), offlineIn = $("offlineText"), tabsText = $("tabsText"), mainTitle = $("mainTitle");
  var kindSel = $("kind"), siteFields = $("siteFields"), widgetFields = $("widgetFields"), widgetBox = $("widget");
  var autoplayBox = $("autoplay"), codeShared = $("codeShared");
  var borderSel = $("border"), roundedBox = $("rounded"), cornerBg = $("cornerBg"), tabsBar = $("tabs"), frameZoom = $("frameZoom");

  var view = null;        // null (ще невідомо) | "edit" | "read" (показ слайдів)
  var current = null;     // збережене налаштування цієї вставки (див. readSettings)
  var editing = false;    // форма відкрита кнопкою «Змінити»; current при цьому зберігається
  var revealed = false;   // заставку вже прибрали кліком
  var activeTab = 0;      // 0 = основна адреса, 1..3 = додаткові вкладки
  var dialogSupported = false;
  var isWeb = false;
  var loadTimer = null;
  var refreshTimer = null;
  var frameLoaded = false;
  var qrDrawnFor = { corner: null, card: null };
  var sharedCodeOk = false;   // PowerPoint дозволяє читати й писати теги презентації (спільний код)

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
  function showQr(which, el, url) {
    if (qrDrawnFor[which] !== url) {
      if (!drawQr(which === "card" ? $("qrBig") : el, url)) return;
      qrDrawnFor[which] = url;
    }
    show(el);
  }

  // ---------- Фатальні стани: без office.js або поза PowerPoint ----------
  function showFatal(text) {
    hide(frameWrap, launch, setup, bar, loadfail, fallback, loading, cover, qrCard, qrCorner);
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
    meta.content = "frame-src " + allowed.join(" ") + ";";
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

  function defaultFailText() { return L("Сторінка не відповідає. Можливо, сайт забороняє вбудовування або немає мережі.", "The page is not responding. The site may forbid embedding, or there is no network."); }
  function showLoadFail(url) {
    hide(loading);
    var own = current && current.offline;
    $("loadfailText").textContent = own || defaultFailText();
    $("loadfailText").style.fontSize = own ? "clamp(16px, 5vmin, 40px)" : "";
    $("loadfailHost").textContent = own ? "" : hostOf(url);
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
    ["wMinutes", "wNames"].forEach(function (id) { $(id).removeAttribute("aria-invalid"); });
    if (text) (field || input).setAttribute("aria-invalid", "true");
  }

  function render() {
    hide(frameWrap, launch, setup, bar, loadfail, fallback, loading, cover, qrCard, qrCorner, tabsBar, widgetBox);
    document.body.style.background = "";
    if (view === null) return; // не малюємо нічого, поки не знаємо режим: інакше форма блимає під час показу

    var showForm = !current || needsCode(current) || (editing && view === "edit");
    if (showForm) {
      unloadFrame();
      if (view === "edit") {
        if (current && !needsCode(current)) show($("cancelBtn")); else hide($("cancelBtn"));
        show(setup);
        // Нова вставка: одразу ставимо курсор у поле адреси.
        if (!current) input.focus();
      }
      // Під час показу без налаштованого URL нічого не показуємо.
      return;
    }

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
    if (current.mode === "button") {
      // Режим кнопки: iframe не використовується, тож «Оновити» лише викликало б хибне «не відповідає».
      unloadFrame();
      $("launchHost").textContent = hostOf(target);
      show(launch);
    } else if (current.mode === "qr") {
      unloadFrame();
      $("qrUrl").textContent = shortUrl(target);
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
      if (current.tabs.length) renderTabs();
    }
    if (current.qrCorner && current.mode !== "qr") showQr("corner", qrCorner, target);

    if (view === "edit") {
      if (reloadable) show($("reloadBtn")); else hide($("reloadBtn"));
      if (dialogSupported) show($("popBtn"));
      if (usesCode(current)) show($("codeBtn")); else hide($("codeBtn"));
      show(bar);
    }
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

  // ---------- Форма ----------
  function fillForm(cur) {
    kindSel.value = cur ? cur.kind : "site";
    var w = cur && cur.widget ? cur.widget : {};
    $("wTitle").value = w.title || "";
    $("wMinutes").value = w.minutes || 5;
    $("wNames").value = w.names || "";
    $("wRepeat").checked = !!w.repeat;
    $("wTeams").value = w.teams || "";
    $("wStep").value = w.step || 1;
    autoplayBox.checked = cur ? cur.autoplay : false;
    codeShared.checked = cur ? cur.codeShared : false;
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
      cur.qrCorner || cur.reveal || cur.refresh || cur.zoom !== 1 || cur.offline || cur.tabs.length || cur.border || cur.rounded || cur.autoplay));
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
    if (kind === "timer") show($("wTimerRow")); else hide($("wTimerRow"));
    if (kind === "picker") show($("wPickerRow")); else hide($("wPickerRow"));
    if (kind === "score") show($("wScoreRow")); else hide($("wScoreRow"));
    if (sharedCodeOk) show($("codeSharedRow")); else hide($("codeSharedRow"));
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
    var kind = savedKind === "timer" || savedKind === "picker" || savedKind === "score" ? savedKind : "site";
    var savedWidget = null;
    try { savedWidget = JSON.parse(settings.get(KEY_WIDGET) || "null"); } catch (e) { savedWidget = null; }
    var siteOk = !!savedUrl && cfg.check(cfg.fillCode(savedUrl, savedCode || "x")).ok;
    if (siteOk || kind !== "site") {
      current = {
        kind: kind,
        widget: savedWidget && typeof savedWidget === "object" ? savedWidget : {},
        url: siteOk ? savedUrl : "",
        code: String(savedCode),
        codeShared: settings.get(KEY_CODE_SHARED) === "1",
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
        cornerBg: cleanColor(settings.get(KEY_CORNER_BG))
      };
      if (needsCode(current)) fillForm(current);
    } else {
      updateForm();
    }

    // ---------- Спільний код для всієї презентації ----------
    // Зберігається в тезі презентації (PowerPoint API 1.3). Якщо PowerPoint цього не дозволяє
    // (старіша версія або обмежені дозволи надбудови), галочка просто не показується.
    function readSharedCode() {
      return PowerPoint.run(function (ctx) {
        var tag = ctx.presentation.tags.getItemOrNullObject(SHARED_CODE_TAG);
        tag.load("value");
        return ctx.sync().then(function () { return tag.isNullObject ? null : String(tag.value || ""); });
      });
    }
    function writeSharedCode(code) {
      return PowerPoint.run(function (ctx) {
        ctx.presentation.tags.add(SHARED_CODE_TAG, code);
        return ctx.sync();
      });
    }
    if (typeof PowerPoint !== "undefined" && PowerPoint.run && req && req.isSetSupported("PowerPointApi", "1.3")) {
      try {
        readSharedCode().then(function (code) {
          sharedCodeOk = true;
          updateForm();
          if (code && current && current.codeShared && code !== current.code) {
            current.code = code;
            render();
          }
        }, function () { sharedCodeOk = false; });
      } catch (e) { sharedCodeOk = false; }
    }

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
    kindSel.addEventListener("change", function () { setError(""); updateForm(); });
    revealBox.addEventListener("change", updateForm);
    zoomSel.addEventListener("change", updateForm);
    roundedBox.addEventListener("change", updateForm);
    tabsText.addEventListener("input", function () { if (err.textContent) setError(""); updateForm(); });
    fillPresets();
    fillRecent();

    function saveAll(next) {
      settings.set(KEY_KIND, next.kind);
      settings.set(KEY_WIDGET, JSON.stringify(next.widget || {}));
      settings.set(KEY_URL, next.url);
      settings.set(KEY_CODE, next.code);
      settings.set(KEY_CODE_SHARED, next.codeShared ? "1" : "");
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
      settings.saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          setError(L("Не вдалося зберегти налаштування у файлі: ", "Couldn't save the settings in the file: ") + (r.error && r.error.message ? r.error.message : L("невідома помилка", "unknown error")) + L(". Перевірте, чи файл не відкрито лише для читання.", ". Check that the file isn't read-only."));
          return;
        }
        if (next.kind === "site" && next.codeShared && sharedCodeOk && next.code) {
          try { writeSharedCode(next.code).then(null, function () { /* лишається код цього блоку */ }); } catch (e2) { /* те саме */ }
        }
        current = next;
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
        var w = { title: $("wTitle").value.trim().slice(0, 60) };
        if (k === "timer") {
          var m = parseFloat($("wMinutes").value);
          if (!(m >= 0.25 && m <= 180)) { setError(L("Тривалість від 0,25 до 180 хвилин.", "Duration must be 0.25–180 minutes."), $("wMinutes")); $("wMinutes").focus(); return; }
          w.minutes = m;
        } else if (k === "picker") {
          w.names = $("wNames").value.slice(0, 5000);
          w.repeat = $("wRepeat").checked;
          if (!w.names.trim()) { setError(L("Додайте хоча б одне ім'я.", "Add at least one name."), $("wNames")); $("wNames").focus(); return; }
        } else {
          w.teams = $("wTeams").value.slice(0, 1000);
          w.step = Math.max(1, Math.min(100, parseInt($("wStep").value, 10) || 1));
        }
        setError("");
        saveAll({
          kind: k, widget: w, url: base.url || "", code: base.code || "", codeShared: !!base.codeShared, autoplay: !!base.autoplay,
          mode: base.mode || "embed", scale: base.scale || 1, scroll: base.scroll || "hide", qrCorner: !!base.qrCorner,
          reveal: !!base.reveal, revealText: base.revealText || "", refresh: base.refresh || 0, zoom: base.zoom || 1,
          fx: base.fx == null ? 50 : base.fx, fy: base.fy || 0, offline: base.offline || "", tabs: base.tabs || [],
          mainTitle: base.mainTitle || "", border: base.border || "", rounded: !!base.rounded, cornerBg: base.cornerBg || "#ffffff"
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
        url: templ,
        code: code,
        codeShared: codeShared.checked && sharedCodeOk,
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
        cornerBg: cleanColor(cornerBg.value)
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
