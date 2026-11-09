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
  var LOAD_TIMEOUT_MS = 10000; // через скільки показати картку «сторінка не відповідає»
  var REFRESH_CHOICES = [0, 10, 30, 60, 300];
  var DEFAULT_REVEAL_TEXT = "Натисніть, щоб показати";

  var cfg = window.WEB_ON_SLIDE_CONFIG;

  function $(id) { return document.getElementById(id); }
  var frame = $("frame"), frameWrap = $("frameWrap"), launch = $("launch"), setup = $("setup"), bar = $("bar");
  var loadfail = $("loadfail"), fallback = $("fallback"), fatal = $("fatal"), loading = $("loading");
  var cover = $("cover"), qrCard = $("qrCard"), qrCorner = $("qrCorner");
  var input = $("url"), codeInput = $("code"), codeRow = $("codeRow"), info = $("info"), err = $("err"), warn = $("warn");
  var modeSel = $("mode"), scaleSel = $("scale"), refreshSel = $("refresh"), allowScroll = $("allowScroll");
  var qrBox = $("qrCornerOn"), revealBox = $("reveal"), revealText = $("revealText"), more = $("more");

  var view = null;        // null (ще невідомо) | "edit" | "read" (показ слайдів)
  var current = null;     // збережене налаштування цієї вставки (див. readSettings)
  var editing = false;    // форма відкрита кнопкою «Змінити»; current при цьому зберігається
  var revealed = false;   // заставку вже прибрали кліком
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
  function cleanMode(v) {
    if (v === "button" && dialogSupported) return "button";
    if (v === "qr") return "qr";
    return "embed";
  }

  // Адреса, яку відкривають люди (у вікні, за QR-кодом): шаблон із підставленим кодом.
  function targetUrl(cur) { return cfg.normalize(cfg.fillCode(cur.url, cur.code)); }
  // Адреса для блоку на слайді: та сама, але перетворена на embed-адресу, якщо сервіс відомий.
  function frameUrl(cur) {
    var t = targetUrl(cur);
    var e = cfg.toEmbed(t);
    return e.service && cfg.check(e.url).ok ? e.url : t;
  }
  function needsCode(cur) { return cfg.hasCode(cur.url) && !cur.code; }

  // Масштаб < 1: iframe більший за блок у 1/scale разів і зменшений transform-ом,
  // тож сайт розкладається як на більшому екрані й вміщується без прокрутки.
  function applyLook(cur) {
    var k = cleanScale(cur.scale);
    frame.style.width = (100 / k) + "%";
    frame.style.height = (100 / k) + "%";
    frame.style.transform = k === 1 ? "" : "scale(" + k + ")";
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
    svg.setAttribute("aria-label", "QR-код: " + text);
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
    showFatal("Не вдалося завантажити налаштування надбудови (config.js). Повідомте адміністратора.");
    return;
  }
  if (typeof Office === "undefined" || !Office.onReady) {
    showFatal("Не вдалося завантажити бібліотеку Office (appsforoffice.microsoft.com). Перевірте мережу або проксі.");
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
    $("loadingText").textContent = hostOf(url) ? "Завантажую " + hostOf(url) + "…" : "Завантажую сторінку…";
    show(loading);
    frame.title = hostOf(url) || "Вбудована сторінка";
    // scrolling="no" ховає смугу прокрутки вбудованої сторінки (Chromium/Edge WebView2).
    // Атрибут має стояти до завантаження, тому виставляємо його перед src.
    var noScroll = !current || current.scroll !== "show";
    if (noScroll) frame.setAttribute("scrolling", "no"); else frame.removeAttribute("scrolling");
    frame.src = url;
    loadTimer = setTimeout(function () {
      if (!frameLoaded) {
        hide(loading);
        $("loadfailHost").textContent = hostOf(url);
        show(loadfail);
      }
    }, LOAD_TIMEOUT_MS);
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
    if (text) (field || input).setAttribute("aria-invalid", "true");
  }

  function render() {
    hide(frameWrap, launch, setup, bar, loadfail, fallback, loading, cover, qrCard, qrCorner);
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
      $("coverBtn").textContent = current.revealText || DEFAULT_REVEAL_TEXT;
      show(cover);
    } else {
      var url = frameUrl(current);
      applyLook(current);
      reloadable = true;
      show(frameWrap);
      if (frame.getAttribute("src") !== url) loadFrame(url);
      else if (!frameLoaded && loadTimer) show(loading);
      if (!refreshTimer) startRefresh(current);
    }
    if (current.qrCorner && current.mode !== "qr") showQr("corner", qrCorner, target);

    if (view === "edit") {
      if (reloadable) show($("reloadBtn")); else hide($("reloadBtn"));
      if (cfg.hasCode(current.url)) show($("codeBtn")); else hide($("codeBtn"));
      show(bar);
    }
  }

  // ---------- Форма ----------
  function fillForm(cur) {
    input.value = cur ? cur.url : "";
    codeInput.value = cur ? cur.code : "";
    modeSel.value = cur ? cur.mode : "embed";
    scaleSel.value = String(cur ? cur.scale : 1);
    if (scaleSel.selectedIndex < 0) scaleSel.value = "1";
    refreshSel.value = String(cur ? cur.refresh : 0);
    allowScroll.checked = cur ? cur.scroll === "show" : false;
    qrBox.checked = cur ? cur.qrCorner : false;
    revealBox.checked = cur ? cur.reveal : false;
    revealText.value = cur && cur.revealText !== DEFAULT_REVEAL_TEXT ? cur.revealText : "";
    // Розгортаємо «Додатково», якщо там щось змінено від типового.
    more.open = !!(cur && (cur.mode !== "embed" || cur.scale !== 1 || cur.scroll === "show" ||
      cur.qrCorner || cur.reveal || cur.refresh));
    warn.textContent = "";
    setError("");
    updateForm();
  }

  // Показує поле коду, підказку про перетворення адреси і вмикає лише доречні для режиму параметри.
  function updateForm() {
    var templ = cfg.normalize(input.value);
    var withCode = cfg.hasCode(templ);
    if (withCode) show(codeRow); else hide(codeRow);

    var mode = modeSel.value;
    var embed = mode === "embed";
    scaleSel.disabled = !embed;
    refreshSel.disabled = !embed;
    allowScroll.disabled = !embed;
    revealBox.disabled = !embed;
    qrBox.disabled = mode === "qr";
    if (embed && revealBox.checked) show($("revealRow")); else hide($("revealRow"));
    $("buttonHint").textContent = mode === "button"
      ? "На слайді буде велика кнопка, що відкриває сайт у вікні. Для сайтів, які забороняють вбудовування."
      : mode === "qr"
        ? "На слайді буде лише великий QR-код і коротка адреса: учасники відкривають сайт на телефонах."
        : "";

    var res = cfg.check(cfg.fillCode(templ, withCode ? (codeInput.value || "x") : ""));
    warn.textContent = res.ok && res.warnSecret
      ? "Схоже, адреса містить токен або ключ. Вона збережеться у файлі презентації відкритим текстом."
      : "";
    var conv = res.ok && embed ? cfg.toEmbed(res.url) : { service: null };
    info.textContent = conv.service && cfg.check(conv.url).ok
      ? "Адресу " + conv.service + " буде показано у форматі для вбудовування."
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
      showFatal("Ця сторінка працює лише всередині PowerPoint як надбудова «Веб на слайді».");
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

    var settings = Office.context.document.settings;
    var savedUrl = settings.get(KEY_URL);
    var savedCode = settings.get(KEY_CODE) || "";
    if (savedUrl && cfg.check(cfg.fillCode(savedUrl, savedCode || "x")).ok) {
      current = {
        url: savedUrl,
        code: String(savedCode),
        mode: cleanMode(settings.get(KEY_MODE)),
        scale: cleanScale(settings.get(KEY_SCALE)),
        scroll: settings.get(KEY_SCROLL) === "show" ? "show" : "hide",
        qrCorner: settings.get(KEY_QR) === "corner",
        reveal: settings.get(KEY_REVEAL) === "1",
        revealText: String(settings.get(KEY_REVEAL_TEXT) || DEFAULT_REVEAL_TEXT),
        refresh: cleanRefresh(settings.get(KEY_REFRESH))
      };
      if (needsCode(current)) fillForm(current);
    } else {
      updateForm();
    }

    var activeViewSupported = !!(req && req.isSetSupported("ActiveView", "1.1"));
    if (activeViewSupported) {
      Office.context.document.getActiveViewAsync(function (r) {
        view = r.status === Office.AsyncResultStatus.Succeeded ? r.value : "edit";
        render();
      });
      Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, function (e) {
        // Кожен новий показ починається із заставки, навіть якщо в редакторі її вже відкрили.
        if (e.activeView !== view) revealed = false;
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
    revealBox.addEventListener("change", updateForm);

    setup.addEventListener("submit", function (e) {
      e.preventDefault();
      var templ = cfg.normalize(input.value);
      var withCode = cfg.hasCode(templ);
      var code = withCode ? codeInput.value.replace(/\s+/g, "") : "";
      var res = cfg.check(cfg.fillCode(templ, code || "x"));
      if (!res.ok) {
        setError(cfg.reasonText[res.reason] || "Ця адреса не дозволена.");
        input.focus();
        return;
      }
      if (withCode && !code) {
        setError("Введіть код сесії.", codeInput);
        codeInput.focus();
        return;
      }
      setError("");
      var next = {
        url: templ,
        code: code,
        mode: cleanMode(modeSel.value),
        scale: cleanScale(scaleSel.value),
        scroll: allowScroll.checked ? "show" : "hide",
        qrCorner: qrBox.checked && modeSel.value !== "qr",
        reveal: revealBox.checked && modeSel.value === "embed",
        revealText: revealText.value.trim() || DEFAULT_REVEAL_TEXT,
        refresh: cleanRefresh(refreshSel.value)
      };
      settings.set(KEY_URL, next.url);
      settings.set(KEY_CODE, next.code);
      settings.set(KEY_MODE, next.mode);
      settings.set(KEY_SCALE, next.scale);
      settings.set(KEY_SCROLL, next.scroll);
      settings.set(KEY_QR, next.qrCorner ? "corner" : "");
      settings.set(KEY_REVEAL, next.reveal ? "1" : "");
      settings.set(KEY_REVEAL_TEXT, next.revealText);
      settings.set(KEY_REFRESH, next.refresh);
      settings.saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          setError("Не вдалося зберегти налаштування у файлі: " + (r.error && r.error.message ? r.error.message : "невідома помилка") + ". Перевірте, чи файл не відкрито лише для читання.");
          return;
        }
        current = next;
        editing = false;
        revealed = false;
        unloadFrame();
        render();
      });
    });

    function openForm(focusEl) {
      fillForm(current);
      editing = true;
      render();
      focusEl.focus();
      if (focusEl.select) focusEl.select();
    }
    $("editBtn").addEventListener("click", function () { openForm(input); });
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
    $("loadfailRetry").addEventListener("click", function () { if (current) loadFrame(frameUrl(current)); });
    $("popBtn").addEventListener("click", function () { if (current) openInDialog(targetUrl(current)); });
    $("launchBtn").addEventListener("click", function () { if (current) openInDialog(targetUrl(current)); });
    $("loadfailOpen").addEventListener("click", function () { if (current) openInDialog(targetUrl(current)); });
    $("fallbackClose").addEventListener("click", function () { hide(fallback); });
    $("copyBtn").addEventListener("click", function () {
      var text = $("fallbackUrl").textContent;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () {
          $("copyBtn").textContent = "Скопійовано";
          setTimeout(function () { $("copyBtn").textContent = "Скопіювати адресу"; }, 1500);
        }, copyFailed);
      } else {
        copyFailed();
      }
      function copyFailed() {
        // Адреса виділяється одним кліком (user-select: all), тож підказуємо скопіювати вручну.
        $("copyBtn").textContent = "Не вдалося. Виділіть адресу й натисніть Ctrl+C";
        setTimeout(function () { $("copyBtn").textContent = "Скопіювати адресу"; }, 3000);
      }
    });
  });
})();
