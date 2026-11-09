// Кожна вставка content add-in має власні налаштування в .pptx,
// тож на кожному слайді може бути свій URL і свій режим.
(function () {
  "use strict";

  var KEY_URL = "url";
  var KEY_MODE = "mode"; // "embed" (iframe) або "button" (кнопка, що відкриває вікно)
  var KEY_SCALE = "scale"; // масштаб сторінки, 0.5–1 (1 = 100%)
  var KEY_SCROLL = "scroll"; // "show" = показувати прокрутку; інакше прихована
  var LOAD_TIMEOUT_MS = 10000; // через скільки показати картку «сторінка не відповідає»

  var cfg = window.WEB_ON_SLIDE_CONFIG;

  function $(id) { return document.getElementById(id); }
  var frame = $("frame"), frameWrap = $("frameWrap"), launch = $("launch"), setup = $("setup"), bar = $("bar");
  var loadfail = $("loadfail"), fallback = $("fallback"), fatal = $("fatal");
  var input = $("url"), asButton = $("asButton"), err = $("err"), warn = $("warn");
  var scaleSel = $("scale"), allowScroll = $("allowScroll");

  var view = null;        // null (ще невідомо) | "edit" | "read" (показ слайдів)
  var current = null;     // { url, mode, scale, scroll } — збережене налаштування цієї вставки
  var editing = false;    // форма відкрита кнопкою «Змінити»; current при цьому зберігається
  var dialogSupported = false;
  var isWeb = false;
  var loadTimer = null;
  var frameLoaded = false;

  function hide() { for (var i = 0; i < arguments.length; i++) arguments[i].classList.add("hidden"); }
  function show() { for (var i = 0; i < arguments.length; i++) arguments[i].classList.remove("hidden"); }

  function cleanScale(v) {
    var n = parseFloat(v);
    return isFinite(n) && n >= 0.25 && n <= 1 ? n : 1;
  }

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

  // ---------- Фатальні стани: без office.js або поза PowerPoint ----------
  function showFatal(text) {
    hide(frameWrap, launch, setup, bar, loadfail, fallback);
    $("fatalText").textContent = text;
    $("fatalUrl").textContent = location.href;
    show(fatal);
  }

  if (!cfg || typeof cfg.check !== "function") {
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

  function loadFrame(url) {
    clearLoadTimer();
    frameLoaded = false;
    hide(loadfail);
    frame.title = hostOf(url) || "Вбудована сторінка";
    // scrolling="no" ховає смугу прокрутки вбудованої сторінки (Chromium/Edge WebView2).
    // Атрибут має стояти до завантаження, тому виставляємо його перед src.
    var noScroll = !current || current.scroll !== "show";
    if (noScroll) frame.setAttribute("scrolling", "no"); else frame.removeAttribute("scrolling");
    frame.src = url;
    loadTimer = setTimeout(function () {
      if (!frameLoaded) {
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
    hide(loadfail);
  });

  function render() {
    hide(frameWrap, launch, setup, bar, loadfail, fallback);
    if (view === null) return; // не малюємо нічого, поки не знаємо режим: інакше форма блимає під час показу

    var showForm = !current || (editing && view === "edit");
    if (showForm) {
      if (view === "edit") {
        if (current) show($("cancelBtn")); else hide($("cancelBtn"));
        show(setup);
      }
      // Під час показу без налаштованого URL нічого не показуємо.
      return;
    }

    if (current.mode === "button") {
      $("launchHost").textContent = hostOf(current.url);
      show(launch);
    } else {
      applyLook(current);
      if (frame.getAttribute("src") !== current.url) loadFrame(current.url);
      show(frameWrap);
    }
    if (view === "edit") show(bar);
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
  Office.onReady(function (info) {
    if (!info || !info.host) {
      showFatal("Ця сторінка працює лише всередині PowerPoint як надбудова «Веб на слайді».");
      return;
    }

    var req = Office.context.requirements;
    dialogSupported = !!(req && req.isSetSupported("DialogApi", "1.1") && Office.context.ui && Office.context.ui.displayDialogAsync);
    isWeb = !!(Office.PlatformType && Office.context.platform === Office.PlatformType.OfficeOnline);
    if (!dialogSupported) {
      hide($("popBtn"));
      // Режим «кнопка» без діалогу не має сенсу: лишаємо лише вбудовування.
      asButton.checked = false;
      asButton.disabled = true;
    }

    applyFrameCsp();

    var settings = Office.context.document.settings;
    var savedUrl = settings.get(KEY_URL);
    var savedMode = settings.get(KEY_MODE);
    if (savedUrl && cfg.check(savedUrl).ok) {
      current = {
        url: savedUrl,
        mode: savedMode === "button" && dialogSupported ? "button" : "embed",
        scale: cleanScale(settings.get(KEY_SCALE)),
        scroll: settings.get(KEY_SCROLL) === "show" ? "show" : "hide"
      };
    }

    var activeViewSupported = !!(req && req.isSetSupported("ActiveView", "1.1"));
    if (activeViewSupported) {
      Office.context.document.getActiveViewAsync(function (r) {
        view = r.status === Office.AsyncResultStatus.Succeeded ? r.value : "edit";
        render();
      });
      Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, function (e) {
        view = e.activeView;
        render();
      });
    } else {
      view = "edit";
      render();
    }

    // Попередження про секрети в адресі під час введення (не блокує)
    input.addEventListener("input", function () {
      var res = cfg.check(input.value);
      warn.textContent = res.ok && res.warnSecret
        ? "Схоже, адреса містить токен або ключ. Вона збережеться у файлі презентації відкритим текстом."
        : "";
    });

    setup.addEventListener("submit", function (e) {
      e.preventDefault();
      var res = cfg.check(input.value);
      if (!res.ok) {
        err.textContent = cfg.reasonText[res.reason] || "Ця адреса не дозволена.";
        return;
      }
      err.textContent = "";
      var mode = asButton.checked && dialogSupported ? "button" : "embed";
      var scale = cleanScale(scaleSel.value);
      var scroll = allowScroll.checked ? "show" : "hide";
      settings.set(KEY_URL, res.url);
      settings.set(KEY_MODE, mode);
      settings.set(KEY_SCALE, scale);
      settings.set(KEY_SCROLL, scroll);
      settings.saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          err.textContent = "Не вдалося зберегти адресу у файлі: " + (r.error && r.error.message ? r.error.message : "невідома помилка") + ". Перевірте, чи файл не відкрито лише для читання.";
          return;
        }
        current = { url: res.url, mode: mode, scale: scale, scroll: scroll };
        editing = false;
        frame.removeAttribute("src");
        render();
      });
    });

    $("editBtn").addEventListener("click", function () {
      input.value = current ? current.url : "";
      asButton.checked = current ? current.mode === "button" : false;
      scaleSel.value = String(current ? current.scale : 1);
      if (scaleSel.selectedIndex < 0) scaleSel.value = "1";
      allowScroll.checked = current ? current.scroll === "show" : false;
      warn.textContent = "";
      err.textContent = "";
      editing = true;
      render();
      input.focus();
    });
    $("cancelBtn").addEventListener("click", function () {
      editing = false;
      err.textContent = "";
      render();
    });
    $("reloadBtn").addEventListener("click", function () { if (current) loadFrame(current.url); });
    $("loadfailRetry").addEventListener("click", function () { if (current) loadFrame(current.url); });
    $("popBtn").addEventListener("click", function () { if (current) openInDialog(current.url); });
    $("launchBtn").addEventListener("click", function () { if (current) openInDialog(current.url); });
    $("loadfailOpen").addEventListener("click", function () { if (current) openInDialog(current.url); });
    $("fallbackClose").addEventListener("click", function () { hide(fallback); });
    $("copyBtn").addEventListener("click", function () {
      var text = $("fallbackUrl").textContent;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () {
          $("copyBtn").textContent = "Скопійовано";
          setTimeout(function () { $("copyBtn").textContent = "Скопіювати адресу"; }, 1500);
        }, function () { /* текст уже виділяється кліком (user-select: all) */ });
      }
    });
  });
})();
