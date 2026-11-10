// Сторінка-перехідник для діалогу Office: перевіряє адресу тими ж правилами, що й index.html, і переходить на неї.
// Поза Office (просто відкрита в браузері) сама не переходить, а показує адресу й кнопку:
// так домен надбудови не можна використати як «перекидач» у фішингових посиланнях.
(function () {
  "use strict";
  var msg = document.getElementById("msg");
  var cfg = window.WEB_ON_SLIDE_CONFIG;
  var target = new URLSearchParams(location.search).get("u");

  function isEn() {
    var l = (typeof Office !== "undefined" && Office.context && Office.context.displayLanguage) || navigator.language || "uk";
    return cfg && (cfg.language === "en" || (cfg.language !== "uk" && !/^(uk|ru|be)\b/i.test(l)));
  }

  function insideOffice(info) {
    if (info && info.host) return true;
    try { return !!(Office.context && Office.context.ui && typeof Office.context.ui.messageParent === "function"); }
    catch (e) { return false; }
  }

  function confirmLink(url) {
    var en = isEn();
    while (msg.firstChild) msg.removeChild(msg.firstChild);
    var p = document.createElement("p");
    p.textContent = en ? "This link opens another site:" : "Це посилання відкриває інший сайт:";
    var host = document.createElement("p");
    host.style.cssText = "font-weight:600;word-break:break-all";
    host.textContent = url;
    var a = document.createElement("a");
    a.href = url;
    a.rel = "noopener noreferrer";
    a.textContent = en ? "Open the site" : "Відкрити сайт";
    msg.appendChild(p); msg.appendChild(host); msg.appendChild(a);
  }

  function go(info) {
    if (!cfg || typeof cfg.check !== "function") {
      msg.textContent = "Не вдалося завантажити налаштування надбудови (config.js). / Couldn't load the add-in settings (config.js).";
      return;
    }
    var res = cfg.check(target);
    if (!res.ok) {
      msg.textContent = cfg.reason ? cfg.reason(res.reason, isEn() ? "en" : "uk") : ((cfg.reasonText && cfg.reasonText[res.reason]) || "Ця адреса не дозволена.");
      return;
    }
    if (insideOffice(info)) location.replace(res.url);
    else confirmLink(res.url);
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(go);
  } else {
    // Без Office.js не можемо переконатися, що це діалог надбудови: лише показуємо адресу.
    go(null);
  }
})();
