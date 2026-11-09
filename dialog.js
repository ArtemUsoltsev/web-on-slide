// Сторінка-перехідник для діалогу Office: перевіряє адресу тими ж правилами, що й index.html, і переходить на неї.
(function () {
  "use strict";
  var msg = document.getElementById("msg");
  var cfg = window.WEB_ON_SLIDE_CONFIG;
  var target = new URLSearchParams(location.search).get("u");

  function go() {
    if (!cfg || typeof cfg.check !== "function") {
      msg.textContent = "Не вдалося завантажити налаштування надбудови (config.js). / Couldn't load the add-in settings (config.js).";
      return;
    }
    var res = cfg.check(target);
    if (res.ok) {
      location.replace(res.url);
    } else {
      var l = (typeof Office !== "undefined" && Office.context && Office.context.displayLanguage) || navigator.language || "uk";
      var en = cfg.language === "en" || (cfg.language !== "uk" && !/^(uk|ru|be)\b/i.test(l));
      msg.textContent = cfg.reason ? cfg.reason(res.reason, en ? "en" : "uk") : ((cfg.reasonText && cfg.reasonText[res.reason]) || "Ця адреса не дозволена.");
    }
  }

  // Office.js тут не обов'язковий: якщо він не завантажився, все одно переходимо.
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(go);
  } else {
    go();
  }
})();
