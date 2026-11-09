// Сторінка-перехідник для діалогу Office: перевіряє адресу тими ж правилами, що й index.html, і переходить на неї.
(function () {
  "use strict";
  var msg = document.getElementById("msg");
  var cfg = window.WEB_ON_SLIDE_CONFIG;
  var target = new URLSearchParams(location.search).get("u");

  function go() {
    if (!cfg || typeof cfg.check !== "function") {
      msg.textContent = "Не вдалося завантажити налаштування надбудови (config.js).";
      return;
    }
    var res = cfg.check(target);
    if (res.ok) {
      location.replace(res.url);
    } else {
      msg.textContent = (cfg.reasonText && cfg.reasonText[res.reason]) || "Ця адреса не дозволена.";
    }
  }

  // Office.js тут не обов'язковий: якщо він не завантажився, все одно переходимо.
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(go);
  } else {
    go();
  }
})();
