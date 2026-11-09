// Дві мови інтерфейсу: українська (основна, прямо в розмітці) та англійська.
// У HTML англійський текст лежить поруч з українським в атрибутах:
//   data-en        → textContent
//   data-en-html   → innerHTML (лише статичні рядки з цього репозиторію)
//   data-en-ph     → placeholder
//   data-en-title  → title
//   data-en-aria   → aria-label
// У JS: L("українською", "in English").
(function () {
  "use strict";
  var lang = "uk";

  function pick(setting, officeLang) {
    if (setting === "uk" || setting === "en") return setting;
    var l = String(officeLang || navigator.language || "uk").toLowerCase();
    return /^(uk|ru|be)\b/.test(l) ? "uk" : "en";
  }

  function apply(root) {
    if (lang !== "en") return;
    var els = (root || document).querySelectorAll("[data-en],[data-en-html],[data-en-ph],[data-en-title],[data-en-aria]");
    for (var i = 0; i < els.length; i++) {
      var el = els[i];
      if (el.hasAttribute("data-en")) el.textContent = el.getAttribute("data-en");
      if (el.hasAttribute("data-en-html")) el.innerHTML = el.getAttribute("data-en-html");
      if (el.hasAttribute("data-en-ph")) el.setAttribute("placeholder", el.getAttribute("data-en-ph"));
      if (el.hasAttribute("data-en-title")) el.setAttribute("title", el.getAttribute("data-en-title"));
      if (el.hasAttribute("data-en-aria")) el.setAttribute("aria-label", el.getAttribute("data-en-aria"));
    }
    document.documentElement.lang = "en";
  }

  window.WOS_I18N = {
    // Викликається після Office.onReady: setting із config.js, officeLang = Office.context.displayLanguage.
    init: function (setting, officeLang) {
      lang = pick(setting, officeLang);
      apply(document);
      return lang;
    },
    lang: function () { return lang; },
    L: function (uk, en) { return lang === "en" && en ? en : uk; }
  };
})();
