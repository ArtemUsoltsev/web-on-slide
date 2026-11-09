// Налаштування безпеки та спільна перевірка адрес для index.html і dialog.html.
// Змінюються на хостингу, без повторного розгортання маніфесту.
// Після зміни підніміть параметр ?v= у index.html і dialog.html, щоб Office не взяв стару копію з кешу.
(function () {
  var cfg = {
    // Порожній список = дозволено будь-який HTTPS-сайт.
    // Якщо безпека хоче білий список, перелічіть origin, наприклад:
    // ["https://www.menti.com", "https://www.mentimeter.com", "https://game.example.com"]
    // Той самий список автоматично потрапляє в CSP frame-src (app.js), тому обмежує
    // не лише першу адресу, а й усі переходи всередині iframe.
    allowedOrigins: [],

    // Сайти, які заборонено вставляти завжди (перевіряється, навіть якщо allowedOrigins порожній).
    // Це захист від випадкової вставки, а не від переходів усередині сторінки.
    blockedOrigins: [],

    // Параметри запиту, що схожі на секрети. Адреса зберігається у файлі .pptx відкритим текстом,
    // тому при таких параметрах показуємо попередження (не блокуємо).
    secretParams: ["token", "key", "secret", "sig", "signature", "auth", "password", "pwd", "access_token"]
  };

  // Дописує https://, якщо схеми немає. "menti.com/abc" -> "https://menti.com/abc".
  cfg.normalize = function (value) {
    var v = String(value == null ? "" : value).trim();
    if (v && !/^[a-z][a-z0-9+.-]*:/i.test(v)) v = "https://" + v;
    return v;
  };

  // Повертає { ok, url, reason, warnSecret }.
  // reason: "empty" | "invalid" | "scheme" | "credentials" | "self" | "blocked" | "notAllowed" | null
  cfg.check = function (value) {
    var url = cfg.normalize(value);
    var result = { ok: false, url: url, reason: null, warnSecret: false };
    if (!url) { result.reason = "empty"; return result; }
    var u;
    try { u = new URL(url); } catch (e) { result.reason = "invalid"; return result; }
    if (u.protocol !== "https:") { result.reason = "scheme"; return result; }
    if (u.username || u.password) { result.reason = "credentials"; return result; }
    // Власний домен надбудови в iframe з allow-same-origin зняв би sandbox і дістався до Office API.
    if (typeof location !== "undefined" && u.origin === location.origin) { result.reason = "self"; return result; }
    var blocked = Array.isArray(cfg.blockedOrigins) ? cfg.blockedOrigins : [];
    var allowed = Array.isArray(cfg.allowedOrigins) ? cfg.allowedOrigins : [];
    if (blocked.indexOf(u.origin) !== -1) { result.reason = "blocked"; return result; }
    if (allowed.length && allowed.indexOf(u.origin) === -1) { result.reason = "notAllowed"; return result; }
    var secret = Array.isArray(cfg.secretParams) ? cfg.secretParams : [];
    u.searchParams.forEach(function (_, name) {
      if (secret.indexOf(name.toLowerCase()) !== -1) result.warnSecret = true;
    });
    result.ok = true;
    return result;
  };

  cfg.isAllowed = function (value) { return cfg.check(value).ok; };

  // Повідомлення для користувача за кодом причини.
  cfg.reasonText = {
    empty: "Введіть адресу сторінки.",
    invalid: "Це не схоже на адресу сайту.",
    scheme: "Потрібна адреса, що починається з https://.",
    credentials: "Адреса містить логін або пароль. Вони потраплять у файл презентації, тому така адреса не приймається.",
    self: "Не можна вставляти сторінку самої надбудови.",
    blocked: "Цей сайт заборонений політикою організації.",
    notAllowed: "Цей сайт не входить до списку дозволених. Зверніться до адміністратора."
  };

  window.WEB_ON_SLIDE_CONFIG = cfg;
})();
