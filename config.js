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
    secretParams: ["token", "key", "secret", "sig", "signature", "auth", "password", "pwd", "access_token"],

    // Мова інтерфейсу: "auto" (українська для Office українською чи російською, інакше англійська), "uk" або "en".
    language: "auto",

    // Готові шаблони блоків для всієї організації (з'являються у списку «Мої шаблони»).
    // data — те саме, що зберігає кнопка «Зберегти цей блок як шаблон», наприклад:
    // { name: "Перерва 10 хв", data: { kind: "timer", w: { minutes: 10, title: "Перерва" } } }
    blockTemplates: [],

    // Швидкі адреси: кнопки під полем адреси. {код} стає окремим полем «Код сесії».
    // Додайте сюди сервіси своєї організації, наприклад { name: "Наша LMS", url: "https://lms.example.com/course/{код}" }.
    presets: [
      { name: "Mentimeter", url: "https://www.menti.com/{код}" },
      { name: "Kahoot", url: "https://kahoot.it/?pin={код}" }
    ]
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
    // Пробіл в адресі майже завжди означає, що вставлено зайвий текст. Chromium такі адреси «лагодить» мовчки.
    if (/\s/.test(url)) { result.reason = "invalid"; return result; }
    var u;
    try { u = new URL(url); } catch (e) { result.reason = "invalid"; return result; }
    if (u.protocol !== "https:") { result.reason = "scheme"; return result; }
    // «menti» без домену браузер перетворив би на https://menti/, і блок довго чекав би на неіснуючий сайт.
    if (u.hostname.indexOf(".") === -1 && u.hostname !== "localhost") { result.reason = "noDomain"; return result; }
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

  // ---------- Код сесії ----------
  // Адреса може містити {код} (або {code}): тоді на слайді є окреме поле «Код сесії»,
  // і перед новою групою тренер міняє лише код. Наприклад: https://www.menti.com/{код}
  // Пробіли з коду прибираються: «1234 5678» → 12345678.
  var CODE_RE = /\{(код|code)\}/gi;
  cfg.hasCode = function (template) { CODE_RE.lastIndex = 0; return CODE_RE.test(String(template || "")); };
  cfg.fillCode = function (template, code) {
    return String(template || "").replace(CODE_RE, encodeURIComponent(String(code == null ? "" : code).replace(/\s+/g, "")));
  };

  // ---------- Розумні посилання ----------
  // Звичайні адреси популярних сервісів забороняють вбудовування, а їхні embed-адреси дозволяють.
  // Перетворення робиться під час показу, тож нове правило тут діє і для вже збережених презентацій.
  // Кожне правило: { name, test(u: URL) -> рядок нової адреси або null }.
  function ytStart(u) {
    var t = u.searchParams.get("t") || u.searchParams.get("start") || "";
    if (/^\d+$/.test(t)) return +t;
    var m = /^(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s?)?$/.exec(t);
    if (!t || !m) return 0;
    return (+m[1] || 0) * 3600 + (+m[2] || 0) * 60 + (+m[3] || 0);
  }
  function ytEmbed(id, u) {
    if (!/^[\w-]{6,20}$/.test(id)) return null;
    var s = ytStart(u);
    return "https://www.youtube.com/embed/" + id + (s ? "?start=" + s : "");
  }
  cfg.embedRules = [
    { name: "YouTube", test: function (u) {
      var h = u.hostname.replace(/^(www|m|music)\./, "");
      if (h === "youtu.be") return ytEmbed(u.pathname.slice(1).split("/")[0], u);
      if (h !== "youtube.com" && h !== "youtube-nocookie.com") return null;
      if (u.pathname === "/watch" && u.searchParams.get("v")) return ytEmbed(u.searchParams.get("v"), u);
      var list = u.searchParams.get("list");
      if (u.pathname === "/playlist" && list && /^[\w-]{10,64}$/.test(list)) return "https://www.youtube.com/embed/videoseries?list=" + list;
      var m = /^\/(shorts|live)\/([\w-]+)/.exec(u.pathname);
      return m ? ytEmbed(m[2], u) : null;
    } },
    { name: "Vimeo", test: function (u) {
      if (u.hostname.replace(/^www\./, "") !== "vimeo.com") return null;
      var m = /^\/(\d+)(?:\/([0-9a-f]+))?\/?$/.exec(u.pathname);
      return m ? "https://player.vimeo.com/video/" + m[1] + (m[2] ? "?h=" + m[2] : "") : null;
    } },
    { name: "Google", test: function (u) {
      if (u.hostname === "drive.google.com") {
        var f = /^\/file\/d\/([\w-]+)\/(view|edit)/.exec(u.pathname);
        return f ? "https://drive.google.com/file/d/" + f[1] + "/preview" : null;
      }
      if (u.hostname !== "docs.google.com") return null;
      // Форма для відповідей: додаємо embedded=true (редактор форми не чіпаємо).
      if (/^\/forms\/d\/e\/[\w-]+\/viewform/.test(u.pathname)) {
        if (u.searchParams.get("embedded") === "true") return null;
        u.searchParams.set("embedded", "true");
        return u.href;
      }
      // Документи, таблиці, презентації: редактор → перегляд (/preview).
      var d = /^\/(document|spreadsheets|presentation)\/d\/([\w-]+)\/(edit|view)?/.exec(u.pathname);
      return d ? "https://docs.google.com/" + d[1] + "/d/" + d[2] + "/preview" : null;
    } },
    { name: "Miro", test: function (u) {
      if (u.hostname.replace(/^www\./, "") !== "miro.com") return null;
      var m = /^\/app\/board\/([^/]+)/.exec(u.pathname);
      return m ? "https://miro.com/app/live-embed/" + m[1] + "/" : null;
    } },
    { name: "Figma", test: function (u) {
      if (u.hostname.replace(/^www\./, "") !== "figma.com") return null;
      if (!/^\/(file|design|proto|board|slides)\//.test(u.pathname)) return null;
      return "https://www.figma.com/embed?embed_host=share&url=" + encodeURIComponent(u.href);
    } },
    { name: "Canva", test: function (u) {
      if (u.hostname.replace(/^www\./, "") !== "canva.com") return null;
      var m = /^\/design\/([\w-]+)\/([\w-]+\/)?(view|watch)\/?$/.exec(u.pathname);
      if (!m || u.searchParams.has("embed")) return null;
      return "https://www.canva.com/design/" + m[1] + "/" + (m[2] || "") + "view?embed";
    } }
  ];

  // Автозапуск без звуку для вбудованих плеєрів YouTube і Vimeo (браузери дозволяють autoplay лише без звуку).
  // Повертає нову адресу або ту саму, якщо це не відомий плеєр.
  cfg.withAutoplay = function (url) {
    var u;
    try { u = new URL(url); } catch (e) { return url; }
    var h = u.hostname.replace(/^www\./, "");
    if ((h === "youtube.com" || h === "youtube-nocookie.com") && /^\/embed\//.test(u.pathname)) {
      u.searchParams.set("autoplay", "1"); u.searchParams.set("mute", "1"); u.searchParams.set("playsinline", "1");
      return u.href;
    }
    if (h === "player.vimeo.com" && /^\/video\//.test(u.pathname)) {
      u.searchParams.set("autoplay", "1"); u.searchParams.set("muted", "1");
      return u.href;
    }
    return url;
  };
  cfg.isPlayer = function (url) { return cfg.withAutoplay(url) !== url; };

  // Повертає { url, service }: service = назва сервісу, якщо адресу перетворено, інакше null.
  cfg.toEmbed = function (url) {
    var u;
    try { u = new URL(url); } catch (e) { return { url: url, service: null }; }
    for (var i = 0; i < cfg.embedRules.length; i++) {
      var out = null;
      try { out = cfg.embedRules[i].test(new URL(u.href)); } catch (e2) { out = null; }
      if (out && out !== url) return { url: out, service: cfg.embedRules[i].name };
    }
    return { url: url, service: null };
  };

  // Повідомлення для користувача за кодом причини.
  cfg.reasonText = {
    empty: "Введіть адресу сторінки.",
    invalid: "Це не схоже на адресу сайту.",
    noDomain: "Схоже, адреса неповна: бракує домену. Наприклад: menti.com",
    scheme: "Потрібна адреса, що починається з https://.",
    credentials: "Адреса містить логін або пароль. Вони потраплять у файл презентації, тому така адреса не приймається.",
    self: "Не можна вставляти сторінку самої надбудови.",
    blocked: "Цей сайт заборонений політикою організації.",
    notAllowed: "Цей сайт не входить до списку дозволених. Зверніться до адміністратора."
  };

  cfg.reasonTextEn = {
    empty: "Enter the page address.",
    invalid: "This doesn't look like a web address.",
    noDomain: "The address looks incomplete: the domain is missing. For example: menti.com",
    scheme: "The address must start with https://.",
    credentials: "The address contains a login or password. They would be saved in the presentation file, so this address is not accepted.",
    self: "You can't embed the add-in's own page.",
    blocked: "This site is blocked by your organization's policy.",
    notAllowed: "This site is not on the allowed list. Contact your administrator."
  };
  // Текст причини потрібною мовою: lang = "uk" | "en".
  cfg.reason = function (code, lang) {
    var t = lang === "en" ? cfg.reasonTextEn : cfg.reasonText;
    return t[code] || (lang === "en" ? "This address is not allowed." : "Ця адреса не дозволена.");
  };

  window.WEB_ON_SLIDE_CONFIG = cfg;
})();
