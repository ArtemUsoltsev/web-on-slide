# Безпека / Security

## Підтримувані версії

Підтримується лише остання версія з гілки `main`, яку віддає хостинг надбудови. Оновлення сайту доходять до всіх користувачів одразу (після зміни параметра кешу `?v=`), тому старі версії не підтримуються окремо.

| Компонент | Поточна версія |
| --- | --- |
| Файли сайту (`?v=`) | 1.9.1 |
| Маніфест (`<Version>`) | 1.2.0.0, `Permissions=Restricted`, `AllowSnapshot=false` |

## Як повідомити про вразливість

Не створюйте публічний issue з описом вразливості. Натомість:

1. Скористайтеся приватним повідомленням GitHub: вкладка **Security → Report a vulnerability** у цьому репозиторії (GitHub Private Vulnerability Reporting), або
2. для співробітників ДТЕК: напишіть власнику репозиторію (Artem Usoltsev, DTEK Academy) у корпоративний месенджер чи пошту.

Опишіть, що саме можна зробити, кроки для відтворення і, якщо можливо, приклад файлу .pptx чи адреси. Ми підтвердимо отримання протягом 3 робочих днів і повідомимо план виправлення. Критичні вразливості виправляються в першу чергу; виправлення на хостингу діє для всіх користувачів без переустановлення надбудови.

## Що вважається вразливістю

- виконання чужого коду в контексті надбудови (наприклад, через підроблений файл .pptx);
- доступ вбудованого сайту до Office API або до вмісту презентації;
- обхід білого списку доменів (`allowedOrigins` у `config.js`);
- використання домену надбудови як відкритого редиректу;
- витік даних із презентації чи комп'ютера на сторонні сервери.

Поведінка самих вбудованих сторонніх сайтів (їхні cookie, трекінг, вміст) не є вразливістю надбудови: адміністратор обмежує її білим списком доменів.

## Документи для перевірки

- [docs/security-review.md](docs/security-review.md): архітектура, потоки даних, дозволи, мережеві звернення, CSP, модель загроз, залишкові ризики й історія аудитів.
- [docs/admin-deployment.md](docs/admin-deployment.md): розгортання, хостинг, проксі, оновлення й відкат.
- [privacy.html](privacy.html): політика конфіденційності.

---

## English summary

Only the latest version on `main`, as served by the add-in host, is supported. Please do not open public issues for vulnerabilities; use GitHub **Security → Report a vulnerability** (private reporting) instead. We acknowledge reports within 3 business days. In scope: code execution in the add-in context (e.g. via a crafted .pptx), an embedded site reaching the Office API or presentation content, bypassing the `allowedOrigins` allowlist, open redirects on the add-in domain, and data leaks to third parties. The full security dossier is in [docs/security-review.md](docs/security-review.md) (Ukrainian).
