# Сторонні компоненти / Third-party notices

У репозиторії є лише такий сторонній код і ресурси. Усі вони лежать локально й віддаються з хостингу надбудови; під час роботи сторонні CDN не використовуються.

| Компонент | Де | Версія | Ліцензія | Джерело |
| --- | --- | --- | --- | --- |
| qrcode-generator (Kazuhiko Arase) | `vendor/qrcode.js` | 1.4.4 | MIT | https://github.com/kazuhikoarase/qrcode-generator |
| Шрифти «Таймера PRO» (Fira Code, Gelasio, IBM Plex Mono, IBM Plex Sans, Inconsolata, Inter, JetBrains Mono, Montserrat, New Rocker, Open Sans, Oswald, Poppins, PT Mono, Roboto) | `timer/fonts/` | @fontsource 5.3.0 | SIL Open Font License 1.1 | https://fontsource.org |
| Шрифт Ubuntu Mono | `timer/fonts/` | @fontsource 5.3.0 | Ubuntu Font Licence 1.0 | https://fontsource.org |

Перелік файлів шрифтів і ліцензій: [timer/fonts/LICENSES.txt](timer/fonts/LICENSES.txt).

Під час роботи надбудова завантажує бібліотеку **Office.js** з `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` (Microsoft). Вона не входить у репозиторій і використовується за умовами Microsoft.

## qrcode-generator — MIT License

```
Copyright (c) 2009 Kazuhiko Arase

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

The word "QR Code" is a registered trademark of DENSO WAVE INCORPORATED.
