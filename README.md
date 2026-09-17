# Gamma Cleaner

**A free, browser-based tool to remove Gamma's repeated watermark/branding images from exported `.pptx` files — 100% client-side, nothing uploaded.**

[![Live Demo](https://img.shields.io/badge/Live%20Demo-8b5cf6?style=for-the-badge)](https://pptx-tools.pages.dev/)
[![MIT License](https://img.shields.io/badge/MIT%20License-blue?style=for-the-badge)](./LICENSE)
![100% Client-Side](https://img.shields.io/badge/100%25%20Client--Side-10b981?style=for-the-badge)

[**➤ Use the live tool**](https://pptx-tools.pages.dev/)

---

## What is this?

When you export a deck from [Gamma](https://gamma.app) as `.pptx`, the file often has a repeated "Made with Gamma" watermark/logo baked into the slide layouts and master — not just one visible badge, but the same image asset duplicated across every layout. That makes it tedious to strip out manually in PowerPoint.

**Gamma Cleaner** scans the internal XML structure of the `.pptx` file, detects the repeated watermark image by hash, and removes it — while trying to preserve your actual slide content and full-slide backgrounds.

Built for personal use, cleaning up decks I exported myself. Sharing it in case it helps others doing the same with their own exports.

> **Disclaimer:** This tool is intended for personal use on presentations you created and exported yourself. It does not interact with Gamma's servers, bypass any login, or unlock paid features — it only edits a local `.pptx` file you already have, the same way you could do manually with a zip editor, just automated. "Gamma" is a trademark of its respective owner; this project is not affiliated with or endorsed by Gamma. Please review Gamma's terms of service and use this responsibly.

---

## Privacy

Your `.pptx` file never leaves your device. No upload server, no account, no file storage — everything runs in your browser.

---

## Features

- 100% client-side, runs directly in the browser
- No file upload, no account required
- Detects Gamma's repeated watermark/logo image via hashing
- Scans slides, slide layouts, and slide masters
- Skips full-slide background-like images to avoid breaking your design
- Downloads a cleaned copy automatically
- Free and open source

---

## How it works

A `.pptx` file is a zip archive of XML files and media assets. Gamma Cleaner:

1. Opens the file locally using JSZip
2. Reads slide, layout, and master relationship files
3. Maps image references to their XML traces
4. Hashes embedded images to detect repeats (the watermark)
5. Filters out full-slide background-like images
6. Removes matching watermark image references
7. Rebuilds the `.pptx` and downloads the cleaned version

---

## Heads up

Results can vary depending on how the deck was exported. If a watermark doesn't fully get removed, or the cleaned `.pptx` doesn't open correctly, [open an issue](https://github.com/Shariqtechie/gamma-cleaner/issues) with the details — I'll take it from there.

- Always review the cleaned file before using it

---

## Tech stack

![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
[![JSZip](https://img.shields.io/badge/JSZip-yellow?style=for-the-badge)](https://stuk.github.io/jszip/)
[![FileSaver.js](https://img.shields.io/badge/FileSaver.js-grey?style=for-the-badge)](https://github.com/eligrey/FileSaver.js/)
![Cloudflare Pages](https://img.shields.io/badge/Cloudflare%20Pages-F38020?style=for-the-badge&logo=cloudflarepages&logoColor=white)

## Contributing

Bug reports and test cases are welcome. If the tool fails on a specific `.pptx` (e.g. a Gamma export format change), open an issue with:

- what you expected
- what happened
- whether the background changed
- whether the watermark remained

Please avoid uploading private or sensitive presentations publicly when filing issues.

## License

MIT — see [LICENSE](./LICENSE)

---

Made by [Shariqtechie](https://github.com/Shariqtechie)
