# resume

A small Node script that generates my resume as a `.docx`, then converts it to a PDF and a PNG preview, and exports the content as JSON — all in one run.

This is my personal resume, not a general-purpose resume builder. I'm sharing the repo because the setup (JS that builds a docx and pipes it through LibreOffice and poppler) might be a decent starting point if you want to do the same thing for yours. Feel free to fork it and swap out the content.

## What it does

1. Builds a `.docx` using the [`docx`](https://www.npmjs.com/package/docx) npm package (fonts, colors, spacing, everything is defined in code, no template file).
2. Converts that `.docx` to a `.pdf` using LibreOffice, headless.
3. Converts the first page of that `.pdf` to a `.png` using `pdftoppm`, for a quick preview image.
4. Writes `resume.json` — the same content as structured data, plus a `generatedAt` timestamp.

All outputs land in the same folder as the script.

## Where the content lives

| File | Holds |
| --- | --- |
| `content.js` | The content, and nothing else. Plain data, zero dependencies, no styling. |
| `index.js` | The rendering: colors, fonts, spacing, page setup — and the pipeline. |

Content and rendering are deliberately separate, so the content is importable
and transformable without pulling in `docx`. If you fork this, edit
`content.js` for wording and `index.js` for anything visual.

## Requirements

- [Node.js](https://nodejs.org/) (18+ is fine)
- [LibreOffice](https://www.libreoffice.org/), for the docx to pdf step
- [poppler-utils](https://poppler.freedesktop.org/), for the pdf to png step (specifically `pdftoppm`)

On Fedora:

```bash
sudo dnf install libreoffice poppler-utils
```

On other distros, swap `dnf` for whatever your package manager is, the package names are usually the same or close to it.

> [!NOTE]
> LibreOffice and poppler-utils are only needed for the PDF and PNG conversion steps. If you're not on Linux/macOS, or you just don't have them installed, the script still generates the `.docx` fine, it'll only fail (with a clear error message) on the PDF and PNG steps, not the whole run.

## Running it

```bash
npm install
node index.js
```

That's it. You'll end up with four files: the `.docx`, a `.pdf`, a `.png` preview, and a `resume.json`.

## If you're using this as a template

Everything, text, styling, colors, spacing, lives directly in `content.js` (text) and `index.js` (styling). There's no config file or CLI flags. To make it yours:

- Replace the content (name, summary, experience, projects, skills) in `content.js`
- Change `NAVY`, `ACCENT`, and `RULE` in `index.js` if you want different colors
- Adjust the page margins and font sizes in `index.js` if your content is longer or shorter than mine (it's tuned to fit exactly one page right now, so adding content will likely push it to a second page unless you tighten spacing elsewhere)

No warranty, no guarantees it fits your content out of the box. It's just what I use.
