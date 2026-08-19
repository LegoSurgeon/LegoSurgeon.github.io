# The Engineering Portfolio

Design and engineering work by **Nicholas Paradizov** — a browsable catalog built with
[Astro](https://astro.build/) and Tailwind CSS.

The Excel workbook is the single source of truth. You write in Excel; the site
regenerates itself.

---

## Quick start (Windows)

Double-click these — no command line needed.

| File | What it does |
| --- | --- |
| **`start-server.bat`** | Reads the workbook, starts the local site, opens your browser. Leave the window open while you work. |
| **`stop-server.bat`** | Stops the server if the window was closed without Ctrl+C and the port is still held. |
| **`publish-site.bat`** | Builds the uploadable site into `publish\`. Copy that folder's contents to any web host. |

The first run of `start-server.bat` installs dependencies and takes a minute.
It needs [Node.js LTS](https://nodejs.org/) installed.

### Command line equivalents

```bash
npm install        # once
npm run dev        # convert workbook + dev server at http://localhost:4321/
npm run build      # convert workbook + static build into dist/
npm run data       # convert the workbook only
npm run data:check # verify src/data/projects.json matches the workbook
```

---

## Adding or editing a project

1. Open `data/Engineering Portfolio.xlsx`.
2. Add a sheet, or edit an existing one.
3. Drop any referenced images into `images/`.
4. Run `start-server.bat` (or `npm run data`).

A new page, a nav entry, a catalog card, and a search entry all appear
automatically. **No code changes are needed to add a project.**

### Sheet format

One sheet per project. Column **A** holds field names, column **B** holds values.

| Column A | Column B |
| --- | --- |
| `Title` | Full title, shown as the page heading |
| `Status` | e.g. `In Development` — becomes a badge |
| `Project Type` | e.g. `FTC Robotics / Modular Mechatronic System` |
| `Role` | Your role on the project |
| `Disciplines` | Comma-separated — becomes chips and catalog filters |
| `Tools` | Comma-separated — becomes chips |
| `Project Overview` | Section separator; leave column B empty |
| `What it is` | Prose |
| `Why I made it` | Prose |
| `What I designed` | Prose |
| `Key engineering challenge` | Prose |
| `What I learned` | Prose |
| `Image 1` … `Image 4` | Any path — **only the filename is used**, resolved against `images/` |

Notes:

- **Image paths are ignored except for the filename.** `C:\Users\nicho\Pictures\shot.png`
  and `shot.png` both resolve to `images/shot.png`.
- **Rows are rendered in workbook order**, so reordering the spreadsheet reorders the page.
- **Unrecognised field names still appear** on the page as extra overview sections
  rather than being dropped, so you can add fields without editing code.
- **More than 4 images work** — just keep numbering (`Image 5`, `Image 6`, …).

### Sheet names drive ordering and grouping

The sheet name is parsed as `<number><letter>. <short name>`:

| Sheet name | Meaning | URL |
| --- | --- | --- |
| `1A. MK2 Modular Chassis` | Project 1, primary page | `/projects/mk2-modular-chassis/` |
| `1B. MK2 Modular Chassis RWMA` | Project 1, sub-page of 1A | `/projects/mk2-modular-chassis-rwma/` |
| `2. MK1 Modular Chassis` | Project 2, standalone | `/projects/mk1-modular-chassis/` |

Sheets sharing a number are grouped: the first is the primary page, and the rest are
linked from it under "More on this project". Add `1C`, `1D`, `1E` and they nest
automatically.

A sheet with a `Title` but no other content renders as a **"Coming Soon"** placeholder,
so planned projects stay visible in the catalog while you write them up.

---

## Publishing

`publish-site.bat` writes a self-contained static site to `publish\`. Upload the
**contents** of that folder to any host — no Node.js required on the server.

Before publishing to a real address, open `publish-site.bat` and edit the two lines
near the top:

```bat
set "SITE_URL=http://localhost:4321"
set "SITE_BASE=/"
```

| Hosting | `SITE_URL` | `SITE_BASE` |
| --- | --- | --- |
| GitHub Pages project site | `https://<user>.github.io` | `/engineering-portfolio` |
| GitHub Pages user site, or a custom domain | `https://example.com` | `/` |

Every internal link and image path is built from `BASE_URL` (see `src/lib/url.js`),
so changing those two lines is all that's required — nothing else needs editing.
A `.nojekyll` file is written automatically so GitHub Pages serves the `_astro/`
folder correctly.

---

## Project structure

```
data/Engineering Portfolio.xlsx   Source of truth — edit this
images/                           Drop image files here
scripts/build-data.mjs            Workbook -> src/data/projects.json
src/data/projects.json            Generated; do not edit by hand
src/lib/projects.js               Read-only accessors over the generated data
src/lib/url.js                    Base-path-aware link helper
src/components/                   Header, Footer, ProjectCard, Gallery, Icon
src/pages/index.astro             Home
src/pages/projects/index.astro    Catalog with search + discipline filters
src/pages/projects/[slug].astro   One page per sheet, generated
src/pages/about.astro             About the author
publish/                          Generated by publish-site.bat — upload this
```

`public/images/` is a generated mirror of `images/` (Astro only serves files under
`public/`). It is git-ignored; `images/` is the folder you edit.

---

## Editing the site itself

Most content comes from the workbook, but a few things are in code:

- **Author name and site title** — `src/lib/projects.js` (`AUTHOR`, `SITE_TITLE`)
- **About page bio and contact links** — top of `src/pages/about.astro`, marked
  `PLACEHOLDER COPY`
- **Colors and fonts** — `src/styles/global.css` and `tailwind.config.mjs`
