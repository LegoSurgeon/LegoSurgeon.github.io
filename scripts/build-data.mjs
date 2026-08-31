#!/usr/bin/env node
/**
 * Converts the portfolio workbook into src/data/projects.json.
 *
 *   node scripts/build-data.mjs          # write src/data/projects.json
 *   node scripts/build-data.mjs --check  # verify the committed JSON is current (CI)
 *
 * Workbook contract — one sheet per project, key/value pairs in columns A and B:
 *
 *   A                            B
 *   ─────────────────────────    ────────────────────────────────
 *   Title                        MK2 Modular Chassis
 *   Status                       In Development
 *   Year Made                    2026
 *   Project Type                 FTC Robotics / Modular …
 *   Role                         Sole chassis and subsystem designer
 *   Disciplines                  Mechanical Design, Mechatronics, …
 *   Tools                        Fusion 360, KiCad, …
 *   Project Overview             ← section separator, blank B cell
 *   What it is                   Modular Chassis MK2 is the second …
 *   Bricklink Page               N/A                  ← catalogued, not rendered
 *   …
 *   Image 1                      C:\Users\…\Screenshot ….png
 *   Image 1 Caption              Full Layout          ← optional, per image
 *   Video 1                      https://youtu.be/lvcCY7FPsXI
 *   Video Caption                Transformation Sequence
 *   Video Thumbnail              C:\Users\…\Screenshot ….png   ← poster frame
 *
 * The workbook is the only source. Captions, video captions and video poster
 * frames all come from these rows and nowhere else.
 *
 * Sheet names carry ordering and grouping: "1A. MK2 Modular Chassis" means
 * project 1, variant A. Variants of one project (1A, 1B, 1C…) are grouped
 * together, with A as the primary and the rest as sub-pages. A sheet with no
 * letter ("2. MK1 Modular Chassis") is a standalone project.
 *
 * SECTION CATALOG — the first sheet lists every section title the portfolio
 * knows about, most of them set to "N/A". That column A is the catalog, and
 * every project carries the *whole* catalog in `overview`: a section the sheet
 * fills in gets its prose, one it doesn't gets "N/A". Each entry is tagged with
 * `visible`, which is false for "N/A" and for blanks. Nothing is dropped here —
 * hiding is a rendering rule, applied by the components, so the JSON stays a
 * faithful copy of the workbook.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
// xlsx ships as CommonJS; the default import is the whole namespace. We hand it
// a Buffer via read() rather than readFile() so it never needs its own fs shim.
import XLSX from 'xlsx';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const WORKBOOK = path.join(ROOT, 'data', 'Engineering Portfolio.xlsx');
// Optional second workbook: prose for the pages that aren't projects.
const EXTRAS = path.join(ROOT, 'data', 'additional-info.xlsx');
const ABOUT_SHEET = 'About Me';
const IMAGE_DIR = path.join(ROOT, 'images');
const OUT_FILE = path.join(ROOT, 'src', 'data', 'projects.json');
// Astro only serves files under public/, but images/ is the folder you drop
// files into. We mirror one into the other so both stay true.
const PUBLIC_IMAGE_DIR = path.join(ROOT, 'public', 'images');

// Excel label -> key on the project object. Comparison is case-insensitive on a
// whitespace-collapsed label, so "project type" and "Project  Type" both match.
const KNOWN_FIELDS = {
  'title': 'title',
  'status': 'status',
  'year made': 'year',
  'project type': 'projectType',
  'role': 'role',
  'disciplines': 'disciplines',
  'tools': 'tools',
};

const SEPARATORS = new Set(['project overview']);

// Comma-separated cells that become chip lists on the page.
const LIST_FIELDS = new Set(['disciplines', 'tools']);

// Row labels that carry media rather than prose, so they never become sections.
const IMAGE_ROW = /^image\s*(\d+)$/;
const IMAGE_CAPTION_ROW = /^image\s*(\d+)\s*caption$/;
const VIDEO_ROW = /^video\s*(\d+)$/;
// The number is optional on these two — the workbook writes "Video Caption",
// not "Video 1 Caption" — and an absent one means the first video.
const VIDEO_CAPTION_ROW = /^video\s*(\d*)\s*caption$/;
const VIDEO_THUMB_ROW = /^video\s*(\d*)\s*(?:thumbnail|poster)$/;

const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
const key = (s) => norm(s).toLowerCase();

/** "N/A", "n/a", "NA" — the workbook's way of saying "not part of this project". */
const isNotApplicable = (s) => /^n\s*\/?\s*a$/i.test(norm(s));

/**
 * A narrative section whose whole body is a file path is not prose — it is a
 * media row whose label was mistyped ("Video Tumbnail"), so it missed the media
 * patterns above and fell through to the section catalog. Left alone it
 * publishes a raw local path onto the page, so it is worth shouting about.
 */
const looksLikePath = (s) => {
  const v = norm(s).replace(/^["']+|["']+$/g, '');
  return /^(?:[A-Za-z]:[\\/]|\\\\|\/)/.test(v) && /\.[A-Za-z0-9]{2,5}$/.test(v);
};

const slugify = (s) =>
  norm(s)
    .toLowerCase()
    .replace(/[''‛]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');

/**
 * "1A. MK2 Modular Chassis" -> { number: 1, variant: 'A', short: 'MK2 Modular Chassis' }
 * "2. MK1 Modular Chassis"  -> { number: 2, variant: '',  short: 'MK1 Modular Chassis' }
 * "Loose Sheet"             -> { number: null, variant: '', short: 'Loose Sheet' }
 */
function parseSheetName(sheetName) {
  const m = norm(sheetName).match(/^(\d+)\s*([A-Za-z]?)\s*[.)-]?\s*(.*)$/);
  if (!m || !m[3]) return { number: null, variant: '', short: norm(sheetName) };
  return { number: Number(m[1]), variant: m[2].toUpperCase(), short: norm(m[3]) };
}

/** Excel holds absolute Windows paths; only the filename is meaningful here. */
function imageFileName(raw) {
  return norm(raw).replace(/^["']+|["']+$/g, '').split(/[\\/]/).pop() ?? '';
}

/**
 * youtu.be/ID · watch?v=ID · /embed/ID · /shorts/ID -> "ID", anything else null.
 * Only the id travels into the JSON, so the page can build a privacy-mode embed
 * URL itself rather than trusting whatever share link Excel happened to hold.
 */
function youtubeId(raw) {
  const value = norm(raw).replace(/^["']+|["']+$/g, '');
  if (!value) return null;
  const m =
    value.match(/(?:youtu\.be\/)([A-Za-z0-9_-]{6,})/) ??
    value.match(/(?:[?&]v=)([A-Za-z0-9_-]{6,})/) ??
    value.match(/(?:\/(?:embed|shorts|live|v)\/)([A-Za-z0-9_-]{6,})/);
  if (m) return m[1];
  // A bare id pasted into the cell, with no URL around it.
  return /^[A-Za-z0-9_-]{11}$/.test(value) ? value : null;
}

/**
 * Column A of every sheet, deduplicated, minus the fixed fields and the media
 * rows: the full list of section titles the portfolio can show. The first sheet
 * spells the catalog out in full, so its order and its capitalisation win; any
 * title only a later sheet uses is appended in the order it is first seen.
 */
function collectSectionCatalog(wb) {
  const catalog = [];
  const seen = new Set();

  for (const sheetName of wb.SheetNames) {
    for (const { label } of readSheetPairs(wb.Sheets[sheetName])) {
      const k = key(label);
      if (!k || seen.has(k)) continue;
      if (SEPARATORS.has(k) || k in KNOWN_FIELDS) continue;
      if (IMAGE_ROW.test(k) || IMAGE_CAPTION_ROW.test(k)) continue;
      if (VIDEO_ROW.test(k) || VIDEO_CAPTION_ROW.test(k) || VIDEO_THUMB_ROW.test(k)) continue;
      seen.add(k);
      catalog.push({ key: k, label: norm(label) });
    }
  }

  return catalog;
}

function readSheetPairs(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });
  const pairs = [];
  for (const row of rows) {
    const label = norm(row?.[0]);
    const value = norm(row?.[1]);
    if (!label && !value) continue;
    pairs.push({ label, value });
  }
  return pairs;
}

function buildProject(sheetName, pairs, availableImages, catalog) {
  const { number, variant, short } = parseSheetName(sheetName);

  // "Image 3 Caption" may sit above or below its "Image 3" row, so captions are
  // collected by index here and merged once every row has been read.
  const captions = new Map();
  const videoCaptions = new Map();
  const videoThumbs = new Map();
  const videoRows = [];

  // Sections this sheet actually spells out, in sheet order. The rest of the
  // catalog is appended as "N/A" once the sheet has been read.
  const sectionText = new Map();
  const sectionOrder = [];

  const project = {
    sheet: sheetName,
    number,
    variant,
    short,
    slug: slugify(short || sheetName),
    title: '',
    status: '',
    year: '',
    projectType: '',
    role: '',
    disciplines: [],
    tools: [],
    overview: [],
    images: [],
    videos: [],
  };

  for (const { label, value } of pairs) {
    const k = key(label);
    if (!k || SEPARATORS.has(k)) continue;

    if (k in KNOWN_FIELDS) {
      project[KNOWN_FIELDS[k]] = LIST_FIELDS.has(k)
        ? value.split(',').map(norm).filter(Boolean)
        : value;
      continue;
    }

    // "Image 1", "Image 12", … — index preserved so gaps don't renumber.
    const img = k.match(IMAGE_ROW);
    if (img) {
      if (!value) continue;
      const file = imageFileName(value);
      project.images.push({
        index: Number(img[1]),
        file,
        src: `/images/${file}`,
        exists: availableImages.has(file.toLowerCase()),
        caption: '',
      });
      continue;
    }

    // "Image 1 Caption" — optional. Matched before the section fallback below so
    // it never leaks into the narrative as a prose block of its own.
    const capt = k.match(IMAGE_CAPTION_ROW);
    if (capt) {
      if (value) captions.set(Number(capt[1]), value);
      continue;
    }

    const vid = k.match(VIDEO_ROW);
    if (vid) {
      videoRows.push({ index: Number(vid[1]), url: value });
      continue;
    }

    const vcapt = k.match(VIDEO_CAPTION_ROW);
    if (vcapt) {
      if (value) videoCaptions.set(Number(vcapt[1] || 1), value);
      continue;
    }

    // "Video Thumbnail" — an explicit poster frame, which is usually not the
    // project's first image. Takes the same absolute-path form as an image row.
    const vthumb = k.match(VIDEO_THUMB_ROW);
    if (vthumb) {
      if (value) videoThumbs.set(Number(vthumb[1] || 1), imageFileName(value));
      continue;
    }

    // Everything else is a narrative section. "N/A" and blanks are kept here
    // too — they are marked invisible below rather than dropped.
    if (!sectionText.has(k)) sectionOrder.push(k);
    sectionText.set(k, value);
  }

  // Sections the sheet fills in first, in its own order, then the rest of the
  // catalog, so every project carries the same set of keys.
  const labels = new Map(catalog.map((c) => [c.key, c.label]));
  for (const k of sectionOrder) {
    const text = sectionText.get(k) ?? '';
    project.overview.push({
      label: labels.get(k) ?? k,
      text: text || 'N/A',
      visible: Boolean(text) && !isNotApplicable(text),
    });
  }
  for (const { key: k, label } of catalog) {
    if (sectionText.has(k)) continue;
    project.overview.push({ label, text: 'N/A', visible: false });
  }

  for (const image of project.images) {
    image.caption = captions.get(image.index) ?? '';
  }
  project.images.sort((a, b) => a.index - b.index);

  // Poster frame: an explicit "Video Thumbnail" row if the sheet has one, and
  // only failing that the project's first image. `posterSource` records which
  // it was, so a fallback poster is visible in the JSON rather than looking
  // like a deliberate choice.
  const firstImage = project.images.find((i) => i.index === 1) ?? project.images[0] ?? null;

  project.videos = videoRows
    .sort((a, b) => a.index - b.index)
    .map(({ index, url }) => {
      const id = youtubeId(url);
      const named = videoThumbs.get(index);
      const file = named ?? (firstImage ? firstImage.file : '');
      return {
        index,
        url,
        youtubeId: id ?? '',
        embedUrl: id ? `https://www.youtube-nocookie.com/embed/${id}` : '',
        caption: videoCaptions.get(index) ?? '',
        poster: file ? `/images/${file}` : '',
        posterFile: file,
        posterExists: availableImages.has(file.toLowerCase()),
        posterSource: named ? 'workbook' : file ? 'first image' : 'none',
        visible: Boolean(id) && !isNotApplicable(url),
      };
    });

  // A sheet counts as populated once it has a title AND something to show.
  project.populated = Boolean(
    project.title &&
      (project.overview.some((o) => o.visible) ||
        project.images.length ||
        project.videos.some((v) => v.visible) ||
        project.status)
  );

  return project;
}

function assignRouting(projects) {
  // Sheets sharing a number are variants of one project; A (or the first one
  // seen) is the primary, later letters become sub-pages beneath it.
  const byNumber = new Map();
  for (const p of projects) {
    const groupKey = p.number === null ? `x-${p.slug}` : String(p.number);
    p.groupKey = groupKey;
    if (!byNumber.has(groupKey)) byNumber.set(groupKey, []);
    byNumber.get(groupKey).push(p);
  }

  for (const group of byNumber.values()) {
    group.sort((a, b) => a.variant.localeCompare(b.variant));
    const [primary, ...variants] = group;
    primary.isPrimary = true;
    primary.parentSlug = null;
    primary.variantSlugs = variants.map((v) => v.slug);
    for (const v of variants) {
      v.isPrimary = false;
      v.parentSlug = primary.slug;
      v.variantSlugs = [];
    }
  }

  // Guard against two sheets slugifying to the same URL.
  const seen = new Map();
  for (const p of projects) {
    if (!seen.has(p.slug)) {
      seen.set(p.slug, p);
      continue;
    }
    const suffix = p.number === null ? 'alt' : `${p.number}${p.variant}`.toLowerCase();
    p.slug = `${p.slug}-${suffix}`;
    console.warn(`  ! duplicate slug from "${p.sheet}" — renamed to "${p.slug}"`);
  }

  return projects;
}

/* ─────────────────────────────── about ─────────────────────────────────── */

/**
 * Drafting notes that were left inside an otherwise finished cell — copy about
 * the copy, addressed to Nick rather than to a reader of the site. They are
 * dropped from the page and reported, so the fix stays "delete the sentence
 * from the sheet" rather than "remember the site quietly hides it".
 *
 * Matched on a prefix so light edits to the tail of the note still match.
 * Delete an entry once the sheet no longer contains it.
 */
const DRAFTING_NOTES = ['I think this works well because'];

/**
 * data/additional-info.xlsx, sheet "About Me" — the same column A / column B
 * shape as a project sheet, but the values are prose sections rather than
 * fields. A cell that is one line of "·"-separated terms becomes a chip list;
 * anything else is split into paragraphs on blank lines.
 */
/**
 * The About content already written to src/data/projects.json, which is
 * committed. additional-info.xlsx is not — it carries a personal email address
 * and phone number — so on any machine without it this is the only copy.
 * Without this, `npm run build` would regenerate the JSON and silently blank
 * the About page on a fresh clone.
 */
function previousAbout() {
  try {
    return JSON.parse(fs.readFileSync(OUT_FILE, 'utf8')).about ?? null;
  } catch {
    return null;
  }
}

function readAbout() {
  const warnings = [];

  if (!fs.existsSync(EXTRAS)) {
    const kept = previousAbout();
    warnings.push(
      kept
        ? `${path.basename(EXTRAS)} not present — kept the about content already in projects.json`
        : `${path.basename(EXTRAS)} not present and no about content in projects.json — about page will be bare`
    );
    return { about: kept, warnings };
  }

  const wb = XLSX.read(fs.readFileSync(EXTRAS), { type: 'buffer' });
  const sheet = wb.Sheets[ABOUT_SHEET];
  if (!sheet) {
    warnings.push(`${path.basename(EXTRAS)} has no "${ABOUT_SHEET}" sheet — about page left empty`);
    return { about: null, warnings };
  }

  // Blank lines separate paragraphs, so the raw cell is read rather than the
  // whitespace-collapsed norm() used everywhere else.
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });
  const sections = [];

  for (const row of rows) {
    const label = norm(row?.[0]);
    const raw = String(row?.[1] ?? '').trim();
    if (!label || !raw || isNotApplicable(raw)) continue;

    let paragraphs = raw
      .split(/\r?\n\s*\r?\n/)
      .map((p) => p.replace(/\s*\r?\n\s*/g, ' ').trim())
      .filter(Boolean);

    const kept = paragraphs.filter(
      (p) => !DRAFTING_NOTES.some((note) => p.startsWith(note))
    );
    if (kept.length !== paragraphs.length) {
      warnings.push(`"${ABOUT_SHEET}" → ${label}: dropped a drafting note left in the cell`);
      paragraphs = kept;
    }
    if (!paragraphs.length) continue;

    // "A · B · C" on one line is a term list, not a sentence.
    const single = paragraphs.length === 1 ? paragraphs[0] : '';
    const items = single.includes('·')
      ? single.split('·').map(norm).filter(Boolean)
      : [];

    sections.push({ label, paragraphs, items });
  }

  if (!sections.length) warnings.push(`"${ABOUT_SHEET}" sheet is empty`);

  return {
    about: { source: path.relative(ROOT, EXTRAS).replace(/\\/g, '/'), sections },
    warnings,
  };
}

/** Mirror images/ into public/images/, copying only what changed. */
function syncImages() {
  fs.mkdirSync(PUBLIC_IMAGE_DIR, { recursive: true });
  const source = fs.existsSync(IMAGE_DIR) ? fs.readdirSync(IMAGE_DIR) : [];
  let copied = 0;

  for (const name of source) {
    const from = path.join(IMAGE_DIR, name);
    const to = path.join(PUBLIC_IMAGE_DIR, name);
    if (!fs.statSync(from).isFile()) continue;
    const fresh =
      fs.existsSync(to) && fs.statSync(to).mtimeMs >= fs.statSync(from).mtimeMs;
    if (fresh) continue;
    fs.copyFileSync(from, to);
    copied++;
  }

  // Drop mirrored files whose source is gone, so deletions propagate.
  const keep = new Set(source);
  let removed = 0;
  for (const name of fs.readdirSync(PUBLIC_IMAGE_DIR)) {
    if (keep.has(name)) continue;
    fs.rmSync(path.join(PUBLIC_IMAGE_DIR, name), { force: true });
    removed++;
  }

  return { copied, removed, total: source.length };
}

function build() {
  if (!fs.existsSync(WORKBOOK)) {
    throw new Error(`Workbook not found: ${path.relative(ROOT, WORKBOOK)}`);
  }

  const availableImages = new Set(
    fs.existsSync(IMAGE_DIR)
      ? fs.readdirSync(IMAGE_DIR).map((f) => f.toLowerCase())
      : []
  );

  const wb = XLSX.read(fs.readFileSync(WORKBOOK), { type: 'buffer' });
  const catalog = collectSectionCatalog(wb);
  const projects = wb.SheetNames.map((name) =>
    buildProject(name, readSheetPairs(wb.Sheets[name]), availableImages, catalog)
  );

  assignRouting(projects);

  const aboutData = readAbout();

  // Workbook order drives site order: number, then variant letter.
  projects.sort(
    (a, b) => (a.number ?? 1e9) - (b.number ?? 1e9) || a.variant.localeCompare(b.variant)
  );

  // Poster frames count as used even though no "Image N" row points at them —
  // otherwise a thumbnail chosen deliberately reads as an orphaned file.
  const usedImages = new Set(
    projects.flatMap((p) => [
      ...p.images.map((i) => i.file.toLowerCase()),
      ...p.videos.map((v) => v.posterFile.toLowerCase()).filter(Boolean),
    ])
  );

  return {
    source: path.relative(ROOT, WORKBOOK).replace(/\\/g, '/'),
    projectCount: projects.length,
    populatedCount: projects.filter((p) => p.populated).length,
    // The catalog is emitted so a reader of the JSON can see the full set of
    // section titles without having to diff two projects against each other.
    sectionCatalog: catalog.map((c) => c.label),
    captionCount: projects.reduce((n, p) => n + p.images.filter((i) => i.caption).length, 0),
    about: aboutData.about,
    projects,
    warnings: {
      missingImages: projects.flatMap((p) =>
        p.images.filter((i) => !i.exists).map((i) => `${p.sheet}: ${i.file}`)
      ),
      unusedImages: [...availableImages].filter((f) => !usedImages.has(f)),
      unparsedVideos: projects.flatMap((p) =>
        p.videos.filter((v) => !v.visible && v.url).map((v) => `${p.sheet}: ${v.url}`)
      ),
      // Captions come from the workbook alone, so a sheet with images and no
      // "Image N Caption" rows ships an uncaptioned gallery. Worth saying out
      // loud rather than letting it look intentional.
      uncaptionedSheets: projects
        .filter((p) => p.images.length && !p.images.some((i) => i.caption))
        .map((p) => `${p.sheet.trim()} (${p.images.length} image(s))`),
      pathLikeSections: projects.flatMap((p) =>
        p.overview
          .filter((o) => o.visible && looksLikePath(o.text))
          .map((o) => `${p.sheet.trim()}: section "${o.label}" holds a file path — mistyped media row?`)
      ),
      fallbackPosters: projects.flatMap((p) =>
        p.videos
          .filter((v) => v.visible && v.posterSource === 'first image')
          .map((v) => `${p.sheet.trim()}: using ${v.posterFile} — add a "Video Thumbnail" row to choose`)
      ),
      about: aboutData.warnings,
    },
  };
}

const data = build();
const json = JSON.stringify(data, null, 2) + '\n';

if (process.argv.includes('--check')) {
  const current = fs.existsSync(OUT_FILE) ? fs.readFileSync(OUT_FILE, 'utf8') : '';
  if (current !== json) {
    console.error('projects.json is out of date — run: npm run data');
    process.exit(1);
  }
  console.log('projects.json is up to date.');
  process.exit(0);
}

fs.mkdirSync(path.dirname(OUT_FILE), { recursive: true });
fs.writeFileSync(OUT_FILE, json);

const img = syncImages();

console.log(`Wrote ${path.relative(ROOT, OUT_FILE).replace(/\\/g, '/')}`);
console.log(`  images/ -> public/images/ : ${img.total} file(s), ${img.copied} copied, ${img.removed} removed`);
console.log(`  ${data.projectCount} sheets — ${data.populatedCount} populated, ${data.projectCount - data.populatedCount} awaiting content`);
console.log(`  section catalog: ${data.sectionCatalog.length} title(s) carried by every project`);
const videoCount = data.projects.reduce((n, p) => n + p.videos.filter((v) => v.visible).length, 0);
console.log(`  ${videoCount} video(s) linked`);
const imageCount = data.projects.reduce((n, p) => n + p.images.length, 0);
console.log(`  ${data.captionCount} of ${imageCount} image(s) captioned from the workbook`);
for (const p of data.projects) {
  const flag = p.populated ? '✓' : '·';
  const parent = p.parentSlug ? `  ↳ under /${p.parentSlug}/` : '';
  console.log(`  ${flag} ${String(p.sheet).padEnd(34)} /${p.slug}/${parent}`);
}
if (data.warnings.missingImages.length) {
  console.log(`  ! ${data.warnings.missingImages.length} referenced image(s) not in images/:`);
  for (const m of data.warnings.missingImages) console.log(`      ${m}`);
}
if (data.warnings.unusedImages.length) {
  console.log(`  ! ${data.warnings.unusedImages.length} file(s) in images/ referenced by no sheet:`);
  for (const m of data.warnings.unusedImages) console.log(`      ${m}`);
}
if (data.warnings.unparsedVideos.length) {
  console.log(`  ! ${data.warnings.unparsedVideos.length} video link(s) with no recognisable YouTube id:`);
  for (const m of data.warnings.unparsedVideos) console.log(`      ${m}`);
}
if (data.warnings.uncaptionedSheets.length) {
  console.log(`  ! ${data.warnings.uncaptionedSheets.length} sheet(s) with images but no "Image N Caption" rows:`);
  for (const m of data.warnings.uncaptionedSheets) console.log(`      ${m}`);
}
if (data.about) {
  console.log(`  ${data.about.source} -> about page: ${data.about.sections.length} section(s)`);
}
if (data.warnings.about.length) {
  console.log(`  ! ${data.warnings.about.length} about-page note(s):`);
  for (const m of data.warnings.about) console.log(`      ${m}`);
}
if (data.warnings.pathLikeSections.length) {
  console.log(`  ! ${data.warnings.pathLikeSections.length} section(s) publishing a raw file path:`);
  for (const m of data.warnings.pathLikeSections) console.log(`      ${m}`);
}
if (data.warnings.fallbackPosters.length) {
  console.log(`  ! ${data.warnings.fallbackPosters.length} video(s) with no chosen poster frame:`);
  for (const m of data.warnings.fallbackPosters) console.log(`      ${m}`);
}
