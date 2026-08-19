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
 *   Project Type                 FTC Robotics / Modular …
 *   Role                         Sole chassis and subsystem designer
 *   Disciplines                  Mechanical Design, Mechatronics, …
 *   Tools                        Fusion 360, KiCad, …
 *   Project Overview             ← section separator, blank B cell
 *   What it is                   Modular Chassis MK2 is the second …
 *   …
 *   Image 1                      C:\Users\…\Screenshot ….png
 *
 * Sheet names carry ordering and grouping: "1A. MK2 Modular Chassis" means
 * project 1, variant A. Variants of one project (1A, 1B, 1C…) are grouped
 * together, with A as the primary and the rest as sub-pages. A sheet with no
 * letter ("2. MK1 Modular Chassis") is a standalone project.
 *
 * Anything not in KNOWN_FIELDS is preserved under `extra`, so new rows added to
 * the workbook survive this script rather than being silently dropped.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
// xlsx ships as CommonJS; the default import is the whole namespace. We hand it
// a Buffer via read() rather than readFile() so it never needs its own fs shim.
import XLSX from 'xlsx';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const WORKBOOK = path.join(ROOT, 'data', 'Engineering Portfolio.xlsx');
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
  'project type': 'projectType',
  'role': 'role',
  'disciplines': 'disciplines',
  'tools': 'tools',
};

// Narrative prose rows, in the order they should appear on the page. Rows are
// emitted in workbook order, so reordering the spreadsheet reorders the page.
const OVERVIEW_FIELDS = {
  'what it is': 'What it is',
  'why i made it': 'Why I made it',
  'what i designed': 'What I designed',
  'key engineering challenge': 'Key engineering challenge',
  'what i learned': 'What I learned',
};

const SEPARATORS = new Set(['project overview']);

// Comma-separated cells that become chip lists on the page.
const LIST_FIELDS = new Set(['disciplines', 'tools']);

const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
const key = (s) => norm(s).toLowerCase();

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

function buildProject(sheetName, pairs, availableImages) {
  const { number, variant, short } = parseSheetName(sheetName);

  const project = {
    sheet: sheetName,
    number,
    variant,
    short,
    slug: slugify(short || sheetName),
    title: '',
    status: '',
    projectType: '',
    role: '',
    disciplines: [],
    tools: [],
    overview: [],
    images: [],
    extra: {},
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

    if (k in OVERVIEW_FIELDS) {
      if (value) project.overview.push({ label: OVERVIEW_FIELDS[k], text: value });
      continue;
    }

    // "Image 1", "Image 12", … — index preserved so gaps don't renumber.
    const img = k.match(/^image\s*(\d+)$/);
    if (img) {
      if (!value) continue;
      const file = imageFileName(value);
      project.images.push({
        index: Number(img[1]),
        file,
        src: `/images/${file}`,
        exists: availableImages.has(file.toLowerCase()),
      });
      continue;
    }

    // Unrecognised row: keep prose under overview so it still reaches the page,
    // and stash the raw value so nothing is lost.
    if (value) {
      project.overview.push({ label: norm(label), text: value });
      project.extra[norm(label)] = value;
    }
  }

  project.images.sort((a, b) => a.index - b.index);

  // A sheet counts as populated once it has a title AND something to show.
  project.populated = Boolean(
    project.title && (project.overview.length || project.images.length || project.status)
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
  const projects = wb.SheetNames.map((name) =>
    buildProject(name, readSheetPairs(wb.Sheets[name]), availableImages)
  );

  assignRouting(projects);

  // Workbook order drives site order: number, then variant letter.
  projects.sort(
    (a, b) => (a.number ?? 1e9) - (b.number ?? 1e9) || a.variant.localeCompare(b.variant)
  );

  const usedImages = new Set(projects.flatMap((p) => p.images.map((i) => i.file.toLowerCase())));

  return {
    source: path.relative(ROOT, WORKBOOK).replace(/\\/g, '/'),
    projectCount: projects.length,
    populatedCount: projects.filter((p) => p.populated).length,
    projects,
    warnings: {
      missingImages: projects.flatMap((p) =>
        p.images.filter((i) => !i.exists).map((i) => `${p.sheet}: ${i.file}`)
      ),
      unusedImages: [...availableImages].filter((f) => !usedImages.has(f)),
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
