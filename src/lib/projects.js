// Read-only accessors over the generated data file. Nothing here reaches into
// the workbook — regenerate src/data/projects.json with `npm run data`.

import data from '../data/projects.json';
import { url } from './url.js';

export const AUTHOR = 'Nicholas Paradizov';
export const SITE_TITLE = 'The Engineering Portfolio';

/** Every sheet in workbook order, primaries and sub-pages alike. */
export const projects = data.projects;

export const projectCount = data.projectCount;
export const populatedCount = data.populatedCount;

/**
 * About-page prose, from the "About Me" sheet of data/additional-info.xlsx.
 * Sections carry `paragraphs`, and a "·"-separated one-liner also carries
 * `items` so it can render as chips instead of a sentence.
 */
export const about = data.about ?? { sections: [] };
export const aboutSections = about.sections ?? [];

/** LinkedIn is the one public contact channel the site advertises. */
export const LINKEDIN_URL = 'https://www.linkedin.com/in/nicholas-paradizov-b131a7357/';

/** "1A", "1B", "2" — the workbook's own label for a sheet. */
export function projectCode(project) {
  if (project.number === null) return '—';
  return `${project.number}${project.variant}`;
}

/** Full title where available, falling back to the sheet's short name. */
export function projectTitle(project) {
  return project.title || project.short;
}

export function projectHref(project) {
  return url(`/projects/${project.slug}/`);
}

/**
 * The N/A rendering rule.
 *
 * Every project carries the whole section catalog, so a section this project
 * has nothing to say about is present in the JSON with the workbook's literal
 * "N/A" and `visible: false`. The data keeps it; the page never shows it. Route
 * every read of `overview` and `videos` through these two, or an "N/A" leaks
 * onto the page as if it were prose.
 */
export function sectionsOf(project) {
  return project.overview.filter((block) => block.visible);
}

export function videosOf(project) {
  return (project.videos ?? []).filter((video) => video.visible);
}

/** First line of real prose — the card blurb and the meta description. */
export function blurbOf(project) {
  const sections = sectionsOf(project);
  return sections.find((s) => s.label === 'What it is')?.text ?? sections[0]?.text ?? '';
}

/** Sub-pages of a project (1B, 1C… under 1A), in workbook order. */
export function variantsOf(project) {
  return project.variantSlugs
    .map((slug) => projects.find((p) => p.slug === slug))
    .filter(Boolean);
}

export function parentOf(project) {
  return project.parentSlug
    ? projects.find((p) => p.slug === project.parentSlug) ?? null
    : null;
}

/** Prev/next across the flat workbook order, for footer navigation. */
export function neighborsOf(project) {
  const i = projects.findIndex((p) => p.slug === project.slug);
  return {
    prev: i > 0 ? projects[i - 1] : null,
    next: i >= 0 && i < projects.length - 1 ? projects[i + 1] : null,
  };
}

/** Distinct values of a field across populated projects — used for filter chips. */
export function distinctValues(field) {
  const seen = new Set();
  for (const p of projects) {
    const value = p[field];
    if (Array.isArray(value)) value.forEach((v) => v && seen.add(v));
    else if (value) seen.add(value);
  }
  return [...seen].sort((a, b) => a.localeCompare(b));
}

/** One entry per project for the Cmd+K search overlay. */
export function searchIndex() {
  return projects.map((p) => ({
    title: `${projectCode(p)} — ${projectTitle(p)}`,
    path: projectHref(p),
    description: p.populated ? p.projectType || blurbOf(p) : 'Coming soon',
    tags: [
      projectCode(p),
      p.short,
      p.projectType,
      p.status,
      ...p.disciplines,
      ...p.tools,
    ].filter(Boolean),
  }));
}
