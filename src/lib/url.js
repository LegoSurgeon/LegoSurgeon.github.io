// Every internal href and asset path goes through url(), so moving the site
// from local (base '/') to a GitHub Pages subpath ('/engineering-portfolio')
// is a config change rather than a find-and-replace.

const BASE = import.meta.env.BASE_URL || '/';

/** url('/projects/') -> '/projects/'  |  '/engineering-portfolio/projects/' */
export function url(pathname = '/') {
  const prefix = BASE.replace(/\/+$/, '');
  const suffix = String(pathname).replace(/^\/+/, '');
  return `${prefix}/${suffix}`;
}

/**
 * Image filenames come from Excel and contain spaces ("Screenshot 2026-08-13
 * 225235.png") and commas ("Aug 18, 2026, 10_31_34 PM.png"), so they need
 * encoding before they are valid in an href.
 *
 * encodeURI, not encodeURIComponent-per-segment: the latter also escapes
 * characters that are perfectly legal in a path, notably the comma, and the
 * static file server behind `astro preview` does not decode a %2C back to a
 * comma — it 404s on a file that is really there. encodeURI escapes the space
 * and leaves the comma alone, which both the browser and the server accept.
 *
 * It leaves # and ? alone too, and those genuinely must be escaped, or the
 * browser reads the rest of the filename as a fragment or query string.
 */
export function assetUrl(pathname = '/') {
  return encodeURI(url(pathname)).replace(/#/g, '%23').replace(/\?/g, '%3F');
}

/** Compares a link target against the current URL, ignoring trailing slashes. */
export function isActive(currentPathname, target) {
  const norm = (p) => (String(p).replace(/\/+$/, '') || '/');
  const here = norm(currentPathname);
  const there = norm(target);
  return here === there || (there !== norm(BASE) && here.startsWith(there + '/'));
}
