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
 * 225235.png"), which are not valid in an href until encoded.
 */
export function assetUrl(pathname = '/') {
  return url(pathname).split('/').map(encodeURIComponent).join('/');
}

/** Compares a link target against the current URL, ignoring trailing slashes. */
export function isActive(currentPathname, target) {
  const norm = (p) => (String(p).replace(/\/+$/, '') || '/');
  const here = norm(currentPathname);
  const there = norm(target);
  return here === there || (there !== norm(BASE) && here.startsWith(there + '/'));
}
