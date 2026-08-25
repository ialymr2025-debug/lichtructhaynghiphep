// Every API call is made same-origin, against the site's own /api path.
//
// Locally that is the Express server, which serves the app and the API on one port.
// In production the frontend sits on Vercel and the API on Render — two different
// hosts — so vercel.json rewrites /api/* through to the Render service. The browser
// still only ever sees one origin.
//
// That indirection is the point. Talking to the Render URL directly would make every
// request cross-origin, which brings two problems: the API sends no CORS headers, so
// the browser discards the response even when the server answered correctly; and the
// session cookie is SameSite=Lax, so it would not be attached to cross-site requests
// at all. Going through the proxy keeps the cookie first-party and sidesteps both.
//
// Hence: no configurable base URL. A wrong value here reintroduces exactly the
// cross-origin failure this is meant to avoid.
export const API_BASE = '';

export const apiCall = (endpoint: string, options?: RequestInit) => {
  return fetch(API_BASE + endpoint, options);
};
