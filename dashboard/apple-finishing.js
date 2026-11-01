'use strict';
(async function finishSignIn() {
  const marker = document.querySelector('meta[name="apple-handoff"]');
  const handoff = marker?.content;
  marker?.remove();
  try {
    if (!handoff) throw new Error();
    const response = await fetch('/api/auth/apple/complete', {
      method: 'POST', credentials: 'same-origin', cache: 'no-store',
      headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ handoff }),
      signal: AbortSignal.timeout(45000),
    });
    if (!response.ok) throw new Error();
    const result = await response.json();
    const target = new URL(result.redirect, location.origin);
    if (typeof result.redirect !== 'string' || target.origin !== location.origin ||
        !['/', '/editor', '/converter', '/apple-welcome', '/apple-delete'].includes(target.pathname)) throw new Error();
    location.replace(target.pathname + target.search);
  } catch {
    document.getElementById('spinner').hidden = true;
    document.getElementById('heading').textContent = 'Sign-in didn’t finish';
    document.getElementById('status').textContent = 'Please start a fresh sign-in. Your saved playbooks are safe.';
    document.getElementById('retry').hidden = false;
  }
})();
