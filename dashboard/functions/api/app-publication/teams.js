import { jsonNoStore, readBoundedUtf8Text } from '../../_lib/auth.js';
import { AppAccessError, managedTeams } from '../../_lib/app-access.js';
import { readAppWebSession } from '../../_lib/app-web-session.js';

export async function onRequestPost({ request, env }) {
  if (request.headers.get('origin') !== new URL(request.url).origin ||
      request.headers.get('content-type')?.split(';')[0].trim() !== 'application/json') {
    return jsonNoStore({ error: 'Invalid request origin' }, { status: 403 });
  }
  try {
    const text = await readBoundedUtf8Text(request, 8192);
    if (text === null) return jsonNoStore({ error: 'Request too large' }, { status: 413 });
    let body;
    try { body = JSON.parse(text); } catch { throw new AppAccessError(400, 'Invalid request'); }
    if (!body || typeof body !== 'object' || Array.isArray(body) ||
        Object.keys(body).some(key => key !== 'continuation')) throw new AppAccessError(400, 'Invalid request');
    const { user, bearer } = await readAppWebSession(request, env);
    return jsonNoStore(await managedTeams(env, bearer, user.account.appleAccountKey, body.continuation ?? null));
  } catch (error) {
    if (error instanceof AppAccessError) return jsonNoStore({ error: error.message }, { status: error.status });
    console.error(JSON.stringify({ event: 'playbook_team_discovery_failed' }));
    return jsonNoStore({ error: 'Team access is temporarily unavailable' }, { status: 503 });
  }
}
