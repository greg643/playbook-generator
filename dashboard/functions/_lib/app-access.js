import { readBoundedUtf8Text } from './auth.js';
import { validGSSBearer } from './gss-bearer.js';

export const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
export const isUUID = value => typeof value === 'string' && UUID_RE.test(value);
export const isTeamID = value => typeof value === 'string' && value.startsWith('T-') && isUUID(value.slice(2));
export class AppAccessError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

export function appBearer(request) {
  const header = request.headers.get('authorization') || '';
  if (!header.startsWith('Bearer ') || !validGSSBearer(header.slice(7))) throw new AppAccessError(401, 'Sign in to GSS again');
  return header.slice(7);
}

async function gssRead(env, bearer, path, body) {
  let base;
  try {
    base = new URL(env.GSS_API_BASE);
    if (base.protocol !== 'https:' || base.username || base.password || base.search || base.hash) throw Error();
  } catch { throw new AppAccessError(503, 'Team access is temporarily unavailable'); }
  let response;
  try {
    response = await fetch(base.href.replace(/\/$/, '') + path, {
      method: 'POST', redirect: 'manual', signal: AbortSignal.timeout(10000),
      headers: { Authorization: `Bearer ${bearer}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
  } catch { throw new AppAccessError(503, 'Team access is temporarily unavailable'); }
  if (!response.ok) {
    await response.body?.cancel();
    if (response.status === 401) throw new AppAccessError(401, 'Sign in to GSS again');
    if ([403, 404].includes(response.status)) throw new AppAccessError(404, 'Playbook not found');
    throw new AppAccessError(503, 'Team access is temporarily unavailable');
  }
  const text = await readBoundedUtf8Text(response, 32768);
  try { if (text === null) throw Error(); return JSON.parse(text); }
  catch { throw new AppAccessError(503, 'Team access is temporarily unavailable'); }
}

// Fresh identity + membership reads, never decode a client JWT or infer Team
// authority from the Playbook cookie/email. Do not cache raw bearer credentials.
export async function verifyTeamAccess(env, bearer, teamId) {
  if (!isTeamID(teamId)) throw new AppAccessError(404, 'Playbook not found');
  if (!validGSSBearer(bearer)) {
    throw new AppAccessError(401, 'Sign in to GSS again');
  }
  const me = await gssRead(env, bearer, '/v2/me', { include_teams: false });
  if (!/^[a-f0-9]{16}$/.test(me?.account_key || '') || me.account_status !== 'enabled' ||
      me.migration_complete !== true) throw new AppAccessError(403, 'GSS account access is not ready');
  const result = await gssRead(env, bearer, '/v2/teams/get', { team_id: teamId });
  const team = result?.team;
  if (!team || team.team_id !== teamId || team.status !== 'active' || !['leader', 'member'].includes(team.role)) {
    throw new AppAccessError(404, 'Playbook not found');
  }
  return { accountKey: me.account_key, teamId, isManager: team.role === 'leader' && team.is_family_member !== true };
}

// One bounded discovery page for the editor; publication still rechecks the
// selected Team directly. Never infer management from names or email addresses.
export async function managedTeams(env, bearer, accountKey, continuation = null) {
  if (continuation !== null && (typeof continuation !== 'string' || continuation.length > 4096 || !continuation)) {
    throw new AppAccessError(400, 'Invalid Team page');
  }
  const me = await gssRead(env, bearer, '/v2/me', { include_teams: false });
  if (me?.account_key !== accountKey || !/^[a-f0-9]{16}$/.test(accountKey || '') ||
      me.account_status !== 'enabled' || me.migration_complete !== true) throw new AppAccessError(401, 'Continue with Apple again');
  const page = await gssRead(env, bearer, '/v2/teams/list', { limit: 25, ...(continuation ? { continuation } : {}) });
  if (!Array.isArray(page?.teams) || page.teams.length > 25 ||
      (page.next !== null && (typeof page.next !== 'string' || !page.next || page.next.length > 4096))) {
    throw new AppAccessError(503, 'Team access is temporarily unavailable');
  }
  const teams = page.teams.filter(team => team && isTeamID(team.team_id) && team.status === 'active' &&
    team.role === 'leader' && team.is_family_member !== true && typeof team.name === 'string' && team.name.trim())
    .map(team => ({ teamId: team.team_id, name: team.name.slice(0, 80),
      season: typeof team.active_season_label === 'string' ? team.active_season_label.slice(0, 80) : '' }));
  return { teams, next: page.next };
}
