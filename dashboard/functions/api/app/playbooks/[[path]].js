import { jsonNoStore } from '../../../_lib/auth.js';
import { AppAccessError, appBearer, isUUID, verifyTeamAccess } from '../../../_lib/app-access.js';
import { livePublisher, publicationManifest, readTeamPublications, snapshotPrefix } from '../../../_lib/app-publications.js';

// GET /api/app/playbooks/{teamId}
// GET /api/app/playbooks/{teamId}/{bookId}/{snapshotId}/{assetId}
// The list contains bounded manifests, so no separate manifest route is needed.
export async function onRequestGet(context) {
  try {
    const path = Array.isArray(context.params.path) ? context.params.path : String(context.params.path || '').split('/');
    if (![1, 4].includes(path.length)) throw new AppAccessError(404, 'Playbook not found');
    const [teamId, bookId, snapshotId, assetId] = path;
    if (path.length === 4 && (!isUUID(bookId) || !isUUID(snapshotId) || !/^[A-Za-z0-9_-]{1,128}$/.test(assetId))) {
      throw new AppAccessError(404, 'Playbook not found');
    }
    const access = await verifyTeamAccess(context.env, appBearer(context.request), teamId);
    const { value: catalog } = await readTeamPublications(context.env, teamId);
    const visible = [];
    for (const entry of catalog.entries) {
      if (!entry.current || (bookId && (entry.bookId !== bookId || entry.current.snapshotId !== snapshotId))) continue;
      if (!await livePublisher(context.env, entry)) continue;
      const manifest = await publicationManifest(context.env, teamId, entry);
      if (!manifest || (manifest.coachesOnly && !access.isManager)) continue;
      if (!bookId) { visible.push(manifest); continue; }
      const descriptor = manifest.assets.find(a => a.id === assetId);
      if (!descriptor) continue;
      const object = await context.env.PLAYBOOK_BUCKET.get(snapshotPrefix(entry.ownerId, bookId, snapshotId) + assetId);
      const hash = object?.checksums?.sha256 && [...new Uint8Array(object.checksums.sha256)].map(n => n.toString(16).padStart(2, '0')).join('');
      if (!object || object.size !== descriptor.bytes || hash !== descriptor.sha256) {
        await object?.body?.cancel(); throw new AppAccessError(503, 'Playbook download is temporarily unavailable');
      }
      return new Response(object.body, { headers: {
        'Content-Type': descriptor.mediaType, 'Content-Length': String(descriptor.bytes),
        'Cache-Control': 'private, no-store', 'X-Content-Type-Options': 'nosniff',
      } });
    }
    if (bookId) throw new AppAccessError(404, 'Playbook not found');
    return jsonNoStore({ schema: 'gss.team-playbook-library.v1', teamId, canPublish: access.isManager,
      defaultBookId: visible.some(m => m.bookId === catalog.defaultBookId) ? catalog.defaultBookId : visible[0]?.bookId || null,
      books: visible });
  } catch (error) {
    if (error instanceof AppAccessError) return jsonNoStore({ error: error.message }, { status: error.status });
    // Do not log request headers, tokens, private play content or upstream text.
    console.error(JSON.stringify({ event: 'app_playbook_read_failed' }));
    return jsonNoStore({ error: 'Playbooks are temporarily unavailable' }, { status: 503 });
  }
}
