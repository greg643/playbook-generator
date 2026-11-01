// Pure app-publication contract. Call only with the owned, server-read source
// revision and verified render descriptors; route authorization is separate.
import { normalizePlaybookName } from './playbooks.js';
import { COACH_LAYOUTS } from './coach-layouts.js';

export const APP_SNAPSHOT_SCHEMA = 'gss.team-playbook-snapshot.v1';
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const ID = /^[A-Za-z0-9_-]{1,128}$/;
const PLAY_ID = /^[A-Za-z0-9_-]{1,122}$/; // Leave room for the image- asset prefix.
const HASH = /^[a-f0-9]{64}$/;
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;
const MAX_PDF_BYTES = 16 * 1024 * 1024;
const MAX_SNAPSHOT_BYTES = 64 * 1024 * 1024;
const integer = (n, max) => Number.isSafeInteger(n) && n > 0 && n <= max;
const matches = (value, pattern) => typeof value === 'string' && pattern.test(value);

function descriptor(asset, id, mediaType) {
  if (!asset || asset.id !== id || asset.mediaType !== mediaType || !matches(asset.sha256, HASH) ||
      !integer(asset.bytes, mediaType === 'image/png' ? MAX_IMAGE_BYTES : MAX_PDF_BYTES)) {
    throw new Error('Missing or invalid rendered asset');
  }
  const result = { id, mediaType, bytes: asset.bytes, sha256: asset.sha256 };
  if (mediaType === 'image/png') {
    if (!integer(asset.width, 8192) || !integer(asset.height, 8192) ||
        asset.width * asset.height > 16_777_216) throw new Error('Invalid image dimensions');
    Object.assign(result, { width: asset.width, height: asset.height });
  }
  // Never propagate backend storage keys, URLs, tokens or arbitrary extra fields.
  return result;
}

export function buildAppSnapshot({ bookId, snapshotId, teamId, name, document,
  expectedRevision, publishedAt, coachesOnly = false, assets, layouts = {} }) {
  if (!matches(bookId, UUID) || !matches(snapshotId, UUID) || typeof teamId !== 'string' ||
      !teamId.startsWith('T-') || !UUID.test(teamId.slice(2)) ||
      typeof name !== 'string' || normalizePlaybookName(name) !== name || typeof coachesOnly !== 'boolean' ||
      typeof publishedAt !== 'string' || !Number.isFinite(Date.parse(publishedAt)) ||
      new Date(publishedAt).toISOString() !== publishedAt) throw new Error('Invalid publication identity');
  if (!document || document.schema !== 2 || !matches(expectedRevision, UUID) ||
      document.revision !== expectedRevision) throw new Error('Reload the saved playbook before publishing');
  if (!Array.isArray(assets) || assets.length > 90 ||
      assets.some(a => !a || !matches(a.id, ID)) || new Set(assets.map(a => a.id)).size !== assets.length) {
    throw new Error('Invalid render asset set');
  }
  const byId = new Map(assets.map(a => [a.id, a]));
  const plays = [], usedAssets = [], coachCards = {}, ids = new Set(), formats = new Set();
  for (const side of ['offense', 'defense']) {
    const source = document[side];
    if (!Array.isArray(source) || source.length > (side === 'offense' ? 64 : 24)) {
      throw new Error('Invalid play selection');
    }
    const numbers = new Set();
    const selected = source.filter(play => {
      if (!play || typeof play !== 'object' || Array.isArray(play)) throw new Error('Invalid play');
      return play.included !== false; // Same compatibility rule as the editor.
    });
    for (const play of selected) {
      if (!matches(play.id, PLAY_ID) || ids.has(play.id) || !integer(play.num, 999) || numbers.has(play.num) ||
          ![5, 6].includes(play.playersPerSide) || typeof play.name !== 'string' ||
          play.name.length > 60 || /[\u0000-\u001f\u007f-\u009f]/u.test(play.name)) {
        throw new Error('Save valid play IDs, numbers and formats before publishing');
      }
      ids.add(play.id); numbers.add(play.num); formats.add(play.playersPerSide);
      const imageId = `image-${play.id}`;
      usedAssets.push(descriptor(byId.get(imageId), imageId, 'image/png'));
      plays.push({ id: play.id, sourcePlayId: play.id, source: 'editor', side, name: play.name, number: play.num,
        displayNumber: side === 'defense' && play.num <= 26 ? String.fromCharCode(64 + play.num) : String(play.num),
        playersPerSide: play.playersPerSide, imageAssetId: imageId });
    }
    if (selected.length) {
      const layout = layouts[side] ?? 16;
      if (!COACH_LAYOUTS.includes(layout)) throw new Error('Invalid coach card layout');
      const pdfId = `coach-${side}`;
      usedAssets.push(descriptor(byId.get(pdfId), pdfId, 'application/pdf'));
      coachCards[side] = { assetId: pdfId, playsPerPage: layout };
    }
  }
  if (!plays.length) throw new Error('Include at least one play before publishing');
  if (usedAssets.length !== assets.length) throw new Error('Render contains unselected or extra assets');
  if (usedAssets.reduce((n, a) => n + a.bytes, 0) > MAX_SNAPSHOT_BYTES) throw new Error('Snapshot is too large');
  return { schema: APP_SNAPSHOT_SCHEMA, bookId, snapshotId, teamId, name,
    sourceRevision: expectedRevision, publishedAt, coachesOnly,
    format: formats.size === 1 ? `${[...formats][0]}v${[...formats][0]}` : 'mixed',
    plays, assets: usedAssets, coachCards };
}
