import { jsonNoStore, readBoundedUtf8Text } from '../../_lib/auth.js';
import { AppAccessError } from '../../_lib/app-access.js';
import { finishPublicationJob } from '../../_lib/app-publication-jobs.js';

export async function onRequestPost({ request, env }) {
  try {
    if (request.headers.get('content-type')?.split(';')[0] !== 'application/json') {
      return jsonNoStore({ error: 'Expected JSON' }, { status: 400 });
    }
    const text = await readBoundedUtf8Text(request, 1024);
    if (text === null) return jsonNoStore({ error: 'Request too large' }, { status: 413 });
    let body;
    try { body = JSON.parse(text); } catch { return jsonNoStore({ error: 'Invalid JSON' }, { status: 400 }); }
    return jsonNoStore(await finishPublicationJob(env, body));
  } catch (error) {
    if (error instanceof AppAccessError) return jsonNoStore({ error: error.message }, { status: error.status });
    console.error(JSON.stringify({ event: 'app_publication_completion_failed' }));
    return jsonNoStore({ error: 'Publication could not finish' }, { status: 503 });
  }
}
