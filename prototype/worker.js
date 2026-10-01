// Minimal Cloudflare Worker: GET /api/home?accountId=1 -> events + friends + settlement input.
// No auth on purpose - this is a latency prototype, do not deploy with real data open.
import { neon } from '@neondatabase/serverless';
import { homeQueries } from './queries.mjs';

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    if (url.pathname !== '/api/home') return new Response('not found', { status: 404 });
    const accountId = Number(url.searchParams.get('accountId') || 1);
    const t = Date.now();
    const sql = neon(env.DATABASE_URL);
    const [events, friends, txs, settlements] = await sql.transaction(homeQueries(sql, accountId));
    return Response.json({ events, friends, txs, settlements, serverMs: Date.now() - t });
  },
};
