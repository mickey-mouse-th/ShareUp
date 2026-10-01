// Usage: DATABASE_URL='postgres://...' ACCOUNT_ID=1 node bench.mjs [runs]
// Measures the getHomeData reads against Neon from THIS machine, so it shows
// DB + network cost only (no Apps Script bridge). Compare with the browser
// snippet in README.md for the Apps Script side.
import { neon } from '@neondatabase/serverless';
import { homeQueries } from './queries.mjs';

const url = process.env.DATABASE_URL;
const accountId = Number(process.env.ACCOUNT_ID || 1);
const runs = Number(process.argv[2] || 20);
if (!url) { console.error('Set DATABASE_URL'); process.exit(1); }

const sql = neon(url);
const ms = (t) => Math.round(performance.now() - t);
const stats = (a) => {
  const s = [...a].sort((x, y) => x - y);
  return `median ${s[Math.floor(s.length / 2)]} ms | p95 ${s[Math.floor(s.length * 0.95)]} ms | min ${s[0]} | max ${s.at(-1)}`;
};

async function parallel() { await Promise.all(homeQueries(sql, accountId)); }
async function batched() { await sql.transaction(homeQueries(sql, accountId)); }

for (const [name, fn] of [['parallel (4 HTTP calls)', parallel], ['batched (1 HTTP call)', batched]]) {
  const t0 = performance.now(); await fn(); const first = ms(t0);
  const times = [];
  for (let i = 0; i < runs; i++) { const t = performance.now(); await fn(); times.push(ms(t)); }
  console.log(`${name}\n  first call: ${first} ms (may include Neon wake-up)\n  warm x${runs}: ${stats(times)}`);
}
