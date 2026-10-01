// The same three reads _pgGetHomeData does in Code.js, expressed for the
// Neon serverless driver. Shared by bench.mjs and worker.js.
export function homeQueries(sql, accountId) {
  return [
    sql`SELECT * FROM event WHERE account_id = ${accountId} ORDER BY cre_at DESC`,
    sql`SELECT id, name, is_self FROM friend WHERE account_id = ${accountId} ORDER BY id`,
    sql`SELECT t.event_id, t.payer_id, s.friend_id, s.amount FROM transaction t
        JOIN split s ON s.transaction_id = t.id JOIN event e ON e.id = t.event_id
        WHERE e.account_id = ${accountId} AND t.excluded_at IS NULL`,
    sql`SELECT s.event_id, s.from_id, s.to_id, s.amount FROM settlement s
        JOIN event e ON e.id = s.event_id WHERE e.account_id = ${accountId}`,
  ];
}
