# Latency prototype

## 1. New stack (Neon via serverless driver)
    npm install
    DATABASE_URL='postgres://...' ACCOUNT_ID=1 npm run bench 30

Optional, closer to production (Worker at the edge):
    npx wrangler secret put DATABASE_URL
    npx wrangler deploy
    curl -w '\n%{time_total}s\n' 'https://shareup-prototype.<you>.workers.dev/api/home?accountId=1'

## 2. Current stack (Apps Script)
Open the deployed web app, log in, then paste in the browser console
(select the iframe context "userCodeAppPanel" first):

    (async()=>{const t=[];for(let i=0;i<10;i++){const s=performance.now();
    await new Promise(r=>google.script.run.withSuccessHandler(r).getHomeData(TOKEN));
    t.push(Math.round(performance.now()-s));}console.log(t,'median',t.sort((a,b)=>a-b)[5])})()

Replace TOKEN with your session token (localStorage / the variable Shared_js.html uses).
Compare: Apps Script per-call vs. bench.mjs "warm" and the Worker curl time.
Use the Neon DB URL from Script Properties (DB_URL/DB_USER/DB_PASS) - never commit it.
