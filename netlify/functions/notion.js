const https = require('https');
exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return { statusCode: 405, body: 'Method not allowed' };
  let body;
  try { body = JSON.parse(event.body); } catch { return { statusCode: 400, body: 'Invalid JSON' }; }
  const { endpoint, method = 'POST', payload, notionKey } = body;
  if (!endpoint || !notionKey) return { statusCode: 400, body: 'Missing params' };
  return new Promise((resolve) => {
    const data = payload ? JSON.stringify(payload) : undefined;
    const options = { hostname: 'api.notion.com', path: `/v1${endpoint}`, method, headers: { 'Authorization': `Bearer ${notionKey}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json', ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {}) } };
    const req = https.request(options, (res) => { let d = ''; res.on('data', c => d += c); res.on('end', () => resolve({ statusCode: res.statusCode, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }, body: d })); });
    req.on('error', (err) => resolve({ statusCode: 500, body: JSON.stringify({ error: err.message }) }));
    if (data) req.write(data);
    req.end();
  });
};
