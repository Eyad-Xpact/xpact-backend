const express = require('express');
const cors = require('cors');
const https = require('https');
const app = express();

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ── Health check ────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'XPACT Proposal API', version: '3.0', ready: true });
});

// ── Score RFP ───────────────────────────────────────────────
app.post('/score-rfp', async (req, res) => {
  const { rfpText } = req.body;
  if (!rfpText) return res.status(400).json({ error: 'rfpText required' });
  const prompt = `You are an expert event management consultant. Analyze this RFP and return ONLY a JSON object with this exact structure:
{"score":<0-100>,"recommendation":"<Pursue|Consider|Pass>","summary":"<2 sentence summary>","criteria":{"budgetFit":<0-20>,"timeline":<0-20>,"eventTypeMatch":<0-20>,"scopeComplexity":<0-20>,"strategicValue":<0-20>},"risks":["<r1>","<r2>","<r3>"],"opportunities":["<o1>","<o2>","<o3>"]}
RFP TEXT:\n${rfpText.slice(0, 4000)}`;
  callClaude(prompt, 1000, (err, text) => {
    if (err) return res.status(500).json({ error: err });
    try { res.json(JSON.parse(text.replace(/```json|```/g, '').trim())); }
    catch(e) { res.status(500).json({ error: 'Parse error' }); }
  });
});

// ── Generate section ─────────────────────────────────────────
app.post('/generate-section', (req, res) => {
  const { prompt, section } = req.body;
  if (!prompt) return res.status(400).json({ error: 'prompt required' });
  callClaude(prompt, 1000, (err, text) => {
    if (err) return res.status(500).json({ error: err });
    res.json({ text, section });
  });
});

function callClaude(prompt, maxTokens, cb) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return cb('API key not configured');
  const body = JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: maxTokens, messages: [{ role: 'user', content: prompt }] });
  const opts = { hostname: 'api.anthropic.com', path: '/v1/messages', method: 'POST', headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'Content-Length': Buffer.byteLength(body) } };
  const req = https.request(opts, (res) => {
    let data = '';
    res.on('data', c => data += c);
    res.on('end', () => {
      try { cb(null, JSON.parse(data).content[0].text); }
      catch(e) { cb('Parse error: ' + data.slice(0, 100)); }
    });
  });
  req.on('error', e => cb(e.message));
  req.write(body);
  req.end();
}

// ============================================================
// RFP DISCOVERY — Forsah public API (no login needed)
// ============================================================

const FORSAH_API = 'forsah-api.910ths.sa';

const EVENT_KEYWORDS = [
  'فعاليات','فعالية','حفل','حفلات','معرض','معارض',
  'مؤتمر','مؤتمرات','ملتقى','احتفال','احتفالية',
  'ندوة','منتدى','إكسبو','أمسية','تنظيم','أوبرا',
  'معرض فني','حفل موسيقي','مسرح','عروض','ترفيه',
  'event','events','conference','exhibition','expo',
  'ceremony','gala','seminar','forum','festival',
  'concert','entertainment','show','stage'
];

function isEventRelated(title, categories) {
  const text = (title || '').toLowerCase();
  const cats = (categories || []).map(c => (c.name && c.name.ar ? c.name.ar : (c.nameAr || c.name || ''))).join(' ').toLowerCase();
  const combined = text + ' ' + cats;
  return EVENT_KEYWORDS.some(kw => combined.includes(kw.toLowerCase()));
}

function scoreRelevance(title, categories) {
  let score = 50;
  const text = ((title || '') + ' ' + (categories || []).map(c => c.nameAr || '').join(' ')).toLowerCase();
  const hits = EVENT_KEYWORDS.filter(kw => text.includes(kw.toLowerCase()));
  score += Math.min(40, hits.length * 10);
  return Math.min(100, score);
}

function forsahGet(path) {
  return new Promise((resolve) => {
    const opts = {
      hostname: FORSAH_API,
      path,
      method: 'GET',
      headers: {
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (compatible; XpactAI/3.0)',
        'Origin': 'https://forsah.sa',
        'Referer': 'https://forsah.sa/'
      },
      timeout: 15000
    };
    const req = https.request(opts, (res) => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => resolve({ ok: res.statusCode < 400, status: res.statusCode, data: raw }));
    });
    req.on('error', () => resolve({ ok: false, data: null }));
    req.on('timeout', () => { req.destroy(); resolve({ ok: false, data: null }); });
    req.end();
  });
}

async function fetchForsahPage(page) {
  const r = await forsahGet(`/api/v1/opportunities?perPage=50&page=${page}`);
  if (!r.ok || !r.data) return null;
  try { return JSON.parse(r.data); }
  catch(e) { return null; }
}

async function fetchAllForsahEvents() {
  const results = [];
  let page = 1;
  let totalPages = 1;

  // Fetch up to 10 pages (500 opportunities) — enough to find all event-related ones
  while (page <= Math.min(totalPages, 10)) {
    console.log(`Fetching Forsah page ${page}/${totalPages}...`);
    const data = await fetchForsahPage(page);

    if (!data || !data.result) break;

    // Set total pages from first response
    if (page === 1 && data.pagination) {
      totalPages = data.pagination.pageCount || 1;
      console.log(`Forsah total: ${data.pagination.totalCount} opportunities across ${totalPages} pages`);
    }

    for (const item of data.result) {
      if (isEventRelated(item.title, item.categories)) {
        results.push({
          id: 'fors_' + item.id,
          title: item.title || '',
          agency: item.publisher ? (item.publisher.nameAr || item.publisher.name || '') : '',
          deadline: item.dueDate || item.closeDate || (item.daysToGo ? `${item.daysToGo} يوم` : ''),
          budget: item.valueRange ? `${(item.valueRange.nameAr || item.valueRange.name || '')} (${(item.valueRange.min || 0).toLocaleString()}–${(item.valueRange.max || 0).toLocaleString()} ر.س)` : '',
          budgetMin: item.valueRange ? (item.valueRange.min || 0) : 0,
          daysLeft: item.daysToGo || null,
          status: item.status || 'مفتوح',
          source: 'Forsah',
          tenderUrl: `https://forsah.sa/opportunities/${item.id}`
        });
      }
    }

    page++;
    // Small delay to be respectful to the API
    await new Promise(r => setTimeout(r, 300));
  }

  return results;
}

// ── Cache ───────────────────────────────────────────────────
let rfpCache = { tenders: [], lastFetch: null, newCount: 0, scanning: false };

async function runDiscovery() {
  if (rfpCache.scanning) {
    console.log('Discovery already running, skipping');
    return;
  }
  rfpCache.scanning = true;
  console.log('Starting Forsah discovery scan...');

  try {
    const prevIds = new Set(rfpCache.tenders.map(t => t.id));
    const fresh = await fetchAllForsahEvents();

    const scored = fresh
      .map(t => ({
        ...t,
        relevanceScore: scoreRelevance(t.title, []),
        isNew: !prevIds.has(t.id)
      }))
      .sort((a, b) => b.relevanceScore - a.relevanceScore);

    rfpCache = {
      tenders: scored,
      lastFetch: new Date().toISOString(),
      newCount: scored.filter(t => t.isNew).length,
      scanning: false
    };

    console.log(`Discovery complete: ${scored.length} event tenders found (${rfpCache.newCount} new)`);
  } catch(e) {
    console.error('Discovery error:', e.message);
    rfpCache.scanning = false;
  }
}

// ── Endpoints ───────────────────────────────────────────────
app.get('/discover-rfps', async (req, res) => {
  // Kick off scan in background, return current cache immediately
  if (!rfpCache.scanning) runDiscovery();

  res.json({
    tenders: rfpCache.tenders,
    lastFetch: rfpCache.lastFetch,
    newCount: rfpCache.newCount,
    total: rfpCache.tenders.length,
    scanning: rfpCache.scanning,
    message: rfpCache.tenders.length === 0 ? 'First scan in progress, check back in 30 seconds' : 'Background refresh started'
  });
});

app.get('/rfp-status', (req, res) => {
  res.json({
    tenders: rfpCache.tenders,
    lastFetch: rfpCache.lastFetch,
    newCount: rfpCache.newCount,
    total: rfpCache.tenders.length,
    scanning: rfpCache.scanning,
    cached: true
  });
});

// Manual full refresh
app.post('/refresh-rfps', async (req, res) => {
  rfpCache.scanning = false; // allow restart
  runDiscovery();
  res.json({ message: 'Full scan started', scanning: true });
});

// ── Auto-scan every 6 hours ──────────────────────────────────
const SIX_HOURS = 6 * 60 * 60 * 1000;
setTimeout(() => {
  runDiscovery();
  setInterval(runDiscovery, SIX_HOURS);
}, 5000); // 5s after startup to let server settle

// ── Start ───────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`XPACT API v3.0 on port ${PORT}`);
});
