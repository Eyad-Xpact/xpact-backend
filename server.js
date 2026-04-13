const express = require('express');
const cors = require('cors');
const https = require('https');
const app = express();

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Health check
app.get('/', (req, res) => {
  res.json({ status: 'XPACT Proposal API', version: '1.0', ready: true });
});

// Score RFP endpoint
app.post('/score-rfp', async (req, res) => {
  const { rfpText } = req.body;
  if (!rfpText) return res.status(400).json({ error: 'rfpText required' });

  const prompt = `You are an expert event management consultant. Analyze this RFP and return ONLY a JSON object with this exact structure:
{
  "score": <number 0-100>,
  "recommendation": "<Pursue|Consider|Pass>",
  "summary": "<2 sentence summary>",
  "criteria": {
    "budgetFit": <0-20>,
    "timeline": <0-20>,
    "eventTypeMatch": <0-20>,
    "scopeComplexity": <0-20>,
    "strategicValue": <0-20>
  },
  "risks": ["<risk1>", "<risk2>", "<risk3>"],
  "opportunities": ["<opp1>", "<opp2>", "<opp3>"]
}

RFP TEXT:
${rfpText.slice(0, 4000)}`;

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'API key not configured' });

  const requestBody = JSON.stringify({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1000,
    messages: [{ role: 'user', content: prompt }]
  });

  const options = {
    hostname: 'api.anthropic.com',
    path: '/v1/messages',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Length': Buffer.byteLength(requestBody)
    }
  };

  const apiReq = https.request(options, (apiRes) => {
    let data = '';
    apiRes.on('data', chunk => data += chunk);
    apiRes.on('end', () => {
      try {
        const parsed = JSON.parse(data);
        const text = parsed.content[0].text;
        const clean = text.replace(/```json|```/g, '').trim();
        const result = JSON.parse(clean);
        res.json(result);
      } catch (e) {
        res.status(500).json({ error: 'Parse error', raw: data.slice(0, 200) });
      }
    });
  });

  apiReq.on('error', err => res.status(500).json({ error: err.message }));
  apiReq.write(requestBody);
  apiReq.end();
});

// Generate proposal sections endpoint
app.post('/generate-section', async (req, res) => {
  const { prompt, section } = req.body;
  if (!prompt) return res.status(400).json({ error: 'prompt required' });

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'API key not configured' });

  const requestBody = JSON.stringify({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1000,
    messages: [{ role: 'user', content: prompt }]
  });

  const options = {
    hostname: 'api.anthropic.com',
    path: '/v1/messages',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Length': Buffer.byteLength(requestBody)
    }
  };

  const apiReq = https.request(options, (apiRes) => {
    let data = '';
    apiRes.on('data', chunk => data += chunk);
    apiRes.on('end', () => {
      try {
        const parsed = JSON.parse(data);
        res.json({ text: parsed.content[0].text, section });
      } catch (e) {
        res.status(500).json({ error: 'Parse error' });
      }
    });
  });

  apiReq.on('error', err => res.status(500).json({ error: err.message }));
  apiReq.write(requestBody);
  apiReq.end();
});

// ============================================================
// RFP DISCOVERY MODULE
// ============================================================

let rfpCache = { tenders: [], lastFetch: null, prevIds: [], newCount: 0 };

const EVENT_KEYWORDS = [
  'فعاليات','فعالية','حفل','حفلات','معرض','معارض',
  'مؤتمر','مؤتمرات','ملتقى','احتفال','احتفالية',
  'ندوة','منتدى','إكسبو','أمسية','تنظيم فعالية',
  'حفلة','يوم وطني','العيد','برنامج ترفيه',
  'event','events','conference','exhibition','expo',
  'ceremony','gala','seminar','forum','celebration',
  'entertainment','festival','concert'
];

function isEventRelated(tender) {
  const text = ((tender.title||'') + ' ' + (tender.agency||'') + ' ' + (tender.category||'')).toLowerCase();
  return EVENT_KEYWORDS.some(kw => text.includes(kw.toLowerCase()));
}

function scoreRelevance(tender) {
  let score = 40;
  const text = ((tender.title||'') + ' ' + (tender.agency||'')).toLowerCase();
  const hits = EVENT_KEYWORDS.filter(kw => text.includes(kw.toLowerCase()));
  score += Math.min(30, hits.length * 8);
  if (tender.budget > 0) score += 10;
  if (tender.budget > 500000) score += 10;
  if (tender.deadline) {
    try {
      const days = Math.floor((new Date(tender.deadline) - Date.now()) / 86400000);
      if (days > 30) score += 10;
      else if (days > 7) score += 5;
      else if (days < 0) score -= 40;
    } catch(e) {}
  }
  return Math.min(100, Math.max(0, score));
}

function httpsFetch(hostname, path, method, postData, extraHeaders) {
  return new Promise((resolve) => {
    const opts = {
      hostname, path, method: method || 'GET',
      headers: {
        'Accept': 'application/json, */*',
        'User-Agent': 'Mozilla/5.0 (compatible; XpactAI/1.0)',
        'Accept-Language': 'ar,en;q=0.9',
        ...extraHeaders
      },
      timeout: 12000
    };
    if (method === 'POST' && postData) {
      opts.headers['Content-Type'] = 'application/json';
      opts.headers['Content-Length'] = Buffer.byteLength(postData);
    }
    const req = https.request(opts, (res) => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => resolve({ ok: res.statusCode < 400, data: raw, status: res.statusCode }));
    });
    req.on('error', () => resolve({ ok: false, data: null, status: 0 }));
    req.on('timeout', () => { req.destroy(); resolve({ ok: false, data: null, status: 0 }); });
    if (method === 'POST' && postData) req.write(postData);
    req.end();
  });
}

async function fetchEtimad() {
  try {
    const r = await httpsFetch(
      'tenders.etimad.sa',
      '/Tender/GetAllActiveTendersTable?PageSize=50&PageIndex=1',
      'GET', null,
      { 'Referer': 'https://tenders.etimad.sa/Tender/AllTendersForVisitor' }
    );
    if (!r.ok || !r.data) return [];
    let p;
    try { p = JSON.parse(r.data); } catch(e) { return []; }
    const list = p.data || p.Data || p.result || p.Result || (Array.isArray(p) ? p : []);
    if (!Array.isArray(list)) return [];
    return list.map(t => ({
      id: 'etim_' + (t.TenderId || t.tenderId || Math.random().toString(36).slice(2,10)),
      title: t.TenderName || t.tenderName || t.TenderTitle || 'بدون عنوان',
      agency: t.AgencyName || t.agencyName || t.EntityName || '',
      category: t.TenderTypeName || t.tenderType || '',
      deadline: t.SubmissionDeadline || t.submissionDeadline || t.LastOfferDate || '',
      budget: parseFloat(t.TenderValue || t.tenderValue || t.BudgetAmount || 0) || 0,
      status: t.TenderStatusName || 'نشط',
      source: 'Etimad',
      tenderUrl: t.TenderId
        ? 'https://tenders.etimad.sa/Tender/Details/' + t.TenderId
        : 'https://tenders.etimad.sa/Tender/AllTendersForVisitor'
    }));
  } catch(e) { console.log('Etimad error:', e.message); return []; }
}

async function fetchFursa() {
  try {
    const r = await httpsFetch('fursa.sa', '/api/v1/rfp?status=open&page=1&limit=50', 'GET', null, {});
    if (!r.ok || !r.data) return [];
    let p;
    try { p = JSON.parse(r.data); } catch(e) { return []; }
    const list = p.data || p.rfps || p.results || (Array.isArray(p) ? p : []);
    if (!Array.isArray(list)) return [];
    return list.map(t => ({
      id: 'furs_' + (t.id || t._id || Math.random().toString(36).slice(2,10)),
      title: t.title || t.rfpTitle || t.name || 'فرصة',
      agency: t.companyName || t.company || t.organization || '',
      category: t.category || t.type || '',
      deadline: t.deadline || t.closingDate || t.endDate || '',
      budget: parseFloat(t.budget || t.value || 0) || 0,
      status: t.status || 'مفتوح',
      source: 'FURSA',
      tenderUrl: (t.id || t._id)
        ? 'https://fursa.sa/rfp/' + (t.id || t._id)
        : 'https://fursa.sa'
    }));
  } catch(e) { console.log('FURSA error:', e.message); return []; }
}

app.get('/discover-rfps', async (req, res) => {
  try {
    const prevIds = new Set(rfpCache.tenders.map(t => t.id));
    const [etimadRaw, fursaRaw] = await Promise.all([fetchEtimad(), fetchFursa()]);
    const filtered = [...etimadRaw, ...fursaRaw].filter(isEventRelated);
    const scored = filtered
      .map(t => ({ ...t, relevanceScore: scoreRelevance(t), isNew: !prevIds.has(t.id) }))
      .sort((a, b) => b.relevanceScore - a.relevanceScore);
    rfpCache = {
      tenders: scored,
      lastFetch: new Date().toISOString(),
      prevIds: [...prevIds],
      newCount: scored.filter(t => t.isNew).length
    };
    res.json({
      tenders: scored,
      lastFetch: rfpCache.lastFetch,
      newCount: rfpCache.newCount,
      total: scored.length,
      sources: { etimad: etimadRaw.length, fursa: fursaRaw.length, afterFilter: filtered.length }
    });
  } catch(e) {
    res.status(500).json({ error: e.message, tenders: [], total: 0 });
  }
});

app.get('/rfp-status', (req, res) => {
  res.json({
    tenders: rfpCache.tenders,
    lastFetch: rfpCache.lastFetch,
    newCount: rfpCache.newCount,
    total: rfpCache.tenders.length,
    cached: true
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log('XPACT API on port ' + PORT));
