const express = require('express');
const cors = require('cors');
const https = require('https');
const puppeteer = require('puppeteer');
const app = express();

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ── Health check ───────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'XPACT Proposal API', version: '2.0', ready: true });
});

// ── Score RFP ──────────────────────────────────────────────
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
        res.json(JSON.parse(clean));
      } catch (e) {
        res.status(500).json({ error: 'Parse error', raw: data.slice(0, 200) });
      }
    });
  });
  apiReq.on('error', err => res.status(500).json({ error: err.message }));
  apiReq.write(requestBody);
  apiReq.end();
});

// ── Generate section ───────────────────────────────────────
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
// RFP DISCOVERY MODULE — Puppeteer-based Etimad scraper
// ============================================================

// ── Persistent browser singleton ──────────────────────────
let _browser = null;

async function getBrowser() {
  if (_browser) {
    try { await _browser.version(); return _browser; } catch(e) { _browser = null; }
  }
  console.log('Launching headless Chrome...');
  _browser = await puppeteer.launch({
    headless: 'new',
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--no-first-run',
      '--no-zygote',
      '--single-process',
      '--disable-gpu',
      '--disable-extensions',
      '--disable-background-networking',
      '--disable-default-apps'
    ]
  });
  console.log('Chrome ready.');
  return _browser;
}

// ── Event keywords ─────────────────────────────────────────
const EVENT_KEYWORDS = [
  'فعاليات','فعالية','حفل','حفلات','معرض','معارض',
  'مؤتمر','مؤتمرات','ملتقى','احتفال','احتفالية',
  'ندوة','منتدى','إكسبو','أمسية','تنظيم فعالية',
  'حفلة','يوم وطني','العيد','برنامج ترفيه','أوبرا',
  'event','events','conference','exhibition','expo',
  'ceremony','gala','seminar','forum','festival','concert'
];

function isEventRelated(text) {
  const t = (text || '').toLowerCase();
  return EVENT_KEYWORDS.some(kw => t.includes(kw.toLowerCase()));
}

function scoreRelevance(title, agency) {
  let score = 40;
  const text = ((title||'') + ' ' + (agency||'')).toLowerCase();
  const hits = EVENT_KEYWORDS.filter(kw => text.includes(kw.toLowerCase()));
  score += Math.min(30, hits.length * 8);
  return Math.min(100, score);
}

// ── Scrape one Etimad page ─────────────────────────────────
async function scrapeEtimadPage(keyword) {
  const browser = await getBrowser();
  const page = await browser.newPage();

  try {
    await page.setUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
      '(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
    );
    await page.setExtraHTTPHeaders({ 'Accept-Language': 'ar-SA,ar;q=0.9,en;q=0.8' });

    const url = keyword
      ? `https://tenders.etimad.sa/Tender/AllTendersForVisitor?PageNumber=1&searchText=${encodeURIComponent(keyword)}`
      : 'https://tenders.etimad.sa/Tender/AllTendersForVisitor?PageNumber=1';

    await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });

    // Wait a moment for any JS rendering
    await new Promise(r => setTimeout(r, 2000));

    const tenders = await page.evaluate(() => {
      const results = [];
      const seen = new Set();

      document.querySelectorAll('a[href*="DetailsForVisitor"]').forEach(a => {
        const href = a.href;
        const title = (a.innerText || '').trim();

        // Skip "التفاصيل" and short/duplicate links
        if (!title || title === 'التفاصيل' || title.length < 8) return;
        if (seen.has(href)) return;
        seen.add(href);

        // Walk up to find the containing card
        let card = a;
        for (let i = 0; i < 8; i++) {
          if (!card.parentElement) break;
          card = card.parentElement;
          if (card.innerText && card.innerText.length > 80) break;
        }
        const rawText = (card.innerText || '').replace(/\s+/g, ' ').trim();

        // Extract publish date
        const dateMatch = rawText.match(/(\d{4}-\d{2}-\d{2})/);

        // Extract agency — usually the last meaningful line
        const lines = rawText.split(/[\n|]/)
          .map(l => l.trim())
          .filter(l => l.length > 5 && l !== title && l !== 'التفاصيل');
        const agency = lines.find(l =>
          l.includes('وزارة') || l.includes('هيئة') || l.includes('أمانة') ||
          l.includes('جامعة') || l.includes('مستشفى') || l.includes('المديرية') ||
          l.includes('الرئاسة') || l.includes('إدارة') || l.includes('مركز') ||
          l.includes('شركة') || l.includes('الشؤون') || l.includes('الديوان')
        ) || lines[lines.length - 1] || '';

        results.push({ title, href, rawText, date: dateMatch ? dateMatch[1] : '', agency });
      });

      return results;
    });

    return tenders;
  } catch (e) {
    console.error('Etimad scrape error:', e.message);
    return [];
  } finally {
    await page.close();
  }
}

// ── Main Etimad fetcher (3 keyword passes) ─────────────────
async function fetchEtimad() {
  try {
    const keywords = ['فعاليات', 'مؤتمرات', 'معارض'];
    const allResults = [];
    const seen = new Set();

    for (const kw of keywords) {
      const results = await scrapeEtimadPage(kw);
      for (const t of results) {
        if (!seen.has(t.href)) {
          seen.add(t.href);
          if (isEventRelated(t.title + ' ' + t.agency)) {
            allResults.push({
              id: 'etim_' + Buffer.from(t.href).toString('base64').slice(0, 16),
              title: t.title,
              agency: t.agency,
              deadline: t.date,
              budget: 0,
              status: 'نشط',
              source: 'Etimad',
              tenderUrl: t.href
            });
          }
        }
      }
    }

    // Also scrape page 1 without keyword to catch anything missed
    const generalResults = await scrapeEtimadPage(null);
    for (const t of generalResults) {
      if (!seen.has(t.href) && isEventRelated(t.title + ' ' + t.agency)) {
        seen.add(t.href);
        allResults.push({
          id: 'etim_' + Buffer.from(t.href).toString('base64').slice(0, 16),
          title: t.title,
          agency: t.agency,
          deadline: t.date,
          budget: 0,
          status: 'نشط',
          source: 'Etimad',
          tenderUrl: t.href
        });
      }
    }

    console.log(`Etimad: found ${allResults.length} event tenders`);
    return allResults;
  } catch (e) {
    console.error('fetchEtimad error:', e.message);
    return [];
  }
}

// ── FURSA (kept as HTTP — adjust if needed) ────────────────
function httpsFetch(hostname, path, extraHeaders) {
  return new Promise((resolve) => {
    const opts = {
      hostname, path, method: 'GET',
      headers: {
        'Accept': 'application/json, */*',
        'User-Agent': 'Mozilla/5.0 (compatible; XpactAI/1.0)',
        'Accept-Language': 'ar,en;q=0.9',
        ...extraHeaders
      },
      timeout: 12000
    };
    const req = https.request(opts, (res) => {
      let raw = '';
      res.on('data', c => raw += c);
      res.on('end', () => resolve({ ok: res.statusCode < 400, data: raw }));
    });
    req.on('error', () => resolve({ ok: false, data: null }));
    req.on('timeout', () => { req.destroy(); resolve({ ok: false, data: null }); });
    req.end();
  });
}

async function fetchFursa() {
  try {
    const r = await httpsFetch('fursa.sa', '/api/v1/rfp?status=open&page=1&limit=50', {});
    if (!r.ok || !r.data) return [];
    const p = JSON.parse(r.data);
    const list = p.data || p.rfps || p.results || (Array.isArray(p) ? p : []);
    if (!Array.isArray(list)) return [];
    return list.map(t => ({
      id: 'furs_' + (t.id || t._id || Math.random().toString(36).slice(2, 10)),
      title: t.title || t.rfpTitle || t.name || 'فرصة',
      agency: t.companyName || t.company || t.organization || '',
      deadline: t.deadline || t.closingDate || t.endDate || '',
      budget: parseFloat(t.budget || t.value || 0) || 0,
      status: t.status || 'مفتوح',
      source: 'FURSA',
      tenderUrl: (t.id || t._id) ? `https://fursa.sa/rfp/${t.id || t._id}` : 'https://fursa.sa'
    }));
  } catch (e) { console.log('FURSA error:', e.message); return []; }
}

// ── In-memory cache ────────────────────────────────────────
let rfpCache = { tenders: [], lastFetch: null, prevIds: [], newCount: 0 };

// ── /discover-rfps ─────────────────────────────────────────
app.get('/discover-rfps', async (req, res) => {
  try {
    const prevIds = new Set(rfpCache.tenders.map(t => t.id));

    const [etimadRaw, fursaRaw] = await Promise.all([fetchEtimad(), fetchFursa()]);
    const all = [...etimadRaw, ...fursaRaw];

    const scored = all
      .map(t => ({
        ...t,
        relevanceScore: scoreRelevance(t.title, t.agency),
        isNew: !prevIds.has(t.id)
      }))
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
      sources: {
        etimad: etimadRaw.length,
        fursa: fursaRaw.length
      }
    });
  } catch (e) {
    res.status(500).json({ error: e.message, tenders: [], total: 0 });
  }
});

// ── /rfp-status (cached, instant) ─────────────────────────
app.get('/rfp-status', (req, res) => {
  res.json({
    tenders: rfpCache.tenders,
    lastFetch: rfpCache.lastFetch,
    newCount: rfpCache.newCount,
    total: rfpCache.tenders.length,
    cached: true
  });
});

// ── Start ──────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log('XPACT API on port ' + PORT);
  // Warm up the browser on startup
  getBrowser().catch(e => console.error('Browser warmup failed:', e.message));
});
