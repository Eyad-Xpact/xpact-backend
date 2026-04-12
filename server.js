'use strict';
const express = require('express');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');
const os      = require('os');
const https   = require('https');

const app = express();
app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '10mb' }));

const IMGS = JSON.parse(fs.readFileSync(path.join(__dirname, 'assets/images.json'), 'utf8'));
const { buildProposal } = require('./proposal_builder.js');

// Health check
app.get('/', (req, res) => {
    res.json({ status: 'XPACT Proposal API', version: '1.0', ready: true });
});

// Translate text to Arabic using Anthropic API
function translateToArabic(text, apiKey) {
    if (!text || text.length < 10) return Promise.resolve(text);
    return new Promise((resolve) => {
          const body = JSON.stringify({
                  model: 'claude-haiku-4-5-20251001',
                  max_tokens: 1000,
                  system: 'Translate the following to formal Arabic (MSA). Output only the translation, no preamble.',
                  messages: [{ role: 'user', content: text }]
          });
          const options = {
                  hostname: 'api.anthropic.com',
                  path: '/v1/messages',
                  method: 'POST',
                  headers: {
                            'Content-Type': 'application/json',
                            'x-api-key': apiKey,
                            'anthropic-version': '2023-06-01',
                            'Content-Length': Buffer.byteLength(body)
                  }
          };
          const req = https.request(options, (res) => {
                  let data = '';
                  res.on('data', chunk => data += chunk);
                  res.on('end', () => {
                            try {
                                        const d = JSON.parse(data);
                                        resolve(d.content && d.content[0] ? d.content[0].text : text);
                            } catch(e) { resolve(text); }
                  });
          });
          req.on('error', () => resolve(text));
          req.write(body);
          req.end();
    });
}

// Check if text is primarily Arabic
function isArabic(text) {
    if (!text) return false;
    const arabicChars = (text.match(/[\u0600-\u06FF]/g) || []).length;
    return arabicChars > text.length * 0.2;
}

// Generate PPTX
app.post('/generate-pptx', async (req, res) => {
    try {
          const data = req.body;
          if (!data || !data.event_name) {
                  return res.status(400).json({ error: 'Missing proposal data' });
          }

      // Detect if this is an Arabic proposal
      const gen = data.generated_sections || {};
          const content = data.content || {};
          const sampleText = gen.executive_summary || gen.understanding || content.objectives_intro || '';
          const proposalIsArabic = isArabic(sampleText);

      // Translate fixed sections if proposal is Arabic and they are in English
      if (proposalIsArabic && data.fixed_sections) {
              const apiKey = process.env.ANTHROPIC_API_KEY;
              if (apiKey) {
                        const fixed = data.fixed_sections;
                        const keys = Object.keys(fixed);
                        for (const key of keys) {
                                    const text = fixed[key];
                                    if (text && text.length > 20 && !isArabic(text)) {
                                                  fixed[key] = await translateToArabic(text, apiKey);
                                    }
                        }
                        data.fixed_sections = fixed;
              }
      }

      const tmpPptx = path.join(os.tmpdir(), 'proposal_' + Date.now() + '.pptx');
          await buildProposal(data, tmpPptx);

      const pptxBuffer = fs.readFileSync(tmpPptx);
          const filename = (data.event_name || 'Proposal').replace(/[^\w\u0600-\u06FF]/g, '_') + '_Proposal.pptx';

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
          res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
          res.send(pptxBuffer);

      try { fs.unlinkSync(tmpPptx); } catch(e) {}
    } catch (err) {
          console.error(err);
          res.status(500).json({ error: err.message });
    }
});


// RFP Qualification & Scoring
app.post('/score-rfp', async (req, res) => {
  try {
    const { rfpText } = req.body;
    if (!rfpText || rfpText.length < 20) {
      return res.status(400).json({ error: 'RFP text required' });
    }
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'API key not configured' });

    const prompt = 'Score this RFP for Xpact (Saudi event management company in Riyadh). Return ONLY valid JSON with no markdown:\n{"score":0-100,"recommendation":"Pursue|Consider|Pass","budget_fit":0-100,"timeline":0-100,"event_type_match":0-100,"scope_complexity":0-100,"strategic_value":0-100,"key_risks":["str","str"],"opportunities":["str","str"],"summary":"2 sentence summary"}\n\nRFP:\n' + rfpText.slice(0, 3000);

    const body = JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 700,
      system: 'Return ONLY valid JSON. No markdown. No explanation.',
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
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const result = await new Promise((resolve, reject) => {
      const request = https.request(options, (response) => {
        let data = '';
        response.on('data', chunk => data += chunk);
        response.on('end', () => {
          try {
            const d = JSON.parse(data);
            const text = d.content && d.content[0] ? d.content[0].text : '';
            const clean = text.replace(/```json/g, '').replace(/```/g, '').trim();
            resolve(JSON.parse(clean));
          } catch(e) { reject(new Error('Parse error: ' + e.message)); }
        });
      });
      request.on('error', reject);
      request.write(body);
      request.end();
    });

    res.json(result);
  } catch (err) {
    console.error('Score RFP error:', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log('XPACT API on port ' + PORT));
