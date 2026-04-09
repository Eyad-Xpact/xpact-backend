'use strict';
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const os = require('os');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Load assets
const IMGS = JSON.parse(fs.readFileSync(path.join(__dirname, 'assets/images.json'), 'utf8'));

// Load buildProposal - inline require from app directory
const { buildProposal } = require('./proposal_builder.js');

// Health check
app.get('/', (req, res) => {
  res.json({ status: 'XPACT Proposal API', version: '1.0', ready: true });
});

// Generate PPTX
app.post('/generate-pptx', async (req, res) => {
  try {
    const data = req.body;
    if (!data || !data.event_name) {
      return res.status(400).json({ error: 'Missing proposal data' });
    }

    const tmpPptx = path.join(os.tmpdir(), 'proposal_' + Date.now() + '.pptx');
    buildProposal(data, tmpPptx);

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('XPACT API on port ' + PORT));
