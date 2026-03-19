const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/gemini', async (req, res) => {
  try {
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      return res.status(500).json({ error: 'GEMINI_API_KEY not set in Render environment variables.' });
    }

    const { body } = req.body;
    if (!body) {
      return res.status(400).json({ error: 'No body provided.' });
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${key}`;
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    const data = await response.json();
    return res.status(response.status).json(data);

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  const key = process.env.GEMINI_API_KEY;
  res.json({
    status: 'ok',
    gemini_key_set: !!key,
    key_preview: key ? key.slice(0, 8) + '...' : 'NOT SET'
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  const key = process.env.GEMINI_API_KEY;
  console.log(`\n✅ PMC Civil AI Agent running on port ${PORT}`);
  console.log(`🔑 GEMINI_API_KEY: ${key ? 'SET ✅' : 'NOT SET ❌'}`);
  console.log(`\nPress Ctrl+C to stop.\n`);
});
