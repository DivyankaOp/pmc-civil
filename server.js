const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// API key from environment variable (set in Render dashboard)
const GEMINI_KEY = process.env.GEMINI_API_KEY;

app.post('/gemini', async (req, res) => {
  const { body } = req.body;
  const key = GEMINI_KEY;
  if (!key) return res.status(500).json({ error: 'API key not configured on server.' });

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${key}`;
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const data = await response.json();
    res.status(response.status).json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅ PMC Civil AI Agent running on port ${PORT}`);
  console.log(`\nPress Ctrl+C to stop.\n`);
});
