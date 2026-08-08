'use strict';
/**
 * PMC Civil AI — Claude API Client
 * Drop-in replacement for fetchGeminiWithRetry
 * Supports: text, images, extended thinking (Pass 3 validation)
 */

const CLAUDE_URL = 'https://api.anthropic.com/v1/messages';

/**
 * Call Claude API with automatic retry on overload
 * @param {Object} opts
 * @param {string} opts.system     - System prompt (your rules + DSR rates)
 * @param {Array}  opts.messages   - Message array [{role, content}]
 * @param {string} opts.model      - Model to use (default: claude-sonnet-4-5)
 * @param {number} opts.maxTokens  - Max tokens (default: 8192)
 * @param {boolean} opts.thinking  - Enable extended thinking for Pass 3
 * @param {number} opts.thinkingBudget - Thinking tokens (default: 8000)
 * @returns {string} - Text response from Claude
 */
async function callClaude({
  system,
  messages,
  model = 'claude-sonnet-4-5',
  maxTokens = 8192,
  thinking = false,
  thinkingBudget = 8000,
}) {
  const key = process.env.CLAUDE_API_KEY;
  if (!key) throw new Error('CLAUDE_API_KEY not set in environment');

  const body = {
    model,
    max_tokens: thinking ? Math.max(maxTokens, thinkingBudget + 4000) : maxTokens,
    system,
    messages,
  };

  // Extended thinking — Claude re-checks its own BOQ calculations
  // Use this for Pass 3 (validation) only — costs more tokens
  if (thinking) {
    body.thinking = {
      type: 'enabled',
      budget_tokens: thinkingBudget,
    };
  }

  const maxRetries = 4;
  let lastError;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const r = await fetch(CLAUDE_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': key,
          'anthropic-version': '2023-06-01',
        },
        body: JSON.stringify(body),
      });

      const data = await r.json();

      if (r.ok && data.content) {
        // Extract text from response (skip thinking blocks)
        const text = data.content
          .filter(b => b.type === 'text')
          .map(b => b.text || '')
          .join('\n');
        return text;
      }

      const errMsg = data?.error?.message || JSON.stringify(data);
      const status = r.status;
      lastError = new Error(`Claude API ${status}: ${errMsg}`);

      // Retry on overload / rate limit
      if (status === 429 || status === 529 || status >= 500) {
        const delay = Math.min(2000 * Math.pow(2, attempt), 30000);
        await new Promise(res => setTimeout(res, delay));
        continue;
      }
      throw lastError;
    } catch (e) {
      lastError = e;
      if (attempt < maxRetries) {
        const delay = Math.min(2000 * Math.pow(2, attempt), 30000);
        await new Promise(res => setTimeout(res, delay));
        continue;
      }
    }
  }

  throw lastError || new Error('Claude API failed after retries');
}

module.exports = { callClaude };
