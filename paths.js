'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = __dirname;

function rootPath(...parts) {
  return path.join(ROOT, ...parts);
}

/** data/Rates.json, data/legend.json, learned files, templates, etc. */
function dataPath(...parts) {
  return path.join(ROOT, 'data', ...parts);
}

/** python/*.py scripts */
function scriptsPath(...parts) {
  return path.join(ROOT, 'python', ...parts);
}

/** lib/*.js helpers (rarely needed — prefer require('../lib/...')) */
function libPath(...parts) {
  return path.join(ROOT, 'lib', ...parts);
}

function publicDir() {
  return path.join(ROOT, 'public');
}

/** Prefer Rates.json; fall back to rates.json for older callers. */
function ratesFilePath() {
  const primary = dataPath('Rates.json');
  if (fs.existsSync(primary)) return primary;
  const alt = dataPath('rates.json');
  if (fs.existsSync(alt)) return alt;
  return primary;
}

/** Canonical learned-symbols file (migrates typo ymbols-learned.json if present). */
function symbolsLearnedPath() {
  return dataPath('symbols-learned.json');
}

module.exports = {
  ROOT,
  rootPath,
  dataPath,
  scriptsPath,
  libPath,
  publicDir,
  ratesFilePath,
  symbolsLearnedPath,
};
