const path = require('path');
const BASE = __dirname;

// dataPath('file.json') → '/app/file.json'
function dataPath(filename) {
  return path.join(BASE, filename);
}

// scriptsPath('script.py') → '/app/script.py'
function scriptsPath(filename) {
  return path.join(BASE, filename);
}

module.exports = { dataPath, scriptsPath };
