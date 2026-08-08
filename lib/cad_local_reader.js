'use strict';
/**
 * Run local CAD-zoom OCR (python) — zero Claude tokens.
 * Input: absolute path to PDF/PNG, or base64 PDF/PNG with mime.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const { scriptsPath } = require('../paths');

function pythonBin() {
  return process.env.PMC_PYTHON || (process.platform === 'win32' ? 'py' : 'python3');
}

function runCadZoomOcrOnFile(filePath, outDir) {
  const script = scriptsPath('cad_zoom_ocr.py');
  const args = process.platform === 'win32'
    ? ['-3', script, filePath, outDir || '']
    : [script, filePath, outDir || ''];
  const bin = pythonBin();
  const cmdArgs = process.platform === 'win32' ? args : args;
  const exe = process.platform === 'win32' ? 'py' : bin;
  const finalArgs = process.platform === 'win32' ? ['-3', script, filePath, ...(outDir ? [outDir] : [])] : [script, filePath, ...(outDir ? [outDir] : [])];
  try {
    const out = execFileSync(exe, finalArgs, {
      timeout: 300000,
      maxBuffer: 80 * 1024 * 1024,
      windowsHide: true,
      env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' },
    });
    return JSON.parse(out.toString('utf8'));
  } catch (e) {
    const stderr = e.stderr?.toString?.() || e.message;
    // try parse stdout anyway
    try {
      const maybe = e.stdout?.toString?.('utf8');
      if (maybe && maybe.trim().startsWith('{')) return JSON.parse(maybe);
    } catch (_) {}
    return { success: false, error: stderr };
  }
}

function runCadZoomOcrFromBase64(b64, mime = 'application/pdf') {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pmc_zoom_'));
  const ext = mime.includes('png') ? '.png' : mime.includes('jpeg') || mime.includes('jpg') ? '.jpg' : '.pdf';
  const filePath = path.join(tmp, `input${ext}`);
  fs.writeFileSync(filePath, Buffer.from(b64, 'base64'));
  const outDir = path.join(tmp, 'crops');
  fs.mkdirSync(outDir, { recursive: true });
  const result = runCadZoomOcrOnFile(filePath, outDir);
  result._tmp = tmp;
  return result;
}

module.exports = {
  runCadZoomOcrOnFile,
  runCadZoomOcrFromBase64,
  pythonBin,
};
