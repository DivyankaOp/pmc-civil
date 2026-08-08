'use strict';
/**
 * In-memory multi-sheet project workspace (session).
 * Survives for process lifetime; UI keeps projectId.
 */

const crypto = require('crypto');

const projects = new Map();
const MAX_SHEETS = 20;
const TTL_MS = 6 * 60 * 60 * 1000; // 6h

function prune() {
  const now = Date.now();
  for (const [id, p] of projects) {
    if (now - (p.updatedAt || p.createdAt) > TTL_MS) projects.delete(id);
  }
}

function createProject(name) {
  prune();
  const id = crypto.randomBytes(8).toString('hex');
  const project = {
    id,
    name: name || `Project ${id.slice(0, 6)}`,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    sheets: [],
  };
  projects.set(id, project);
  return project;
}

function getProject(id) {
  prune();
  return projects.get(id) || null;
}

function touch(p) {
  p.updatedAt = Date.now();
}

function addSheet(projectId, sheet) {
  let p = getProject(projectId);
  if (!p) p = createProject();
  if (p.sheets.length >= MAX_SHEETS) {
    throw new Error(`Max ${MAX_SHEETS} sheets per project`);
  }
  const row = {
    id: crypto.randomBytes(4).toString('hex'),
    filename: sheet.filename || 'sheet.pdf',
    addedAt: Date.now(),
    drawing_type: sheet.drawing_type || null,
    schedule_rows: sheet.schedule_rows || 0,
    markdown: sheet.markdown || '',
    extracted: sheet.extracted || null,
    measure: sheet.measure || null,
    question: sheet.question || '',
    file_mb: sheet.file_mb || null,
    combined_text: (sheet.combined_text || '').slice(0, 100000),
  };
  p.sheets.push(row);
  touch(p);
  return { project: summarize(p), sheet: row };
}

function summarize(p) {
  return {
    id: p.id,
    name: p.name,
    sheet_count: p.sheets.length,
    updatedAt: p.updatedAt,
    sheets: p.sheets.map(s => ({
      id: s.id,
      filename: s.filename,
      drawing_type: s.drawing_type,
      schedule_rows: s.schedule_rows,
      file_mb: s.file_mb,
      addedAt: s.addedAt,
    })),
  };
}

function mergeProjectMarkdown(projectId) {
  const p = getProject(projectId);
  if (!p) return null;
  const parts = [`# Project takeoff — ${p.name}`, `Sheets: **${p.sheets.length}**`, ''];
  for (let i = 0; i < p.sheets.length; i++) {
    const s = p.sheets[i];
    parts.push(`---`);
    parts.push(`## Sheet ${i + 1}: ${s.filename}`);
    parts.push(`Type: **${s.drawing_type || '—'}** · Schedule rows: **${s.schedule_rows || 0}**`);
    parts.push('');
    parts.push(s.markdown || '_No takeoff yet_');
    parts.push('');
  }
  return {
    markdown: parts.join('\n'),
    project: summarize(p),
  };
}

function removeSheet(projectId, sheetId) {
  const p = getProject(projectId);
  if (!p) return null;
  p.sheets = p.sheets.filter(s => s.id !== sheetId);
  touch(p);
  return summarize(p);
}

module.exports = {
  createProject,
  getProject,
  addSheet,
  summarize,
  mergeProjectMarkdown,
  removeSheet,
};
