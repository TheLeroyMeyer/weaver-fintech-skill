/**
 * setup.js — Weaver Fintech Asset Setup
 * ───────────────────────────────────────
 * Run this ONCE per Claude session before generating any deck:
 *   node setup.js
 *
 * This writes the two SVG logos to /home/claude/assets/ from
 * base64 strings embedded in SKILL.md. Photos are loaded from
 * GitHub raw URLs and need no local setup.
 *
 * ⚠️  The base64 strings below are auto-generated from the real
 *     Weaver brand assets. Do not edit them manually.
 */

const fs   = require('fs');
const path = require('path');

// Read base64 strings from SKILL.md at runtime
// (Claude injects these when generating a deck — see SKILL.md Setup block)
const ASSETS_DIR = '/home/claude/assets';
fs.mkdirSync(ASSETS_DIR, { recursive: true });

// These values are populated by Claude from SKILL.md when running
// If running manually, paste the base64 strings from SKILL.md here:
const LOGO_FULL_B64 = process.env.LOGO_FULL_B64 || '';
const LOGO_MARK_B64 = process.env.LOGO_MARK_B64 || '';

if (!LOGO_FULL_B64 || !LOGO_MARK_B64) {
  console.error('❌ Logo base64 strings not found.');
  console.error('   Run this via Claude (it will inject the strings from SKILL.md),');
  console.error('   or set LOGO_FULL_B64 and LOGO_MARK_B64 environment variables.');
  process.exit(1);
}

fs.writeFileSync(path.join(ASSETS_DIR, 'image1.svg'), Buffer.from(LOGO_FULL_B64, 'base64'));
fs.writeFileSync(path.join(ASSETS_DIR, 'image2.svg'), Buffer.from(LOGO_MARK_B64, 'base64'));

console.log('✅ Logos written to', ASSETS_DIR);
console.log('   image1.svg —', fs.statSync(path.join(ASSETS_DIR, 'image1.svg')).size, 'bytes');
console.log('   image2.svg —', fs.statSync(path.join(ASSETS_DIR, 'image2.svg')).size, 'bytes');
console.log('');
console.log('Ready. Now run: node generate-pitch.js');
