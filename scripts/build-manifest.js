#!/usr/bin/env node
// Usage: node scripts/build-manifest.js https://<owner>.github.io/<repo>/
'use strict';
const fs = require('fs');
const path = require('path');

let base = process.argv[2];
if (!base || !/^https:\/\/[^\s"<>]+$/i.test(base)) {
  console.error('Usage: npm run manifest -- https://<owner>.github.io/<repo>/');
  process.exit(1);
}
if (!base.endsWith('/')) base += '/';

const root = path.join(__dirname, '..');
const tpl = fs.readFileSync(path.join(root, 'manifest.template.xml'), 'utf8');
const out = tpl.replace(/^<!-- Template\..*-->\n/m, '').split('{{BASE_URL}}').join(base);
fs.writeFileSync(path.join(root, 'manifest.xml'), out);
console.log('Wrote manifest.xml for ' + base);
