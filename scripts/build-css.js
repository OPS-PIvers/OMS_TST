#!/usr/bin/env node
/**
 * Builds the Tailwind CSS for Index.html and inlines it into the page.
 *
 *   npm run build:css     rewrite the block in Index.html
 *   npm run check:css     fail if the block is out of date (CI runs this)
 *
 * Apps Script serves Index.html as a single file, so the CSS lives inside it,
 * between the BEGIN/END markers below. Everything else in the page is scanned
 * for class names; the generated block itself is left out of the scan.
 */

const fs = require('fs');
const path = require('path');
const postcss = require('postcss');
const tailwindcss = require('tailwindcss');

const ROOT = path.join(__dirname, '..');
const INDEX = path.join(ROOT, 'Index.html');
const BEGIN = '/* BEGIN generated Tailwind CSS: do not edit, run `npm run build:css` */';
const END = '/* END generated Tailwind CSS */';

function splitPage(html) {
  const a = html.indexOf(BEGIN);
  const b = html.indexOf(END);
  if (a === -1 || b === -1 || b < a) {
    throw new Error('Index.html is missing the generated Tailwind CSS markers.');
  }
  return { before: html.slice(0, a + BEGIN.length), after: html.slice(b) };
}

async function buildCss(pageWithoutCss) {
  const config = Object.assign({}, require(path.join(ROOT, 'tailwind.config.js')), {
    content: [{ raw: pageWithoutCss, extension: 'html' }]
  });
  const result = await postcss([tailwindcss(config)]).process(
    '@tailwind base;\n@tailwind components;\n@tailwind utilities;\n', { from: undefined });
  // Collapsed onto one line so the block doesn't bury the rest of the page (it
  // is ~45 KB). The CSS has no multi-line strings, so this changes nothing else.
  return result.css.replace(/\n\s*/g, ' ').trim();
}

async function main() {
  const check = process.argv.includes('--check');
  const html = fs.readFileSync(INDEX, 'utf8');
  const { before, after } = splitPage(html);
  const css = await buildCss(before + '\n' + after);
  const next = before + '\n' + css + '\n' + after;

  if (check) {
    if (next !== html) {
      console.error('Index.html has out-of-date Tailwind CSS. Run `npm run build:css` and commit the result.');
      process.exit(1);
    }
    console.log('Tailwind CSS in Index.html is up to date.');
    return;
  }
  fs.writeFileSync(INDEX, next);
  console.log('Wrote ' + css.length + ' bytes of Tailwind CSS into Index.html.');
}

main().catch(err => { console.error(err); process.exit(1); });
