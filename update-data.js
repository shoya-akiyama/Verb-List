const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const dir = __dirname;
const xlsxPath = path.join(dir, 'List.xlsx');
const htmlPath = path.join(dir, 'index.html');

// ===== 1. Read Excel =====
const wb = XLSX.readFile(xlsxPath);
console.log('Sheets:', wb.SheetNames.join(', '));

const categories = {};
const headerRe = /^(word|column|英単語|英語|単語|自動詞|他動詞)/i;

for (const name of wb.SheetNames) {
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
  const words = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const word = String(r[2] || '').trim();
    if (!word) continue;
    if (i === 0 && headerRe.test(word)) continue;
    words.push({
      jp_i: String(r[0] || '').trim(),
      jp_t: String(r[1] || '').trim(),
      word: word,
      ex_i: String(r[3] || '').trim(),
      ex_t: String(r[4] || '').trim(),
      past: String(r[5] || '').trim(),
      ipa:  String(r[6] || '').trim()
    });
  }
  if (words.length) {
    categories[name] = words;
    console.log('  ' + name + ': ' + words.length + ' words');
  }
}

const catNames = Object.keys(categories);

// ===== 2. Generate data lines =====
const catLine = '  const CATEGORIES = ' + JSON.stringify(categories) + ';';

const firstCat = categories[catNames[0]];
const embLine = '    const EMBEDDED_ITEMS = ' + JSON.stringify(firstCat) + ';';

// MODE_MAP
const modeMap = {};
catNames.forEach((name, i) => { modeMap[String(i + 1)] = name; });
const modeMapLine = '  const MODE_MAP = ' + JSON.stringify(modeMap) + ';';

// HTML buttons
let buttonsHTML = '';
catNames.forEach((name, i) => {
  const m = i + 1;
  buttonsHTML += '        <button type="button" class="modeBtn" data-mode="' + m + '">\n';
  buttonsHTML += '          <div class="modeBtnTop"><span class="modeBtnName">' + name + '</span><span class="modeBtnCount"></span></div>\n';
  buttonsHTML += '          <div class="modeBtnBar"><div class="modeBtnFill" data-mode="' + m + '"></div></div>\n';
  buttonsHTML += '        </button>\n';
});
// Bookmark button
buttonsHTML += '        <button type="button" class="modeBtn modeBtnBm" data-mode="bm" id="bmModeBtn"><div class="modeBtnTop"><span class="modeBtnName">★ 苦手単語</span><span class="modeBtnCount" id="bmModeCount"></span></div></button>';

// getCatStats mode list
const modeList = catNames.map((_, i) => '"' + (i + 1) + '"').join(',');

// ===== 3. Update index.html =====
let html = fs.readFileSync(htmlPath, 'utf8');
let changed = 0;

// Replace CATEGORIES + EMBEDDED_ITEMS (two consecutive lines)
html = html.replace(
  /^[ \t]*const CATEGORIES = .*$\n^[ \t]*const EMBEDDED_ITEMS = .*$/m,
  catLine + '\n' + embLine
);
changed++;

// Replace MODE_MAP
html = html.replace(
  /^[ \t]*const MODE_MAP = \{.*?\};$/m,
  modeMapLine
);
changed++;

// Replace category buttons in HTML
html = html.replace(
  /(<div class="modeBtns" id="modeBtns">\n)([\s\S]*?)(<\/div>\n\s*<div class="searchSection">)/,
  '$1' + buttonsHTML + '\n      $3'
);
changed++;

// Replace mode list in getCatStats/updateDashboard: for (const m of ["1","2",...])
html = html.replace(
  /for \(const m of \[("[0-9]+"(?:,"[0-9]+")*)\]\)/,
  'for (const m of [' + modeList + '])'
);
changed++;

fs.writeFileSync(htmlPath, html, 'utf8');
console.log('\nindex.html updated! (' + changed + ' sections replaced)');
console.log('Categories: ' + catNames.join(', '));
console.log('\nTotal: ' + Object.values(categories).reduce((s, c) => s + c.length, 0) + ' words');
