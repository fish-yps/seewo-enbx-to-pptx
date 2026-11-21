const { DOMParser } = require('@xmldom/xmldom');
const JSZip = require('jszip');
const fs = require('fs');
const { createCanvas } = require('canvas');
const c = createCanvas(10, 10), ctx = c.getContext('2d');
const F = { '宋体': 'Noto Serif CJK SC', '微软雅黑': 'Noto Sans CJK SC', '楷体': 'Noto Serif CJK SC' };
(async () => {
  const z = await JSZip.loadAsync(fs.readFileSync(process.argv[2]));
  const files = Object.keys(z.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a, b) => parseInt(a.match(/\d+/)) - parseInt(b.match(/\d+/)));
  let n = 0;
  for (const [i, f] of files.entries()) {
    const doc = new DOMParser({ onError: () => {} }).parseFromString(await z.file(f).async('string'), 'text/xml');
    for (const sp of doc.getElementsByTagName('p:sp')) {
      const spPr = sp.getElementsByTagName('p:spPr')[0]; if (!spPr) continue;
      const off = spPr.getElementsByTagName('a:off')[0], ext = spPr.getElementsByTagName('a:ext')[0];
      const y = +off.getAttribute('y') / 9525, h = +ext.getAttribute('cy') / 9525, w = +ext.getAttribute('cx') / 9525;
      if (y + h > 722) { console.log('第' + (i + 1) + '页 框超出页面: y=' + y.toFixed(0) + ' h=' + h.toFixed(0)); n++; }
    }
  }
  console.log('超出页面的文本框数量:', n);
})();
