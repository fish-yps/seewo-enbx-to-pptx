const { DOMParser } = require('@xmldom/xmldom');
const JSZip = require('jszip');
const { createCanvas, loadImage } = require('canvas');
const fs = require('fs');
const path = require('path');
function arg(what) {
  const p = process.argv[2];
  if (!p) {
    console.error(`用法: node scripts/render.js <${what}> [输出目录，默认 ./png]`);
    process.exit(1);
  }
  return p;
}
const OUTDIR = process.argv[3] || 'png';
const EMU = 9525, W = 1280, H = 720;
const FONTMAP = { '宋体': 'Noto Serif CJK SC', '微软雅黑': 'Noto Sans CJK SC', '楷体': 'Noto Serif CJK SC', '黑体': 'Noto Sans CJK SC', 'Times New Roman': 'Liberation Serif', 'Arial': 'Liberation Sans', 'Calibri': 'Liberation Sans' };
function col(el) { const c = el.getElementsByTagName('a:srgbClr')[0]; if (!c) return null; const a = el.getElementsByTagName('a:alpha')[0]; const al = a ? +a.getAttribute('val') / 1e5 : 1; return { hex: '#' + c.getAttribute('val'), a: al }; }
(async () => {
  const z = await JSZip.loadAsync(fs.readFileSync(arg('pptx')));
  const imgs = {};
  for (const f of Object.keys(z.files)) if (f.startsWith('ppt/media/') && !z.files[f].dir) imgs[f] = await loadImage(Buffer.from(await z.file(f).async('nodebuffer')));
  const order = Object.keys(z.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort((a, b) => parseInt(a.match(/\d+/)) - parseInt(b.match(/\d+/)));
  fs.mkdirSync(OUTDIR, { recursive: true });
  for (const [i, f] of order.entries()) {
    const doc = new DOMParser({ onError: () => {} }).parseFromString(await z.file(f).async('string'), 'text/xml').documentElement;
    const rels = await z.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`).async('string');
    const relMap = {};
    for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)) relMap[m[1]] = m[2].replace('../', 'ppt/');
    const c = createCanvas(W, H), ctx = c.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    const bg = doc.getElementsByTagName('p:bg')[0];
    if (bg) {
      const blip = bg.getElementsByTagName('a:blip')[0];
      if (blip) { const img = imgs[relMap[blip.getAttribute('r:embed')]]; const s = Math.max(W / img.width, H / img.height); ctx.drawImage(img, (W - img.width * s) / 2, (H - img.height * s) / 2, img.width * s, img.height * s); }
      else { const sf = bg.getElementsByTagName('a:solidFill')[0]; const cc = col(sf); if (cc) { ctx.fillStyle = cc.hex; ctx.fillRect(0, 0, W, H); } }
    }
    for (const sp of [...doc.getElementsByTagName('p:sp'), ...doc.getElementsByTagName('p:pic')]) {
      const spPr = sp.getElementsByTagName('p:spPr')[0];
      const off = spPr.getElementsByTagName('a:off')[0], ext = spPr.getElementsByTagName('a:ext')[0];
      const x = +off.getAttribute('x') / EMU, y = +off.getAttribute('y') / EMU, w = +ext.getAttribute('cx') / EMU, h = +ext.getAttribute('cy') / EMU;
      if (sp.localName === 'pic') {
        const img = imgs[relMap[sp.getElementsByTagName('a:blip')[0].getAttribute('r:embed')]];
        const sr = sp.getElementsByTagName('a:srcRect')[0];
        if (sr) { const l = +sr.getAttribute('l') / 1e5, t = +sr.getAttribute('t') / 1e5, r = +sr.getAttribute('r') / 1e5, b = +sr.getAttribute('b') / 1e5; ctx.drawImage(img, img.width * l, img.height * t, img.width * (1 - l - r), img.height * (1 - t - b), x, y, w, h); }
        else ctx.drawImage(img, x, y, w, h);
        continue;
      }
      const prst = spPr.getElementsByTagName('a:prstGeom')[0]?.getAttribute('prst') || 'rect';
      if (spPr.getElementsByTagName('a:noFill').length === 0) { const sf = spPr.getElementsByTagName('a:solidFill')[0]; if (sf) { const cc = col(sf); ctx.globalAlpha = cc.a; ctx.fillStyle = cc.hex; ctx.beginPath(); if (prst === 'ellipse') ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, 7); else ctx.rect(x, y, w, h); ctx.fill(); ctx.globalAlpha = 1; } }
      const body = sp.getElementsByTagName('p:txBody')[0];
      if (!body) continue;
      const bp = body.getElementsByTagName('a:bodyPr')[0];
      const lI = +bp.getAttribute('lIns') / EMU, tI = +bp.getAttribute('tIns') / EMU, rI = +bp.getAttribute('rIns') / EMU;
      const innerW = w - lI - rI;
      let cy = y + tI;
      for (const p of body.getElementsByTagName('a:p')) {
        const pPr = p.getElementsByTagName('a:pPr')[0];
        const algn = pPr?.getAttribute('algn') || 'l';
        const runs = [...p.getElementsByTagName('a:r')];
        if (!runs.length) continue;
        // 每个 run 单独测量宽度绘制（避免启发式字宽把文字推出框外）
        const rInfo = runs.map((r) => {
          const rp = r.getElementsByTagName('a:rPr')[0];
          const sz = +rp.getAttribute('sz') / 100;
          const fam = (rp.getElementsByTagName('a:ea')[0]?.getAttribute('typeface')) || '';
          const latin = (rp.getElementsByTagName('a:latin')[0]?.getAttribute('typeface')) || '';
          const cc = col(rp.getElementsByTagName('a:solidFill')[0]) || { hex: '#000000', a: 1 };
          const b = rp.getAttribute('b') === '1', it = rp.getAttribute('i') === '1';
          const t = r.getElementsByTagName('a:t')[0]?.textContent || '';
          return { sz, fam, latin, cc, b, it, t, maxSz: sz };
        });
        if (!rInfo.some((q) => q.t.length)) continue;
        const maxSz = Math.max(...rInfo.map((q) => q.sz));
        const lh = maxSz * 1.2;
        // 拆成字符 tokens，保留样式，便于换行
        const toks = [];
        for (const q of rInfo) {
          ctx.font = `${q.it ? 'italic ' : ''}${q.b ? 'bold ' : ''}${q.sz}px "${FONTMAP[q.fam] || 'Noto Sans CJK SC'}"`;
          for (const ch of q.t) toks.push({ ch, q, w: ctx.measureText(ch).width });
        }
        const lines = [];
        let cur = [], curW = 0;
        for (const tk of toks) {
          if (tk.ch === '\n' || tk.ch === '\r') { lines.push({ cur, curW }); cur = []; curW = 0; continue; }
          if (curW + tk.w > innerW && cur.length) { lines.push({ cur, curW }); cur = []; curW = 0; }
          cur.push(tk); curW += tk.w;
        }
        lines.push({ cur, curW });
        for (const ln of lines) {
          let lx = x + lI;
          if (algn === 'ctr') lx = x + lI + (innerW - ln.curW) / 2;
          else if (algn === 'r') lx = x + lI + innerW - ln.curW;
          cy += lh;
          for (const tk of ln.cur) {
            ctx.font = `${tk.q.it ? 'italic ' : ''}${tk.q.b ? 'bold ' : ''}${tk.q.sz}px "${FONTMAP[tk.q.fam] || 'Noto Sans CJK SC'}"`;
            ctx.globalAlpha = tk.q.cc.a;
            ctx.fillStyle = tk.q.cc.hex;
            ctx.textBaseline = 'alphabetic';
            ctx.fillText(tk.ch, lx, cy - lh * 0.22);
            ctx.globalAlpha = 1;
            lx += tk.w;
          }
        }
      }
    }
    fs.writeFileSync(path.join(OUTDIR, `slide${String(i).padStart(2, '0')}.png`), c.toBuffer('image/png'));
  }
  console.log(`已渲染 ${order.length} 页 -> ${OUTDIR}`);
})();
