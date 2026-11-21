// 结构校验：关系、内容类型、形状合法性
// 用法: node scripts/verify.js <pptx>
const JSZip = require('jszip');
const fs = require('fs');
function required(what) {
  const p = process.argv[2];
  if (!p) {
    console.error(`用法: node scripts/verify.js <${what}>`);
    process.exit(1);
  }
  return p;
}
(async () => {
  const z = await JSZip.loadAsync(fs.readFileSync(required('pptx')));
  const get = (p) => z.file(p) ? z.file(p).async('string') : Promise.resolve(null);
  const ct = await get('[Content_Types].xml');
  const overrides = [...ct.matchAll(/PartName="([^"]+)"/g)].map((m) => m[1]);
  const defaults = [...ct.matchAll(/Extension="([^"]+)"/g)].map((m) => m[1]);
  let errors = [];
  for (const o of overrides) if (!z.file(o.slice(1))) errors.push('missing part for override: ' + o);
  const pres = await get('ppt/presentation.xml');
  const presRels = await get('ppt/_rels/presentation.xml.rels');
  const presRelIds = new Set([...presRels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
  for (const m of pres.matchAll(/r:id="([^"]+)"/g)) if (!presRelIds.has(m[1])) errors.push('presentation.xml dangling ' + m[1]);
  const slidePaths = Object.keys(z.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  for (const sp of slidePaths) {
    const relPath = sp.replace('slides/', 'slides/_rels/') + '.rels';
    const rels = (await get(relPath)) || '';
    const ids = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
    const xml = await get(sp);
    for (const m of xml.matchAll(/r:(?:id|embed)="([^"]+)"/g)) if (!ids.has(m[1])) errors.push(sp + ' dangling ' + m[1]);
    for (const m of rels.matchAll(/Target="([^"]+)"/g)) {
      const t = m[1].replace('../', 'ppt/');
      if (!z.file(t)) errors.push(sp + ' missing target ' + t);
    }
    // 每个 shape 必须有 xfrm 且有 ext
    const sps = [...xml.matchAll(/<p:(sp|pic)>[\s\S]*?<\/p:\1>/g)].map((m) => m[0]);
    for (const s of sps) if (!/<a:xfrm/.test(s)) errors.push(sp + ' shape without xfrm');
    // id 唯一
    const idsInSlide = [...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]);
    if (new Set(idsInSlide).size !== idsInSlide.length) errors.push(sp + ' duplicate cNvPr ids');
    // 所有 rPr 的 sz
    for (const m of xml.matchAll(/sz="(\d+)"/g)) if (+m[1] < 100 || +m[1] > 400000) errors.push(sp + ' bad sz ' + m[1]);
  }
  console.log('slides:', slidePaths.length, 'overrides:', overrides.length, 'exts:', defaults.join(','));
  console.log(errors.length ? 'ERRORS:\n' + errors.slice(0, 30).join('\n') : 'OK: no structural errors');
})();
