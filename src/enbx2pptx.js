#!/usr/bin/env node
// 希沃 ENBX 课件 -> PPTX 转换器
// 读取从希沃云拉取的课件目录（Slides/*.xml + Resources/*），生成 .pptx
//
// 用法:
//   node src/enbx2pptx.js <enbx 目录> [输出.pptx] [--title 标题] [--author 作者]
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { DOMParser } = require('@xmldom/xmldom');

function parseArgs(argv) {
  const opts = { input: null, output: null, title: '', author: '' };
  const rest = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--title') opts.title = argv[++i];
    else if (a === '--author') opts.author = argv[++i];
    else if (a === '-h' || a === '--help') opts.help = true;
    else rest.push(a);
  }
  opts.input = rest[0] || path.join(__dirname, '..', 'enbx');
  opts.output = rest[1] || path.join(path.dirname(opts.input), `${path.basename(opts.input)}.pptx`);
  return opts;
}
// ---------------- 通用 ----------------
const PX = 9525; // 1 px = 9525 EMU (96 dpi)
const SLIDE_W = 1280;
const SLIDE_H = 720;

function pcf(x) { return x.replace(/></g, '>\n<'); }

function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function kids(node, name) {
  const out = [];
  if (!node) return out;
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (!name || c.nodeName === name)) out.push(c);
  }
  return out;
}
function kid(node, name) {
  return kids(node, name)[0] || null;
}
function txt(node, name, dflt = '') {
  const k = kid(node, name);
  if (!k) return dflt;
  return (k.textContent || '').trim();
}
function num(node, name, dflt = 0) {
  const v = parseFloat(txt(node, name, ''));
  return Number.isFinite(v) ? v : dflt;
}
function parseFile(file) {
  const src = fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '');
  const doc = new DOMParser({ onError: () => {} }).parseFromString(src, 'text/xml');
  return doc.documentElement;
}

// ---------------- 字号 ----------------
// 希沃的 <FontSize> 是画布像素(px)，PowerPoint 的 sz 是磅(pt)。
// 希沃字号表：12号=17.33px、16号=22.67px、28号=40px、40号=58.67px，即 pt = px * 18 / 26。
const PX_TO_PT = 18 / 26;
function toPt(px) {
  const pt = px * PX_TO_PT;
  return Math.max(100, Math.min(400000, Math.round(pt * 100)));
}

// ---------------- 颜色 ----------------
function parseColor(s) {
  if (!s) return null;
  const t = s.trim().replace(/^#/, '');
  if (/^[0-9a-fA-F]{8}$/.test(t)) {
    return { hex: t.slice(2).toUpperCase(), alpha: parseInt(t.slice(0, 2), 16) / 255 };
  }
  if (/^[0-9a-fA-F]{6}$/.test(t)) return { hex: t.toUpperCase(), alpha: 1 };
  return null;
}
function colorXml(c, mult = 1) {
  if (!c) return '<a:srgbClr val="000000"/>';
  const a = Math.max(0, Math.min(1, (c.alpha === undefined ? 1 : c.alpha) * mult));
  if (a >= 0.999) return `<a:srgbClr val="${c.hex}"/>`;
  return `<a:srgbClr val="${c.hex}"><a:alpha val="${Math.round(a * 100000)}"/></a:srgbClr>`;
}
// 从带有 <ColorBrush> 子树里取颜色
function brushColor(node) {
  if (!node) return null;
  const cb = kid(node, 'ColorBrush');
  return cb ? parseColor(cb.textContent) : null;
}

// ---------------- 元素级解析 ----------------
// 把 slide 里的元素统一成中间结构（供 pptx 生成与自检渲染共用）
function readSlide(file) {
  const root = parseFile(file);
  const out = { width: num(root, 'Width', SLIDE_W), height: num(root, 'Height', SLIDE_H), background: null, elements: [] };

  const bg = kid(root, 'Background');
  if (bg) {
    const brush = kid(bg, 'ImageBrush');
    const cb = kid(bg, 'ColorBrush');
    if (brush) {
      out.background = { kind: 'image', src: txt(brush, 'Source').replace('id://', ''), stretch: txt(brush, 'Stretch', 'Fill') };
    } else if (cb) {
      out.background = { kind: 'color', color: parseColor(cb.textContent) };
    }
  }

  const elementsNode = kid(root, 'Elements');
  for (const el of kids(elementsNode)) {
    const common = {
      id: txt(el, 'Id'),
      x: num(el, 'X'),
      y: num(el, 'Y'),
      w: num(el, 'Width'),
      h: num(el, 'Height'),
      rot: num(el, 'Rotation'),
      tag: el.nodeName,
    };
    if (el.nodeName === 'Text' || el.nodeName === 'Shape') {
      const rich = kid(el, 'RichText');
      common.text = rich ? readRichText(rich) : null;
    }
    if (el.nodeName === 'Picture') {
      common.src = txt(el, 'Source').replace('id://', '');
      common.name = txt(el, 'PictureName');
      const region = kid(el, 'DisplayRegion');
      const rect = txt(region, 'Rectangle', '');
      const meta = kid(el, 'MetaData');
      const psize = txt(meta, 'PictureSize', '');
      common.rect = rect ? rect.split(',').map(Number) : null;
      common.pictureSize = psize ? psize.split(',').map(Number) : null;
      common.opacity = num(el, 'Alpha', 1);
    }
    if (el.nodeName === 'Shape') {
      const geo = kid(el, 'Geometry');
      const preset = kid(geo, 'PresetGeometry');
      common.geometry = preset ? txt(preset, 'GeometryType', 'Rectangle') : 'Rectangle';
      common.fill = brushColor(kid(el, 'Background'));
      common.opacity = num(el, 'Opacity', 1);
      common.line = brushColor(kid(el, 'Foreground'));
      common.lineType = txt(el, 'LineType', 'None');
      common.thickness = num(el, 'Thickness', 0);
    }
    out.elements.push(common);
  }
  return out;
}

// RichText -> { lines:[{align, spacing, runs:[...]}], hasLines:bool, plain:string }
function readRichText(rich) {
  const linesNode = kid(rich, 'TextLines');
  if (linesNode) {
    return { hasLines: true, lines: kids(linesNode, 'TextLine').map(readLine) };
  }
  // 无排版框的纯文本（Shape / 特殊元素）
  const t = kid(rich, 'Text');
  return { hasLines: false, plain: t ? t.textContent || '' : '' };
}

function readLine(line) {
  const runsNode = kid(line, 'TextRuns');
  const runs = kids(runsNode, 'TextRun').map((r) => {
    const fam = kid(r, 'FontFamily');
    const t = kid(r, 'Text');
    return {
      text: t ? t.textContent : '',
      size: num(r, 'FontSize', 18),
      bold: txt(r, 'FontWeight', 'Normal') === 'Bold',
      italic: txt(r, 'FontStyle', 'Normal') === 'Italic',
      underline: /Underline/i.test(txt(r, 'Decoration', '')),
      family: fam ? txt(fam, 'Source', '') : '',
      color: brushColor(kid(r, 'Foreground')) || { hex: '000000', alpha: 1 },
      bgColor: brushColor(kid(r, 'Background')),
    };
  });
  return {
    align: txt(line, 'TextAlignment', 'Left'),
    spacing: parseFloat(txt(line, 'LineSpacing', '1')) || 1,
    spaceBefore: num(line, 'SpaceBefore', 0),
    spaceAfter: num(line, 'SpaceAfter', 0),
    marginLeft: num(line, 'MarginLeft', 0),
    indent: num(line, 'Indent', 0),
    indentLevel: parseInt(txt(line, 'IndentLevel', '0'), 10) || 0,
    marker: txt(line, 'TextMarker', 'None'),
    direction: txt(line, 'Direction', 'LeftToRight'),
    runs,
  };
}

// ---------------- pptx XML ----------------
// 希沃里空 TextRun 常带 Times New Roman，但真正渲染的是中文字体；按要求把西文字体归一到常见中文字体
const LATIN_TO_CJK = { 'Times New Roman': '宋体', 'Arial': '微软雅黑', 'Calibri': '微软雅黑' };
function normFamily(f) {
  if (!f) return '';
  return LATIN_TO_CJK[f] || f;
}

function runPr(r) {
  const attrs = [`sz="${toPt(r.size)}"`, r.bold ? 'b="1"' : '', r.italic ? 'i="1"' : '', r.underline ? 'u="sng"' : '']
    .filter(Boolean)
    .join(' ');
  let x = `<a:rPr lang="zh-CN" altLang="en-US" ${attrs} dirty="0">`;
  if (r.bgColor && r.bgColor.alpha > 0.01) x += `<a:highlight>${colorXml(r.bgColor)}</a:highlight>`;
  x += `<a:solidFill>${colorXml(r.color)}</a:solidFill>`;
  const fam = normFamily(r.family);
  x += `<a:latin typeface="${esc(fam || 'Arial')}"/><a:ea typeface="${esc(fam)}"/><a:cs typeface=""/>`;
  x += '</a:rPr>';
  return x;
}

function algn(a) {
  switch ((a || '').toLowerCase()) {
    case 'center':
      return 'ctr';
    case 'right':
      return 'r';
    case 'justify':
      return 'just';
    default:
      return 'l';
  }
}

function pPrXml(line, inBox) {
  const lvl = Math.min(line.indentLevel, 8);
  const marL = Math.round(line.marginLeft * PX);
  const indent = Math.round((line.indent || 0) * PX * -1);
  let pr = `<a:pPr marL="${marL}" indent="${indent}" lvl="${lvl}" algn="${algn(line.align)}">`;
  pr += `<a:spcBef><a:spcPts val="${Math.round(line.spaceBefore * 100)}"/></a:spcBef>`;
  pr += `<a:spcAft><a:spcPts val="${Math.round(line.spaceAfter * 100)}"/></a:spcAft>`;
  if (line.spacing && Math.abs(line.spacing - 1) > 0.01) pr += `<a:lnSpc><a:spcPct val="${Math.round(line.spacing * 100000)}"/></a:lnSpc>`;
  if (line.marker === 'Number') pr += '<a:buFont typeface="Arial"/><a:buAutoNum type="arabicPeriod"/>';
  else if (line.marker === 'Bullet') pr += '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>';
  else pr += '<a:buNone/>';
  pr += '</a:pPr>';
  return pr;
}

// 一行里的 runs -> 段落数组（按 \n 切分，保留样式）
// 一个 <TextLine> 是"一段文字"：文字过长时希沃不会预先折行，而是靠文本框自动换行。
// 所以这里只把显式换行拆成多次 <a:p>，长段落交给 PowerPoint 自己排版。
function paragraphsFromLine(line) {
  const ps = [];
  let cur = '';
  const flush = () => {
    ps.push(`<a:p>${pPrXml(line, true)}${cur}</a:p>`);
    cur = '';
  };
  for (const r of line.runs) {
    const parts = String(r.text).replace(/\r\n?/g, '\n').split('\n');
    parts.forEach((t, i) => {
      if (i > 0) flush();
      if (t === '') return;
      cur += `<a:r>${runPr(r)}<a:t>${esc(t)}</a:t></a:r>`;
    });
  }
  flush();
  return ps;
}

function textBodyXml(richText, boxEl, shapeRun) {
  const ps = [];
  if (richText.hasLines) {
    for (const line of richText.lines) ps.push(...paragraphsFromLine(line));
  } else {
    // 无排版框的纯文本（Shape 内文字）：按换行拆段
    const style = shapeRun || { text: '', size: 18, bold: false, italic: false, underline: false, family: '', color: { hex: '000000', alpha: 1 } };
    const t = String(richText.plain || '').replace(/\r\n?/g, '\n');
    for (const ln of t.split('\n')) ps.push(`<a:p><a:pPr algn="l"/><a:r>${runPr(style)}<a:t>${esc(ln)}</a:t></a:r></a:p>`);
  }
  if (!ps.length) ps.push('<a:p><a:endParaRPr lang="zh-CN"/></a:p>');
  return (
    `<p:txBody><a:bodyPr wrap="square" lIns="45720" tIns="45720" rIns="45720" bIns="45720" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>` +
    ps.join('') +
    '</p:txBody>'
  );
}

function xfrm(el, w, h) {
  const rotAttr = el.rot && Math.abs(el.rot) > 0.01 ? ` rot="${Math.round(el.rot * 60000)}"` : '';
  return `<a:xfrm${rotAttr}><a:off x="${Math.round(el.x * PX)}" y="${Math.round(el.y * PX)}"/><a:ext cx="${Math.max(1, Math.round(w * PX))}" cy="${Math.max(1, Math.round(h * PX))}"/></a:xfrm>`;
}

function srcRect(pic) {
  if (!pic.rect || !pic.pictureSize || pic.pictureSize.length < 2) return '';
  const [x, y, w, h] = pic.rect;
  const [W, H] = pic.pictureSize;
  if (!W || !H) return '';
  const l = Math.max(0, x) / W;
  const t = Math.max(0, y) / H;
  const r = Math.max(0, W - (x + w)) / W;
  const b = Math.max(0, H - (y + h)) / H;
  if (l + t + r + b < 0.001) return '';
  return `<a:srcRect l="${Math.round(l * 1e5)}" t="${Math.round(t * 1e5)}" r="${Math.round(r * 1e5)}" b="${Math.round(b * 1e5)}"/>`;
}

let idSeq = 1;
const nid = () => ++idSeq;

// ---------------- 主题 ----------------
function themeXml() {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>` +
    `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>` +
    `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>` +
    `</a:themeElements></a:theme>`
  );
}

// ---------------- 主流程 ----------------
async function main(opts) {
  const ROOT = opts.input;
  const OUT = opts.output;
  const slidesDir = path.join(ROOT, 'Slides');
  const slideFiles = fs
    .readdirSync(slidesDir)
    .filter((f) => /^Slide_\d+\.xml$/.test(f))
    .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));

  const resDir = path.join(ROOT, 'Resources');
  const resByKey = new Map();
  if (fs.existsSync(resDir)) {
    for (const f of fs.readdirSync(resDir)) resByKey.set(f.replace(/\.[^.]+$/, ''), f);
  }

  const media = [];
  const mediaKey = new Map();
  function addMedia(key) {
    if (mediaKey.has(key)) return mediaKey.get(key);
    const fname = resByKey.get(key);
    if (!fname) return null;
    const buf = fs.readFileSync(path.join(resDir, fname));
    let ext = path.extname(fname).slice(1).toLowerCase();
    if (!ext) ext = sniff(buf);
    const entry = { zip: `image${media.length + 1}.${ext}`, buf, ext };
    media.push(entry);
    mediaKey.set(key, entry);
    return entry;
  }

  const slides = [];
  for (const sf of slideFiles) {
    const data = readSlide(path.join(slidesDir, sf));
    idSeq = 1;
    const rels = [];
    let bgXml = '';
    if (data.background) {
      if (data.background.kind === 'image') {
        const m = addMedia(data.background.src);
        if (m) {
          const rid = `rId${rels.length + 1}`;
          rels.push({ rid, target: `../media/${m.zip}` });
          bgXml = `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
        }
      } else if (data.background.color && data.background.color.alpha > 0.01) {
        bgXml = `<p:bg><p:bgPr><a:solidFill>${colorXml(data.background.color)}</a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
      }
    }

    let body = '';
    for (const el of data.elements) {
      if (el.tag === 'Text') {
        if (!el.text) continue;
        body +=
          `<p:sp><p:nvSpPr><p:cNvPr id="${nid()}" name="TextBox ${nid()}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>${xfrm(el, el.w, el.h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
          textBodyXml(el.text, el) +
          '</p:sp>';
      } else if (el.tag === 'Picture') {
        const m = addMedia(el.src);
        if (!m) continue;
        const rid = `rId${rels.length + 1}`;
        rels.push({ rid, target: `../media/${m.zip}` });
        body +=
          `<p:pic><p:nvPicPr><p:cNvPr id="${nid()}" name="${esc(el.name || 'Picture')}"/><p:cNvPicPr><a:picLocks noChangeAspect="0"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="${rid}"/>${srcRect(el)}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
          `<p:spPr>${xfrm(el, el.w, el.h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      } else if (el.tag === 'Shape') {
        const prst = el.geometry === 'Ellipse' ? 'ellipse' : el.geometry === 'Triangle' ? 'triangle' : 'rect';
        const fill = el.fill && el.fill.alpha > 0.01 ? `<a:solidFill>${colorXml(el.fill, el.opacity)}</a:solidFill>` : '<a:noFill/>';
        const ln =
          el.line && el.line.alpha > 0.01 && el.lineType !== 'None' && el.thickness > 0
            ? `<a:ln w="${Math.round(el.thickness * PX)}"><a:solidFill>${colorXml(el.line)}</a:solidFill></a:ln>`
            : '<a:ln><a:noFill/></a:ln>';
        body +=
          `<p:sp><p:nvSpPr><p:cNvPr id="${nid()}" name="Shape ${nid()}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr>${xfrm(el, el.w, el.h)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fill}${ln}</p:spPr>` +
          (el.text ? textBodyXml(el.text, el) : '') +
          '</p:sp>';
      }
    }

    slides.push({
      xml:
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        `<p:cSld>${bgXml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${body}</p:spTree></p:cSld>` +
        `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      rels,
    });
  }

  const zip = new JSZip();
  const exts = [...new Set(media.map((m) => m.ext))];
  const defaultExts = exts.map((e) => `<Default Extension="${e}" ContentType="${mime(e)}"/>`).join('');
  const n = slides.length;
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${defaultExts}` +
      `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
      `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
      `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
      slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('') +
      `<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>` +
      `<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>` +
      `<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>` +
      `</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`
  );
  zip.file(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(opts.title || '希沃课件')}</dc:title><dc:creator>${esc(opts.author || "")}</dc:creator></cp:coreProperties>`
  );
  zip.file(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Slides>${n}</Slides></Properties>`
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
      slides.map((_, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>` +
      `<Relationship Id="rId${n + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>` +
      `<Relationship Id="rId${n + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
      `<Relationship Id="rId${n + 5}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>` +
      `</Relationships>`
  );
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">` +
      `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')}</p:sldIdLst>` +
      `<p:sldSz cx="${SLIDE_W * PX}" cy="${SLIDE_H * PX}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`
  );
  zip.file('ppt/theme/theme1.xml', themeXml());
  zip.file('ppt/presProps.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`);
  zip.file('ppt/viewProps.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr/><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`);
  zip.file('ppt/tableStyles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`
  );
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`
  );
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`
  );
  slides.forEach((s, i) => {
    zip.file(`ppt/slides/slide${i + 1}.xml`, pcf(s.xml));
    zip.file(
      `ppt/slides/_rels/slide${i + 1}.xml.rels`,
      pcf(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
          s.rels.map((r) => `<Relationship Id="${r.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${r.target}"/>`).join('') +
          `</Relationships>`
      )
    );
  });
  for (const m of media) zip.file(`ppt/media/${m.zip}`, m.buf);

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  fs.writeFileSync(OUT, buf);
  console.log(`已生成: ${OUT}  页数=${n}  图片=${media.length}  大小=${(buf.length / 1024).toFixed(0)} KB`);
}

function sniff(b) {
  if (b[0] === 0xff && b[1] === 0xd8) return 'jpg';
  if (b[0] === 0x89 && b[1] === 0x50) return 'png';
  if (b.slice(0, 4).toString() === 'GIF8') return 'gif';
  if (b.slice(8, 12).toString() === 'WEBP') return 'webp';
  return 'png';
}
function mime(e) {
  return { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' }[e] || 'application/octet-stream';
}

const opts = parseArgs(process.argv.slice(2));
if (opts.help) {
  console.log('用法: node src/enbx2pptx.js <enbx 目录> [输出.pptx] [--title 标题] [--author 作者]');
  process.exit(0);
}
main(opts).catch((e) => {
  console.error(e);
  process.exit(1);
});
