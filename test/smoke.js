// 冒烟测试：用最小课件样本跑一遍转换，校验产出是合法 pptx 且内容正确
// 用法: node test/smoke.js
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const JSZip = require('jszip');

const root = path.join(__dirname, '..');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'enbx-smoke-'));

// ---- 构造最小 ENBX 结构：两页，第一页纯色背景 + 文本框，第二页图片 ----
const png1x1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64'
);
fs.mkdirSync(path.join(tmp, 'Slides'), { recursive: true });
fs.mkdirSync(path.join(tmp, 'Resources'), { recursive: true });
fs.writeFileSync(path.join(tmp, 'Resources', 'abc123.png'), png1x1);

const slide0 = `<?xml version="1.0" encoding="utf-8"?>
<Slide>
  <Id>s0</Id>
  <Width>1280</Width>
  <Height>720</Height>
  <Background><ColorBrush>#FFFFFFFF</ColorBrush></Background>
  <Elements>
    <Text>
      <RichText>
        <TextLines>
          <TextLine>
            <LineSpacing>1</LineSpacing>
            <TextAlignment>Center</TextAlignment>
            <TextRuns>
              <TextRun>
                <Text>第一页标题</Text>
                <FontSize>48</FontSize>
                <FontWeight>Bold</FontWeight>
                <FontFamily><Source>宋体</Source></FontFamily>
                <Foreground><ColorBrush>#FFFF0000</ColorBrush></Foreground>
              </TextRun>
            </TextRuns>
          </TextLine>
        </TextLines>
      </RichText>
      <Id>t1</Id><X>0</X><Y>300</Y><Width>1280</Width><Height>120</Height><Rotation>0</Rotation>
    </Text>
  </Elements>
</Slide>`;

const slide1 = `<?xml version="1.0" encoding="utf-8"?>
<Slide>
  <Id>s1</Id>
  <Width>1280</Width>
  <Height>720</Height>
  <Elements>
    <Picture>
      <Source>id://abc123</Source>
      <PictureName>dot.png</PictureName>
      <Id>p1</Id><X>100</X><Y>100</Y><Width>200</Width><Height>200</Height><Rotation>0</Rotation>
    </Picture>
  </Elements>
</Slide>`;

fs.writeFileSync(path.join(tmp, 'Slides', 'Slide_0.xml'), '\uFEFF' + slide0);
fs.writeFileSync(path.join(tmp, 'Slides', 'Slide_1.xml'), '\uFEFF' + slide1);

const out = path.join(tmp, 'out.pptx');
execFileSync(process.execPath, [path.join(root, 'src/enbx2pptx.js'), tmp, out, '--title', '测试课件'], {
  cwd: root,
  stdio: 'inherit',
});

(async () => {
  const z = await JSZip.loadAsync(fs.readFileSync(out));
  const assert = (cond, msg) => {
    if (!cond) {
      console.error('FAIL:', msg);
      process.exit(1);
    }
  };
  const slidePaths = Object.keys(z.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  assert(slidePaths.length === 2, `应有 2 页，实际 ${slidePaths.length}`);

  const raw1 = await z.file('ppt/slides/slide1.xml').async('string');
  const s1 = raw1.replace(/>\n</g, '><');
  assert(s1.includes('第一页标题'), '第一页应包含标题文字');
  assert(s1.includes('sz="4800"'), '字号 48pt 应写成 4800');
  assert(s1.includes('FF0000'), '红色文字颜色应保留');
  assert(s1.includes('<a:solidFill><a:srgbClr val="FFFFFF"/>'), '白色背景应保留');

  const rels2 = await z.file('ppt/slides/_rels/slide2.xml.rels').async('string');
  assert(/ppt\/media|\.\.\/media/.test(rels2), '第二页应引用图片');
  assert(z.file('ppt/media/image1.png'), '图片应写入 ppt/media');

  const pres = (await z.file('ppt/presentation.xml').async('string'));
  assert(/cx="12192000"/.test(pres), '16:9 尺寸应为 12192000 EMU');
  assert((pres.match(/<p:sldId /g) || []).length === 2, 'presentation.xml 应有两个 sldId');

  const ct = await z.file('[Content_Types].xml').async('string');
  assert(ct.includes('Extension="png"'), 'Content_Types 应声明 png');

  // 结构校验脚本
  execFileSync(process.execPath, [path.join(root, 'scripts/verify.js'), out], { cwd: root, stdio: 'inherit' });

  console.log('OK: 冒烟测试通过');
  fs.rmSync(tmp, { recursive: true, force: true });
})();
