#!/usr/bin/env node
// 希沃白板分享链接 -> 课件原始文件（ENBX 结构目录）
//
// 用法:
//   node scripts/fetch.js <分享链接> [输出目录]
//   例: node scripts/fetch.js "https://xxx.r302.cc/QzapBdX?platform=enmobile&channel=copylink" ./enbx
//
// 原理:
//   1. 短链重定向到 easinote.seewo.com/m/linkShare?id=<shareLinkUid>
//   2. GET  /com/apis?api=SHARE_DETAIL&shareLinkUid=...      取 coursewareId / version
//   3. GET  /enow/open/api/v1/courseware/detail?shareLinkUid=...&cid=...&version=...&turnMp4=true&w=1200&h=800
//      得到课件内全部文件清单（每页 Slide_N.xml 及 Resources 图片）
//   4. 按清单把文件下载到 output/Slides、output/Resources 等
const fs = require('fs');
const path = require('path');

const UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36';
const PROXY = process.env.HTTPS_PROXY || process.env.https_proxy || '';

function usage() {
  console.log('用法: node scripts/fetch.js <分享链接> [输出目录]');
  process.exit(1);
}

// 用 undici 的 ProxyAgent 支持 -x 代理（Node 18+ 内置 fetch）
async function makeFetch() {
  if (!PROXY) return fetch;
  const { ProxyAgent, setGlobalDispatcher } = await import('undici');
  setGlobalDispatcher(new ProxyAgent(PROXY));
  return fetch;
}

async function jget(f, url, headers = {}) {
  const res = await f(url, { headers: { 'user-agent': UA, 'x-auth-refer': 'ENOW', ...headers } });
  const text = await res.text();
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error(`接口返回非 JSON (${res.status}): ${text.slice(0, 200)}`);
  }
  return json;
}

function shareIdFrom(input) {
  const m = /[?&]id=([0-9a-f]{32})/i.exec(input);
  if (m) return m[1];
  if (/^[0-9a-f]{32}$/i.test(input.trim())) return input.trim();
  return null;
}

async function main() {
  const [input, outDirArg] = process.argv.slice(2);
  if (!input) usage();
  const f = await makeFetch();

  // 1. 解析短链
  let url = input.trim();
  let shareId = shareIdFrom(url);
  if (!shareId) {
    const res = await f(url, { headers: { 'user-agent': UA }, redirect: 'follow' });
    shareId = shareIdFrom(res.url || '');
    if (!shareId) throw new Error('无法从链接中解析出 shareLinkUid: ' + res.url);
  }

  // 2. 分享详情
  const detail = await jget(f, `https://easinote.seewo.com/com/apis?api=SHARE_DETAIL&shareLinkUid=${shareId}`);
  if (detail.error_code !== 0) throw new Error(`SHARE_DETAIL 失败: ${detail.message} (code ${detail.error_code})`);
  const cw = detail.data.courseware;
  const name = cw.coursewareName || 'courseware';
  const version = cw.version;
  console.log(`课件: ${name}  coursewareId=${cw.coursewareId}  version=${version}  作者=${(detail.data.shareUser || {}).name || ''}`);

  // 3. 文件清单
  const q =
    `shareLinkUid=${shareId}&cid=${encodeURIComponent(cw.coursewareId)}` +
    `&version=${version}&turnMp4=true&w=1200&h=800`;
  const manifest = await jget(f, `https://enow.seewo.com/enow/open/api/v1/courseware/detail?${q}`);
  if (manifest.error_code && manifest.error_code !== 0) throw new Error(`清单接口失败: ${manifest.message}`);
  const files = manifest.files || [];
  if (!files.length) throw new Error('清单为空，可能课件已失效或无权限');
  console.log(`共 ${files.length} 个文件，开始下载...`);

  // 4. 下载
  const outDir = outDirArg || path.join(process.cwd(), 'enbx');
  let ok = 0;
  const failed = [];
  for (const file of files) {
    const rel = (file.path || file.name || '').replace(/\\/g, '/');
    if (!rel) continue;
    const dest = path.join(outDir, rel);
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    const res = await f(file.url, { headers: { 'user-agent': UA } });
    if (!res.ok) {
      failed.push(rel);
      continue;
    }
    fs.writeFileSync(dest, Buffer.from(await res.arrayBuffer()));
    ok++;
  }
  fs.writeFileSync(
    path.join(outDir, 'manifest.json'),
    JSON.stringify({ shareLinkUid: shareId, coursewareId: cw.coursewareId, name, version, author: (detail.data.shareUser || {}).name || '', files: files.length }, null, 2)
  );
  console.log(`完成: ${ok}/${files.length} 个文件 -> ${outDir}`);
  if (failed.length) {
    console.log('下载失败:', failed.join(', '));
    process.exit(2);
  }
  console.log(`课件名（供转换脚本使用）: ${name}`);
}

main().catch((e) => {
  console.error('错误:', e.message);
  process.exit(1);
});
