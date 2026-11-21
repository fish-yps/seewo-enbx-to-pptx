# seewo-enbx-to-pptx

把**希沃白板（EasiNote / EN5）课件分享链接**里的课件下载下来，转成微软 PowerPoint 的 `.pptx`，不需要安装希沃白板客户端。

希沃的课件（`.enbx`）本质是一个 ZIP：`Slide_N.xml` 记录每页的文字/位置/字号/颜色，`Resources/` 放图片，`Board.xml` 记录页序。本项目直接从希沃服务器取这些原始文件，再按 OOXML 规范重新打包成 pptx，所以文字是**可编辑的真实文本框**，不是截图。

## 环境要求

- Node.js 18+（推荐 20/24）
- 网络能访问 `easinote.seewo.com` 与 `enow.seewo.com`
- `scripts/render.js` 需要 `canvas`（已在 devDependencies，装有中文字体时才能正确渲染）；其余功能不依赖它

```bash
npm install
npm test          # 冒烟测试（用内置最小样本，不联网）
```

## 用法

### 1. 下载课件原始文件

```bash
node scripts/fetch.js "<希沃分享链接>" ./enbx

# 需要走代理时
HTTPS_PROXY=http://127.0.0.1:7890 node scripts/fetch.js "<分享链接>" ./enbx
```

分享链接形如 `https://xxx.r302.cc/xxxxxx?platform=enmobile&channel=copylink` 或 `https://easinote.seewo.com/m/linkShare?id=<shareLinkUid>`。

下载完成后目录结构：

```
enbx/
├── Board.xml            # 页序
├── Document.xml         # 课件名、作者、版本
├── Reference.xml        # 资源 id -> 文件
├── Slides/Slide_0.xml …
├── Resources/…          # 背景图、插图
└── manifest.json        # 本次抓取的元信息
```

### 2. 转成 PPTX

```bash
node src/enbx2pptx.js ./enbx 课件.pptx --title "2416模拟三" --author "席凤英"
```

不传输出路径时，默认输出到 `enbx` 同级目录、以目录名命名。

### 3. 自检（可选）

```bash
node scripts/verify.js 课件.pptx          # 关系/内容类型/坐标结构校验
node scripts/render.js 课件.pptx ./png    # 逐页渲染成 PNG 方便目检（需要 canvas 与中文字体）
```

## 转换规则

| 希沃元素 | PPTX 对应 |
| --- | --- |
| 画布 1280×720 | 13.333in×7.5in（16:9），1px = 9525 EMU |
| `<FontSize>40`（px） | 字号按 `pt = px × 18/26` 换算：40px → 27.69pt（希沃的「28号」）；直接当 pt 会大出 44% |
| `<Background><ImageBrush>` | 页面背景图片（Stretch=Fill 时按填充裁切） |
| `<Text>` + `<RichText>` | 文本框；`<TextLine>` 逐行转段落，`<TextRun>` 逐段保留字号/加粗/斜体/颜色/字体/底色 |
| `<Picture>` + `<DisplayRegion>` | 图片，按 `Rectangle` 换算成 `srcRect` 裁剪 |
| `<Shape>` | 矩形/椭圆等预设形状，带填充与描边 |
| `#AARRGGBB` | OOXML 颜色（保留透明度） |
| 文本框绝对坐标 | `a:off`/`a:ext` 绝对定位，不依赖占位符 |

解析用 `@xmldom/xmldom` 走真正的 XML DOM，而不是正则切分——希沃导出的 XML 是 CRLF 换行，正则很容易把 `</Text>` 和后面的 `<Text>` 误吞成一段文本，导致整页内容重复。

## 支持范围与限制

- 纯文字、图片、基础形状的课件可以完整还原。
- 希沃的**动画、触发交互、课堂活动（选择题/分类等）、思维导图、音视频**不会转换，只会保留静态的第一帧内容。
- 字体依赖本机安装（宋体、微软雅黑、楷体等）；没有对应字体时会回退。
- 一个 `<TextLine>` 是一整段文字（希沃不预折行），长段落交给 PowerPoint 自动换行。
- 部分元素在希沃里靠看不见的空 `TextRun` 撑开间距，转成 pptx 后行距可能与原版略有出入。
- 本项目只处理**公开分享链接**，不涉及登录账号或私有课件。

## 用到的希沃接口

```
GET  /com/apis?api=SHARE_DETAIL&shareLinkUid=<id>                     # 课件 id / 版本 / 作者
GET  /enow/open/api/v1/courseware/detail?shareLinkUid=<id>&cid=<cwId>
     &version=<v>&turnMp4=true&w=1200&h=800                           # 文件清单 + 签名下载链接
```

## 许可证

MIT
