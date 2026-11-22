# docx-paper (中文文档)

[![Live Demo](https://img.shields.io/badge/在线体验-GitHub%20Pages-2563eb.svg)](https://zkwasm.github.io/docx-paper/demo/)
[![npm version](https://img.shields.io/npm/v/docx-paper.svg)](https://www.npmjs.com/package/docx-paper)
[![npm downloads](https://img.shields.io/npm/dm/docx-paper.svg)](https://www.npmjs.com/package/docx-paper)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](./LICENSE)
[![Compatibility](https://img.shields.io/badge/OpenXML%20Benchmark-100%2F100%20Passed-brightgreen.svg)](#)

**[English](./README.md) | [简体中文](./README.zh-CN.md) | [✨ 在线体验 Demo](https://zkwasm.github.io/docx-paper/demo/)**

**`docx-paper`** 是一个面向 AI 时代的**高保真、纯前端 `.docx` (Office Open XML) 页面渲染与打印引擎**。

在“人们不再手写 Word，只看别人发来的文档或让 AI 修改文档”的今天，传统的纯前端库（如 `mammoth.js` 或原版 `docx-preview`）在面对中文排版、理工科试卷公式和复杂公文时往往出现公式空白、排版错乱或页面粘连。`docx-paper` 针对真实世界的中文与理工科文档进行了深度重构，核心特性包括：

1. **纯前端 WMF / EMF 二进制矢量公式渲染引擎**：
   - 无需后端、无需安装 LibreOffice 或 Microsoft Office，直接在浏览器内存里解析 Word 中的 `<w:object>`（MathType `Equation.DSMT4` 嵌入公式）与 `.wmf` / `.emf` Windows 矢量图元，毫秒级转换为高清矢量 `<svg>`（支持分式线、根号多边形、`Symbol` 希腊字母、大括号拼接以及 `MT Extra` 向量顶标箭头 $\overrightarrow{AB}$）。
2. **标准 OMML 数学公式引擎（`<m:oMath>`）**：
   - 将 Word 原生 OMML 公式词法分析并编译为标准 W3C MathML，利用现代浏览器原生 C++ MathML 引擎零延迟渲染（完整覆盖 `<m:sSubSup>` 上下标、`<m:acc>` 顶标符号、`<m:limUpp>`/`<m:limLow>` 极限、`<m:rad>` 根式、`<m:nary>` 大型积分求和、`<m:m>` 矩阵等）。
3. **深度中文 OpenXML 排版支持**：
   - 支持中文特有的 `w:firstLineChars` / `w:hangingChars` / `w:leftChars` 字符单位缩进；
   - 支持 `w:beforeLines` / `w:afterLines` 行单位段前段后间距，以及 `w:lineRule="atLeast"` / `"auto"` 行距校准；
   - 支持中文分散对齐（`distribute`）、字下着重号（`<w:em>`）、拼音指南（`<w:ruby>`）、字符字宽缩放（`<w:w>`）、带圈/带括号中文自动编号（`① ② ③`、`⑴ ⑵ ⑶`、`⒈ ⒉ ⒊`、`㊀ ㊁ ㊂`）及 `Wingdings`/`Symbol` 勾选框符号映射；
   - 内置 macOS / Windows 中文字体自动回退链（宋体、黑体、楷体、仿宋、等线、华文中宋、方正小标宋）。
4. **中文三线表与复杂表格（`insideH` / `insideV`）**：
   - 独立解析表格外框与内部网格线（`insideH` / `insideV`），完美还原学术论文与报告中的“三线表”，并支持表格内部 `<w:sdt>` 结构化内容控件与 VML 自适应文本框（`<v:textbox>`）。
5. **流式物理分页与原生级打印**：
   - 自动按 A4 / B5 / 横向宽表真实纸张高度切分为独立的白色纸张卡片（带页间阴影间隙与右上角真实页码角标）；
   - **首屏秒开前 2 页** + 首屏前 28 个公式，剩余页面与公式在后台逐帧静默流式挂载，700+ 公式的长试卷也能瞬间打开；
   - 内置 `@media print` 样式，打印时自动隐藏页码角标并保证公式、表格行跨页不被劈开。

---

## 快速接入

### 方式一：NPM 安装（React / Vue / Next.js / Vite）

```bash
npm install docx-paper jszip
```

```js
import { renderDocx } from 'docx-paper';

const viewer = await renderDocx(fileArrayBuffer, document.getElementById('viewer'), {
  paginate: true,
  progressive: true
});
```

### 方式二：浏览器 CDN 直接引入（无需打包工具）

```html
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/docx-paper@1.0.0/dist/docx-paper.js"></script>

<div id="viewport" style="height: 100vh; overflow: auto; background: #e2e8f0; padding: 24px;">
  <div id="viewer"></div>
</div>

<script>
  const viewer = await DocxPaper.renderDocx(fileArrayBuffer, document.getElementById('viewer'), {
    scrollContainer: document.getElementById('viewport'),
    paginate: true,       // 自动物理纸张切页与页间间隙
    progressive: true,    // 首屏秒开前 2 页，后续页后台流式加载
    showPageBadge: true,  // 每页右上角显示「第 X / Y 页」
    onPageChange: ({ page, total, formulaCount }) => {
      console.log(`当前第 ${page} / ${total} 页，共 ${formulaCount} 处公式`);
    }
  });

  // 翻页、缩放与打印控制：
  // viewer.nextPage();
  // viewer.prevPage();
  // viewer.setZoom(1.2);
  // viewer.print();
</script>
```

---

## 与主流开源库对比

| 特性维度 | `mammoth.js` | 原版 `docx-preview` (`docxjs`) | **`docx-paper` (本项目)** |
| :--- | :---: | :---: | :---: |
| **页面级排版与页边距** | ❌ 仅提取纯文本语义 | ✅ 支持 | ✅ **支持（含缺省页边距标准回退）** |
| **MathType / OLE 嵌入公式 (`<w:object>`)** | ❌ 完全丢弃 | ❌ 完全丢弃 | ✅ **100% 完整解析还原** |
| **`.wmf` / `.emf` 二进制矢量图元** | ❌ 浏览器裂图 | ❌ 浏览器裂图 | ✅ **内置纯 JS 二进制转矢量 SVG / BMP** |
| **OMML 原生数学公式 (`<m:oMath>`)** | ❌ 丢失 | ⚠️ 存在 `<ms>` 双引号 Bug、缺 `sSubSup`/`acc` | ✅ **完整 W3C MathML 词法与结构编译** |
| **中文排版 (`firstLineChars`/着重号/拼音)** | ❌ 丢失 | ❌ 丢失 | ✅ **完整支持 + 中文字体回退链** |
| **中文“三线表”与内框线 (`insideH`/`insideV`)** | ❌ 无样式 | ❌ 内外框混用（Issue #221） | ✅ **外框与单元格内框独立解析** |
| **真实物理切页与流式加载** | ❌ 单页长文 | ⚠️ 无缓存分页符时粘成一长条，或产生空白页 | ✅ **按纸张高度自动切分独立卡片 + 首屏流式秒开** |

---

## 已知边界与尚未实现的特性（Known Limitations）

由于 `docx-paper` 的设计目标是保持**极致轻量（Gzip 后仅 ~90 KB）、零后端依赖、纯前端毫秒级启动**，而非加载 30MB+ 的 LibreOffice WebAssembly 虚拟机，以下少数 Word 专有特性目前尚未实现或存在边界限制：

1. **不支持旧版二进制 `.doc` 格式（Word 97–2003 BIFF 二进制结构）**：
   - 仅支持基于 OpenXML (ZIP) 规范的现代 `.docx` 文档。旧版 `.doc` 文件需先转换为 `.docx`。
2. **原生 SmartArt 与复杂 DrawingML 图表（`<dgm:relIds>`、`<c:chart>`）**：
   - Word 桌面端使用专有 C++ 算法动态计算 SmartArt 节点布局与 3D 图表网格。若文档内未包含 `<mc:Fallback>` 预览图，纯 DrawingML 动态图表暂不支持实时重算绘制。
3. **WordArt 3D 艺术字弯曲与立体光影特效（`<w14:textOutline>`、`<a:prstTxWarp>`）**：
   - 保留文字内容、字号与颜色，但不模拟贝塞尔曲线路径扭曲和三维棱台/光照效果。
4. **VBA 宏（`.docm`）与 ActiveX 控件**：
   - 出于浏览器安全沙箱与跨平台限制，忽略所有 VBA 宏脚本与 Windows 专有 ActiveX 控件。
5. **跨操作系统缺失字体时的字距微差**：
   - 若客户端系统未安装文档指定的 Windows 专有商业字体（如特定方正字库），回退到 macOS/Linux 系统字体后，浏览器的 CSS 断行算法可能与桌面版 Word 存在每行 1~2 个字符的折行位置微差。

## 开源协议

Apache-2.0
