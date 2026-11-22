# docx-paper (中文文档)

**[English](./README.md) | [简体中文](./README.zh-CN.md)**

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

## 3 行代码快速接入

```html
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<script src="./dist/docx-paper.js"></script>

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

## 开源协议

Apache-2.0
