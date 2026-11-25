---
name: courseware-review-pdf
description: 把课件幻灯片（.ppt/.pptx）批量倒成一份可打印的复习讲义 PDF：课件原图整页截图 + 逐题详解 + 表格/SVG 优先。使用场景——用户给一批课件 PPT，要求“生成复习 PDF”“把课件转成 PDF”“题目全部摘抄并附答案”“保留原图/直接截图”“做成能打印的讲义”。覆盖 .ppt 转换、全量导图、题目定位、HTML+CSS 排版、Chrome 出 PDF（带页码）、逐页校验的完整链路。
whenToUse: 输入是课件幻灯片（.ppt/.pptx），输出是可翻阅/可打印的复习资料或讲义 PDF，且要求保留原始图形与题目。
---

# 课件 PPT → 复习讲义 PDF

这套流程已完整跑通过一次：8 个 `.ppt` / 679 页课件 → 108 页讲义 / 70 张原图 / 72 道例题全解 / 15 MB。
本目录 `scripts/` 与 `assets/` 里是当时验证过的脚本和样式，**直接拷进项目复用，不要重写**。

## 0. 开工前先确认三件事

1. **交付物形态**：页数量级（10 页速查 / 100 页全解）、是否要封面+目录+附录。
2. **素材边界**：哪几个课件、是否含习题文件（常见情况：`习题/` 目录是空的，要提前说明并给出替代方案）。
3. **原图要求**：默认**必须整页截图**——这是这类需求的核心，不要用文字复述电路图。

如果用户没说，按下面这套默认契约执行，并在交付时说明：

> ① 课件里的**每一道题**都摘录，答案详细分步；课件没给答案或给得不全的，自己补齐。
> ② **能画图、能列表的，不用文字描述**（真值表、卡诺图、状态表、对比表、流程图）。
> ③ **直接贴 PPT 原页截图**，图下注明「课件 第X章 Pxx」，并补解释性文字。

## 1. 环境自检（先跑，别假设）

```powershell
# 有没有 Office / LibreOffice
Test-Path "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
Test-Path "C:\Program Files\LibreOffice\program\soffice.exe"
# 浏览器（出 PDF 用）
Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe"
# Python 依赖
& $PY -c "import pptx, PIL; print('ok')"
```

- 有 PowerPoint COM 就**优先用 PowerPoint**：`.ppt` 用 `SaveAs(24)` 转 `.pptx` 保真度最高；LibreOffice 仅在无 Office 时兜底。
- 无 poppler / gs / magick 也能做 PDF 目视校验——用 `scripts/pdf2png.ps1`（WinRT `Windows.Data.Pdf`，零依赖）。
- `$PY` 用 `load_workspace_dependencies` 返回的 bundled python，`python -c "import pptx"` 失败时别急着 pip install。

## 2. 流水线（六步）

### A. 转换 + 全量导图

```powershell
powershell -ExecutionPolicy Bypass -File scripts\ppt2png.ps1 -InputDir <课件目录>
# -> <课件目录>/_pptx/chNN.pptx, chNN.name.txt
# -> <课件目录>/_slides/chNN/001.png ...   (1600x1200)
```

- 一次导出**全部**幻灯片，不要边看边导：后面定位题目、核对公式、挑插图全都要用到。
- 600+ 页约 5–8 分钟，放后台跑，同时做 B 步。
- 命名 `chNN` 与原始中文文件名解耦，避免中文路径在脚本里出问题；映射记在 `chNN.name.txt`。

### B. 抽文本、建索引

```powershell
& $PY scripts\extract.py --pptx <课件目录>\_pptx --out <课件目录>\_text
```

产出 `chNN.json`、`chNN.txt`、`index.txt`。**索引是后续所有判断的地基**：每页一行，带 `P`（有图）/`T`（有表）标记和开头 110 字。

**关键：用 `read` 工具读这些 UTF-8 文件，不要靠控制台回显。** Windows 控制台是 GBK，中文会糊成乱码，白跑一轮。

### C. 定位题目

先用 `index.txt` 全量扫一遍（一章一次 `read`，带 offset），再对可疑页用正则细筛：

```
例|例题|习题|练习|思考|作业  →  强信号
求|试画|画出|写出|化简|设计|分析|判断  →  弱信号，两条以上才算
```

三条经验：

- **课件后半页常是“续页”**：一道例题跨 2–8 页（状态图 → 卡诺图 → 驱动方程 → 逻辑图），要把整段页号一起记下来，别只抓标题页。
- **公式里的非号**：这份课件的 PPT 用 `A'` 写法，文本抽取能完整保留；但也有课件把非号画成独立线框，抽取后就丢了——**凡是公式密集的题，务必 `read_image` 看一眼原文**再落笔。
- **章末“习题 X.Y”只有题号没有题面**（很常见）。不要瞎猜教材题，改为：列出题号 + 补写同类型自测题并给答案，并在使用说明里讲清楚。

### D. 定结构

实测好用的骨架：

```
封面（不透明卡片式，别做满出血，见踩坑）
目录（章节 + 例题数 + 各章关键词，标注每题来源页码）
每章：
  ├ 章首导语（本章在全书的位置 + 最易失分的 2–3 点）
  ├ 知识点速览（全是表格 + 1–2 张 SVG 流程图）
  ├ 例题卡片 ×N  ← 主体
  │    ├ 标题条：【例 4-3】…  |  课件 第4章 P18~P20
  │    ├ 原题（一句话重述，便于检索）
  │    ├ 整页原图截图（1~3 张，按解题顺序）
  │    └ 解析：思路 → 分步（①②③）→ 结论框 → 易错提醒框
  └ 章末自测（表格：题号 / 题目 / 答案）
附录：器件速查表、易错点清单、综合自测
```

每道例题大约占 1–1.5 页；72 道题 + 8 章知识块 ≈ 100 页，可作为排期基准。

### E. 处理图片 + 写 HTML

```powershell
& $PY scripts\prep_img.py --slides <课件目录>\_slides --out _src\img --sel ch03:031 ch03:038 ...
```

- 白点拉伸（`--thr 232`）把 PPT 模板的网格底纹压成纯白：体积从 ~400 KB 降到 ~150 KB，打印也更干净。
- 宽度 1200 px（≈235 dpi @13 cm），质量 80。70 张图约 11 MB，成书 15 MB，可接受。
- 只转要用的页；`--all` 仅在课件很短或用户明确要求全贴时使用。

HTML 片段按 `NN_*.html` 命名（`00_front` … `09_appendix`），配 `assets/print.css`。**样式直接用这份，别从零写**。

### F. 出 PDF

```powershell
& $NODE scripts\print_pdf.mjs _src "..\xxx_复习讲义.pdf" "数字电路与逻辑设计 · 期末复习讲义"
```

`print_pdf.mjs` 做三件事：拼接分章片段 + 内联 CSS → 用 CDP 起 Chrome → `Page.printToPDF`。
**必须走 CDP，不要用 `chrome --headless --print-to-pdf`**：命令行版无法自定义页脚，页码只能靠 CDP 的 `footerTemplate`。

### G. 校验（五件事，一件都别省）

```powershell
& scripts\pdf2png.ps1 <pdf> _prev 200      # 逐页栅格化
```

1. **总数**：`pages: N` 是否符合预期量级。
2. **近空页**：`Get-ChildItem _prev\*.png | ? Length -lt 160000` —— 有就说明排版炸了（分页断裂/图片没加载）。**目标 0 张。**
3. **缺图**：正则抓 HTML 里全部 `src="..."`，逐个 `Test-Path`。**缺图在 PDF 里是静默的破图标，不查就交付了。**
4. **抽样目视**：封面、目录、每章首页、含 SVG 的页、代码块页各挑 1–2 页 `read_image`。
5. **内容抽查**：随机抽 3 道题，把 PDF 上的公式和课件原图逐字对一遍。

## 3. 版式规范（`assets/print.css` 里已实现）

| 项 | 取值 | 为什么 |
|---|---|---|
| 纸张 / 页边距 | A4，`0.52 / 0.58 / 0.55 / 0.55` inch | 给 CDP 页脚留出空间 |
| 正文 | 微软雅黑 10.2pt，行高 1.75 | 中文密排可读 |
| 公式 | Cambria/Times 斜体，沿用课件的 `A'` 写法 | 和原图一致，学生不额外换算 |
| 章节标题 | `break-before: page` + 深蓝色块 | 章章独立，便于翻查 |
| 表格 | 深蓝表头 + 斑马纹 + `break-inside: avoid` | 表被劈开最难读 |
| 例题卡片 | 渐变标题条 + 浅边框，**允许跨页** | 见踩坑 3 |
| 结论 / 易错 | `.res` 绿框 / `.box.warn` 红框 | 视觉锚点，复习时一眼抓到 |

## 4. 内容规范

- **解析要写“为什么”，不是复述答案**。课件已给出答案的，补中间步骤；课件只给结论的，补推导；课件写得跳跃的，补一句“这一步用了哪个公式”。
- **主动核对课件答案**。这次就发现两处课件错误（某题答案的第 3 项印错，代入某个最小项后与真值表矛盾；两处的例题编号串位）。发现后在正文里标明“课件此处印作 X，按卡诺图应为 Y，建议以卡诺图为准”——这比照抄更有价值。
- **公式统一写法**：`A'` 表示非，`⊕` 异或，`∑m(…)` 最小项和，`∑d(…)` 无关项和；下标用 `<sub>`。
- **每章末留自测表**（题号 / 题目 / 答案三列，一眼可对答案）。
- **附录要有“一页纸”**：器件速查表 + 易错点 20 条 + 综合自测。考前最后 30 分钟只翻附录。

## 5. 踩坑清单（都是这次真踩到的）

1. **封面别做满出血**。打印时负 margin 会被裁掉，`position: fixed` 的页脚也不会稳稳出现在每页底部。封面改成**页面内的大色块卡片**，页码交给 CDP `footerTemplate`。
2. **页码只能用 `footerTemplate`**。CSS 的 `@page { @bottom-center }` Chrome 不支持；`position:fixed` 在分页时错位。
3. **`.card { break-inside: avoid }` 会制造大量近空页**（实测 121 页里有十几页只剩几行）。正确做法：卡片允许跨页，只对 `标题条` 加 `break-after: avoid`、对 `表格/图/pre` 加 `break-inside: avoid`。改完 121 → 105 页。
4. **中文标题被两端对齐拉散**。`body{text-align:justify}` 会作用到标题，出现“扩 展 题 三 步 速 算 法”。给 `h1,h2,h3,.box .h,.card>.hd` 统一 `text-align:left`。
5. **PowerShell 5.1 把 `.ps1` 当 ANSI 读**。脚本里写中文路径 = 找不到文件。要么脚本内**全 ASCII**（推荐，用参数传路径），要么存成 UTF-8 **带 BOM**。
6. **`$presentation.Visible = $true` 会报 MsoTriState 转换错误**，写 `$ppt.Visible = -1` 并 try/catch。
7. **控制台输出是 GBK**。所有中文分析结果写文件、再用 `read` 读；不要在 `python ... 2>&1 | Out-String` 里看中文。
8. **`python-pptx` 读不了 `.ppt`**（只能 `.pptx`）。必须先转换。
9. **`Slide.Export` 尺寸**：4:3 课件写 1600×1200；16:9 用 1600×900，否则图会被拉变形。
10. **缺图是静默失败**。`<img src>` 指错路径，PDF 里只留一个破图标，肉眼粗扫很容易漏——必须用脚本核对（见校验 3）。
11. **图片不压缩会让 PDF 虚胖**。79 张原图 PNG 转 1400px JPEG-86 是 27 MB；改 1200px + 白点拉伸 + q80 后 11 MB。
12. **别在 `<pre>` 里用中文引号混排代码**，等宽字体下会错位；代码块统一 `Consolas` + 灰底 + 左侧蓝条。

## 6. 交付与清理

- 交付前跑完第 G 节五项校验，`present` 一次 PDF 即可（源文件在文末一句带过）。
- 中间产物很占地方（679 张 PNG + pptx ≈ 230 MB）。交付后：
  - 保留 `_src/`（分章 HTML + 图 + CSS，约 11 MB）用于返工；
  - 保留 `_build/`（`print_pdf.mjs` + `pdf2png.ps1`）；
  - 删掉 `_slides/`、`_pptx/`、`_text/`、`.chromeprofile/`、`_prev/`、draft PDF。
- 顺手把 `_src/img/` 里**没被引用**的图删掉（按 HTML 里的 `src` 求差集），通常能再省 1–2 MB。

## 7. 复用清单

| 步骤 | 脚本 / 资产 | 备注 |
|---|---|---|
| 转格式 + 全量导图 | `scripts/ppt2png.ps1` | 需 PowerPoint COM；`-Width/-Height` 跟版式走 |
| 抽文本建索引 | `scripts/extract.py` | 递归 group、读备注、出 `index.txt` |
| 图片压缩 | `scripts/prep_img.py` | `--sel ch03:031` 或 `--all` |
| 版式 | `assets/print.css` | 直接拷成项目 `style.css` |
| 出 PDF（带页码） | `scripts/print_pdf.mjs` | Chrome + CDP，`footerTemplate` 出页码 |
| PDF 目视校验 | `scripts/pdf2png.ps1` | WinRT，零依赖，替代 poppler |
