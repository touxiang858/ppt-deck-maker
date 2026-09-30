# 给 DeepSeek（网页版）的提示词 · PPT 渲染台专用

## 怎么用（三步）

1. 打开 DeepSeek 网页版，**新建一个对话**。
2. 把本文件**从「↓↓↓ 下面整段粘贴 ↓↓↓」开始到结尾**，整段复制粘进去；
   然后把最下面的【我的需求】填好，直接发送。
3. 拿到回复后，**整段复制**存成一个 `.json` 文件，拖进「PPT 渲染台」窗口，点「开始」。

> 素材目录先用渲染台的「素材清单」扫一次；本文件的【可用素材】段落里也会自动带上文件名清单。

---

## ↓↓↓ 下面整段粘贴 ↓↓↓

你是资深的 PPT 结构师。请严格按我给的规则，输出**一个 deck-spec JSON**。

### 一、输出格式（最重要的要求，先看这条）

1. **JSON 必须放在 ```json 代码块里**，像这样：

```json
{ "assets_dir": "...", "output": "...", "deck": { ... }, "pages": [ ... ] }
```

2. 代码块外面可以写一两句话说明，但**代码块里面必须是完整、能被机器直接解析的 JSON**：
   - 不要写注释（`//`、`#`、`/* */` 都不行）
   - 不要用省略号、不要用「...」占位，每一页都要写全
   - 不要出现中文引号，双引号一律用英文 " "
   - 确认首尾大括号闭合，最后一个字符是 }
3. 如果你觉得内容太多一次写不完，**宁可少写几页并告诉我「还差几页」**，也不要在 JSON 中间截断。
   （截断的 JSON 渲染台读不了，等于白写。）

### 二、deck-spec 规则（页面类型全表）

# PPT 渲染台 · deck-spec 预设全表（v2 · 25 种页面）

> 把这份文件整段发给 DeepSeek，让它按这里的类型产出 deck-spec JSON；
> 再把这个 JSON 拖进「PPT 渲染台」窗口，就能出 pptx。
> **只允许用下面列出的 type**，别自创。

---

## 一、顶层结构

```json
{
  "assets_dir": "C:/素材/图片文件夹",
  "output": "D:/成品.pptx",
  "deck": {
    "title": "整套的标题",
    "subtitle": "副标题",
    "theme": { "bg": "22384E", "ink": "F2EDE3", "sub": "C9BFA9", "gold": "C9A053",
               "red": "A63A2E", "hair": "3A5566", "muted": "9AA8B5", "font": "微软雅黑" }
  },
  "pages": [ { "type": "cover", "title": "……" } ]
}
```

- 颜色写 6 位十六进制**不带 #**。theme 不写就用上面这套深蓝金默认值。
- 图片只写**文件名**（不要路径），文件放在 `assets_dir` 下。
- 每页都能加：`badge`（左上角红色序号，如 "03"）、`pill`（右上角金色标签）、`footer`（左下页脚文字）。

**页面顺序惯例**：`cover` → `toc` → 每章 `section` + 若干内容页 → `contact` 或 `finale` 收尾。

---

## 二、25 种页面类型

### 骨架类（4 种）

| type | 用途 | 必填 | 常用可选 |
|---|---|---|---|
| `cover` | 封面 | title | kicker, subtitle, tagline, image, strip |
| `toc` | 目录（两栏编号列表） | title, rows | current（要高亮的章节名）, badge |
| `section` | 章节过渡页（大序号+大标题） | title | badge, subtitle, image（整页背景图） |
| `contact` | 结尾联系页 | title | slogan, lines, image（二维码）, image_caption |

```json
{ "type": "cover", "kicker": "重阳节文创纪念品系列", "title": "瑞鹤九九",
  "subtitle": "菊酒登高糕 · 新品", "tagline": "一块糕，过完整个重阳节。",
  "strip": [["菊金", "#D9A441"], ["米白", "#F3EDDF"]], "image": "瑞鹤图.jpg" }

{ "type": "toc", "badge": "00", "title": "目录", "current": "产品矩阵",
  "rows": [["文化母题", "重阳三俗与一味"], ["产品矩阵", "四条产品线怎么分工"]] }

{ "type": "section", "badge": "01", "title": "文化母题",
  "subtitle": "先把「重阳」讲清楚，再谈产品", "image": "场景图.jpg" }

{ "type": "contact", "title": "谢谢观看", "slogan": "一块糕，过完整个重阳节。",
  "lines": ["商务合作：hi@example.com", "样品索取：138-0000-0000"],
  "image": "二维码.png", "image_caption": "扫码看完整系列" }
```

### 文字内容类（6 种）

| type | 用途 | 必填 | 常用可选 |
|---|---|---|---|
| `bullets` | 要点页（一列「小标题+说明」） | title, rows | image, image_caption |
| `imageText` | 图文左右页（最常用） | title, image, rows | image_side: left/right, image_caption |
| `cards` | 2–4 张并列卡片（优势/板块/卖点） | title, cards | — |
| `compare` | 2–3 栏对照（优缺点/方案A·B/前后） | title, columns | note（结论一句话） |
| `quote` | 金句/引用页 | quote | source, image（整页背景） |
| `table` | 表格（规格/参数/清单） | title, head, rows | col_widths, row_h, note |

```json
{ "type": "bullets", "badge": "02", "title": "重阳三俗与一味",
  "rows": [["登高", "九月九日登高望远，取「步步高」的口彩。"],
           ["佩茱萸", "茱萸囊佩于臂，古法驱邪。"]] }

{ "type": "imageText", "title": "产品形态", "image": "产品.png",
  "image_side": "right", "image_caption": "单块装实拍",
  "rows": [["九层压制", "九层对应「九九」。"], ["菊酒入馅", "菊花酒渍枣泥作芯。"]] }

{ "type": "cards", "title": "四条产品线",
  "cards": [ { "icon": "礼", "title": "登高礼盒", "body": "九块装，主打节日送礼。" },
             { "icon": "尝", "title": "尝鲜单块", "body": "便利店渠道，拉新客。" } ] }

{ "type": "compare", "title": "两种包装方案怎么选",
  "columns": [ { "title": "方案 A · 纸盒卡扣", "tone": "up",
                 "points": ["成本低 32%", "打样周期 3 天"] },
               { "title": "方案 B · 翻盖磁吸", "tone": "down",
                 "points": ["成本高，客单价可提 40%", "工期多 5 天"] } ],
  "note": "结论：首批走 A 控风险，礼盒渠道直接上 B 做形象款。" }

{ "type": "quote", "quote": ["把「高」做成一口能吃到的东西，", "比写在包装上更有说服力。"],
  "source": "产品复盘会 · 2026.09", "image": "场景图.jpg" }

{ "type": "table", "title": "规格与成本",
  "head": ["规格", "净含量", "出厂价", "建议零售", "毛利率"],
  "rows": [["尝鲜单块", "45 g", "¥ 3.2", "¥ 6.0", "47%"],
           ["九块礼盒", "405 g", "¥ 48.0", "¥ 98.0", "51%"]],
  "note": "口径：含包装与物流摊销，不含渠道扣点。" }
```

### 结构与关系类（6 种）

| type | 用途 | 必填 | 常用可选 |
|---|---|---|---|
| `timeline` | 时间轴（节点上下交替） | title, nodes | — |
| `process` | 流程/步骤（编号卡片+箭头） | title, steps | — |
| `matrix` | 四象限（SWOT / 优先级 / 业务划分） | title, quadrants | x_label, y_label, note |
| `pyramid` | 金字塔/层级（自下而上） | title, levels | base_note |
| `stats` | 数据看板（2–4 个大数字） | title, stats | source |
| `chart` | 图表 + 右侧结论 | title, chart | note |

```json
{ "type": "timeline", "title": "从打样到上架",
  "nodes": [ { "time": "第 1 周", "title": "配方定稿", "desc": "三轮盲测" },
             { "time": "第 2 周", "title": "开模打样", "desc": "九层压制测试" } ] }

{ "type": "process", "title": "打样流程",
  "steps": [["备料", "分批称重，留样"], ["压制", "九层依次压制"], ["蒸制", "中心温度 92℃"]] }

{ "type": "matrix", "title": "优先级矩阵", "x_label": "落地难度 →", "y_label": "市场收益",
  "quadrants": [ { "name": "立刻做", "desc": "成本低、两周内可上架。" },
                 { "name": "重点投入", "desc": "收益最高，但要供应链配合。" },
                 { "name": "先放着", "desc": "收益不确定，压资金。" },
                 { "name": "顺手做", "desc": "依附现有渠道，边际成本低。" } ] }

{ "type": "pyramid", "title": "品牌价值层级", "base_note": "越往上越抽象，越往下越可验证",
  "levels": [ { "name": "产品事实", "desc": "九层压制、菊酒入馅" },
              { "name": "使用场景", "desc": "重阳送礼、茶馆伴手" },
              { "name": "情绪价值", "desc": "「步步高」的好彩头" },
              { "name": "文化母题", "desc": "重阳 · 瑞鹤 · 菊" } ] }

{ "type": "stats", "title": "首批小批量结果",
  "source": "口径：上架后 30 天，含全部渠道",
  "stats": [ { "value": "12,400", "label": "总销量（块）", "note": "礼盒 + 散卖" },
             { "value": "38%", "label": "礼盒渠道复购", "note": "团购客户占七成" } ] }

{ "type": "chart", "title": "各渠道首月销量",
  "chart": { "type": "bar", "categories": ["便利店", "茶馆堂食", "企业团购"],
             "series": [ { "name": "礼盒（盒）", "values": [420, 260, 1480] },
                         { "name": "单块（十块）", "values": [260, 180, 400] } ] },
  "note": ["企业团购是绝对主力。", "", "单块的拉新作用明显：", "38% 的礼盒客户先买过单块。"] }
```

`chart.type` 可选：`bar`（柱/条，`bar_dir: "col"|"bar"`）、`line`、`area`、`pie`、`doughnut`。
饼图/环形图配 `show_value: true` 会把数值标在扇区上。

### 图片类（4 种）

| type | 用途 | 必填 | 常用可选 |
|---|---|---|---|
| `product` | 产品页（1–2 图 + 参数行 + 参考图） | title, images, rows | ref{src,caption} |
| `case` | 案例页（左图 + 右文 + 底部数据条） | title, image, rows | metrics, image_caption |
| `team` | 团队/人物页 | title, images | images[].label/.sub/.note |
| `gallery` | 图集网格（3/4/6 图 + 图注） | title, images | captions[] |

```json
{ "type": "product", "title": "菊酒登高糕",
  "images": [ { "src": "产品.png", "label": "正面" }, { "src": "场景.png", "label": "场景" } ],
  "rows": [["配方", "菊酒渍枣泥作芯"], ["尺寸", "45×45×38 mm"]],
  "ref": { "src": "参考.jpg", "caption": "参考实拍" } }

{ "type": "case", "title": "课题三 · 城市文具落地", "image": "实拍.jpg",
  "image_caption": "线下快闪店实拍",
  "rows": [["挑战", "同质化严重，缺在地文化母题。"],
           ["做法", "城市地标拆解为图形系统。"],
           ["结果", "首月售罄，71% 主动集齐一整套。"]],
  "metrics": [["3 周", "从提案到上架"], ["+64%", "客单价提升"], ["71%", "整套收集率"]] }

{ "type": "team", "title": "这一版是谁做的",
  "images": [ { "src": "a.jpg", "label": "配方", "sub": "三轮盲测定甜度" },
              { "src": "b.jpg", "label": "包装", "sub": "封套与结构打样" } ] }

{ "type": "gallery", "title": "视觉素材", "images": ["a.jpg", "b.png", "c.png", "d.jpg"],
  "captions": ["瑞鹤图原作", "产品主图", "场景图", "实拍参考"] }
```

### 文创专有类（5 种，一期就有）

| type | 用途 | 必填 | 常用可选 |
|---|---|---|---|
| `intro` | 图文介绍页（左图 + 右侧「标签+正文」多行） | title, rows | image, image_caption |
| `scene` | 整页大图页 | image, chip | title_line |
| `symbols` | 符号转译（左侧多行 + 右侧 2 图） | title, rows, images | fig_caps[] |
| `colors` | 系列用色（色块 + 比例条） | title, swatches, rows | ratio[[名,色,占比]], ratio_note |
| `finale` | 收尾页（缩略图条 + 金句） | title, thumbs, quote | strip, badge, pill |

```json
{ "type": "scene", "image": "场景.jpg", "chip": "03  菊酒登高糕 · 场景",
  "title_line": "一块糕，过完整个重阳节。" }

{ "type": "colors", "title": "系列用色",
  "swatches": [["菊金", "#D9A441"], ["米白", "#F3EDDF"], ["枣褐", "#6B3A28"]],
  "rows": [["主色", "菊金用于标题与强调"], ["辅色", "米白大面积留白"]],
  "ratio": [["菊金", "#D9A441", 30], ["米白", "#F3EDDF", 55], ["枣褐", "#6B3A28", 15]],
  "ratio_note": "系列用色比例" }
```

---

## 三、给 DeepSeek 的约束（可直接抄）

1. 只能使用上面 25 个 `type`，字段名严格按表里的写法。
2. `rows` 一律是 `[["小标题","正文"], …]` 两元数组；
   `cards/columns/nodes/steps/quadrants/levels/stats` 既可用对象，也可用数组简写
   （按表里「必填」的顺序填，如 `steps: [["备料","分批称重"], …]`）。
3. 图片只写文件名，不要写路径；文件名必须和 `assets_dir` 下的实际文件一致。
4. 一页只讲一件事：正文控制在 **每条 20–40 字**，全页不超过 5 条。
5. 数据必须自洽：`stats` 的数字要和 `chart`/`table` 对得上；
   `chart` 的多个 series 单位要统一、量级要接近，否则柱子会看不出对比。
6. 长文档的节奏建议：每 4–6 页插一个 `section`；每页最多 1 个 `stats` 或 `chart`；
   `quote` 用来换气，不要连续出现。
7. 输出**纯 JSON**，不要加解释文字和 markdown 代码围栏。

> 引擎自带两道兜底：行数多了会**自动压行距**、正文太长会**自动缩字号**，
> 所以偶尔写长一点不会糊出框；但缩到 7.5pt 就到底了，**该短还是要短**。

### 三、素材怎么引用

- 图片**只写文件名**，不要写路径、不要写 http 地址。
- 文件名**必须**来自我给你的素材清单，不能自己编。
- `assets_dir` 填我给你的那个素材目录路径（正斜杠或反斜杠都行）。

### 四、内容约束

1. 一页只讲一件事；`rows` 里每条正文控制 **20–40 字**，一页不超过 5 条。
2. 只能使用上面列出的 `type`，字段名严格照抄，**不要自创类型或字段**。
3. 数据要自洽：`stats` 的数字和 `chart`、`table` 能对上；
   同一个 `chart` 里多个系列的单位要统一、量级要接近。
4. 节奏建议：开篇 `cover` → `toc` → 每 4–6 页插一个 `section` → 收尾 `contact` 或 `finale`；
   `quote` 用来换气，不要连着两页都用。
5. 深色底、浅色字的默认主题已经配好，不用自己配色；要换色只改 `deck.theme` 里那几个字段。

### 五、我的需求（这一节请按我的填写来）

【主题】：（例如：罗马介绍）
【页数】：（例如：25 页）
【受众】：（例如：中学生 / 公司内部培训 / 客户提案）
【语言风格】：（例如：简洁书面语 / 口语化 / 带一点幽默）
【特别要求】：（例如：要有时间轴和数据页；不要出现英文；结尾放联系方式）

### 六、可用素材（文件名清单）

【素材目录】：
（这里放素材目录路径）

【文件名】：
（这里放文件名清单，每行一个）

---

现在请按上面全部规则，输出那份 deck-spec JSON。**记得用 ```json 代码块包起来。**