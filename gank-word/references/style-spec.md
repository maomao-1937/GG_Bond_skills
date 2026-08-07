# 万人教育 Word 排版规范 — 样式精确取值参考

## 字号换算表

中文字号制到磅再到 Word XML 半磅值（`w:sz` / `w:szCs` 用半磅 half-point）：

| 中文名 | 磅(pt) | 半磅(half-pt) |
|--------|--------|---------------|
| 初号   | 42     | 84            |
| 小初   | 36     | 72            |
| 一号   | 26     | 52            |
| 小一   | 24     | 48            |
| 二号   | 22     | 44            |
| 小二   | 18     | 36            |
| 三号   | 16     | 32            |
| 小三   | 15     | 30            |
| 四号   | 14     | 28            |
| 小四   | 12     | 24            |
| 五号   | 10.5   | 21            |
| 小五   | 9      | 18            |

用户说的 "48号" → 48pt → 半磅 96。

## 度量换算

- 磅 → 二十分之一磅（twips）：`pt × 20`
- cm → twips：`cm × 567`（精确 567.0）
- 1.5 倍行距：`w:line="360" w:lineRule="auto"`（240 = 单倍）
- 段前 17.5pt → twips = 350
- 段后 16pt → twips = 320
- 页眉距边界 1.3cm → twips ≈ 737
- 页脚距边界 0.9cm → twips ≈ 510

## 各样式的精确 XML 属性

### TOCHeading（目录标题）

```xml
<w:style w:type="paragraph" w:styleId="TOCHeading">
  <w:pPr>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:eastAsia="华文新魏" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
    <w:sz w:val="96"/>
    <w:szCs w:val="96"/>
    <w:color w:val="000000"/>
  </w:rPr>
</w:style>
```

不加粗（无 `<w:b/>`）。

### TOC Entry（目录条目 — toc 1..toc 9）

TOC 条目样式是 Word 生成目录字段时自动创建的，styleId 随版本/语言而异：英文版 `TOC 1`~`TOC 9`，中文版可能是 `目录 1`~`目录 9`。脚本用 `TOC_ENTRY_PATTERN`（大小写不敏感）匹配。

脚本对 TOC 条目样式只做 **targeted 字体更新**——只改 `w:rFonts` 与去粗体，spacing、缩进、制表位（点式前导符）、字号、颜色一律保留，保证目录原有版式不破坏。

```xml
<!-- 对每个 toc-N 样式执行的 rPr 定向更新 -->
<w:rPr>
  <w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <!-- 若存在 w:b / w:bCs 则删除 -->
</w:rPr>
```

字体规格：

- 中文字体（eastAsia）：宋体
- 西文字体（ascii / hAnsi）：Times New Roman
- 粗体：显式移除（无 `<w:b/>` / `<w:bCs/>`）
- 其余属性（字号、颜色、间距、缩进、制表位）保持 Word 原样

TOC 条目**段落**同样只清理 run 级直接字体与粗体（含 hyperlink 内嵌套的 run），段落样式仍为原来的 toc-N，pStyle 不改。

### Heading 1（一级标题）

```xml
<w:pPr>
  <w:jc w:val="center"/>
  <w:spacing w:before="350" w:after="320"/>
</w:pPr>
<w:rPr>
  <w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <w:b/>
  <w:bCs/>
  <w:sz w:val="32"/>
  <w:szCs w:val="32"/>
  <w:color w:val="000000"/>
</w:rPr>
```

### Heading 2（二级标题）

```xml
<w:pPr>
  <w:jc w:val="left"/>
  <w:spacing w:line="360" w:lineRule="auto"/>
</w:pPr>
<w:rPr>
  <w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <w:b/>
  <w:bCs/>
  <w:sz w:val="30"/>
  <w:szCs w:val="30"/>
  <w:color w:val="000000"/>
</w:rPr>
```

### Heading 3（三级标题）

```xml
<w:pPr>
  <w:jc w:val="left"/>
  <w:spacing w:line="360" w:lineRule="auto"/>
</w:pPr>
<w:rPr>
  <w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <w:b/>
  <w:bCs/>
  <w:sz w:val="28"/>
  <w:szCs w:val="28"/>
  <w:color w:val="000000"/>
</w:rPr>
```

### Normal（正文）

```xml
<w:pPr>
  <w:jc w:val="both"/>
  <w:spacing w:line="360" w:lineRule="auto"/>
</w:pPr>
<w:rPr>
  <w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <w:sz w:val="24"/>
  <w:szCs w:val="24"/>
  <w:color w:val="000000"/>
</w:rPr>
```

### Header（页眉）

```xml
<w:pPr>
  <w:jc w:val="center"/>
</w:pPr>
<w:rPr>
  <w:rFonts w:eastAsia="华文新魏" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <w:sz w:val="21"/>
  <w:szCs w:val="21"/>
  <w:color w:val="000000"/>
</w:rPr>
```

### Footer（页脚）

Same as Header.

### WanrenExplanation（解析）

解析段落**不使用专属段落样式**，而是：

- 段落样式设为 **Normal**（宋体，小四，两端对齐，1.5 倍行距，黑色）
- 「【万人教育解析】」标签所在 run 的 `w:rPr/w:rFonts` 设为 `w:eastAsia="华文新魏"`
- 标签可能被 Word 拆成多个 run：脚本按文本位置判断，凡是覆盖标签前缀的 run 都保持华文新魏；其余 run 删除直接字体、继承 Normal 的宋体

等价 run 级设置：

```xml
<w:rPr>
  <w:rFonts w:eastAsia="华文新魏"/>
</w:rPr>
```

## sectPr 页眉页脚距离

```xml
<w:pgMar ... w:header="737" w:footer="510" .../>
```

## Word 标题样式名映射

脚本识别标题时，需要兼容中文版 Word 和英文版 Word 的样式名：

| 标准 styleId | 中文 `w:name` | 英文 `w:name` |
|---|---|---|
| Heading1 / 1 | 标题 1 | heading 1 |
| Heading2 / 2 | 标题 2 | heading 2 |
| Heading3 / 3 | 标题 3 | heading 3 |
| TOCHeading   | 目录标题 | TOC Heading |

注意：`w:styleId` 值可能为 `Heading1`、`1`、`heading1` 等变体，以及中文 Word 可能用纯数字或 `a0`+数字形式。脚本需按 `w:name` 内容做 case-insensitive 匹配后再处理。
