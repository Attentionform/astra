# 📐 ISO/IEC 29500 OOXML Layout & Canvas Renderer Mathematics Specification

## 1. Unit Systems & Precision Conversion Architecture

Microsoft WordprocessingML (ISO/IEC 29500) uses several distinct unit systems for different visual properties. When building a native Canvas visual renderer, all OOXML integer values must be converted to Canvas CSS pixels ($96 \text{ DPI}$) with high floating-point precision.

### 1.1 Summary of OOXML Units & Scale Factors

| OOXML Unit | ISO/IEC Tag / Context | Mathematical Definition | Conversion to Inches | Canvas Pixels ($96 \text{ DPI}$) |
| :--- | :--- | :--- | :--- | :--- |
| **Twip** | `<w:pgMar>`, `<w:spacing>`, `<w:ind>`, `<w:tblCellMar>` | $1/20 \text{ of a point}$ | $1/1440 \text{ in}$ | $1 \text{ twip} = \frac{96}{1440} = \mathbf{\frac{1}{15} \text{ px}} \approx 0.0666667 \text{ px}$ |
| **Half-Point** | `<w:sz>`, `<w:szCs>` (Font sizes) | $1/2 \text{ of a point}$ | $1/144 \text{ in}$ | $1 \text{ hp} = \frac{1}{2} \times \frac{96}{72} = \mathbf{\frac{2}{3} \text{ px}} \approx 0.666667 \text{ px}$ |
| **Eighth-Point** | `<w:bdr w:sz="...">`, `<w:tcBorders>` | $1/8 \text{ of a point}$ | $1/576 \text{ in}$ | $1 \text{ ep} = \frac{1}{8} \times \frac{96}{72} = \mathbf{\frac{1}{6} \text{ px}} \approx 0.166667 \text{ px}$ |
| **EMU** | DrawingML `<wp:extent>`, `<a:off>`, `<a:ext>` | English Metric Unit | $1/914,400 \text{ in}$ | $1 \text{ EMU} = \frac{96}{914,400} = \mathbf{\frac{1}{9525} \text{ px}} \approx 0.000104987 \text{ px}$ |
| **Line 240ths** | `<w:spacing w:line="...">` (auto rule) | $1/240 \text{ of single line}$ | Relative | $\text{Multiplier} = \frac{\text{w:line}}{240}$ |
| **Degree 60000ths** | `<a:xfrm rot="...">` (Rotation) | $1/60,000 \text{ of a degree}$ | Angular | $\theta_{\text{rad}} = \frac{\text{rot}}{60000} \times \frac{\pi}{180} = \frac{\text{rot} \cdot \pi}{10,800,000}$ |

---

## 2. Unit Conversion Constants & Implementation Functions

```typescript
/**
 * ISO/IEC 29500 Layout Unit Conversion Engine
 * Assumes Standard Web / Canvas Display Resolution of 96 DPI.
 */
export const DPI = 96;
export const PTS_PER_INCH = 72;
export const TWIPS_PER_INCH = 1440;
export const EMUS_PER_INCH = 914400;

export const TWIPS_PER_PX = TWIPS_PER_INCH / DPI; // 15
export const PX_PER_TWIP = 1 / TWIPS_PER_PX;       // 0.06666666666666667

export const EMUS_PER_PX = EMUS_PER_INCH / DPI;   // 9525
export const PX_PER_EMU = 1 / EMUS_PER_PX;         // 0.00010498687664041995

// Conversion Functions
export const twipsToPx = (twips: number): number => twips / 15;
export const pxToTwips = (px: number): number => Math.round(px * 15);

export const halfPtsToPx = (halfPts: number): number => (halfPts / 2) * (DPI / PTS_PER_INCH); // halfPts * (2/3)
export const eighthPtsToPx = (eighthPts: number): number => (eighthPts / 8) * (DPI / PTS_PER_INCH); // eighthPts * (1/6)

export const emusToPx = (emus: number): number => emus / 9525;
export const pxToEmus = (px: number): number => Math.round(px * 9525);

export const emuRotToRadians = (rot: number): number => (rot / 60000) * (Math.PI / 180);
```

---

## 3. Font Metrics & Line Height Calculations

### 3.1 Font Metrics Breakdown
When rendering a run of text (`<w:r>`) inside a paragraph (`<w:p>`), baseline placement and line height are governed by both font metrics and `<w:spacing>` properties.

Given a font with metrics:
- $UPM$: Units Per Em (e.g. 2048 for Calibri/Arial, 1000 for PostScript fonts)
- $A$: Typographic Ascent (`sTypoAscender` or `hhea.ascent`)
- $D$: Typographic Descent (`sTypoDescender` or `hhea.descent`, signed value)
- $G$: Typographic Line Gap (`sTypoLineGap` or `hhea.lineGap`)
- $F_{\text{pt}}$: Font size in points ($F_{\text{pt}} = \text{w:sz} / 2$)

The scale factor $S$ for converting font metric units to canvas pixels is:
$$S = \frac{F_{\text{pt}}}{UPM} \times \frac{96}{72} = \frac{F_{\text{pt}} \times 4}{UPM \times 3}$$

Computed Font Pixel Dimensions:
$$\text{Ascent}_{\text{px}} = A \times S$$
$$\text{Descent}_{\text{px}} = |D| \times S$$
$$\text{LineGap}_{\text{px}} = G \times S$$
$$\text{Natural Line Height}_{\text{px}} = (A + |D| + G) \times S$$

### 3.2 Spacing Rules (`<w:spacing w:line="..." w:lineRule="..."/>`)

1. **`w:lineRule="auto"` (or missing)**:
   - `w:line` specifies line height in $1/240\text{th}$ of a single line.
   - Default when omitted: `w:line="240"` ($1.0 \times$ natural line height).
   $$H_{\text{line}} = \text{Natural Line Height}_{\text{px}} \times \frac{\text{w:line}}{240}$$

2. **`w:lineRule="exact"`**:
   - `w:line` specifies exact line height in twips.
   $$H_{\text{line}} = \text{w:line} \times \frac{1}{15} \text{ px}$$
   - *Baseline position*: $Y_{\text{baseline}} = Y_{\text{line\_top}} + \text{Ascent}_{\text{px}}$. Text is centered vertically within the exact line height if line height exceeds font height, or clipped/overlapped if smaller.

3. **`w:lineRule="atLeast"`**:
   - `w:line` specifies minimum line height in twips.
   $$H_{\text{line}} = \max\left(\text{Natural Line Height}_{\text{px}}, \, \text{w:line} \times \frac{1}{15}\right)$$

---

## 4. Paragraph Spacing & Collapse Geometry

### 4.1 Paragraph Spacing Attributes (`<w:spacing>`)
- `w:before`: twips before paragraph top.
- `w:after`: twips after paragraph bottom.
- `w:contextualSpacing`: boolean (`true`/`1` or `false`/`0`).

### 4.2 Spacing Combination & Collapse Rules (ISO/IEC 29500-1 §17.3.1.33)

Unlike CSS where adjacent vertical margins collapse ($\max(M_1, M_2)$), **MS Word additively combines paragraph spacing**:

$$\text{Inter-Paragraph Gap}_{\text{px}} = \left( \text{after}(P_i) + \text{before}(P_{i+1}) \right) \times \frac{1}{15}$$

#### Exceptions to Additive Spacing:
1. **Contextual Spacing (`<w:contextualSpacing w:val="true"/>`)**:
   If $P_i$ and $P_{i+1}$ share the exact same style ID (`w:pStyle`) AND `contextualSpacing` is enabled:
   $$\text{Inter-Paragraph Gap}_{\text{px}} = 0$$

2. **Top-of-Page / Top-of-Cell Suppression**:
   - The first paragraph $P_1$ at the top of a page body or inside a table cell **suppresses `w:before`** ($\text{before}(P_1) = 0$), unless `<w:beforeAutospacing w:val="1"/>` is explicitly set or compatibility flag `doNotSuppressParagraphCorrection` is set.
   - The last paragraph in a table cell suppresses `w:after`.

---

## 5. Page Geometry, Margins & Section Breakdown

### 5.1 Page Dimensions (`<w:pgSz>`)
Standard Letter Page:
- `w:w="12240"` twips ($8.5 \text{ in} = 816 \text{ px}$)
- `w:h="15840"` twips ($11.0 \text{ in} = 1056 \text{ px}$)

If `w:orient="landscape"`, swap width and height.

### 5.2 Margins & Dynamic Header/Footer Adjustments (`<w:pgMar>`)
Given section margin definitions in twips:
- `top`, `bottom`, `left`, `right`, `gutter`, `header`, `footer`

#### Printable Content Bounds Calculation:
$$X_{\text{content\_start}} = (\text{left} + \text{gutter}) \times \frac{1}{15}$$
$$W_{\text{content}} = (w - \text{left} - \text{right} - \text{gutter}) \times \frac{1}{15}$$
$$Y_{\text{content\_start}} = \max\left(\text{top} \times \frac{1}{15}, \, \text{header} \times \frac{1}{15} + H_{\text{header\_rendered}}\right)$$
$$Y_{\text{content\_end}} = \min\left((h - \text{bottom}) \times \frac{1}{15}, \, (h - \text{footer}) \times \frac{1}{15} - H_{\text{footer\_rendered}}\right)$$
$$H_{\text{content}} = Y_{\text{content\_end}} - Y_{\text{content\_start}}$$

> **Key Word Behavior**: If rendered Header height exceeds $(\text{top} - \text{header})$, MS Word dynamically pushes the body content start downward ($Y_{\text{content\_start}}$).

---

## 6. Table Grid, Cell Padding & Row Height Mathematics

### 6.1 Table Grid (`<w:tblGrid>`) & Column Widths
Grid column array $[W_1, W_2, \dots, W_N]$ in twips. Total table width $W_{\text{tbl\_px}} = \sum_{k=1}^N \frac{W_k}{15}$.

### 6.2 Table Cell Padding (`<w:tblCellMar>` / `<w:tcMar>`)
Default MS Word cell left/right padding is $108 \text{ twips}$ ($7.2 \text{ pt} = 9.6 \text{ px}$).

For cell $C_{i, j}$ spanning grid columns $j$ through $j + S - 1$ with grid span $S$:
$$W_{\text{cell\_outer}} = \sum_{k=j}^{j+S-1} \frac{W_k}{15}$$
$$W_{\text{cell\_inner}} = W_{\text{cell\_outer}} - (\text{pad}_{\text{left}} + \text{pad}_{\text{right}})$$

### 6.3 Table Row Height Rules (`<w:trHeight>`)
- `w:hRule="exact"`: Row height fixed to $\text{w:val} / 15 \text{ px}$. Content clipping applies if overflow occurs.
- `w:hRule="atLeast"`: $H_{\text{row}} = \max\left(\frac{\text{w:val}}{15}, \, \max_{\text{cells}} H_{\text{cell\_content}}\right)$.
- `w:hRule="auto"` (default): $H_{\text{row}} = \max_{\text{cells}} H_{\text{cell\_content}}$.

---

## 7. Graphic Transform Mathematics (EMUs to Canvas Matrix)

DrawingML containers specify positional offset (`<a:off>`), extent (`<a:ext>`), rotation (`<a:xfrm rot="...">`), and flipping (`flipH`, `flipV`).

### 7.1 Coordinate Matrix Transformation Code

```typescript
export interface DrawingTransform {
  offXEmus: number;
  offYEmus: number;
  extCXEmus: number;
  extCYEmus: number;
  rot60000ths?: number; // Rotation in 1/60000th of a degree
  flipH?: boolean;
  flipV?: boolean;
}

export function applyDrawingTransform(
  ctx: CanvasRenderingContext2D,
  transform: DrawingTransform
): { widthPx: number; heightPx: number } {
  const xPx = emusToPx(transform.offXEmus);
  const yPx = emusToPx(transform.offYEmus);
  const wPx = emusToPx(transform.extCXEmus);
  const hPx = emusToPx(transform.extCYEmus);

  const rad = transform.rot60000ths ? emuRotToRadians(transform.rot60000ths) : 0;
  const scaleX = transform.flipH ? -1 : 1;
  const scaleY = transform.flipV ? -1 : 1;

  ctx.save();
  // Translate center of object for rotation and scale flipping
  ctx.translate(xPx + wPx / 2, yPx + hPx / 2);
  if (rad !== 0) ctx.rotate(rad);
  if (scaleX !== 1 || scaleY !== 1) ctx.scale(scaleX, scaleY);
  ctx.translate(-wPx / 2, -hPx / 2);

  return { widthPx: wPx, heightPx: hPx };
}
```

---

## 8. Pagination Geometry & Line Flow Algorithm

### 8.1 Widow / Orphan Control (`<w:widowControl/>`)
When breaking a paragraph $P$ of $M$ lines across Page $N$ and Page $N+1$:
- If $M = 2$ and paragraph overflows Page $N$, move **all 2 lines** to Page $N+1$.
- If $M \ge 3$:
  - Ensure at least 2 lines remain on Page $N$ (Orphan constraint).
  - Ensure at least 2 lines are pushed to Page $N+1$ (Widow constraint).
  - If Page $N$ only has room for 1 line, move 1 extra line down to Page $N+1$ so Page $N$ holds 0 lines (entire paragraph moves) or Page $N$ holds 2 lines if preceding content allows.

### 8.2 Keep With Next (`<w:keepNext/>`)
If paragraph $P_k$ has `<w:keepNext/>`, check if $P_{k+1}$ fits on Page $N$. If $P_{k+1}$ is pushed to Page $N+1$, $P_k$ must also be relocated to Page $N+1$.

---

## 9. Comprehensive Native Canvas Rendering Pipeline Engine

```typescript
export interface ParagraphLayoutInput {
  text: string;
  fontSizePt: number;
  fontFamily: string;
  lineRule: 'auto' | 'exact' | 'atLeast';
  lineVal: number; // e.g. 240 for auto 1.0x, or twips for exact/atLeast
  spacingBeforeTwips: number;
  spacingAfterTwips: number;
  contextualSpacing: boolean;
  styleId: string;
}

export interface RenderedLine {
  text: string;
  widthPx: number;
  heightPx: number;
  ascentPx: number;
}

/**
 * Calculates paragraph line boxes and exact spacing geometry for native canvas rendering.
 */
export function layoutParagraph(
  ctx: CanvasRenderingContext2D,
  input: ParagraphLayoutInput,
  maxAvailableWidthPx: number
): { lines: RenderedLine[]; totalHeightPx: number } {
  const fontSizePx = (input.fontSizePt * 96) / 72;
  ctx.font = `${fontSizePx}px ${input.fontFamily}`;

  // Approximate Typographic Ascent (0.8 * fontSizePx) and Descent (0.2 * fontSizePx) if raw font metrics unavailable
  const approxAscent = fontSizePx * 0.8;
  const approxNaturalHeight = fontSizePx * 1.2;

  let lineMultiplier = 1.0;
  let computedLineHeightPx = approxNaturalHeight;

  if (input.lineRule === 'auto') {
    lineMultiplier = input.lineVal / 240;
    computedLineHeightPx = approxNaturalHeight * lineMultiplier;
  } else if (input.lineRule === 'exact') {
    computedLineHeightPx = twipsToPx(input.lineVal);
  } else if (input.lineRule === 'atLeast') {
    computedLineHeightPx = Math.max(approxNaturalHeight, twipsToPx(input.lineVal));
  }

  // Word Wrapping Logic
  const words = input.text.split(' ');
  const lines: RenderedLine[] = [];
  let currentLine = '';

  for (const word of words) {
    const testLine = currentLine ? `${currentLine} ${word}` : word;
    const metrics = ctx.measureText(testLine);
    if (metrics.width > maxAvailableWidthPx && currentLine !== '') {
      lines.push({
        text: currentLine,
        widthPx: ctx.measureText(currentLine).width,
        heightPx: computedLineHeightPx,
        ascentPx: approxAscent,
      });
      currentLine = word;
    } else {
      currentLine = testLine;
    }
  }
  if (currentLine) {
    lines.push({
      text: currentLine,
      widthPx: ctx.measureText(currentLine).width,
      heightPx: computedLineHeightPx,
      ascentPx: approxAscent,
    });
  }

  const totalLinesHeight = lines.length * computedLineHeightPx;
  return { lines, totalHeightPx: totalLinesHeight };
}
```
