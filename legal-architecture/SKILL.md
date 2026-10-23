---
name: legal-architecture
description: 生成專業的法律結構可視化圖，輸出為自包含 HTML 文件（內嵌 SVG，淺色主題，適合列印和嵌入文件）。適用場景：(1) 用戶要求"畫結構圖""生成可視化""做流程圖""法律圖示"時；(2) 涉及訴訟推理結構（當事人→證據→事實→法律→判決）；(3) 契約關係圖（甲乙方權利義務、履約節點、違約責任）；(4) 法規程序圖（聲請→審查→決定→救濟）；(5) 證據鏈圖、請求權基礎分析圖、法律體系層級圖。Use when the user asks for litigation flow diagrams, contract relationship diagrams, legal reasoning structure diagrams, evidence chain diagrams, or any legal visualization.
license: MIT
metadata:
  version: "1.1"
  author: JeeC (adapted from architecture-diagram by Cocoon AI)
---

# Legal Diagram Skill

Create professional legal structure diagrams as self-contained HTML files with inline SVG. Designed for court document analysis, contract review, legal knowledge management, and client presentations. Output is light-themed and print-ready.

## Design System

### Color Palette

Use these semantic colors for legal component types:

| Component Type | Fill (rgba) | Stroke | Use Case |
|---|---|---|---|
| 當事人/主體 (Party) | `rgba(235,245,255,0.95)` | `#1e40af` (法院藍) | 原告、被告、甲方、乙方、第三人 |
| 事實/行為 (Fact) | `rgba(248,250,252,0.95)` | `#475569` (石板灰) | 侵權行為、違約事實、案件經過 |
| 證據 (Evidence) | `rgba(255,251,235,0.95)` | `#a16207` (古卷黃) | 書證、物證、證人證言、鑑定意見 |
| 法律規範 (Law) | `rgba(240,253,244,0.95)` | `#15803d` (法條綠) | 法條、司法解釋、行政法規 |
| 程序節點 (Procedure) | `rgba(245,243,255,0.95)` | `#7c3aed` (程序紫) | 起訴、開庭、詰問、裁定 |
| 裁判/結論 (Judgment) | `rgba(254,242,242,0.95)` | `#991b1b` (責任紅) | 判決主文、責任認定、風險結論 |
| 爭議焦點 (Issue) | `rgba(255,247,237,0.95)` | `#c2410c` (焦點橙) | 爭議焦點、待查明事實 |

### Typography

為了與法律書狀套件保持視覺連貫，圖表應優先使用「標楷體」處理中文，西文則維持 Inter：
```html
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  body { font-family: 'Inter', '標楷體', 'KaiTi', 'BiauKai', serif; }
</style>
```

Font sizes: 13px (Title), 11px (Primary), 10px (Secondary), 9px (Annotations).

### Visual Elements

**Background:** `#f8fafc` with subtle light grid:
```svg
<pattern id="grid" width="40" height="40" patternUnits="userSpaceOnUse">
  <path d="M 40 0 L 0 0 0 40" fill="none" stroke="#f1f5f9" stroke-width="0.8"/>
</pattern>
```

**Component boxes:** Rounded rectangles (`rx="6"`) with 1.5px stroke.
Always draw an opaque white background rect FIRST to mask arrows behind boxes:
```svg
<rect x="X" y="Y" width="W" height="H" rx="6" fill="white"/>
<rect x="X" y="Y" width="W" height="H" rx="6" fill="rgba(219,234,254,0.9)" stroke="#2563eb" stroke-width="1.5"/>
```

**Arrows — standard flow:**
```svg
<marker id="arr" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto">
  <polygon points="0 0, 8 3, 0 6" fill="#94a3b8"/>
</marker>
<line x1="..." y1="..." x2="..." y2="..." stroke="#94a3b8" stroke-width="1.2" marker-end="url(#arr)"/>
```

**Arrows — judgment emphasis (red):**
```svg
<marker id="arr-red" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto">
  <polygon points="0 0, 8 3, 0 6" fill="#b91c1c"/>
</marker>
```

**Cross-layer or indirect logical connections:** Use dashed lines:
```svg
stroke-dasharray="4,3"
```

**Arrow z-order:** ALWAYS draw all arrows BEFORE component boxes in the SVG. SVG renders in document order, so arrows drawn first appear behind boxes drawn later.

### Layout Patterns

Choose the layout based on content type:

**1. 訴訟流程圖 (Litigation Flow)** — top-to-bottom layered with group boxes

For: case analysis, judgment structure, evidence → fact → law → judgment chain.

**PREFERRED PATTERN — dashed group boxes + single vertical arrows:**

Each reasoning stage is enclosed in a dashed group box with a badge label straddling its top-left border. Stages connect via a single purely vertical arrow in the center (x=500 for a 1000-wide viewBox). No node-to-node arrows across stages — this avoids visual clutter and false logical implications.

```
┌─────────────────────────────────┐
│ [當事人]  原告 ●●● 被告          │  ← group box, dashed border
└─────────────────────────────────┘
              ↓  (single arrow at center x)
┌─────────────────────────────────┐
│ [爭議焦點] 焦點一  焦點二         │
└─────────────────────────────────┘
              ↓
  ... etc through 證據 → 事實認定 → 法律適用 → 判決結果
```

Group box SVG pattern (badge label straddles top-left border):
```svg
<!-- Group box -->
<rect x="8" y="127" width="984" height="82" rx="8"
      fill="rgba(255,237,213,0.12)" stroke="#fb923c"
      stroke-width="1.2" stroke-dasharray="6,4"/>
<!-- Badge: rect behind text, centered on top border (y = group top - 8) -->
<rect x="14" y="119" width="64" height="16" rx="4"
      fill="rgba(255,237,213,0.95)" stroke="#fb923c" stroke-width="0.8"/>
<text x="46" y="131" fill="#9a3412" font-size="9" font-weight="600"
      text-anchor="middle" font-family="Inter,system-ui">爭議焦點</text>
```

Vertical gap between group boxes: 28–32px. Place one arrow per gap:
```svg
<line x1="500" y1="[group_bottom+2]" x2="500" y2="[next_group_top-2]"
      stroke="#64748b" stroke-width="1.5" marker-end="url(#arr)"/>
<!-- Final arrow to judgment box uses red: stroke="#b91c1c" stroke-width="2" marker-end="url(#arr-red)" -->
```

Within each group box, items are arranged horizontally (no intra-group arrows needed — spatial proximity implies relationship). Reserve arrows for cross-group logical flow only.

**2. 契約關係圖 (Contract Relationship)** — hub-and-spoke or two-column
For: contract parties, rights/obligations flow, key clause mapping
```
甲方 ←→ 契約核心義務 ←→ 乙方
         ↓
    履約條件/違約責任
```

**3. 法律體系圖 (Legal Framework)** — tree or hierarchical
For: legal concepts, regulatory hierarchy, claim basis analysis (請求權基礎)
```
上位法 → 下位法 → 施行細則/司法解釋
```

### Spacing Rules

**CRITICAL:** Avoid element overlaps:
- Standard component height: 55-60px for single-line, 75-90px for multi-line
- Minimum vertical gap between rows: 38px
- Minimum horizontal gap between same-row items: 15px
- Use equal start/end margins: startX = (viewBoxWidth - totalRowWidth) / 2

**Row width calculation:**
```
totalRowWidth = N * boxWidth + (N-1) * gap
startX = (viewBoxWidth - totalRowWidth) / 2
```

### Legend Placement

Place legend BELOW all diagram content, at least 20px below the lowest element. Use a light rounded-rectangle background. Arrange legend items horizontally.

### Page Structure

1. **Header** — title (with colored status dot), subtitle (case number / date / court)
2. **Diagram card** — white card with border-shadow, contains SVG
3. **Info cards** — 3-column grid with metadata (parties, evidence summary, judgment key points)
4. **Footer** — case number and skill attribution

## 工作流程 (Workflow)

> [!IMPORTANT]
> 絕對禁止直接拋出大量 SVG 或 HTML 原始碼。必須嚴格遵循以下三步驟：

### Step 1: 結構梳理與草圖確認 (Checkpoint)
收到需求後，AI 必須**先用純文字（如 ASCII 樹狀圖或列點）**呈現預計繪製的節點清單、分層佈局與關聯狀態，並**暫停執行，等待使用者確認結構是否正確**。

### Step 2: 座標試算與文字防溢出處理 (Edge Case Handling)
獲得確認後開始計算 SVG 座標。必須強制套用下述換行規則：
- 單一區塊內文字若超過 15 個中文字，**強制拆分為多個 `<tspan x=".." dy="..">` 斷行顯示**，或給予合理截斷，絕不可讓文字溢出外框。
- 需預留足夠高度，隨字串行數動態延展 Component box 與 Group box 的 `height`。

### Step 3: 寫入 HTML 檔案並交付
將完整包裝好的字串使用 `write_to_file` 等工具，直接儲存為 `.html` 檔案到使用者的工作目錄（Workspace），隨後提供檔案路徑與開啟建議給使用者即可，不需將原始碼印於對話視窗。

## Output Format

Always produce a single self-contained `.html` file:
- Embedded CSS only (Google Fonts link allowed)
- Inline SVG (no external images or scripts)
- Light background (`#f8fafc`) — suitable for printing and embedding in Word/PPT
- No JavaScript required

File should open and render correctly in any modern browser, and look professional when printed.
