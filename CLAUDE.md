# Obra Autodocs

Figma plugin that auto-generates labeled grid documentation for component variant sets.

## Tech Stack

- Vanilla ES5 JavaScript — no build step, no TypeScript
- `code.js` — plugin sandbox logic (all layout, label creation, variant analysis)
- `ui.html` — plugin UI panel (iframe, communicates via postMessage)
- `manifest.json` — Figma plugin config

## How It Works

1. User selects a component set (or a variant inside one)
2. Plugin analyzes variant positions → clusters into grid (rows/columns)
3. Classifies each property as column-axis (top labels) or row-axis (left labels)
4. Creates a wrapper frame `"❖ ComponentName"` containing the component set + label frames
5. Labels use `"PropertyName: Value"` format, with brackets for outer (grouped) properties

## Groupify (Grid Arrangement Tool)

When a component set's variants aren't already in a uniform grid, the **Groupify** tab helps the user arrange them into one. It:

1. Reads all variant properties and their values
2. Lets the user assign each property to ROW, COLUMN, or SECTION axis
3. Builds the full cartesian product grid and positions variants accordingly

**Empty cells:** Some component sets have non-existent variant combinations (e.g., a Nav Item where Active=True only has State=Default, not Hover/Focus). This is expected — it shows either intentional omissions or flaws in the component properties. The grid must properly account for these empty cells with correct dimensions, so that autodocs can later generate correct labels and grid lines over the groupified result.

## Key Data Structures

- `colClusters` / `rowClusters` — left/top edge positions of each column/row
- `colWidths` / `rowHeights` — width/height of each column/row
- These are used for label positioning, grid line placement, and wrapper sizing
- **Important:** When the uniform cell grid kicks in, these represent **cell extents** (not variant bounds)

## Uniform Cell Grid (Small Variant Handling)

When label text is wider than the variant itself (e.g. 16px radio button with "State: Error Focus" label at 93px), the plugin:

1. Computes `cellWidth = max(originalPitch, maxLabelWidth)`
2. Repositions all variants centered in uniform cells
3. Updates `colClusters`/`colWidths` to cell extents
4. Resizes the component set

This ensures labels span full cell width, edge cells match middle cells, and grid lines align with cell boundaries. Triggered only when `maxLabelWidth > maxVariantWidth`.

## Debug Tools

Two debug buttons in the UI log to the Figma console:

- **"Log grid coords"** → `debugGridCoords()` — variant positions, column/row clusters, grid line midpoints
- **"Log label coords"** → `debugLabelCoords()` — cell boundary analysis with off-center detection (`⚠ OFF-CENTER`), label frame positions and text placement

## Constants (top of code.js)

| Constant | Value | Purpose |
|----------|-------|---------|
| `BRACKET_THICKNESS` | 14 | Bracket frame width/height |
| `BRACKET_CAP_LENGTH` | 7 | Short end lines on brackets |
| `BRACKET_LABEL_ROW_HEIGHT` | 42 | Label row with bracket |
| `SIMPLE_LABEL_ROW_HEIGHT` | 25 | Label row without bracket |
| `LABEL_FONT_SIZE` | 11 | Inter Regular |
| `LABEL_GAP` | 8 | Space between text and bracket |
| `CLUSTER_TOLERANCE` | 5 | Pixel tolerance for grid clustering |
| `BOOL_ACCURACY_MAX_COMBINATIONS` | 1000 | Auto-disable accuracy above this threshold |

## Boolean Grid Mode

When boolean visibility is set to "Grid" display mode, boolean properties become an additional axis in the main documentation grid — rather than a separate section below it. The component set itself serves as the "default" boolean state (no instances needed), and non-default boolean combinations create instances with a subtle background tint.

```
                    [Checked: On]      [Checked: Off]
                    ┌─────────────────┬─────────────────┐
[Show line 2:       │ (component set  │ (component set  │
 False]             │  lives here)    │  lives here)    │
                    ├─ ─ ─ ─ ─ ─ ─ ─ ┼─ ─ ─ ─ ─ ─ ─ ─ ┤
[Show line 2:       │░░ instance ░░░░░│░░ instance ░░░░░│
 True]              │░░ (tinted bg) ░░│░░ (tinted bg) ░░│
                    └─────────────────┴─────────────────┘
```

**Axis auto-detection:** If the component set is wider than tall, boolean combos go on the row axis (expand vertically). Otherwise they go on the column axis (expand horizontally). The boolean property is always the outermost property on its axis, with brackets when sharing the axis with variant properties.

**Key design decisions:**
- The CS occupies the "default" group, labeled "Base" — no redundant instances
- Non-default groups get a background tint (`DOC_COLOR` at 3% opacity) for visual distinction
- `BOOL_GROUP_GAP` (24px) separates boolean groups visually
- **Nested boolean groups** (multiple bool props): double-line dividers at both edges of the gap for visual weight
- **Simple single-boolean** components: single-line divider only — double lines look heavy for simple cases like checkboxes
- Inner variant labels are repeated for each boolean group
- Grid lines extend across the full expanded grid
- The Documentation frame is **not locked** — users need to reach in and grab components for inspection

**Improved accuracy mode** (enabled by default):
Boolean toggles often change instance sizes (e.g. "Show description: True" makes a 23px header 47px). Without measurement, instances are placed at original variant positions and can overflow their grid cells.

The solution measures the real instances created by `createBooleanInstanceGroup` to compute adjusted cell layouts — rows and columns expand to fit the largest instance in each position. Instances are then centered within their adjusted cells. Grid lines and labels also use the adjusted per-group layout.

**Performance safeguard:** Improved accuracy is auto-disabled when `groups * variants > BOOL_ACCURACY_MAX_COMBINATIONS` (default 1000). This protects users from very slow regeneration on large component sets with many boolean combinations. The threshold is a constant at the top of `code.js` that can be adjusted.

## Overflow Detection & Accommodation (Edge Case)

Some variants contain absolutely positioned children that visually extend beyond the variant's frame bounds (e.g., dropdown menus, tooltips, popovers). In Figma, `layoutPositioning = 'ABSOLUTE'` on a child means it's positioned outside normal flow.

**Example:** An Autocomplete component with 4 variants — 2 have a dropdown menu (56px trigger + 124px menu = 180px visual height), 2 are just 56px. The grid uses per-variant visual sizes: some rows are 180px, others are 56px.

**Detection:** `detectVariantOverflow(variant)` scans children (up to depth 3) for absolute-positioned nodes extending beyond the variant frame. `detectComponentSetOverflow(cs)` aggregates across all variants.

**Accommodation:** When overflow is detected during generation:
1. A `_visualSizeOverrides` map is built with per-variant visual sizes (frame + overflow)
2. `effectiveSize()` checks this map, so `clusterVariantAxis()` uses visual sizes for bounding boxes
3. `colWidths`/`rowHeights` automatically reflect the visual extents
4. The designer's original variant positions are preserved — no repositioning
5. After `analyzeLayout()` completes, overrides are cleared

**UI warning:** The Organize (Groupify) tab shows a warning banner and disables Apply/Auto-align buttons when overflow is detected, since rearranging variants with overflow requires manual positioning. A `figma.notify()` toast also appears during generation.

## Standalone Documentation Mode

When "Standalone documentation" is checked, the plugin generates docs as a completely separate frame — the component set is never moved, modified, or touched in any way.

**How it works:**
1. Layout analysis runs on the CS as-is (same clustering, axis detection, label calculations)
2. A wrapper frame `"❖ ComponentName"` is created **below** the CS (48px gap), marked with `pluginData('standaloneDoc', 'true')`
3. For each variant, `variant.createInstance()` creates an instance positioned in the documented grid layout
4. Uniform cell grid adjustments are computed and applied to instance positions (not to original variants)
5. Labels, grid lines, brackets, boolean grid sections, and extra sections are created inside the wrapper

**Key differences from regular mode:**
- CS stays in its original parent — not reparented into the wrapper
- No auto-layout disabling, no variant repositioning, no CS resize
- No strokes added to the CS — a `'Grid Border'` frame provides the dashed border
- Instances automatically update when source components change
- Removal is simple: just delete the wrapper (no CS restoration needed)

**Wrapper discovery:** `findStandaloneWrapper(node)` searches siblings for a frame with `pluginData('standaloneDoc') === 'true'`. This distinguishes standalone wrappers from regular wrappers (which are parents of the CS, found by `findExistingWrapper()`).

## Mixed-Size Variant Clustering

When a component set contains variants of very different sizes (e.g. 40px icon buttons alongside 120px text buttons), the clustering algorithm uses three strategies to correctly identify columns/rows:

1. **Chain-based clustering** (`clusterValues`): Compares new values against the *highest* value in the current cluster, not the first. This prevents splitting when center values gradually drift across a range wider than `CLUSTER_TOLERANCE` (e.g. centers 474→477→480.5→485 — each consecutive gap ≤5px but total span is 11px).

2. **Bounding-box overlap merge** (`clusterVariantAxis`): After initial clustering, adjacent clusters whose variant bounding boxes physically overlap are merged. This catches cases where even chain-based clustering can't bridge the gap (e.g. centers 669 vs 675.5 — gap 6.5px — but bounding boxes [608,729] and [641,710] clearly overlap).

3. **Per-cluster tolerances** (`tols` array): Merged clusters return per-cluster tolerances so `findClusterIndex()` can match all member variants, even when their reference values span wider than the default `CLUSTER_TOLERANCE=5`.

4. **Range-based instance matching**: Boolean accuracy code uses `pos >= min - tol && pos <= min + width + tol` instead of `|pos - min| < tol` to correctly match instances to wide clusters.

## Known Limitations

- **Boolean × Nested Instance interdependency:** Boolean visibility and nested instances are generated as independent sections. When a boolean property gates the visibility of a nested instance (e.g. Apple's Alert component where `showFields` toggles Text Field instances), the combined grid doesn't show every meaningful combination. The nested instances section forces all booleans to `true` as a workaround, but doesn't cross-reference boolean states with nested instance variants.
- **Groupify optimizer capped at 8 properties:** `groupifyOptimizeDirections` uses brute-force 2^N bitmask enumeration. For component sets with 9+ properties (e.g. Windows UI Kit ComboBox with 10 props), the optimizer is skipped and position-based inference is used instead.

## Key Functions

- `analyzeLayout(cs, enumProps)` — clusters variants into grid, determines axis properties
- `generate(options)` — main generation flow (async, needs font loading)
- `createColLabel()` / `createRowLabel()` — create label frames with text + optional bracket
- `getColGroups()` / `getRowGroups()` — group contiguous columns/rows by property value
- `measureTextWidth(text, fontSize)` — creates temporary text node to measure width
- `clusterVariantAxis(variants, posKey, sizeKey)` — alignment-aware clustering with chain merge and overlap detection
