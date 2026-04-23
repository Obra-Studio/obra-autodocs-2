// Obra Autodocs — Figma Plugin
// Generates labeled grid documentation for component variant sets

// ─── Constants ───────────────────────────────────────────────────────────────

const BRACKET_THICKNESS = 14;
const BRACKET_CAP_LENGTH = 7;
const BRACKET_LABEL_ROW_HEIGHT = 42;
const SIMPLE_LABEL_ROW_HEIGHT = 25;
const LABEL_PADDING = 10;
const LABEL_FONT_SIZE = 11;
const DEFAULT_COLOR = { r: 0.592, g: 0.278, b: 1.0 };     // #9747FF
const BRACKET_STROKE_WEIGHT = 1;
const CLUSTER_TOLERANCE = 5;
const LABEL_GAP = 8;
const GRID_STROKE_WEIGHT = 1;
const GRID_DASH_PATTERN = [4, 4];
const BOOL_ACCURACY_MAX_COMBINATIONS = 1000;
const ITEM_LABEL_GAP = 6;      // gap between variant bottom and per-item label
const ITEM_LABEL_HEIGHT = 14;  // reserved height for per-item label text inside cell
const TITLE_FONT_SIZE = 18;
const TITLE_GAP = 16;          // gap between title and documentation frame
const DESC_GAP = 16;
const DOC_LINK_FONT_SIZE = 11;
const DOC_LINK_GAP = 8;
const DESC_PADDING = 12;
const DESC_FONT_SIZE = 12;
const DESC_LINE_HEIGHT = 1.5;

var DEBUG = false;

// Mutable doc color — set per generation
var DOC_COLOR = { r: DEFAULT_COLOR.r, g: DEFAULT_COLOR.g, b: DEFAULT_COLOR.b };

// Font cache — load eagerly at startup to avoid blocking generation
var _fontLoaded = false;
var _fontLoadPromise = null;

// Reusable text measurement node + cache
var _measureNode = null;
var _measureCache = {};

// Per-variant visual size overrides (for absolute-positioned overflow accommodation)
// When set, effectiveSize() returns the visual size (frame + overflow) instead of frame size.
var _visualSizeOverrides = null; // { [variantId]: { width, height } }
var _selectionGeneration = 0; // guards against stale async sendSelectionInfo results
var _selectionDebounceTimer = null;
var _hidePropertyNames = false; // When true, labels show only values (no "PropName: " prefix)
var _autoLayoutExtras = false; // When true, Property Combinations uses auto layout instead of manual positioning

// ─── Per-Component Settings Persistence ─────────────────────────────────────

function saveComponentSettings(cs, options) {
  if (!cs) return;
  try {
    var settings = {
      showGrid: options.showGrid || false,
      color: options.color || '#9747FF',
      showBooleanVisibility: options.showBooleanVisibility || false,
      booleanCombination: options.booleanCombination || 'individual',
      booleanDisplayMode: options.booleanDisplayMode || 'list',
      enabledBooleanProps: options.enabledBooleanProps || null,
      booleanScope: options.booleanScope || null,
      showNestedInstances: options.showNestedInstances || false,
      nestedInstancesMode: options.nestedInstancesMode || 'representative',
      enabledNestedInstances: options.enabledNestedInstances || null,
      standaloneDoc: options.standaloneDoc || false,
      hidePropertyNames: options.hidePropertyNames || false,
      showTitle: options.showTitle || false,
      showDescription: options.showDescription || false,
      showDocLink: options.showDocLink || false,
      allowSpanning: options.allowSpanning || false,
      autoLayoutExtras: options.autoLayoutExtras || false,
      variableModes: options.variableModes || null,
      docMode: options.docMode || 'base',
    };
    cs.setPluginData('lastSettings', JSON.stringify(settings));
    if (DEBUG) console.log('[Settings] Saved for', cs.name, JSON.stringify(settings));
  } catch (e) {
    console.warn('[Settings] Failed to save:', e.message);
  }
}

function getComponentSettings(cs) {
  if (!cs) return null;
  try {
    var data = cs.getPluginData('lastSettings');
    if (!data) return null;
    return JSON.parse(data);
  } catch (e) {
    console.warn('[Settings] Failed to read:', e.message);
    return null;
  }
}

// ─── Shadcn Style Filter ────────────────────────────────────────────────────
// Hide or delete variants by a "Shadcn style" variant property value. Hiding
// sets visible=false on matching variants, stashes them below the active grid,
// and regenerates docs; the filtered variants are skipped by analyzeLayout and
// all doc builders.

var SHADCN_STYLE_PROP_NAMES = ['shadcn style'];

function findShadcnStylePropKey(cs) {
  if (!cs) return null;
  try {
    var groupProps = cs.variantGroupProperties;
    if (!groupProps) return null;
    for (var key in groupProps) {
      if (!groupProps.hasOwnProperty(key)) continue;
      if (SHADCN_STYLE_PROP_NAMES.indexOf(key.toLowerCase()) !== -1) return key;
    }
  } catch (e) {}
  return null;
}

function getHiddenStyles(cs) {
  if (!cs) return [];
  try {
    var raw = cs.getPluginData('hiddenStyles');
    if (!raw) return [];
    var parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed;
  } catch (e) { return []; }
}

function setHiddenStyles(cs, arr) {
  if (!cs) return;
  try { cs.setPluginData('hiddenStyles', JSON.stringify(arr || [])); } catch (e) {}
}

function getShadcnStyleInfo(cs) {
  var propKey = findShadcnStylePropKey(cs);
  if (!propKey) return null;
  var variants = cs.children.filter(function(c) { return c.type === 'COMPONENT'; });
  var counts = {};
  var order = [];
  for (var i = 0; i < variants.length; i++) {
    var parsed = dvParseVariantName(variants[i].name);
    var val = parsed[propKey];
    if (val === undefined) continue;
    if (counts[val] === undefined) { counts[val] = 0; order.push(val); }
    counts[val]++;
  }
  if (order.length === 0) return null;
  var hiddenList = getHiddenStyles(cs);
  // Prune stale hidden entries (values that no longer exist)
  var pruned = hiddenList.filter(function(v) { return counts[v] !== undefined; });
  if (pruned.length !== hiddenList.length) {
    setHiddenStyles(cs, pruned);
    hiddenList = pruned;
  }
  return {
    propName: propKey,
    values: order.map(function(v) {
      return { name: v, count: counts[v], hidden: hiddenList.indexOf(v) !== -1 };
    }),
  };
}

// Active variants = COMPONENT children that are visible. Hidden-by-style
// variants have visible=false so they're skipped during layout + generation
// but remain as members of the CS (so Figma keeps the variant property alive).
function getActiveVariants(cs) {
  return cs.children.filter(function(c) {
    return c.type === 'COMPONENT' && c.visible !== false;
  });
}

async function applyHiddenStyles(hiddenValues, regenOptions) {
  var cs = getComponentSet();
  if (!cs) { figma.notify('Select a component set first'); return; }
  var propKey = findShadcnStylePropKey(cs);
  if (!propKey) { figma.notify('No "Shadcn style" property on this component set'); return; }

  var hiddenSet = {};
  for (var i = 0; i < hiddenValues.length; i++) hiddenSet[hiddenValues[i]] = true;

  var allVariants = cs.children.filter(function(c) { return c.type === 'COMPONENT'; });
  var wouldHide = 0;
  for (var v = 0; v < allVariants.length; v++) {
    var parsed = dvParseVariantName(allVariants[v].name);
    if (parsed[propKey] !== undefined && hiddenSet[parsed[propKey]]) wouldHide++;
  }
  if (wouldHide === allVariants.length) {
    figma.notify('Cannot hide every style — at least one must stay visible');
    return;
  }

  // Apply visibility
  for (var j = 0; j < allVariants.length; j++) {
    var p = dvParseVariantName(allVariants[j].name);
    var styleVal = p[propKey];
    var shouldHide = styleVal !== undefined && hiddenSet[styleVal];
    if (allVariants[j].visible !== !shouldHide) {
      allVariants[j].visible = !shouldHide;
    }
  }

  setHiddenStyles(cs, hiddenValues);

  // Regenerate docs if there's an existing wrapper
  var wrapper = findExistingWrapper(cs);
  if (wrapper && regenOptions) {
    await generate(regenOptions);
  } else {
    sendSelectionInfo();
  }
}

async function deleteStyleVariants(styleValue, regenOptions) {
  var cs = getComponentSet();
  if (!cs) { figma.notify('Select a component set first'); return; }
  var propKey = findShadcnStylePropKey(cs);
  if (!propKey) { figma.notify('No "Shadcn style" property on this component set'); return; }

  var allVariants = cs.children.filter(function(c) { return c.type === 'COMPONENT'; });
  var toDelete = [];
  var wouldKeep = 0;
  for (var i = 0; i < allVariants.length; i++) {
    var parsed = dvParseVariantName(allVariants[i].name);
    if (parsed[propKey] === styleValue) toDelete.push(allVariants[i]);
    else wouldKeep++;
  }
  if (toDelete.length === 0) { figma.notify('No variants match style "' + styleValue + '"'); return; }
  if (wouldKeep === 0) { figma.notify('Cannot delete — this would empty the component set'); return; }

  for (var d = 0; d < toDelete.length; d++) {
    try { toDelete[d].remove(); } catch (e) { console.warn('[Delete] Failed on', toDelete[d].name, e.message); }
  }

  // Prune from hidden list
  var hiddenList = getHiddenStyles(cs).filter(function(v) { return v !== styleValue; });
  setHiddenStyles(cs, hiddenList);

  figma.notify('Deleted ' + toDelete.length + ' variant' + (toDelete.length === 1 ? '' : 's'));

  var wrapper = findExistingWrapper(cs);
  if (wrapper && regenOptions) {
    await generate(regenOptions);
  } else {
    sendSelectionInfo();
  }
}

function hexToRgb(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  return { r, g, b };
}

function formatLabel(propName, value) {
  return _hidePropertyNames ? value : propName + ': ' + value;
}

function createTitleText(name) {
  var titleText = figma.createText();
  titleText.name = 'Title';
  titleText.characters = name;
  titleText.fontSize = TITLE_FONT_SIZE;
  titleText.fontName = { family: 'Inter', style: 'Semi Bold' };
  titleText.fills = [{ type: 'SOLID', color: DOC_COLOR }];
  titleText.constraints = { horizontal: 'MIN', vertical: 'MIN' };
  return titleText;
}

// ─── Markdown-to-Figma rich text ─────────────────────────────────────────────

/**
 * Parse a Figma component description (markdown-like) into plain text + style ranges.
 * Supports: **bold**, *italic*, ***bold italic***, `inline code`, [link](url),
 *           - bullet lists, 1. numbered lists
 * Returns { text: string, ranges: [{ start, end, bold, italic, code, link }] }
 */
function parseDescriptionMarkdown(raw) {
  var lines = raw.replace(/\r\n/g, '\n').split('\n');
  var plainParts = [];
  var ranges = [];
  var pos = 0;

  for (var li = 0; li < lines.length; li++) {
    var line = lines[li];
    if (li > 0) { plainParts.push('\n'); pos += 1; }

    // Bullet list: "- text" or "• text" or "* text" (but not bold **)
    var bulletMatch = line.match(/^(\s*)[-•*](?!\*)\s+(.*)/);
    if (bulletMatch) {
      var indent = bulletMatch[1] || '';
      plainParts.push(indent + '• ');
      pos += indent.length + 2;
      line = bulletMatch[2];
    }

    // Numbered list: "1. text"
    var numMatch = line.match(/^(\s*)(\d+)\.\s+(.*)/);
    if (!bulletMatch && numMatch) {
      var nIndent = numMatch[1] || '';
      var num = numMatch[2];
      plainParts.push(nIndent + num + '. ');
      pos += nIndent.length + num.length + 2;
      line = numMatch[3];
    }

    // Parse inline styles within the line
    // Order matters: bold italic (***), bold (**), italic (*), code (`), links
    var i = 0;
    while (i < line.length) {
      // Bold italic: ***text***
      if (line[i] === '*' && line[i+1] === '*' && line[i+2] === '*') {
        var closeBI = line.indexOf('***', i + 3);
        if (closeBI !== -1) {
          var biText = line.substring(i + 3, closeBI);
          ranges.push({ start: pos, end: pos + biText.length, bold: true, italic: true });
          plainParts.push(biText);
          pos += biText.length;
          i = closeBI + 3;
          continue;
        }
      }
      // Bold: **text**
      if (line[i] === '*' && line[i+1] === '*') {
        var closeB = line.indexOf('**', i + 2);
        if (closeB !== -1) {
          var bText = line.substring(i + 2, closeB);
          ranges.push({ start: pos, end: pos + bText.length, bold: true });
          plainParts.push(bText);
          pos += bText.length;
          i = closeB + 2;
          continue;
        }
      }
      // Italic: *text*
      if (line[i] === '*' && line[i+1] !== '*') {
        var closeI = line.indexOf('*', i + 1);
        if (closeI !== -1) {
          var iText = line.substring(i + 1, closeI);
          ranges.push({ start: pos, end: pos + iText.length, italic: true });
          plainParts.push(iText);
          pos += iText.length;
          i = closeI + 1;
          continue;
        }
      }
      // Inline code: `text`
      if (line[i] === '`') {
        var closeC = line.indexOf('`', i + 1);
        if (closeC !== -1) {
          var cText = line.substring(i + 1, closeC);
          ranges.push({ start: pos, end: pos + cText.length, code: true });
          plainParts.push(cText);
          pos += cText.length;
          i = closeC + 1;
          continue;
        }
      }
      // Link: [text](url)
      if (line[i] === '[') {
        var closeBracket = line.indexOf(']', i + 1);
        if (closeBracket !== -1 && line[closeBracket + 1] === '(') {
          var closeParen = line.indexOf(')', closeBracket + 2);
          if (closeParen !== -1) {
            var linkLabel = line.substring(i + 1, closeBracket);
            var linkUrl = line.substring(closeBracket + 2, closeParen);
            ranges.push({ start: pos, end: pos + linkLabel.length, link: linkUrl });
            plainParts.push(linkLabel);
            pos += linkLabel.length;
            i = closeParen + 1;
            continue;
          }
        }
      }
      // Plain character
      plainParts.push(line[i]);
      pos += 1;
      i += 1;
    }
  }

  return { text: plainParts.join(''), ranges: ranges };
}

/**
 * Apply parsed style ranges to a Figma text node.
 * Must be called after setting .characters and base font/size.
 */
function applyRichTextStyles(textNode, ranges) {
  for (var ri = 0; ri < ranges.length; ri++) {
    var r = ranges[ri];
    if (r.start >= r.end) continue;

    if (r.bold && r.italic) {
      textNode.setRangeFontName(r.start, r.end, { family: 'Inter', style: 'Bold Italic' });
    } else if (r.bold) {
      textNode.setRangeFontName(r.start, r.end, { family: 'Inter', style: 'Bold' });
    } else if (r.italic) {
      textNode.setRangeFontName(r.start, r.end, { family: 'Inter', style: 'Italic' });
    }

    if (r.code) {
      textNode.setRangeFontName(r.start, r.end, { family: 'Inter', style: 'Regular' });
      textNode.setRangeFills(r.start, r.end, [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.8 }]);
    }

    if (r.link) {
      textNode.setRangeHyperlink(r.start, r.end, { type: 'URL', value: r.link });
      textNode.setRangeTextDecoration(r.start, r.end, 'UNDERLINE');
    }
  }
}

function createDescriptionFrame(description, maxWidth) {
  var descFrame = figma.createFrame();
  descFrame.name = 'Description';
  descFrame.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.04 }];
  descFrame.cornerRadius = 4;
  descFrame.clipsContent = false;
  descFrame.layoutMode = 'VERTICAL';
  descFrame.paddingTop = DESC_PADDING;
  descFrame.paddingBottom = DESC_PADDING;
  descFrame.paddingLeft = DESC_PADDING;
  descFrame.paddingRight = DESC_PADDING;
  descFrame.primaryAxisSizingMode = 'AUTO';
  descFrame.counterAxisSizingMode = 'FIXED';

  var parsed = parseDescriptionMarkdown(description.trim());

  var descText = figma.createText();
  descText.characters = parsed.text;
  descText.fontSize = DESC_FONT_SIZE;
  descText.lineHeight = { value: DESC_FONT_SIZE * DESC_LINE_HEIGHT, unit: 'PIXELS' };
  descText.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.6 }];

  applyRichTextStyles(descText, parsed.ranges);

  descFrame.appendChild(descText);
  descText.layoutSizingHorizontal = 'FILL';

  // Set width after children are added so auto-layout height can settle
  descFrame.resize(Math.max(maxWidth, 240), descFrame.height);

  descFrame.constraints = { horizontal: 'MIN', vertical: 'MIN' };
  return descFrame;
}

function getDocLink(node) {
  if (node.documentationLinks && node.documentationLinks.length > 0 && node.documentationLinks[0].uri) {
    return node.documentationLinks[0].uri;
  }
  return null;
}

function createDocLinkText(uri) {
  var linkText = figma.createText();
  linkText.name = 'Documentation Link';
  linkText.characters = uri;
  linkText.fontSize = DOC_LINK_FONT_SIZE;
  linkText.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.6 }];
  linkText.hyperlink = { type: 'URL', value: uri };
  linkText.textDecoration = 'UNDERLINE';
  linkText.constraints = { horizontal: 'MIN', vertical: 'MIN' };
  return linkText;
}

// ─── Plugin Init ─────────────────────────────────────────────────────────────

var _command = figma.command || 'open';
var _labelFontFamily = 'Inter';

// Eagerly load Inter font — always needed as fallback
_fontLoadPromise = Promise.all([
  figma.loadFontAsync({ family: 'Inter', style: 'Regular' }),
  figma.loadFontAsync({ family: 'Inter', style: 'Semi Bold' }),
  figma.loadFontAsync({ family: 'Inter', style: 'Bold' }),
  figma.loadFontAsync({ family: 'Inter', style: 'Italic' }),
  figma.loadFontAsync({ family: 'Inter', style: 'Bold Italic' })
]).then(function() {
  _fontLoaded = true;
  console.log('[Fonts] Inter Regular/Semi Bold/Bold/Italic/Bold Italic loaded');
}).catch(function(e) {
  console.error('[Fonts] Failed to load Inter:', e.message);
});

function loadLabelFont(family) {
  if (!family || family === 'Inter') {
    _labelFontFamily = 'Inter';
    return Promise.resolve();
  }
  return figma.loadFontAsync({ family: family, style: 'Regular' }).then(function() {
    _labelFontFamily = family;
    console.log('[Fonts] Custom font loaded:', family);
  }).catch(function(e) {
    console.warn('[Fonts] Failed to load "' + family + '", falling back to Inter:', e.message);
    _labelFontFamily = 'Inter';
    figma.notify('Font "' + family + '" not available — using Inter instead.', { timeout: 4000 });
  });
}

if (_command === 'open') {
  figma.showUI(__html__, { width: 340, height: 500 });

  // Load settings
  (async function() {
    try {
      var settings = await figma.clientStorage.getAsync("obra-autodocs-settings");
      if (settings) {
        figma.ui.postMessage({
          type: "settings-init",
          spacingPresets: settings.spacingPresets || null,
          colorPresets: settings.colorPresets || null,
          fontFamily: settings.fontFamily || null,
          docLabel: settings.docLabel || null,
          booleanDisplayMode: settings.booleanDisplayMode || null
        });
      }
    } catch(e) {}
    try {
      var vmPrefs = await figma.clientStorage.getAsync("obra-autodocs-variable-mode-prefs");
      if (vmPrefs) {
        figma.ui.postMessage({ type: "variable-mode-prefs-init", prefs: vmPrefs });
      }
    } catch(e) {}
  })();

  figma.on('selectionchange', () => {
    if (_selectionDebounceTimer) clearTimeout(_selectionDebounceTimer);
    _selectionDebounceTimer = setTimeout(function() {
      _selectionDebounceTimer = null;
      sendSelectionInfo();
      sendDeepVariantInfo();
    }, 150);
    sendGroupifySelectionData();
  });
  sendSelectionInfo();
  sendGroupifySelectionData();
  sendDeepVariantInfo();
} else {
  // Menu commands: hidden UI (needed for postMessage in generate)
  figma.showUI(__html__, { visible: false });

  (async function() {
    var cs = getComponentSet();
    var standalone = !cs ? getStandaloneComponent() : null;
    var node = cs || standalone;
    if (!node) {
      figma.notify('Please select a component set or component.', { error: true });
      figma.closePlugin();
      return;
    }
    var opts = { showGrid: !standalone };
    try {
      var savedSettings = await figma.clientStorage.getAsync("obra-autodocs-settings");
      if (savedSettings && savedSettings.fontFamily) opts.fontFamily = savedSettings.fontFamily;
    } catch(e) {}
    if (_command === 'generate-boolean' || _command === 'generate-both') {
      opts.showBooleanVisibility = true;
    }
    if (_command === 'generate-nested' || _command === 'generate-both') {
      opts.showNestedInstances = true;
    }
    if (_command === 'remove') {
      var wrapper = findExistingWrapper(node);
      var hasPropstar = cs ? detectGusPropstar(cs) : false;
      if (!wrapper && !hasPropstar) {
        figma.notify('No docs found to remove.', { error: true });
        figma.closePlugin();
        return;
      }
      try {
        if (hasPropstar) {
          removeGusPropstar(cs);
        } else {
          removeDocs();
        }
      } catch (e) {
        figma.notify('Error: ' + (e.message || 'Unknown error'), { error: true });
      }
      figma.closePlugin();
      return;
    }
    try {
      await generate(opts);
    } catch (e) {
      figma.notify('Error: ' + (e.message || 'Unknown error'), { error: true });
    }
    figma.closePlugin();
  })();
}

figma.ui.onmessage = async (msg) => {
  if (msg.type === 'resize') {
    figma.ui.resize(msg.width || 340, msg.height);
  }
  if (msg.type === 'generate') {
    try {
      console.log('[Generate] msg keys:', Object.keys(msg).join(', '));
      console.log('[Generate] msg:', JSON.stringify(msg, null, 2));
      await generate({ showGrid: msg.showGrid || false, color: msg.color, gridAlignment: msg.gridAlignment || 'auto', showBooleanVisibility: msg.showBooleanVisibility || false, booleanCombination: msg.booleanCombination || 'individual', booleanDisplayMode: msg.booleanDisplayMode || 'list', booleanImprovedAccuracy: msg.booleanImprovedAccuracy !== false, enabledBooleanProps: msg.enabledBooleanProps || null, booleanScope: msg.booleanScope || null, showNestedInstances: msg.showNestedInstances || false, nestedInstancesMode: msg.nestedInstancesMode || 'representative', enabledNestedInstances: msg.enabledNestedInstances || null, nestedBaseVariantIndices: msg.nestedBaseVariantIndices || null, standaloneDoc: msg.standaloneDoc || false, fontFamily: msg.fontFamily || null, variableModes: msg.variableModes || null, docLabel: msg.docLabel || null, hidePropertyNames: msg.hidePropertyNames || false, showTitle: msg.showTitle || false, showDescription: msg.showDescription || false, showDocLink: msg.showDocLink || false, allowSpanning: msg.allowSpanning || false, autoLayoutExtras: msg.autoLayoutExtras || false, docMode: msg.docMode || 'base' });
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'remove') {
    try {
      removeDocs();
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'toggleGrid') {
    try {
      toggleGrid(msg.showGrid);
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'changeColor') {
    try {
      changeDocsColor(msg.color);
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'migrateNode') {
    try {
      // Navigate to the page
      var targetPage = figma.root.children.find(function(p) { return p.id === msg.pageId; });
      if (!targetPage) throw new Error('Page not found.');
      await targetPage.loadAsync();
      await figma.setCurrentPageAsync(targetPage);

      // Find the PropStar wrapper
      var wrapper = await figma.getNodeByIdAsync(msg.nodeId);
      if (!wrapper) throw new Error('Node not found.');

      // Find the component set inside the wrapper
      var cs = null;
      for (var i = 0; i < wrapper.children.length; i++) {
        if (wrapper.children[i].type === 'COMPONENT_SET') {
          cs = wrapper.children[i];
          break;
        }
      }
      if (!cs) throw new Error('No component set found inside PropStar wrapper.');

      // Select the CS so getComponentSet() works
      figma.currentPage.selection = [cs];
      figma.viewport.scrollAndZoomIntoView([cs]);

      // Remove PropStar docs
      removeGusPropstar(cs);

      // Generate Obra docs
      await generate({ showGrid: true, color: '#9747FF' });

      figma.ui.postMessage({ type: 'migrate-done', rowIndex: msg.rowIndex });
    } catch (e) {
      console.error('[Migrate]', e);
      figma.ui.postMessage({ type: 'migrate-error', rowIndex: msg.rowIndex, message: e.message });
      figma.notify('Migration failed: ' + e.message, { error: true });
    }
  }
  if (msg.type === 'removeGusPropstar') {
    try {
      var cs = getComponentSet();
      if (cs) removeGusPropstar(cs);
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'setDebug') {
    DEBUG = !!msg.debug;
    console.log('[Debug] Verbose logging:', DEBUG ? 'ON' : 'OFF');
  }
  if (msg.type === 'debugGrid') {
    try {
      debugGridCoords();
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'debugLabels') {
    try {
      debugLabelCoords();
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'debugBooleanGrid') {
    try {
      debugBooleanGrid();
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'debugLogSettings') {
    console.log('[Settings Snapshot]', JSON.stringify(msg.snapshot, null, 2));
  }
  if (msg.type === 'navigateToNode') {
    var targetPage = figma.root.children.find(function(p) { return p.id === msg.pageId; });
    if (targetPage) {
      await targetPage.loadAsync();
      await figma.setCurrentPageAsync(targetPage);
      var node = await figma.getNodeByIdAsync(msg.nodeId);
      if (node) {
        figma.currentPage.selection = [node];
        figma.viewport.scrollAndZoomIntoView([node]);
      }
    }
  }
  if (msg.type === 'scanDocs') {
    await scanFileForDocs();
  }
  if (msg.type === 'groupify-reinfer') {
    sendGroupifySelectionData(msg.preferredAxis);
  }
  if (msg.type === 'groupify-apply') {
    applyGroupifyLayout(msg.config).catch(function(e) {
      figma.ui.postMessage({ type: 'groupify-error', message: 'Error: ' + e.message });
    });
  }
  if (msg.type === 'groupify-quick-align') {
    groupifyQuickAlign(msg).catch(function(e) {
      figma.ui.postMessage({ type: 'groupify-error', message: 'Error: ' + e.message });
    });
  }
  if (msg.type === 'settings-save') {
    figma.clientStorage.setAsync("obra-autodocs-settings", { spacingPresets: msg.spacingPresets, colorPresets: msg.colorPresets, fontFamily: msg.fontFamily, docLabel: msg.docLabel, booleanDisplayMode: msg.booleanDisplayMode });
  }
  if (msg.type === 'save-variable-mode-prefs') {
    figma.clientStorage.setAsync("obra-autodocs-variable-mode-prefs", msg.prefs);
  }
  if (msg.type === 'close-plugin') {
    figma.closePlugin();
  }
  // ─── Deep variant selector ────────────────────────────────────
  if (msg.type === 'dv-select-variants') {
    dvSelectVariants(msg.filters, msg.extend);
  }
  if (msg.type === 'dv-remove-from-selection') {
    dvRemoveFromSelection(msg.filters);
  }
  // ─── Shadcn style filter ─────────────────────────────────────
  if (msg.type === 'set-hidden-styles') {
    try {
      await applyHiddenStyles(msg.hidden || [], msg.regenOptions || null);
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  if (msg.type === 'delete-style-variants') {
    try {
      await deleteStyleVariants(msg.value, msg.regenOptions || null);
    } catch (e) {
      figma.ui.postMessage({ type: 'error', message: e.message || 'Unknown error' });
    }
  }
  // ─── Navigate tab ────────────────────────────────────────────
  if (msg.type === 'nav-move-up') { navMoveUp(); }
  if (msg.type === 'nav-move-down') { navMoveDown(); }
  if (msg.type === 'nav-move-left') { navMoveLeft(); }
  if (msg.type === 'nav-move-right') { navMoveRight(); }
  if (msg.type === 'nav-select-row') { navSelectRow(); }
  if (msg.type === 'nav-select-column') { navSelectColumn(); }
  if (msg.type === 'nav-add-next-row') { navAddNextRow(); }
  if (msg.type === 'nav-add-next-column') { navAddNextColumn(); }
  if (msg.type === 'nav-add-prev-row') { navAddPrevRow(); }
  if (msg.type === 'nav-add-prev-column') { navAddPrevColumn(); }
};

// ─── Deep Variant Selector ────────────────────────────────────────────────────

var _dvComponentSet = null;
var _dvAllVariants = [];

function dvParseVariantName(name) {
  var props = {};
  var parts = name.split(',');
  for (var i = 0; i < parts.length; i++) {
    var trimmed = parts[i].trim();
    var equalIndex = trimmed.indexOf('=');
    if (equalIndex > -1) {
      var propName = trimmed.substring(0, equalIndex).trim();
      var propValue = trimmed.substring(equalIndex + 1).trim();
      props[propName] = propValue;
    }
  }
  return props;
}

function dvFindComponentSet(node) {
  if (!node) return null;
  if (node.type === 'COMPONENT_SET') return node;
  if (node.type === 'COMPONENT' && node.parent && node.parent.type === 'COMPONENT_SET') return node.parent;
  var current = node;
  while (current && current.parent) {
    if (current.type === 'COMPONENT' && current.parent.type === 'COMPONENT_SET') return current.parent;
    current = current.parent;
  }
  return null;
}

function dvFindSelectedVariant(node, componentSet) {
  if (!node || !componentSet) return null;
  if (node.type === 'COMPONENT' && node.parent === componentSet) return node;
  var current = node;
  while (current && current.parent) {
    if (current.type === 'COMPONENT' && current.parent === componentSet) return current;
    current = current.parent;
  }
  return null;
}

function dvGetPathFromVariant(node, variant) {
  var path = [];
  var current = node;
  while (current && current !== variant) {
    var parent = current.parent;
    if (!parent || !('children' in parent)) return null;
    var idx = parent.children.indexOf(current);
    if (idx === -1) return null;
    path.unshift(idx);
    current = parent;
  }
  if (current !== variant) return null;
  return path;
}

function dvResolvePathInVariant(variant, path) {
  var node = variant;
  for (var i = 0; i < path.length; i++) {
    if (!node || !('children' in node)) return null;
    var child = node.children[path[i]];
    if (!child) return null;
    node = child;
  }
  return node;
}

function dvAnalyzeSelection(currentSelection, componentSet) {
  var paths = [];
  var seen = {};
  var hasVariantLevel = false;
  for (var i = 0; i < currentSelection.length; i++) {
    var node = currentSelection[i];
    var variant = dvFindSelectedVariant(node, componentSet);
    if (!variant) continue;
    if (node === variant) { hasVariantLevel = true; continue; }
    var path = dvGetPathFromVariant(node, variant);
    if (!path) continue;
    var key = path.join('.');
    if (!seen[key]) { seen[key] = true; paths.push(path); }
  }
  return { paths: paths, hasVariantLevel: hasVariantLevel };
}

function sendDeepVariantInfo() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) {
    // Don't clear cached CS — selection may be empty because UI iframe has focus
    figma.ui.postMessage({ type: 'deep-variant-info', selectedVariantProps: null });
    return;
  }

  var cs = dvFindComponentSet(selection[0]);
  if (!cs) {
    _dvComponentSet = null;
    _dvAllVariants = [];
    figma.ui.postMessage({ type: 'deep-variant-info', selectedVariantProps: null });
    return;
  }

  _dvComponentSet = cs;
  _dvAllVariants = cs.children.filter(function(child) { return child.type === 'COMPONENT'; });

  var selectedVariant = dvFindSelectedVariant(selection[0], cs);
  var selectedProps = selectedVariant ? dvParseVariantName(selectedVariant.name) : null;

  // Collect all possible values per property
  var allPropertyValues = {};
  for (var i = 0; i < _dvAllVariants.length; i++) {
    var vProps = dvParseVariantName(_dvAllVariants[i].name);
    var vPropNames = Object.keys(vProps);
    for (var j = 0; j < vPropNames.length; j++) {
      var pName = vPropNames[j];
      if (!allPropertyValues[pName]) allPropertyValues[pName] = [];
      if (allPropertyValues[pName].indexOf(vProps[pName]) === -1) {
        allPropertyValues[pName].push(vProps[pName]);
      }
    }
  }

  figma.ui.postMessage({
    type: 'deep-variant-info',
    selectedVariantProps: selectedProps,
    allPropertyValues: allPropertyValues
  });
}

function dvEnsureComponentSet() {
  if (_dvComponentSet && _dvAllVariants.length > 0) return true;
  var selection = figma.currentPage.selection;
  for (var i = 0; i < selection.length; i++) {
    var cs = dvFindComponentSet(selection[i]);
    if (cs) {
      _dvComponentSet = cs;
      _dvAllVariants = cs.children.filter(function(child) { return child.type === 'COMPONENT'; });
      return true;
    }
  }
  return false;
}

function dvVariantMatchesFilters(variant, filters) {
  var props = dvParseVariantName(variant.name);
  var propertyNames = Object.keys(filters);
  for (var i = 0; i < propertyNames.length; i++) {
    var propName = propertyNames[i];
    var allowedValues = filters[propName];
    if (allowedValues.length === 0) continue;
    if (allowedValues.indexOf(props[propName]) === -1) return false;
  }
  return true;
}

function dvSelectVariants(filters, extend) {
  if (!dvEnsureComponentSet()) {
    figma.notify('No component set selected');
    return;
  }

  var matchingVariants = _dvAllVariants.filter(function(v) {
    return dvVariantMatchesFilters(v, filters);
  });

  if (matchingVariants.length === 0) {
    figma.notify('No variants match the criteria');
    return;
  }

  if (!extend) {
    figma.currentPage.selection = matchingVariants;
    figma.notify('Selected ' + matchingVariants.length + ' variant' + (matchingVariants.length === 1 ? '' : 's'));
    return;
  }

  var currentSelection = figma.currentPage.selection;
  var analysis = dvAnalyzeSelection(currentSelection, _dvComponentSet);
  var existingIds = {};
  for (var i = 0; i < currentSelection.length; i++) existingIds[currentSelection[i].id] = true;

  var newNodes = [];

  if (analysis.paths.length > 0) {
    // Sublayer-level selection: pull the same paths from each matching variant.
    for (var j = 0; j < matchingVariants.length; j++) {
      var variant = matchingVariants[j];
      for (var k = 0; k < analysis.paths.length; k++) {
        var resolved = dvResolvePathInVariant(variant, analysis.paths[k]);
        if (resolved && !existingIds[resolved.id]) {
          existingIds[resolved.id] = true;
          newNodes.push(resolved);
        }
      }
      if (analysis.hasVariantLevel && !existingIds[variant.id]) {
        existingIds[variant.id] = true;
        newNodes.push(variant);
      }
    }
    figma.currentPage.selection = currentSelection.concat(newNodes);
    figma.notify('Added ' + newNodes.length + ' layer' + (newNodes.length === 1 ? '' : 's'));
  } else {
    // Variant-level (or empty) selection.
    for (var m = 0; m < matchingVariants.length; m++) {
      if (!existingIds[matchingVariants[m].id]) newNodes.push(matchingVariants[m]);
    }
    figma.currentPage.selection = currentSelection.concat(newNodes);
    figma.notify('Added ' + newNodes.length + ' variant' + (newNodes.length === 1 ? '' : 's'));
  }
}

function dvRemoveFromSelection(filters) {
  var currentSelection = figma.currentPage.selection;
  if (currentSelection.length === 0) {
    figma.notify('Nothing selected');
    return;
  }

  dvEnsureComponentSet();
  if (!_dvComponentSet) {
    figma.notify('No component set selected');
    return;
  }

  var matchingVariantIds = {};
  _dvAllVariants.forEach(function(variant) {
    if (dvVariantMatchesFilters(variant, filters)) matchingVariantIds[variant.id] = true;
  });

  var sawNonVariant = false;
  var newSelection = currentSelection.filter(function(node) {
    var variant = dvFindSelectedVariant(node, _dvComponentSet);
    if (!variant) return true;
    if (node !== variant) sawNonVariant = true;
    return !matchingVariantIds[variant.id];
  });

  var removedCount = currentSelection.length - newSelection.length;
  figma.currentPage.selection = newSelection;

  if (removedCount > 0) {
    var unit = sawNonVariant ? 'layer' : 'variant';
    figma.notify('Removed ' + removedCount + ' ' + unit + (removedCount === 1 ? '' : 's'));
  } else {
    figma.notify('No matching variants in selection');
  }
}

// ─── Grid Navigation ─────────────────────────────────────────────────────────

var NAV_TOLERANCE = 5;

function navMoveUp() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select an element first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Selected element has no siblings'); return; }

  var centerYPositions = [];
  for (var i = 0; i < selection.length; i++) {
    var cy = selection[i].y + selection[i].height / 2;
    if (centerYPositions.indexOf(cy) === -1) centerYPositions.push(cy);
  }
  var isRow = centerYPositions.length === 1 && selection.length > 1;
  var minCenterY = Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cy = selection[i].y + selection[i].height / 2;
    if (cy < minCenterY) minCenterY = cy;
  }

  var aboveYs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var childCY = child.y + child.height / 2;
    if (childCY < minCenterY - NAV_TOLERANCE && aboveYs.indexOf(childCY) === -1) aboveYs.push(childCY);
  }
  aboveYs.sort(function(a, b) { return b - a; });
  if (aboveYs.length === 0) { figma.notify('No elements above'); return; }
  var nextCY = aboveYs[0];

  if (isRow) {
    var rowNodes = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.y + child.height / 2 - nextCY) <= NAV_TOLERANCE) rowNodes.push(child);
    }
    figma.currentPage.selection = rowNodes;
    figma.notify('Selected ' + rowNodes.length + ' items in row above');
  } else {
    var targetCX = selection[0].x + selection[0].width / 2;
    var candidates = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.y + child.height / 2 - nextCY) <= NAV_TOLERANCE) candidates.push(child);
    }
    var bestMatch = candidates[0], bestDist = Math.abs(candidates[0].x + candidates[0].width / 2 - targetCX);
    for (var i = 1; i < candidates.length; i++) {
      var dist = Math.abs(candidates[i].x + candidates[i].width / 2 - targetCX);
      if (dist < bestDist) { bestDist = dist; bestMatch = candidates[i]; }
    }
    figma.currentPage.selection = [bestMatch];
  }
}

function navMoveDown() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select an element first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Selected element has no siblings'); return; }

  var centerYPositions = [];
  for (var i = 0; i < selection.length; i++) {
    var cy = selection[i].y + selection[i].height / 2;
    if (centerYPositions.indexOf(cy) === -1) centerYPositions.push(cy);
  }
  var isRow = centerYPositions.length === 1 && selection.length > 1;
  var maxCenterY = -Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cy = selection[i].y + selection[i].height / 2;
    if (cy > maxCenterY) maxCenterY = cy;
  }

  var belowYs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var childCY = child.y + child.height / 2;
    if (childCY > maxCenterY + NAV_TOLERANCE && belowYs.indexOf(childCY) === -1) belowYs.push(childCY);
  }
  belowYs.sort(function(a, b) { return a - b; });
  if (belowYs.length === 0) { figma.notify('No elements below'); return; }
  var nextCY = belowYs[0];

  if (isRow) {
    var rowNodes = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.y + child.height / 2 - nextCY) <= NAV_TOLERANCE) rowNodes.push(child);
    }
    figma.currentPage.selection = rowNodes;
    figma.notify('Selected ' + rowNodes.length + ' items in row below');
  } else {
    var targetCX = selection[0].x + selection[0].width / 2;
    var candidates = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.y + child.height / 2 - nextCY) <= NAV_TOLERANCE) candidates.push(child);
    }
    var bestMatch = candidates[0], bestDist = Math.abs(candidates[0].x + candidates[0].width / 2 - targetCX);
    for (var i = 1; i < candidates.length; i++) {
      var dist = Math.abs(candidates[i].x + candidates[i].width / 2 - targetCX);
      if (dist < bestDist) { bestDist = dist; bestMatch = candidates[i]; }
    }
    figma.currentPage.selection = [bestMatch];
  }
}

function navMoveLeft() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select an element first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Selected element has no siblings'); return; }

  var centerXPositions = [];
  for (var i = 0; i < selection.length; i++) {
    var cx = selection[i].x + selection[i].width / 2;
    if (centerXPositions.indexOf(cx) === -1) centerXPositions.push(cx);
  }
  var isColumn = centerXPositions.length === 1 && selection.length > 1;
  var minCenterX = Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cx = selection[i].x + selection[i].width / 2;
    if (cx < minCenterX) minCenterX = cx;
  }

  var leftXs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var childCX = child.x + child.width / 2;
    if (childCX < minCenterX - NAV_TOLERANCE && leftXs.indexOf(childCX) === -1) leftXs.push(childCX);
  }
  leftXs.sort(function(a, b) { return b - a; });
  if (leftXs.length === 0) { figma.notify('No elements to the left'); return; }
  var nextCX = leftXs[0];

  if (isColumn) {
    var colNodes = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.x + child.width / 2 - nextCX) <= NAV_TOLERANCE) colNodes.push(child);
    }
    figma.currentPage.selection = colNodes;
    figma.notify('Selected ' + colNodes.length + ' items in column to the left');
  } else {
    var targetCY = selection[0].y + selection[0].height / 2;
    var candidates = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.x + child.width / 2 - nextCX) <= NAV_TOLERANCE) candidates.push(child);
    }
    var bestMatch = candidates[0], bestDist = Math.abs(candidates[0].y + candidates[0].height / 2 - targetCY);
    for (var i = 1; i < candidates.length; i++) {
      var dist = Math.abs(candidates[i].y + candidates[i].height / 2 - targetCY);
      if (dist < bestDist) { bestDist = dist; bestMatch = candidates[i]; }
    }
    figma.currentPage.selection = [bestMatch];
  }
}

function navMoveRight() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select an element first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Selected element has no siblings'); return; }

  var centerXPositions = [];
  for (var i = 0; i < selection.length; i++) {
    var cx = selection[i].x + selection[i].width / 2;
    if (centerXPositions.indexOf(cx) === -1) centerXPositions.push(cx);
  }
  var isColumn = centerXPositions.length === 1 && selection.length > 1;
  var maxCenterX = -Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cx = selection[i].x + selection[i].width / 2;
    if (cx > maxCenterX) maxCenterX = cx;
  }

  var rightXs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var childCX = child.x + child.width / 2;
    if (childCX > maxCenterX + NAV_TOLERANCE && rightXs.indexOf(childCX) === -1) rightXs.push(childCX);
  }
  rightXs.sort(function(a, b) { return a - b; });
  if (rightXs.length === 0) { figma.notify('No elements to the right'); return; }
  var nextCX = rightXs[0];

  if (isColumn) {
    var colNodes = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.x + child.width / 2 - nextCX) <= NAV_TOLERANCE) colNodes.push(child);
    }
    figma.currentPage.selection = colNodes;
    figma.notify('Selected ' + colNodes.length + ' items in column to the right');
  } else {
    var targetCY = selection[0].y + selection[0].height / 2;
    var candidates = [];
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type !== 'TEXT' && Math.abs(child.x + child.width / 2 - nextCX) <= NAV_TOLERANCE) candidates.push(child);
    }
    var bestMatch = candidates[0], bestDist = Math.abs(candidates[0].y + candidates[0].height / 2 - targetCY);
    for (var i = 1; i < candidates.length; i++) {
      var dist = Math.abs(candidates[i].y + candidates[i].height / 2 - targetCY);
      if (dist < bestDist) { bestDist = dist; bestMatch = candidates[i]; }
    }
    figma.currentPage.selection = [bestMatch];
  }
}

function navSelectRow() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select an element first'); return; }
  var selected = selection[0];
  var parent = selected.parent;
  if (!parent || !('children' in parent)) { figma.notify('Selected element has no siblings'); return; }
  var targetCY = selected.y + selected.height / 2;
  var rowNodes = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type !== 'TEXT' && Math.abs(child.y + child.height / 2 - targetCY) <= NAV_TOLERANCE) rowNodes.push(child);
  }
  figma.currentPage.selection = rowNodes;
  figma.notify('Selected ' + rowNodes.length + ' items in row');
}

function navSelectColumn() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select an element first'); return; }
  var selected = selection[0];
  var parent = selected.parent;
  if (!parent || !('children' in parent)) { figma.notify('Selected element has no siblings'); return; }
  var targetCX = selected.x + selected.width / 2;
  var colNodes = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type !== 'TEXT' && Math.abs(child.x + child.width / 2 - targetCX) <= NAV_TOLERANCE) colNodes.push(child);
  }
  figma.currentPage.selection = colNodes;
  figma.notify('Selected ' + colNodes.length + ' items in column');
}

function navAddNextRow() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select elements first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Cannot find parent container'); return; }
  var maxCenterY = -Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cy = selection[i].y + selection[i].height / 2;
    if (cy > maxCenterY) maxCenterY = cy;
  }
  var allYs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var cy = child.y + child.height / 2;
    if (allYs.indexOf(cy) === -1) allYs.push(cy);
  }
  allYs.sort(function(a, b) { return a - b; });
  var nextCY = null;
  for (var i = 0; i < allYs.length; i++) {
    if (allYs[i] > maxCenterY + NAV_TOLERANCE) { nextCY = allYs[i]; break; }
  }
  if (nextCY === null) { figma.notify('No more rows below'); return; }
  var nextRowNodes = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type !== 'TEXT' && Math.abs(child.y + child.height / 2 - nextCY) <= NAV_TOLERANCE) nextRowNodes.push(child);
  }
  var newSelection = selection.slice();
  for (var i = 0; i < nextRowNodes.length; i++) newSelection.push(nextRowNodes[i]);
  figma.currentPage.selection = newSelection;
  figma.notify('Added ' + nextRowNodes.length + ' items from next row');
}

function navAddNextColumn() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select elements first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Cannot find parent container'); return; }
  var maxCenterX = -Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cx = selection[i].x + selection[i].width / 2;
    if (cx > maxCenterX) maxCenterX = cx;
  }
  var allXs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var cx = child.x + child.width / 2;
    if (allXs.indexOf(cx) === -1) allXs.push(cx);
  }
  allXs.sort(function(a, b) { return a - b; });
  var nextCX = null;
  for (var i = 0; i < allXs.length; i++) {
    if (allXs[i] > maxCenterX + NAV_TOLERANCE) { nextCX = allXs[i]; break; }
  }
  if (nextCX === null) { figma.notify('No more columns to the right'); return; }
  var nextColNodes = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type !== 'TEXT' && Math.abs(child.x + child.width / 2 - nextCX) <= NAV_TOLERANCE) nextColNodes.push(child);
  }
  var newSelection = selection.slice();
  for (var i = 0; i < nextColNodes.length; i++) newSelection.push(nextColNodes[i]);
  figma.currentPage.selection = newSelection;
  figma.notify('Added ' + nextColNodes.length + ' items from next column');
}

function navAddPrevRow() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select elements first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Cannot find parent container'); return; }
  var minCenterY = Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cy = selection[i].y + selection[i].height / 2;
    if (cy < minCenterY) minCenterY = cy;
  }
  var allYs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var cy = child.y + child.height / 2;
    if (allYs.indexOf(cy) === -1) allYs.push(cy);
  }
  allYs.sort(function(a, b) { return b - a; });
  var prevCY = null;
  for (var i = 0; i < allYs.length; i++) {
    if (allYs[i] < minCenterY - NAV_TOLERANCE) { prevCY = allYs[i]; break; }
  }
  if (prevCY === null) { figma.notify('No more rows above'); return; }
  var prevRowNodes = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type !== 'TEXT' && Math.abs(child.y + child.height / 2 - prevCY) <= NAV_TOLERANCE) prevRowNodes.push(child);
  }
  var newSelection = selection.slice();
  for (var i = 0; i < prevRowNodes.length; i++) newSelection.push(prevRowNodes[i]);
  figma.currentPage.selection = newSelection;
  figma.notify('Added ' + prevRowNodes.length + ' items from previous row');
}

function navAddPrevColumn() {
  var selection = figma.currentPage.selection;
  if (selection.length === 0) { figma.notify('Please select elements first'); return; }
  var parent = selection[0].parent;
  if (!parent || !('children' in parent)) { figma.notify('Cannot find parent container'); return; }
  var minCenterX = Infinity;
  for (var i = 0; i < selection.length; i++) {
    var cx = selection[i].x + selection[i].width / 2;
    if (cx < minCenterX) minCenterX = cx;
  }
  var allXs = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'TEXT') continue;
    var cx = child.x + child.width / 2;
    if (allXs.indexOf(cx) === -1) allXs.push(cx);
  }
  allXs.sort(function(a, b) { return b - a; });
  var prevCX = null;
  for (var i = 0; i < allXs.length; i++) {
    if (allXs[i] < minCenterX - NAV_TOLERANCE) { prevCX = allXs[i]; break; }
  }
  if (prevCX === null) { figma.notify('No more columns to the left'); return; }
  var prevColNodes = [];
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type !== 'TEXT' && Math.abs(child.x + child.width / 2 - prevCX) <= NAV_TOLERANCE) prevColNodes.push(child);
  }
  var newSelection = selection.slice();
  for (var i = 0; i < prevColNodes.length; i++) newSelection.push(prevColNodes[i]);
  figma.currentPage.selection = newSelection;
  figma.notify('Added ' + prevColNodes.length + ' items from previous column');
}

// ─── Selection Detection (Step 2) ────────────────────────────────────────────

function getComponentSet() {
  const sel = figma.currentPage.selection;
  if (sel.length !== 1) {
    return null;
  }
  const node = sel[0];
  if (node.type === 'COMPONENT_SET') {
    return node;
  }
  if (node.type === 'COMPONENT' && node.parent && node.parent.type === 'COMPONENT_SET') {
    return node.parent;
  }
  // Walk up the tree to find a wrapper frame, then find the component set inside it
  let current = node;
  while (current) {
    if (current.type === 'FRAME') {
      // Check for ❖ wrapper or GusPropstar wrapper (has "labels" GROUP or "instances" FRAME)
      var isWrapper = current.name.startsWith('❖');
      if (!isWrapper) {
        for (const ch of current.children) {
          if (ch.type === 'GROUP' && ch.name.toLowerCase() === 'labels') { isWrapper = true; break; }
          if (ch.type === 'FRAME' && ch.name.toLowerCase() === 'instances') { isWrapper = true; break; }
        }
      }
      if (isWrapper) {
        // Found a wrapper — look for the component set child
        for (const child of current.children) {
          if (child.type === 'COMPONENT_SET') {
            return child;
          }
        }
      }
    }
    current = current.parent;
  }
  return null;
}

function isStandaloneComponent(node) {
  return node.type === 'COMPONENT' && (!node.parent || node.parent.type !== 'COMPONENT_SET');
}

function getStandaloneComponent() {
  var sel = figma.currentPage.selection;
  if (sel.length !== 1) return null;
  var node = sel[0];

  // Direct standalone component selection
  if (isStandaloneComponent(node)) {
    return node;
  }

  // Walk up to find a ❖ wrapper containing a standalone component
  var current = node;
  while (current) {
    if (current.type === 'FRAME' && current.name.startsWith('❖')) {
      for (var i = 0; i < current.children.length; i++) {
        if (isStandaloneComponent(current.children[i])) {
          return current.children[i];
        }
      }
    }
    current = current.parent;
  }
  return null;
}

async function sendSelectionInfo() {
  var gen = ++_selectionGeneration;
  var _t0 = Date.now();
  try {
  const cs = getComponentSet();
  if (!cs) {
    // Try standalone component
    var standalone = getStandaloneComponent();
    if (!standalone) {
      figma.ui.postMessage({ type: 'selection', componentSet: null });
      return;
    }

    var saWrapper = findExistingWrapper(standalone);
    var saBoolProps = getBooleanComponentProperties(standalone);
    var saNestedSets = await findNestedComponentSetsForComponent(standalone);
    if (gen !== _selectionGeneration) return; // stale, newer selection in progress
    var saNestedInfo = saNestedSets.map(function(ns) {
      var propValues = [];
      if (ns.variantGroupProperties) {
        for (var pk in ns.variantGroupProperties) {
          if (ns.variantGroupProperties.hasOwnProperty(pk)) {
            propValues.push(ns.variantGroupProperties[pk].length);
          }
        }
      }
      return { name: ns.componentSetName, propertyValues: propValues };
    });

    figma.ui.postMessage({
      type: 'selection',
      componentSet: standalone.name,
      isStandalone: true,
      properties: [],
      hasWrapper: !!saWrapper,
      hasGrid: false,
      hasGusPropstar: false,
      hasBooleanProps: saBoolProps.length > 0,
      booleanPropNames: saBoolProps.map(function(p) { return p.name; }),
      variantCount: 1,
      hasNestedInstances: saNestedSets.length > 0,
      nestedInstancesInfo: saNestedInfo,
      gridWarning: null,
    });
    return;
  }
  const enumProps = getEnumProperties(cs);
  var wrapper = findExistingWrapper(cs);
  var standaloneWrapper = findStandaloneWrapper(cs);
  var variableModesWrapper = findVariableModesWrapper(cs);
  var effectiveWrapper = wrapper || standaloneWrapper;

  // Check grid uniformity — only warn for very sparse grids (>50% empty cells).
  // Many component sets intentionally have empty cells for variant combinations
  // that don't exist (e.g. icon-only buttons don't have text style variants).
  var gridWarning = null;
  var variants = cs.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
  if (variants.length > 1) {
    var colAxis = clusterVariantAxis(variants, 'x', 'width');
    var rowAxis = clusterVariantAxis(variants, 'y', 'height');
    var cols = colAxis.clusters.length;
    var rows = rowAxis.clusters.length;
    var expected = cols * rows;
    var fillRate = variants.length / expected;
    if (variants.length !== expected && fillRate < 0.5) {
      gridWarning = variants.length + ' variants in a ' + cols + '\u00d7' + rows + ' grid (' + (expected - variants.length) + ' empty cells)';
    }
  }

  // Detect single-type grid (one property spread across multi-row × multi-col grid)
  // Use raw position clustering (clusterValues) instead of clusterVariantAxis, because
  // clusterVariantAxis merges overlapping bounding boxes — a spanning variant that covers
  // 2 columns would collapse them into 1, preventing detection.
  var isSingleTypeGrid = false;
  if (enumProps.length === 1 && variants.length > 1) {
    var stCols = clusterValues(variants.map(function(v) { return v.x; })).clusters.length;
    var stRows = clusterValues(variants.map(function(v) { return v.y; })).clusters.length;
    if (stCols > 1 && stRows > 1) {
      isSingleTypeGrid = true;
    }
  }

  // Check for absolute-positioned overflow (e.g. dropdown menus)
  var overflowInfo = detectComponentSetOverflow(cs);

  var boolProps = getBooleanComponentProperties(cs);

  // Check for grid inside whichever wrapper exists
  var hasGrid = false;
  if (effectiveWrapper) {
    // Grid may be inside a Documentation frame or directly in wrapper
    hasGrid = effectiveWrapper.children.some(function(c) {
      if (c.type === 'FRAME' && c.name === 'Grid') return true;
      if (c.type === 'FRAME' && c.name === 'Documentation') {
        return c.children.some(function(dc) { return dc.type === 'FRAME' && dc.name === 'Grid'; });
      }
      return false;
    });
  }

  // Derive existing doc options from layer names in the wrapper
  var existingDocMeta = null;
  if (effectiveWrapper) {
    existingDocMeta = deriveDocMeta(effectiveWrapper);
  }

  // Read saved per-component settings (if any)
  var savedSettings = getComponentSettings(cs);

  // Detect Shadcn style property (for the hide/delete-by-style panel)
  var shadcnStyleInfo = getShadcnStyleInfo(cs);

  // Detect variable collections and modes
  var variableCollections = [];
  try {
    var allCollections = await figma.variables.getLocalVariableCollectionsAsync();
    variableCollections = allCollections
      .filter(function(c) { return c.modes.length > 1; })
      .map(function(c) { return { id: c.id, name: c.name, modes: c.modes.map(function(m) { return { modeId: m.modeId, name: m.name }; }) }; });
  } catch(e) {
    console.warn('[Variables] Could not read variable collections:', e.message);
  }

  // Send selection info immediately (without nested instance data)
  figma.ui.postMessage({
    type: 'selection',
    componentSet: cs.name,
    properties: enumProps.map(p => ({ name: p.name, values: p.values })),
    hasWrapper: !!wrapper,
    hasStandaloneWrapper: !!standaloneWrapper,
    hasVariableModesWrapper: !!variableModesWrapper,
    hasGrid: hasGrid,
    hasGusPropstar: !!detectGusPropstar(cs),
    hasBooleanProps: boolProps.length > 0,
    booleanPropNames: boolProps.map(function(p) { return p.name; }),
    variantCount: variants.length,
    hasNestedInstances: false,
    nestedInstancesInfo: [],
    gridWarning: gridWarning,
    hasOverflow: overflowInfo.hasOverflow,
    isSingleTypeGrid: isSingleTypeGrid,
    variableCollections: variableCollections,
    existingDocMeta: existingDocMeta,
    savedSettings: savedSettings,
    shadcnStyleInfo: shadcnStyleInfo,
  });

  // Async: check for nested component instances, then update UI if still current
  var nestedSets = await findNestedComponentSets(cs);
  if (gen !== _selectionGeneration) return; // stale, newer selection in progress

  if (nestedSets.length > 0) {
    var nestedInfo = nestedSets.map(function(ns) {
      var propValues = [];
      if (ns.variantGroupProperties) {
        for (var pk in ns.variantGroupProperties) {
          if (ns.variantGroupProperties.hasOwnProperty(pk)) {
            propValues.push(ns.variantGroupProperties[pk].length);
          }
        }
      }
      var parentVariants = null;
      if (ns.parentVariantIndices && ns.parentVariantIndices.length > 0) {
        parentVariants = ns.parentVariantIndices.map(function(idx) {
          return { index: idx, name: variants[idx] ? variants[idx].name : 'Variant ' + idx };
        });
      }
      return { name: ns.componentSetName, propertyValues: propValues, parentVariants: parentVariants };
    });

    // Update with nested instance data
    figma.ui.postMessage({
      type: 'selection',
      componentSet: cs.name,
      properties: enumProps.map(p => ({ name: p.name, values: p.values })),
      hasWrapper: !!wrapper,
      hasStandaloneWrapper: !!standaloneWrapper,
      hasGrid: hasGrid,
      hasGusPropstar: !!detectGusPropstar(cs),
      hasBooleanProps: boolProps.length > 0,
      booleanPropNames: boolProps.map(function(p) { return p.name; }),
      variantCount: variants.length,
      hasNestedInstances: true,
      nestedInstancesInfo: nestedInfo,
      gridWarning: gridWarning,
      hasOverflow: overflowInfo.hasOverflow,
      isSingleTypeGrid: isSingleTypeGrid,
      variableCollections: variableCollections,
      existingDocMeta: existingDocMeta,
      savedSettings: savedSettings,
      shadcnStyleInfo: shadcnStyleInfo,
    });
  }

  if (DEBUG) console.log('[Perf] sendSelectionInfo: ' + (Date.now() - _t0) + 'ms (' + variants.length + ' variants)');
  } catch (e) {
    console.error('[sendSelectionInfo] Error:', e.message, e.stack);
    figma.ui.postMessage({ type: 'selection', componentSet: null });
  }
}

// ─── Property Extraction & Filtering (Step 3) ────────────────────────────────

function isBooleanProperty(values) {
  if (values.length !== 2) return false;
  const sorted = values.map(v => v.toLowerCase()).sort();
  const boolPairs = [['false', 'true'], ['no', 'yes'], ['off', 'on']];
  return boolPairs.some(pair => sorted[0] === pair[0] && sorted[1] === pair[1]);
}

function isNestedProperty(name) {
  return name.includes('/');
}

function getEnumProperties(componentSet, options = {}) {
  const groupProps = componentSet.variantGroupProperties;
  const result = [];

  // Values that still have at least one visible variant. Used to drop values
  // whose variants are all hidden via the Shadcn style filter.
  let visibleValuesByProp = null;
  const visibleVariants = componentSet.children.filter(function(c) {
    return c.type === 'COMPONENT' && c.visible !== false;
  });
  const allVariants = componentSet.children.filter(function(c) { return c.type === 'COMPONENT'; });
  if (visibleVariants.length !== allVariants.length) {
    visibleValuesByProp = {};
    for (let i = 0; i < visibleVariants.length; i++) {
      const parsed = dvParseVariantName(visibleVariants[i].name);
      for (const pname in parsed) {
        if (!visibleValuesByProp[pname]) visibleValuesByProp[pname] = {};
        visibleValuesByProp[pname][parsed[pname]] = true;
      }
    }
  }

  for (const [name, info] of Object.entries(groupProps)) {
    if (isNestedProperty(name)) continue;
    let values = info.values;
    if (visibleValuesByProp) {
      const seen = visibleValuesByProp[name] || {};
      values = values.filter(function(v) { return seen[v]; });
      if (values.length === 0) continue;
    }
    result.push({ name, values });
  }
  return result;
}

// ─── Boolean Component Properties ────────────────────────────────────────────

function getBooleanComponentProperties(componentSet) {
  var propDefs = componentSet.componentPropertyDefinitions;
  if (!propDefs) return [];
  var result = [];
  for (var key in propDefs) {
    if (propDefs.hasOwnProperty(key) && propDefs[key].type === 'BOOLEAN') {
      result.push({ key: key, name: key.replace(/#.*$/, ''), defaultValue: propDefs[key].defaultValue });
    }
  }
  return result;
}

function determineBooleanAxis(csWidth, csHeight) {
  return csWidth >= csHeight ? 'row' : 'col';
}

function createBackgroundTint(x, y, width, height) {
  var rect = figma.createRectangle();
  rect.name = 'Boolean Tint';
  rect.x = x;
  rect.y = y;
  rect.resize(width, height);
  rect.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.03 }];
  rect.locked = true;
  return rect;
}

function createBooleanInstanceGroup(sourceNode, propsToSet, offsetX, offsetY) {
  var standalone = isStandaloneComponent(sourceNode);
  var variants = standalone ? [sourceNode] : sourceNode.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
  var instances = [];
  var maxExtentX = 0;
  var maxExtentY = 0;
  for (var i = 0; i < variants.length; i++) {
    try {
      var variant = variants[i];
      var instance = variant.createInstance();
      instance.setProperties(propsToSet);
      var relX = standalone ? 0 : variant.x;
      var relY = standalone ? 0 : variant.y;
      instance.x = offsetX + relX;
      instance.y = offsetY + relY;
      var extX = relX + instance.width;
      var extY = relY + instance.height;
      if (extX > maxExtentX) maxExtentX = extX;
      if (extY > maxExtentY) maxExtentY = extY;
      instances.push(instance);
    } catch (e) {
      console.log('[Boolean Grid] Error creating instance:', e.message);
    }
  }
  return { instances: instances, width: maxExtentX, height: maxExtentY };
}

// ─── Nested Instance Detection ───────────────────────────────────────────────

// Recursively find all INSTANCE nodes inside a node
function findInstanceNodes(node) {
  var result = [];
  if (!('children' in node)) return result;
  for (var i = 0; i < node.children.length; i++) {
    var child = node.children[i];
    if (child.type === 'INSTANCE') {
      result.push(child);
    }
    result = result.concat(findInstanceNodes(child));
  }
  return result;
}

// Shared helper: resolves exposed instance nodes into nested component set results.
// Used by both findNestedComponentSets and findNestedComponentSetsForComponent.
async function resolveNestedComponentSets(exposed, debugLabel, variantIndicesByName) {
  var seen = {};
  var result = [];

  for (var i = 0; i < exposed.length; i++) {
    var inst = exposed[i];
    // Skip instances whose layer name starts with '.' (local/private to the component)
    if (inst.name.charAt(0) === '.') continue;
    var mc;
    try {
      mc = await Promise.race([
        inst.getMainComponentAsync(),
        new Promise(function(_, reject) { setTimeout(function() { reject(new Error('timeout')); }, 2000); })
      ]);
    } catch (e) {
      console.log('[NestedInstances] getMainComponentAsync timed out for "' + inst.name + '"');
      continue;
    }
    if (!mc) continue;
    if (mc.parent && mc.parent.type === 'COMPONENT_SET') {
      var nestedCS = mc.parent;
      if (seen[nestedCS.id]) continue;
      // Skip component sets whose name starts with '.' (local/private)
      if (nestedCS.name.charAt(0) === '.') { seen[nestedCS.id] = true; continue; }
      seen[nestedCS.id] = true;

      var nestedVariants = nestedCS.children.filter(function(c) { return c.type === 'COMPONENT'; });
      if (nestedVariants.length <= 1) continue;

      var nestedGroupProps = nestedCS.variantGroupProperties;
      var groupPropData = {};
      for (var propName in nestedGroupProps) {
        if (nestedGroupProps.hasOwnProperty(propName)) {
          groupPropData[propName] = nestedGroupProps[propName].values;
        }
      }

      var defaultVariant = nestedVariants.filter(function(v) { return v.id === mc.id; })[0];
      var defaultProps = defaultVariant ? defaultVariant.variantProperties : (nestedVariants[0] ? nestedVariants[0].variantProperties : {});

      result.push({
        componentSetId: nestedCS.id,
        componentSetName: nestedCS.name,
        instanceName: inst.name,
        currentVariantId: mc.id,
        variants: nestedVariants.map(function(v) { return { id: v.id, name: v.name }; }),
        variantGroupProperties: groupPropData,
        defaultVariantProperties: defaultProps,
        parentVariantIndices: (variantIndicesByName && variantIndicesByName[inst.name]) || null
      });
    }
  }

  if (DEBUG) {
    console.log('[NestedInstances] Found', result.length, 'nested component sets in', debugLabel);
    for (var j = 0; j < result.length; j++) {
      console.log('[NestedInstances]   "' + result[j].componentSetName + '" (' + result[j].variants.length + ' variants), instance layer: "' + result[j].instanceName + '", parent variant indices:', result[j].parentVariantIndices);
    }
  }

  return result;
}

// Filter exposed instances to only top-level ones (not nested inside other instances)
function filterTopLevelExposed(exposed, rootNode) {
  return exposed.filter(function(inst) {
    var node = inst.parent;
    while (node && node !== rootNode) {
      if (node.type === 'INSTANCE') {
        return false;
      }
      node = node.parent;
    }
    return true;
  });
}

// Find nested component sets by scanning for instances with isExposedInstance === true.
// Only includes instances the designer explicitly exposed in the component set.
async function findNestedComponentSets(componentSet) {
  var variants = componentSet.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
  if (variants.length === 0) return [];

  // Scan all variants for exposed instances (some may only exist in certain variants)
  // Track which parent variant indices contain each exposed instance
  var exposedByComponent = {};
  var variantIndicesByName = {};
  for (var vi = 0; vi < variants.length; vi++) {
    var allInstances = findInstanceNodes(variants[vi]);
    var variantExposed = allInstances.filter(function(inst) { return inst.isExposedInstance === true; });
    variantExposed = filterTopLevelExposed(variantExposed, variants[vi]);
    // Skip instances whose layer name starts with '.' (local/private to the component)
    variantExposed = variantExposed.filter(function(inst) { return inst.name.charAt(0) !== '.'; });

    var seenInVariant = {};
    for (var ei = 0; ei < variantExposed.length; ei++) {
      var key = variantExposed[ei].name;
      if (!exposedByComponent[key]) {
        exposedByComponent[key] = variantExposed[ei];
        variantIndicesByName[key] = [];
      }
      if (!seenInVariant[key]) {
        variantIndicesByName[key].push(vi);
        seenInVariant[key] = true;
      }
    }
  }
  var exposed = [];
  for (var ek in exposedByComponent) {
    if (exposedByComponent.hasOwnProperty(ek)) {
      exposed.push(exposedByComponent[ek]);
    }
  }
  if (DEBUG) { console.log('[NestedInstances] top-level exposed across all variants:', exposed.length); }

  return resolveNestedComponentSets(exposed, componentSet.name, variantIndicesByName);
}

// Variant of findNestedComponentSets for a standalone component (not a component set)
async function findNestedComponentSetsForComponent(component) {
  var allInstances = findInstanceNodes(component);
  var exposed = allInstances.filter(function(inst) { return inst.isExposedInstance === true; });
  exposed = filterTopLevelExposed(exposed, component);
  // Skip instances whose layer name starts with '.' (local/private to the component)
  exposed = exposed.filter(function(inst) { return inst.name.charAt(0) !== '.'; });

  if (DEBUG) { console.log('[NestedInstances] top-level exposed in standalone:', exposed.length); }

  return resolveNestedComponentSets(exposed, 'standalone ' + component.name);
}

// Find a variant in a list by matching target property values against the variant name
function findVariantByProperties(variants, targetProps) {
  for (var i = 0; i < variants.length; i++) {
    var v = variants[i];
    var props = {};
    var parts = v.name.split(',');
    for (var j = 0; j < parts.length; j++) {
      var eq = parts[j].indexOf('=');
      if (eq > 0) {
        props[parts[j].substring(0, eq).trim()] = parts[j].substring(eq + 1).trim();
      }
    }
    var match = true;
    for (var key in targetProps) {
      if (targetProps.hasOwnProperty(key) && props[key] !== targetProps[key]) {
        match = false;
        break;
      }
    }
    if (match) return v;
  }
  return null;
}

// ─── Layout Analysis (Step 4) ────────────────────────────────────────────────

function clusterValues(positions) {
  var sorted = Array.from(new Set(positions)).sort(function(a, b) { return a - b; });
  var clusters = [];
  var highs = []; // track highest value in each cluster for chain comparison
  for (var i = 0; i < sorted.length; i++) {
    // Chain comparison: compare against the LAST (highest) value in the current
    // cluster, not the first. This prevents splitting when values gradually drift
    // across a range wider than CLUSTER_TOLERANCE (e.g. mixed-size variant centers).
    if (clusters.length === 0 || sorted[i] - highs[highs.length - 1] > CLUSTER_TOLERANCE) {
      clusters.push(sorted[i]);
      highs.push(sorted[i]);
    } else {
      highs[highs.length - 1] = sorted[i];
    }
  }
  return { clusters: clusters, highs: highs };
}

// Cluster variants into columns/rows using smart alignment detection.
// Detect variants with absolutely positioned children that overflow the frame.
// Returns { hasOverflow, overflowBottom, overflowRight } for a single variant.
function detectVariantOverflow(variant) {
  var overflowBottom = 0;
  var overflowRight = 0;
  function scanChildren(node, depth) {
    if (depth > 3) return; // limit recursion depth
    if (!node.children) return;
    for (var i = 0; i < node.children.length; i++) {
      var child = node.children[i];
      if (child.layoutPositioning === 'ABSOLUTE') {
        var childBottom = child.y + child.height - variant.height;
        var childRight = child.x + child.width - variant.width;
        if (childBottom > overflowBottom) overflowBottom = childBottom;
        if (childRight > overflowRight) overflowRight = childRight;
      }
      scanChildren(child, depth + 1);
    }
  }
  scanChildren(variant, 0);
  return {
    hasOverflow: overflowBottom > 20 || overflowRight > 20,
    overflowBottom: overflowBottom,
    overflowRight: overflowRight
  };
}

// Detect overflow across all variants in a component set.
// Returns { hasOverflow, maxOverflowBottom, maxOverflowRight, affectedVariants[] }
function detectComponentSetOverflow(componentSet) {
  var variants = componentSet.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
  var maxBottom = 0, maxRight = 0;
  var affected = [];
  for (var i = 0; i < variants.length; i++) {
    var ov = detectVariantOverflow(variants[i]);
    if (ov.hasOverflow) {
      affected.push({ name: variants[i].name, overflowBottom: ov.overflowBottom, overflowRight: ov.overflowRight });
      if (ov.overflowBottom > maxBottom) maxBottom = ov.overflowBottom;
      if (ov.overflowRight > maxRight) maxRight = ov.overflowRight;
    }
  }
  if (affected.length > 0 && DEBUG) {
    console.log('[Overflow] Detected ' + affected.length + ' variant(s) with absolutely positioned overflow:');
    for (var j = 0; j < affected.length; j++) {
      console.log('[Overflow]   "' + affected[j].name + '" → bottom:+' + Math.round(affected[j].overflowBottom) + 'px right:+' + Math.round(affected[j].overflowRight) + 'px');
    }
  }
  return {
    hasOverflow: affected.length > 0,
    maxOverflowBottom: maxBottom,
    maxOverflowRight: maxRight,
    affectedVariants: affected
  };
}

// Tries left-edge, center, and right-edge clustering; picks the one that
// produces fewest clusters (= the common edge variants share).
// Returns { clusters: [refValue, ...], mins: [minEdge, ...], maxes: [maxEdge, ...], mode: 'left'|'center'|'right' }
function effectiveSize(v, sizeKey) {
  if (_visualSizeOverrides && _visualSizeOverrides[v.id]) {
    return Math.max(_visualSizeOverrides[v.id][sizeKey], 1);
  }
  return Math.max(v[sizeKey], 1);
}

function clusterVariantAxis(variants, posKey, sizeKey) {
  var lefts = variants.map(function(v) { return v[posKey]; });
  var centers = variants.map(function(v) { return v[posKey] + effectiveSize(v, sizeKey) / 2; });
  var rights = variants.map(function(v) { return v[posKey] + effectiveSize(v, sizeKey); });

  var leftResult = clusterValues(lefts);
  var centerResult = clusterValues(centers);
  var rightResult = clusterValues(rights);

  // Pick alignment with fewest clusters (= best common edge).
  // On tie, prefer left (most common Figma default) > center > right.
  var best = 'left', bestClusters = leftResult.clusters, bestHighs = leftResult.highs, bestRefs = lefts;
  if (centerResult.clusters.length < bestClusters.length) {
    best = 'center'; bestClusters = centerResult.clusters; bestHighs = centerResult.highs; bestRefs = centers;
  }
  if (rightResult.clusters.length < bestClusters.length) {
    best = 'right'; bestClusters = rightResult.clusters; bestHighs = rightResult.highs; bestRefs = rights;
  }

  if (DEBUG) {
    console.log('[Cluster] ' + posKey + ' axis: left=' + leftResult.clusters.length + ' center=' + centerResult.clusters.length + ' right=' + rightResult.clusters.length + ' → using ' + best);
  }

  // For each cluster, compute the bounding box from actual variant edges.
  // Use the full chain range [cluster, high] to capture all chained members.
  var mins = [];
  var maxes = [];
  for (var i = 0; i < bestClusters.length; i++) {
    var minEdge = Infinity;
    var maxEdge = -Infinity;
    for (var j = 0; j < variants.length; j++) {
      if (bestRefs[j] >= bestClusters[i] - CLUSTER_TOLERANCE && bestRefs[j] <= bestHighs[i] + CLUSTER_TOLERANCE) {
        var left = variants[j][posKey];
        var right = left + effectiveSize(variants[j], sizeKey);
        if (left < minEdge) minEdge = left;
        if (right > maxEdge) maxEdge = right;
      }
    }
    mins.push(minEdge);
    maxes.push(maxEdge);
  }

  // Merge adjacent clusters whose bounding boxes overlap. This handles mixed-size
  // variants (e.g. 40px icons + 120px buttons in the same column) where even
  // chain-based clustering can't bridge the center gap but the variants clearly
  // occupy the same physical column.
  var didMerge = true;
  while (didMerge) {
    didMerge = false;
    for (var mi = 0; mi < bestClusters.length - 1; mi++) {
      if (maxes[mi] >= mins[mi + 1]) {
        if (DEBUG) {
          console.log('[Cluster] Merging overlapping clusters ' + mi + ' (ref=' + bestClusters[mi] +
            ', bounds=[' + mins[mi] + ',' + maxes[mi] + ']) and ' + (mi + 1) + ' (ref=' + bestClusters[mi + 1] +
            ', bounds=[' + mins[mi + 1] + ',' + maxes[mi + 1] + '])');
        }
        mins[mi] = Math.min(mins[mi], mins[mi + 1]);
        maxes[mi] = Math.max(maxes[mi], maxes[mi + 1]);
        bestClusters.splice(mi + 1, 1);
        mins.splice(mi + 1, 1);
        maxes.splice(mi + 1, 1);
        didMerge = true;
        break;
      }
    }
  }

  // After merging, recompute bounding boxes from scratch and compute per-cluster
  // tolerances. Merged clusters may span wider than CLUSTER_TOLERANCE, so
  // findClusterIndex needs per-cluster tolerances to match all member variants.
  var tols = [];
  for (var ti = 0; ti < bestClusters.length; ti++) {
    var tMinEdge = Infinity;
    var tMaxEdge = -Infinity;
    var maxDist = 0;
    for (var tj = 0; tj < variants.length; tj++) {
      var vLeft = variants[tj][posKey];
      var vRight = vLeft + effectiveSize(variants[tj], sizeKey);
      // Variant belongs to this cluster if its position range overlaps the bounding box
      if (vLeft < maxes[ti] + CLUSTER_TOLERANCE && vRight > mins[ti] - CLUSTER_TOLERANCE) {
        if (vLeft < tMinEdge) tMinEdge = vLeft;
        if (vRight > tMaxEdge) tMaxEdge = vRight;
        var dist = Math.abs(bestRefs[tj] - bestClusters[ti]);
        if (dist > maxDist) maxDist = dist;
      }
    }
    if (tMinEdge !== Infinity) { mins[ti] = tMinEdge; maxes[ti] = tMaxEdge; }
    tols.push(Math.max(CLUSTER_TOLERANCE, maxDist));
  }

  return { clusters: bestClusters, mins: mins, maxes: maxes, mode: best, tols: tols };
}

function findClusterIndex(clusters, value, tols) {
  for (var i = 0; i < clusters.length; i++) {
    var tol = tols ? tols[i] : CLUSTER_TOLERANCE;
    if (Math.abs(clusters[i] - value) <= tol) return i;
  }
  return -1;
}

// Get the reference value for a variant based on the axis alignment mode
function getVariantRef(v, posKey, sizeKey, mode) {
  if (mode === 'center') return v[posKey] + effectiveSize(v, sizeKey) / 2;
  if (mode === 'right') return v[posKey] + effectiveSize(v, sizeKey);
  return v[posKey]; // 'left' default
}

function analyzeLayout(componentSet, enumProps, gridAlignment, allowSpanning) {
  const variants = componentSet.children.filter(c => c.type === 'COMPONENT' && c.visible !== false);
  if (DEBUG) {
    console.log('[Step 2] Variants found:', variants.length);
    console.log('[Step 2] Variant positions:', variants.map(v => ({ name: v.name, x: v.x, y: v.y, w: v.width, h: v.height })));
  }

  // Transpose variant grid if the forced alignment doesn't match the physical layout
  if (gridAlignment !== 'auto' && variants.length > 1) {
    const preColCount = clusterVariantAxis(variants, 'x', 'width').clusters.length;
    const preRowCount = clusterVariantAxis(variants, 'y', 'height').clusters.length;

    const needsTranspose =
      (gridAlignment === 'x' && preColCount < preRowCount) ||
      (gridAlignment === 'y' && preRowCount < preColCount);

    if (needsTranspose) {
      if (DEBUG) {
        console.log('[Step 2] Transposing variant grid for alignment:', gridAlignment,
          '(was', preColCount, 'cols ×', preRowCount, 'rows)',
          '| layoutMode:', componentSet.layoutMode);
      }

      if (componentSet.layoutMode === 'VERTICAL') {
        componentSet.layoutMode = 'HORIZONTAL';
      } else if (componentSet.layoutMode === 'HORIZONTAL') {
        componentSet.layoutMode = 'VERTICAL';
      } else {
        // No auto-layout (NONE) — reposition variants into transposed grid
        const preColAxis = clusterVariantAxis(variants, 'x', 'width');
        const preRowAxis = clusterVariantAxis(variants, 'y', 'height');

        // Map each variant to its grid position
        const preGrid = variants.map(v => ({
          node: v,
          col: findClusterIndex(preColAxis.clusters, getVariantRef(v, 'x', 'width', preColAxis.mode), preColAxis.tols),
          row: findClusterIndex(preRowAxis.clusters, getVariantRef(v, 'y', 'height', preRowAxis.mode), preRowAxis.tols),
        }));

        // Determine gap from the axis with multiple items
        let gap = 48;
        if (preRowCount > 1) {
          gap = preRowAxis.mins[1] - preRowAxis.maxes[0];
        } else if (preColCount > 1) {
          gap = preColAxis.mins[1] - preColAxis.maxes[0];
        }
        if (gap < 0) gap = 48;

        const padding = Math.min(preColAxis.mins[0], preRowAxis.mins[0]);
        const newColCount = preRowCount;
        const newRowCount = preColCount;

        // Max variant width per new column (= variants from old row)
        const newColWidths = [];
        for (let nc = 0; nc < newColCount; nc++) {
          const inOldRow = preGrid.filter(g => g.row === nc);
          newColWidths.push(Math.max(...inOldRow.map(g => Math.max(g.node.width, 1))));
        }

        // Max variant height per new row (= variants from old column)
        const newRowHeights = [];
        for (let nr = 0; nr < newRowCount; nr++) {
          const inOldCol = preGrid.filter(g => g.col === nr);
          newRowHeights.push(Math.max(...inOldCol.map(g => Math.max(g.node.height, 1))));
        }

        // Compute new column x-positions (accumulated widths + gap)
        const newColX = [padding];
        for (let nc = 1; nc < newColCount; nc++) {
          newColX.push(newColX[nc - 1] + newColWidths[nc - 1] + gap);
        }

        // Compute new row y-positions
        const newRowY = [padding];
        for (let nr = 1; nr < newRowCount; nr++) {
          newRowY.push(newRowY[nr - 1] + newRowHeights[nr - 1] + gap);
        }

        // Move each variant: old row → new column, old column → new row
        for (const g of preGrid) {
          g.node.x = newColX[g.row];
          g.node.y = newRowY[g.col];
        }

        // Resize component set
        const lastCol = newColCount - 1;
        const lastRow = newRowCount - 1;
        componentSet.resize(
          newColX[lastCol] + newColWidths[lastCol] + padding,
          newRowY[lastRow] + newRowHeights[lastRow] + padding
        );

        if (DEBUG) {
          console.log('[Step 2] Transposed to', newColCount, 'cols ×', newRowCount, 'rows',
            '| gap:', gap, '| padding:', padding, '| CS:', componentSet.width, '×', componentSet.height);
        }
      }
    }
  }

  // Cluster by detected alignment (left, center, or right — whichever gives fewest clusters)
  let colAxis = clusterVariantAxis(variants, 'x', 'width');
  let rowAxis = clusterVariantAxis(variants, 'y', 'height');
  let colClusters = colAxis.mins;   // left edge of each column
  let rowClusters = rowAxis.mins;   // top edge of each row
  let colRefs = colAxis.clusters;   // reference values for column matching
  let rowRefs = rowAxis.clusters;   // reference values for row matching
  if (DEBUG) {
    console.log('[Step 2] Column clusters (' + colClusters.length + ' cols, mode=' + colAxis.mode + '):', colClusters);
    console.log('[Step 2] Row clusters (' + rowClusters.length + ' rows, mode=' + rowAxis.mode + '):', rowClusters);
    console.log('[Step 2] Grid dimensions:', colClusters.length, 'cols ×', rowClusters.length, 'rows');
  }

  // Build grid: map each variant to its (col, row) using detected alignment
  let grid = variants.map(v => ({
    node: v,
    col: findClusterIndex(colRefs, getVariantRef(v, 'x', 'width', colAxis.mode), colAxis.tols),
    row: findClusterIndex(rowRefs, getVariantRef(v, 'y', 'height', rowAxis.mode), rowAxis.tols),
    props: Object.fromEntries(
      Object.entries(v.variantProperties).map(([k, val]) => [k, val])
    ),
    colSpan: 1,
    rowSpan: 1,
  }));

  // Pre-index grid by row and column for fast lookups
  var gridByRow = {};
  var gridByCol = {};
  for (var gi = 0; gi < grid.length; gi++) {
    var g = grid[gi];
    if (!gridByRow[g.row]) gridByRow[g.row] = [];
    gridByRow[g.row].push(g);
    if (!gridByCol[g.col]) gridByCol[g.col] = [];
    gridByCol[g.col].push(g);
  }

  if (DEBUG) {
    console.log('[Step 2] Grid mapping:', grid.map(g => ({ name: g.node.name, col: g.col, row: g.row, props: g.props })));
  }

  // For each enum property, determine axis
  const colAxisProps = [];
  const rowAxisProps = [];

  for (const prop of enumProps) {
    let variesAlongX = false;
    let variesAlongY = false;

    // Check if property varies across columns within same row
    for (let r = 0; r < rowClusters.length; r++) {
      const inRow = gridByRow[r] || [];
      const valuesInRow = new Set(inRow.map(g => g.props[prop.name]));
      if (valuesInRow.size > 1) variesAlongX = true;
    }

    // Check if property varies across rows within same column
    for (let c = 0; c < colClusters.length; c++) {
      const inCol = gridByCol[c] || [];
      const valuesInCol = new Set(inCol.map(g => g.props[prop.name]));
      if (valuesInCol.size > 1) variesAlongY = true;
    }

    // Diagonal layout: each variant at unique x AND y, so no variation detected within rows/cols.
    // Fall back to checking overall spread to pick the dominant axis.
    if (!variesAlongX && !variesAlongY && prop.values.length > 1) {
      const xSpread = colClusters[colClusters.length - 1] - colClusters[0];
      const ySpread = rowClusters[rowClusters.length - 1] - rowClusters[0];
      if (xSpread >= ySpread) {
        variesAlongX = true;
      } else {
        variesAlongY = true;
      }
      if (DEBUG) {
        console.log('[Step 2] Property "' + prop.name + '": diagonal layout detected, xSpread=' + xSpread + ', ySpread=' + ySpread + ' → fallback to ' + (variesAlongX ? 'COLUMN' : 'ROW') + ' axis');
      }
    }

    if (DEBUG) {
      console.log('[Step 2] Property "' + prop.name + '": variesAlongX=' + variesAlongX + ', variesAlongY=' + variesAlongY + ' → ' + (
        variesAlongX && !variesAlongY ? 'COLUMN axis' :
        variesAlongY && !variesAlongX ? 'ROW axis' :
        variesAlongX && variesAlongY ? 'BOTH (defaulting to ROW)' :
        'ROW axis (single value)'
      ));
    }

    if (variesAlongX && !variesAlongY) {
      colAxisProps.push(prop);
    } else if (variesAlongY && !variesAlongX) {
      rowAxisProps.push(prop);
    } else if (variesAlongX && variesAlongY) {
      // Ambiguous — put on rows by default
      rowAxisProps.push(prop);
      prop._bothAxes = true;
    } else {
      // Single value — still show as a row label (left side)
      rowAxisProps.push(prop);
    }
  }

  // Detect single-type grid: one property spread across a multi-row × multi-col grid.
  // Instead of forcing it onto one axis (which produces nonsensical labels), flag it
  // so the generator uses per-item labels below each variant.
  //
  // Also handles spanning variants: when a variant is wider/taller than one cell,
  // clusterVariantAxis merges overlapping bounding boxes and collapses columns.
  // We use raw position clustering (clusterValues) to recover the true grid structure.
  var singleTypeGrid = false;
  if (allowSpanning && enumProps.length === 1 && variants.length > 1) {
    // Check raw positions — clusterValues doesn't do bounding-box overlap merge
    var rawColResult = clusterValues(variants.map(function(v) { return v.x; }));
    var rawRowResult = clusterValues(variants.map(function(v) { return v.y; }));
    var rawColCount = rawColResult.clusters.length;
    var rawRowCount = rawRowResult.clusters.length;

    if (rawColCount > 1 && rawRowCount > 1) {
      singleTypeGrid = true;
      colAxisProps.length = 0;
      rowAxisProps.length = 0;

      // If overlap merge collapsed columns/rows, rebuild grid from raw positions
      if (rawColCount > colClusters.length || rawRowCount > rowClusters.length) {
        if (DEBUG) { console.log('[Step 2] Spanning variant detected — rebuilding grid from raw positions (' + rawColCount + ' cols × ' + rawRowCount + ' rows, was ' + colClusters.length + ' × ' + rowClusters.length + ')'); }

        // Rebuild column data from raw position clusters
        var rawColClusters = rawColResult.clusters; // left edges
        var rawColHighs = rawColResult.highs;
        // Assign each variant to its starting column (closest left-edge cluster)
        // and compute column bounds from non-spanning variants
        var rawColMins = [];
        var rawColMaxes = [];
        var rawColTols = [];
        for (var rci = 0; rci < rawColClusters.length; rci++) {
          rawColMins.push(Infinity);
          rawColMaxes.push(-Infinity);
          rawColTols.push(CLUSTER_TOLERANCE);
        }

        // Rebuild row data similarly
        var rawRowClusters = rawRowResult.clusters;
        var rawRowHighs = rawRowResult.highs;
        var rawRowMins = [];
        var rawRowMaxes = [];
        var rawRowTols = [];
        for (var rri = 0; rri < rawRowClusters.length; rri++) {
          rawRowMins.push(Infinity);
          rawRowMaxes.push(-Infinity);
          rawRowTols.push(CLUSTER_TOLERANCE);
        }

        // Find column/row for each variant and detect spanning
        var rawGrid = [];
        for (var rvi = 0; rvi < variants.length; rvi++) {
          var rv = variants[rvi];
          var rvLeft = rv.x;
          var rvRight = rv.x + rv.width;
          var rvTop = rv.y;
          var rvBottom = rv.y + rv.height;

          // Find starting column (closest cluster to variant's left edge)
          var rvCol = 0;
          var rvColDist = Math.abs(rvLeft - rawColClusters[0]);
          for (var rcci = 1; rcci < rawColClusters.length; rcci++) {
            var d = Math.abs(rvLeft - rawColClusters[rcci]);
            if (d < rvColDist) { rvCol = rcci; rvColDist = d; }
          }

          // Find starting row
          var rvRow = 0;
          var rvRowDist = Math.abs(rvTop - rawRowClusters[0]);
          for (var rrci = 1; rrci < rawRowClusters.length; rrci++) {
            var d2 = Math.abs(rvTop - rawRowClusters[rrci]);
            if (d2 < rvRowDist) { rvRow = rrci; rvRowDist = d2; }
          }

          // Detect column span: how many columns does this variant cover?
          var rvColSpan = 1;
          for (var rcs = rvCol + 1; rcs < rawColClusters.length; rcs++) {
            if (rvRight > rawColClusters[rcs] + CLUSTER_TOLERANCE) {
              rvColSpan++;
            } else {
              break;
            }
          }

          // Detect row span
          var rvRowSpan = 1;
          for (var rrs = rvRow + 1; rrs < rawRowClusters.length; rrs++) {
            if (rvBottom > rawRowClusters[rrs] + CLUSTER_TOLERANCE) {
              rvRowSpan++;
            } else {
              break;
            }
          }

          rawGrid.push({
            node: rv,
            col: rvCol,
            row: rvRow,
            colSpan: rvColSpan,
            rowSpan: rvRowSpan,
            props: Object.fromEntries(
              Object.entries(rv.variantProperties).map(function(e) { return [e[0], e[1]]; })
            ),
          });

          // Update column bounds (only from non-spanning variants for accurate widths)
          if (rvColSpan === 1) {
            if (rvLeft < rawColMins[rvCol]) rawColMins[rvCol] = rvLeft;
            if (rvRight > rawColMaxes[rvCol]) rawColMaxes[rvCol] = rvRight;
          }
          if (rvRowSpan === 1) {
            if (rvTop < rawRowMins[rvRow]) rawRowMins[rvRow] = rvTop;
            if (rvBottom > rawRowMaxes[rvRow]) rawRowMaxes[rvRow] = rvBottom;
          }
        }

        // Fill in any columns/rows that only have spanning variants (use cluster position + tolerance)
        for (var rfci = 0; rfci < rawColClusters.length; rfci++) {
          if (rawColMins[rfci] === Infinity) {
            rawColMins[rfci] = rawColClusters[rfci];
            rawColMaxes[rfci] = rawColClusters[rfci] + 50; // fallback width
          }
        }
        for (var rfri = 0; rfri < rawRowClusters.length; rfri++) {
          if (rawRowMins[rfri] === Infinity) {
            rawRowMins[rfri] = rawRowClusters[rfri];
            rawRowMaxes[rfri] = rawRowClusters[rfri] + 50;
          }
        }

        // Replace clustering data
        colClusters = rawColMins;
        colRefs = rawColClusters;
        colAxis = { clusters: rawColClusters, mins: rawColMins, maxes: rawColMaxes, mode: 'left', tols: rawColTols };
        rowClusters = rawRowMins;
        rowRefs = rawRowClusters;
        rowAxis = { clusters: rawRowClusters, mins: rawRowMins, maxes: rawRowMaxes, mode: 'left', tols: rawRowTols };

        // Replace grid and indices
        grid = rawGrid;
        gridByRow = {};
        gridByCol = {};
        for (var rgii = 0; rgii < grid.length; rgii++) {
          var rg = grid[rgii];
          if (!gridByRow[rg.row]) gridByRow[rg.row] = [];
          gridByRow[rg.row].push(rg);
          if (!gridByCol[rg.col]) gridByCol[rg.col] = [];
          gridByCol[rg.col].push(rg);
        }

        if (DEBUG) {
          console.log('[Step 2] Rebuilt grid with spanning:', grid.map(function(g) {
            return { name: g.node.name, col: g.col, row: g.row, colSpan: g.colSpan, rowSpan: g.rowSpan };
          }));
        }
      }

      if (DEBUG) { console.log('[Step 2] Single-type grid detected: "' + enumProps[0].name + '" with ' + enumProps[0].values.length + ' values in ' + rawColCount + '×' + rawRowCount + ' grid'); }
    }
  }

  // Determine ordering: outer properties change less frequently
  orderByFrequency(colAxisProps, grid, colClusters, 'col');
  orderByFrequency(rowAxisProps, grid, rowClusters, 'row');

  if (DEBUG) {
    console.log('[Step 2] Final column-axis properties (top labels):', colAxisProps.map(function(p) { return p.name; }));
    console.log('[Step 2] Final row-axis properties (left labels):', rowAxisProps.map(function(p) { return p.name; }));
  }

  return { grid, gridByRow, gridByCol, colClusters, rowClusters, colAxisProps, rowAxisProps, variants, colMaxes: colAxis.maxes, rowMaxes: rowAxis.maxes, colTols: colAxis.tols, rowTols: rowAxis.tols, singleTypeGrid: singleTypeGrid, singleTypeProp: singleTypeGrid ? enumProps[0] : null };
}

function orderByFrequency(axisProps, grid, clusters, axis) {
  // Sort so that properties with fewer changes (larger groups) come first (outer)
  axisProps.sort((a, b) => {
    const changesA = countChanges(a.name, grid, clusters, axis);
    const changesB = countChanges(b.name, grid, clusters, axis);
    return changesA - changesB; // fewer changes = outer
  });
}

function countChanges(propName, grid, clusters, axis) {
  // Count how many times the value changes along the axis
  let changes = 0;
  const key = axis === 'col' ? 'col' : 'row';
  const otherKey = axis === 'col' ? 'row' : 'col';

  // Pick the first "line" along the other axis
  const firstOther = grid.reduce((min, g) => Math.min(min, g[otherKey]), Infinity);
  const line = grid
    .filter(g => g[otherKey] === firstOther)
    .sort((a, b) => a[key] - b[key]);

  for (let i = 1; i < line.length; i++) {
    if (line[i].props[propName] !== line[i - 1].props[propName]) {
      changes++;
    }
  }
  return changes;
}

// ─── Measure Text Width (Step 5) ─────────────────────────────────────────────

function initMeasureNode() {
  _measureNode = figma.createText();
  _measureNode.fontName = { family: _labelFontFamily, style: 'Regular' };
  _measureNode.fontSize = LABEL_FONT_SIZE;
  _measureCache = {};
}

function disposeMeasureNode() {
  if (_measureNode) {
    _measureNode.remove();
    _measureNode = null;
  }
  _measureCache = {};
}

function measureTextWidth(text, fontSize) {
  var key = text + '|' + fontSize;
  if (_measureCache[key] !== undefined) return _measureCache[key];
  if (_measureNode.fontSize !== fontSize) {
    _measureNode.fontSize = fontSize;
  }
  _measureNode.characters = text;
  var w = _measureNode.width;
  _measureCache[key] = w;
  return w;
}

// ─── Bracket Creation (Step 7) ───────────────────────────────────────────────

function createVerticalBracket(height) {
  var vec = figma.createVector();
  vec.name = 'Bracket';
  // C-shape: top cap → spine → bottom cap
  vec.vectorPaths = [{
    windingRule: 'NONZERO',
    data: 'M ' + BRACKET_CAP_LENGTH + ' 0 L 0 0 L 0 ' + height + ' L ' + BRACKET_CAP_LENGTH + ' ' + height
  }];
  vec.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
  vec.strokeWeight = BRACKET_STROKE_WEIGHT;
  vec.fills = [];
  return vec;
}

function createHorizontalBracket(width) {
  var vec = figma.createVector();
  vec.name = 'Bracket';
  // U-shape: left cap down → spine across → right cap down
  vec.vectorPaths = [{
    windingRule: 'NONZERO',
    data: 'M 0 ' + BRACKET_CAP_LENGTH + ' L 0 0 L ' + width + ' 0 L ' + width + ' ' + BRACKET_CAP_LENGTH
  }];
  vec.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
  vec.strokeWeight = BRACKET_STROKE_WEIGHT;
  vec.fills = [];
  return vec;
}

// ─── Label Creation (Step 6b, 6c) ────────────────────────────────────────────

function createTextNode(text, fontSize) {
  const node = figma.createText();
  node.fontName = { family: _labelFontFamily, style: 'Regular' };
  node.fontSize = fontSize || LABEL_FONT_SIZE;
  node.characters = text;
  node.fills = [{ type: 'SOLID', color: DOC_COLOR }];
  return node;
}

function createRowLabel(value, x, y, width, height, withBracket) {
  const label = figma.createFrame();
  label.name = 'Label';
  label.resize(width, height);
  label.x = x;
  label.y = y;
  label.fills = [];
  label.clipsContent = false;

  const textNode = createTextNode(value);
  label.appendChild(textNode);

  if (withBracket) {
    const bracketHeight = height - LABEL_PADDING * 2;
    const bracket = createVerticalBracket(bracketHeight > 0 ? bracketHeight : height);
    bracket.x = width - BRACKET_THICKNESS;
    bracket.y = LABEL_PADDING;
    label.appendChild(bracket);

    // Center text vertically, right-align to the left of the bracket
    textNode.x = Math.max(0, width - BRACKET_THICKNESS - LABEL_GAP - textNode.width);
    textNode.y = (height - textNode.height) / 2;
  } else {
    // Center text vertically, right-align with spacing from grid edge
    textNode.x = Math.max(0, width - textNode.width - 6);
    textNode.y = (height - textNode.height) / 2;
  }

  return label;
}

function createColLabel(value, x, y, width, height, withBracket) {
  const label = figma.createFrame();
  label.name = 'Label';
  label.resize(width, height);
  label.x = x;
  label.y = y;
  label.fills = [];
  label.clipsContent = false;

  const textNode = createTextNode(value);
  label.appendChild(textNode);

  if (withBracket) {
    const bracketWidth = width - LABEL_PADDING * 2;
    const bracket = createHorizontalBracket(bracketWidth > 0 ? bracketWidth : width);
    bracket.x = LABEL_PADDING;
    bracket.y = height - BRACKET_THICKNESS;
    label.appendChild(bracket);

    // Center text horizontally, position above the bracket
    textNode.x = (width - textNode.width) / 2;
    textNode.y = 0;
  } else {
    // Center text horizontally
    textNode.x = (width - textNode.width) / 2;
    textNode.y = (height - textNode.height) / 2;
  }

  return label;
}

// ─── Per-Item Labels (single-type grid) ─────────────────────────────────────

function createItemLabel(value, x, y, cellWidth) {
  var label = figma.createFrame();
  label.name = 'Label';
  label.resize(cellWidth, ITEM_LABEL_HEIGHT);
  label.x = x;
  label.y = y;
  label.fills = [];
  label.clipsContent = false;

  var textNode = createTextNode(value);
  textNode.x = (cellWidth - textNode.width) / 2;
  textNode.y = (ITEM_LABEL_HEIGHT - textNode.height) / 2;
  label.appendChild(textNode);

  return label;
}

// Build uniform grid layout for single-type grids. Computes cell dimensions,
// updates colClusters/colWidths/rowClusters/rowHeights in-place, and returns
// { cellW, cellH, maxVarH } for downstream use.
// positionFn(gridEntry, x, y) is called for each variant with its new position.
function buildSingleTypeUniformGrid(grid, gridByCol, colClusters, colWidths, rowClusters, rowHeights, singleTypeProp, positionFn) {
  var pad = 24;
  var innerLabelH = ITEM_LABEL_GAP + ITEM_LABEL_HEIGHT;

  // Compute uniform cell width: max of non-spanning variant widths and label widths
  var cellW = 0;
  for (var cw = 0; cw < colWidths.length; cw++) {
    if (colWidths[cw] > cellW) cellW = colWidths[cw];
  }
  for (var c = 0; c < colClusters.length; c++) {
    var inCol = gridByCol[c] || [];
    for (var ci = 0; ci < inCol.length; ci++) {
      var lw = measureTextWidth(inCol[ci].props[singleTypeProp.name], LABEL_FONT_SIZE) + LABEL_GAP * 2;
      var effLW = inCol[ci].colSpan > 1 ? lw / inCol[ci].colSpan : lw;
      if (effLW > cellW) cellW = effLW;
    }
  }

  var colGap = pad;
  var maxVarH = Math.max.apply(null, rowHeights);
  var cellH = maxVarH + innerLabelH;
  var rowGap = pad;

  // Update column positions
  for (var c2 = 0; c2 < colClusters.length; c2++) {
    colClusters[c2] = pad + c2 * (cellW + colGap);
    colWidths[c2] = cellW;
  }

  // Update row positions — cell height includes label space
  for (var r = 0; r < rowClusters.length; r++) {
    rowClusters[r] = pad + r * (cellH + rowGap);
    rowHeights[r] = cellH;
  }

  // Reposition variants (handling spanning)
  for (var gi = 0; gi < grid.length; gi++) {
    var g = grid[gi];
    var vw = g.node.width;
    var vh = g.node.height;
    var varAreaCenter = rowClusters[g.row] + maxVarH / 2;
    var newX, newY = varAreaCenter - vh / 2;

    if (g.colSpan > 1) {
      var spanLeft = colClusters[g.col];
      var spanRight = colClusters[g.col + g.colSpan - 1] + cellW;
      newX = (spanLeft + spanRight) / 2 - vw / 2;
    } else {
      newX = colClusters[g.col] + cellW / 2 - vw / 2;
    }
    positionFn(g, newX, newY);
  }

  return { cellW: cellW, cellH: cellH, maxVarH: maxVarH, pad: pad };
}

// Create per-item labels for a single-type grid, appended to parentFrame.
function createSingleTypeGridLabels(parentFrame, grid, singleTypeProp, colClusters, colWidths, rowClusters, maxVarH, offsetX, offsetY) {
  for (var i = 0; i < grid.length; i++) {
    var g = grid[i];
    var value = g.props[singleTypeProp.name];
    var cellLeft = colClusters[g.col];
    var labelW;
    if (g.colSpan > 1) {
      var spanRight = colClusters[g.col + g.colSpan - 1] + colWidths[g.col + g.colSpan - 1];
      labelW = spanRight - cellLeft;
    } else {
      labelW = colWidths[g.col];
    }
    var labelX = offsetX + cellLeft;
    var labelY = offsetY + rowClusters[g.row] + maxVarH + ITEM_LABEL_GAP;
    parentFrame.appendChild(createItemLabel(value, labelX, labelY, labelW));
  }
}

// ─── Grid Lines ──────────────────────────────────────────────────────────────

function createGridLines(wrapper, colClusters, rowClusters, colWidths, rowHeights, csX, csY, csWidth, csHeight, spanningGrid) {
  const gridFrame = figma.createFrame();
  gridFrame.name = 'Grid';
  gridFrame.fills = [];
  gridFrame.clipsContent = false;

  // Grid covers the full component set area
  gridFrame.resize(csWidth, csHeight);
  gridFrame.x = csX;
  gridFrame.y = csY;

  // Build a set of rows that each vertical column-boundary should skip (due to spanning)
  // spanningGrid is an array of { col, row, colSpan, rowSpan } entries
  var colSkipRows = null;
  if (spanningGrid) {
    colSkipRows = {}; // colBoundary index → set of row indices to skip
    for (var sgi = 0; sgi < spanningGrid.length; sgi++) {
      var sg = spanningGrid[sgi];
      if (sg.colSpan > 1) {
        // This variant spans from sg.col to sg.col + sg.colSpan - 1
        // Skip vertical lines at column boundaries within this span
        for (var sbc = sg.col + 1; sbc < sg.col + sg.colSpan; sbc++) {
          if (!colSkipRows[sbc]) colSkipRows[sbc] = {};
          for (var sbr = sg.row; sbr < sg.row + sg.rowSpan; sbr++) {
            colSkipRows[sbc][sbr] = true;
          }
        }
      }
    }
  }

  // Inner horizontal lines only (skip first and last — no outer border)
  for (let r = 1; r < rowClusters.length; r++) {
    const prevRowBottom = rowClusters[r - 1] + rowHeights[r - 1];
    const midY = (prevRowBottom + rowClusters[r]) / 2;
    const line = figma.createLine();
    line.resize(csWidth, 0);
    line.x = 0;
    line.y = midY;
    line.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
    line.strokeWeight = GRID_STROKE_WEIGHT;
    line.dashPattern = GRID_DASH_PATTERN;
    gridFrame.appendChild(line);
  }

  // Inner vertical lines (skip first and last — no outer border)
  for (let c = 1; c < colClusters.length; c++) {
    const prevColRight = colClusters[c - 1] + colWidths[c - 1];
    const midX = (prevColRight + colClusters[c]) / 2;

    if (colSkipRows && colSkipRows[c]) {
      // Build skip regions: for each skipped row, the vertical line must be absent
      // from the midpoint above the row to the midpoint below (or grid edge for first/last).
      // This matches where horizontal grid lines are drawn.
      var skipRanges = [];
      for (var vr = 0; vr < rowClusters.length; vr++) {
        if (!colSkipRows[c][vr]) continue;
        var skipTop = vr === 0 ? 0 :
          (rowClusters[vr - 1] + rowHeights[vr - 1] + rowClusters[vr]) / 2;
        var skipBottom = vr === rowClusters.length - 1 ? csHeight :
          (rowClusters[vr] + rowHeights[vr] + rowClusters[vr + 1]) / 2;
        // Merge with previous range if overlapping
        if (skipRanges.length > 0 && skipTop <= skipRanges[skipRanges.length - 1].bottom) {
          skipRanges[skipRanges.length - 1].bottom = skipBottom;
        } else {
          skipRanges.push({ top: skipTop, bottom: skipBottom });
        }
      }

      // Draw segments in the non-skipped regions
      var drawStart = 0;
      for (var sri = 0; sri < skipRanges.length; sri++) {
        if (skipRanges[sri].top > drawStart) {
          var segLine = figma.createLine();
          segLine.rotation = -90;
          segLine.resize(skipRanges[sri].top - drawStart, 0);
          segLine.x = midX;
          segLine.y = drawStart;
          segLine.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
          segLine.strokeWeight = GRID_STROKE_WEIGHT;
          segLine.dashPattern = GRID_DASH_PATTERN;
          gridFrame.appendChild(segLine);
        }
        drawStart = skipRanges[sri].bottom;
      }
      // Final segment after last skip
      if (drawStart < csHeight) {
        var segLineEnd = figma.createLine();
        segLineEnd.rotation = -90;
        segLineEnd.resize(csHeight - drawStart, 0);
        segLineEnd.x = midX;
        segLineEnd.y = drawStart;
        segLineEnd.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
        segLineEnd.strokeWeight = GRID_STROKE_WEIGHT;
        segLineEnd.dashPattern = GRID_DASH_PATTERN;
        gridFrame.appendChild(segLineEnd);
      }
    } else {
      // Full-height vertical line (no spanning in this column boundary)
      const line = figma.createLine();
      line.rotation = -90;
      line.resize(csHeight, 0);
      line.x = midX;
      line.y = 0;
      line.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
      line.strokeWeight = GRID_STROKE_WEIGHT;
      line.dashPattern = GRID_DASH_PATTERN;
      gridFrame.appendChild(line);
    }
  }

  wrapper.appendChild(gridFrame);
  gridFrame.locked = true;
  return gridFrame;
}

// ─── Remove Old Labels (Step 9) ──────────────────────────────────────────────

// Derive doc generation options from existing wrapper layer names
function deriveDocMeta(wrapper) {
  var meta = { modes: [], booleanProps: [], nestedInstances: [], hasBooleanGrid: false };

  function extractSectionHeaders(sectionNode) {
    // Extract header text from label text nodes inside a boolean/nested section
    var headers = [];
    if (!sectionNode || !sectionNode.children) return headers;
    for (var i = 0; i < sectionNode.children.length; i++) {
      var child = sectionNode.children[i];
      // Header text nodes are direct TEXT children with the property values
      if (child.type === 'TEXT' && child.characters) {
        var text = child.characters.trim();
        // Extract the property name from "PropName: Value" format
        var colonIdx = text.indexOf(': ');
        if (colonIdx !== -1) {
          var propName = text.substring(0, colonIdx);
          if (headers.indexOf(propName) === -1) headers.push(propName);
        } else if (text.length > 0 && text.length < 60) {
          if (headers.indexOf(text) === -1) headers.push(text);
        }
      }
    }
    return headers;
  }

  function scan(node) {
    if (!node || !node.children) return;
    for (var i = 0; i < node.children.length; i++) {
      var child = node.children[i];
      var name = child.name || '';

      // Variable modes: "Mode: <name>" or "Mode: <name> (collection1, collection2)"
      if (name.indexOf('Mode: ') === 0) {
        meta.modes.push(name.substring(6));
      }

      // Boolean visibility section — extract property names from headers
      if (name === 'Boolean Visibility') {
        var boolHeaders = extractSectionHeaders(child);
        for (var bh = 0; bh < boolHeaders.length; bh++) {
          if (meta.booleanProps.indexOf(boolHeaders[bh]) === -1) meta.booleanProps.push(boolHeaders[bh]);
        }
      }

      // Boolean grid
      if (name === 'Boolean Grid') meta.hasBooleanGrid = true;

      // Nested instances section — extract instance names from headers
      if (name === 'Nested Instances') {
        var nestedHeaders = extractSectionHeaders(child);
        for (var nh = 0; nh < nestedHeaders.length; nh++) {
          if (meta.nestedInstances.indexOf(nestedHeaders[nh]) === -1) meta.nestedInstances.push(nestedHeaders[nh]);
        }
      }

      // Recurse into containers that hold sub-elements
      if (name === 'Documentation' || name === 'Modes' || name === 'Property Combinations') {
        scan(child);
      }
    }
  }
  scan(wrapper);
  return meta;
}

function findExistingWrapper(componentSet) {
  // Check if the component set is already inside a wrapper frame
  const parent = componentSet.parent;
  if (parent && parent.type === 'FRAME' && parent.name.startsWith('❖ ')) {
    // If wrapper has a stored source ID, verify it matches this CS
    var storedId = parent.getPluginData('sourceComponentSetId');
    if (storedId && storedId !== componentSet.id) return null;
    // Retroactively store ID for old wrappers (pre-fix) so duplicates are caught
    if (!storedId) parent.setPluginData('sourceComponentSetId', componentSet.id);
    return parent;
  }
  return null;
}

// Find a standalone doc wrapper — a sibling frame with pluginData 'standaloneDoc'
function findStandaloneWrapper(node) {
  // Search siblings (legacy placement inside the same parent)
  var parent = node.parent;
  if (!parent) return null;
  for (var i = 0; i < parent.children.length; i++) {
    var sibling = parent.children[i];
    if (sibling !== node && sibling.type === 'FRAME' && sibling.getPluginData('standaloneDoc') === 'true') {
      var storedId = sibling.getPluginData('sourceComponentSetId');
      if (storedId && storedId !== node.id) continue;
      return sibling;
    }
  }
  // Search grandparent's children (new placement: next to the parent frame)
  var grandparent = parent.parent;
  if (!grandparent) return null;
  for (var gi = 0; gi < grandparent.children.length; gi++) {
    var gSibling = grandparent.children[gi];
    if (gSibling !== parent && gSibling.type === 'FRAME' && gSibling.getPluginData('standaloneDoc') === 'true') {
      var gStoredId = gSibling.getPluginData('sourceComponentSetId');
      if (gStoredId && gStoredId !== node.id) continue;
      return gSibling;
    }
  }
  return null;
}

// Find a variable modes doc wrapper — tagged with 'variableModesDoc' plugin data.
// Searches the wrapper's parent (page level) since VM docs sit next to the regular wrapper.
function findVariableModesWrapper(cs) {
  var wrapper = findExistingWrapper(cs);
  var searchParents = [];
  // Search siblings of the regular wrapper (page level)
  if (wrapper && wrapper.parent) searchParents.push(wrapper.parent);
  // Also search siblings of CS itself and its parent's parent
  if (cs.parent) {
    searchParents.push(cs.parent);
    if (cs.parent.parent) searchParents.push(cs.parent.parent);
  }
  for (var spi = 0; spi < searchParents.length; spi++) {
    var sp = searchParents[spi];
    for (var i = 0; i < sp.children.length; i++) {
      var child = sp.children[i];
      if (child.type === 'FRAME' && child.getPluginData('variableModesDoc') === 'true') {
        var storedId = child.getPluginData('sourceComponentSetId');
        if (storedId && storedId !== cs.id) continue;
        return child;
      }
    }
  }
  return null;
}

// ─── GusPropstar Detection & Removal ─────────────────────────────────────────

// Detect competing GusPropstar docs around a component set.
// Key signals:
// - Grid auto layout (layoutWrap === 'WRAP') — GusPropstar uses this, we do not
// - GROUP named "labels" or FRAME named "instances" — GusPropstar artifacts
// Old format: sibling frame with empty name containing GROUP "labels" + FRAME "instances"
// New format: parent ❖ frame using grid auto layout, or containing "labels"/"instances"
function detectGusPropstar(componentSet) {
  var parent = componentSet.parent;
  if (!parent) return null;

  // Check if inside a ❖ wrapper that is a GusPropstar frame
  if (parent.type === 'FRAME' && parent.name.startsWith('❖')) {
    // Grid layout is the strongest signal — we never use it
    if (parent.layoutMode === 'GRID') {
      return 'new';
    }
    // Also check for GusPropstar artifacts (GROUP "labels" or FRAME "instances")
    for (var i = 0; i < parent.children.length; i++) {
      var child = parent.children[i];
      if (child.type === 'GROUP' && child.name.toLowerCase() === 'labels') return 'new';
      if (child.type === 'FRAME' && child.name.toLowerCase() === 'instances') return 'new';
    }
  }

  // Check for unnamed wrapper: parent frame (any name) directly contains
  // GROUP "labels" or FRAME "instances" as siblings of the component set
  if (parent.type === 'FRAME' && parent.children) {
    for (var n = 0; n < parent.children.length; n++) {
      var ch = parent.children[n];
      if (ch === componentSet) continue;
      if (ch.type === 'GROUP' && ch.name.toLowerCase() === 'labels') return 'new';
      if (ch.type === 'FRAME' && ch.name.toLowerCase() === 'instances') return 'new';
    }
  }

  // Check for old format: sibling empty-name frame with "labels" GROUP / "instances" FRAME
  if (parent.children) {
    for (var j = 0; j < parent.children.length; j++) {
      var sibling = parent.children[j];
      if (sibling === componentSet) continue;
      if (sibling.type === 'FRAME' && (sibling.name === '' || sibling.name === ' ')) {
        for (var k = 0; k < sibling.children.length; k++) {
          var sc = sibling.children[k];
          if (sc.type === 'GROUP' && sc.name.toLowerCase() === 'labels') return 'old';
          if (sc.type === 'FRAME' && sc.name.toLowerCase() === 'instances') return 'old';
        }
      }
    }
  }

  return null;
}

function removeGusPropstar(componentSet) {
  var format = detectGusPropstar(componentSet);
  if (!format) {
    figma.ui.postMessage({ type: 'error', message: 'No GusPropstar docs found.' });
    return;
  }

  var parent = componentSet.parent;

  if (format === 'new') {
    // Component set is inside a ❖ wrapper — remove all non-component children, then unwrap
    var wrapper = parent;
    var wrapperParent = wrapper.parent;
    var wrapperIndex = 0;
    for (var i = 0; i < wrapperParent.children.length; i++) {
      if (wrapperParent.children[i] === wrapper) { wrapperIndex = i; break; }
    }
    var wrapperX = wrapper.x;
    var wrapperY = wrapper.y;

    // Remove all non-component children (labels, instances, grids, etc.)
    var children = Array.from(wrapper.children);
    for (var j = 0; j < children.length; j++) {
      if (children[j].type !== 'COMPONENT_SET' && children[j].type !== 'COMPONENT') {
        children[j].remove();
      }
    }

    // Move component set out of wrapper
    wrapperParent.insertChild(wrapperIndex, componentSet);
    componentSet.x = wrapperX;
    componentSet.y = wrapperY;

    // Reset absolute positioning if needed
    if (componentSet.layoutPositioning === 'ABSOLUTE') {
      componentSet.layoutPositioning = 'AUTO';
    }

    // Remove the now-empty wrapper
    wrapper.remove();
  }

  if (format === 'old') {
    // Remove the sibling empty-name doc frame
    for (var k = 0; k < parent.children.length; k++) {
      var sibling = parent.children[k];
      if (sibling === componentSet) continue;
      if (sibling.type === 'FRAME' && (sibling.name === '' || sibling.name === ' ')) {
        var hasArtifacts = false;
        for (var m = 0; m < sibling.children.length; m++) {
          var sc = sibling.children[m];
          if (sc.type === 'GROUP' && sc.name.toLowerCase() === 'labels') hasArtifacts = true;
          if (sc.type === 'FRAME' && sc.name.toLowerCase() === 'instances') hasArtifacts = true;
        }
        if (hasArtifacts) {
          sibling.remove();
          break;
        }
      }
    }
  }

  // Reset component set stroke to default Figma purple with dashes
  componentSet.strokes = [{ type: 'SOLID', color: { r: 0.592, g: 0.278, b: 1.0 }, opacity: 1 }];
  componentSet.dashPattern = [4, 4];

  figma.currentPage.selection = [componentSet];
  figma.viewport.scrollAndZoomIntoView([componentSet]);
  figma.ui.postMessage({ type: 'done', message: 'GusPropstar docs removed.' });
  sendSelectionInfo();
}

function removeDocs() {
  var cs = getComponentSet();
  var node = cs;
  if (!node) {
    node = getStandaloneComponent();
  }
  if (!node) {
    figma.ui.postMessage({ type: 'error', message: 'Please select a component set or component.' });
    return;
  }

  // Standalone docs are separate frames — skip them, they must be deleted manually
  var standaloneWrapper = findStandaloneWrapper(node);
  if (standaloneWrapper && !findExistingWrapper(node)) {
    figma.ui.postMessage({ type: 'error', message: 'Standalone docs cannot be removed via the plugin. Delete the frame manually.' });
    return;
  }

  const wrapper = findExistingWrapper(node);
  if (!wrapper) {
    figma.ui.postMessage({ type: 'error', message: 'No docs found to remove.' });
    return;
  }

  // Restore CS strokes if they were saved (boolean grid mode removes them)
  var savedStrokes = node.getPluginData('originalStrokes');
  if (savedStrokes) {
    try {
      var strokeData = JSON.parse(savedStrokes);
      node.strokes = strokeData.strokes;
      node.strokeWeight = strokeData.strokeWeight;
      node.dashPattern = strokeData.dashPattern;
      node.strokeAlign = strokeData.strokeAlign;
      if (strokeData.cornerRadius !== undefined) node.cornerRadius = strokeData.cornerRadius;
    } catch (e) {
      console.log('[removeDocs] Error restoring strokes:', e.message);
    }
    node.setPluginData('originalStrokes', '');
  }

  // Restore CS fills if they were saved (generation clears them for grid line visibility)
  var savedFills = node.getPluginData('originalFills');
  if (savedFills) {
    try {
      node.fills = JSON.parse(savedFills);
    } catch (e) {
      console.log('[removeDocs] Error restoring fills:', e.message);
    }
    node.setPluginData('originalFills', '');
  }

  // Move node back to the wrapper's parent at the wrapper's position
  const wrapperParent = wrapper.parent;
  const wrapperIndex = wrapperParent.children.indexOf(wrapper);
  const wrapperX = wrapper.x;
  const wrapperY = wrapper.y;

  wrapperParent.insertChild(wrapperIndex, node);
  node.x = wrapperX;
  node.y = wrapperY;

  // Reset absolute positioning if it was set for auto-layout wrapper
  if (node.layoutPositioning === 'ABSOLUTE') {
    node.layoutPositioning = 'AUTO';
  }

  // Remove the wrapper (and all labels/grid inside it)
  wrapper.remove();

  // Also remove variable modes wrapper if it exists
  var vmDocsWrapper = findVariableModesWrapper(node);
  if (vmDocsWrapper) {
    vmDocsWrapper.remove();
    console.log('[removeDocs] Variable modes wrapper removed.');
  }

  figma.currentPage.selection = [node];
  figma.viewport.scrollAndZoomIntoView([node]);

  figma.ui.postMessage({ type: 'done', message: 'Docs removed.' });
  sendSelectionInfo();
}

function changeDocsColor(hex) {
  var cs = getComponentSet();
  if (!cs) cs = getStandaloneComponent();
  if (!cs) return;

  var wrapper = findExistingWrapper(cs) || findStandaloneWrapper(cs);
  if (!wrapper) return;

  var color = hexToRgb(hex);
  DOC_COLOR = color;

  // Recursively update colors on doc nodes only (skip instances and components)
  function updateNode(node) {
    // Never recurse into instances or components — those are actual design content
    if (node.type === 'INSTANCE' || node.type === 'COMPONENT' || node.type === 'COMPONENT_SET') {
      return;
    }
    // Update text fills
    if (node.type === 'TEXT') {
      node.fills = [{ type: 'SOLID', color: color }];
    }
    // Update line/vector strokes
    if (node.type === 'LINE' || node.type === 'VECTOR') {
      node.strokes = [{ type: 'SOLID', color: color }];
    }
    // Update dashed border strokes on Grid Border frames
    if (node.type === 'FRAME' && node.name === 'Grid Border' && node.strokes && node.strokes.length > 0) {
      node.strokes = [{ type: 'SOLID', color: color }];
    }
    // Recurse into children
    if (node.children) {
      for (var i = 0; i < node.children.length; i++) {
        updateNode(node.children[i]);
      }
    }
  }

  for (var i = 0; i < wrapper.children.length; i++) {
    var child = wrapper.children[i];
    if (child.type === 'FRAME' && child.name !== cs.name && child.type !== 'COMPONENT_SET') {
      updateNode(child);
    }
  }

  // Update the component set's stroke color with dashed border
  if (cs.strokes && cs.strokes.length > 0) {
    var newStrokes = [];
    for (var j = 0; j < cs.strokes.length; j++) {
      var s = cs.strokes[j];
      newStrokes.push({ type: s.type, color: color, opacity: s.opacity });
    }
    cs.strokes = newStrokes;
    cs.dashPattern = [4, 4];
  }
}

function toggleGrid(showGrid) {
  const cs = getComponentSet();
  if (!cs) return;

  const wrapper = findExistingWrapper(cs) || findStandaloneWrapper(cs);
  if (!wrapper) return;

  // Find the container where grid frames live (Documentation frame or wrapper itself)
  var container = wrapper;
  for (var ci = 0; ci < wrapper.children.length; ci++) {
    if (wrapper.children[ci].type === 'FRAME' && wrapper.children[ci].name === 'Documentation') {
      container = wrapper.children[ci];
      break;
    }
  }

  // Remove all existing grid frames
  var gridsToRemove = [];
  for (var gi = 0; gi < container.children.length; gi++) {
    if (container.children[gi].type === 'FRAME' && container.children[gi].name === 'Grid') {
      gridsToRemove.push(container.children[gi]);
    }
  }
  for (var ri = 0; ri < gridsToRemove.length; ri++) {
    gridsToRemove[ri].remove();
  }

  if (showGrid) {
    // Recalculate grid from current variant layout using alignment-aware clustering
    var variants = cs.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
    var colAxis = clusterVariantAxis(variants, 'x', 'width');
    var rowAxis = clusterVariantAxis(variants, 'y', 'height');
    var colClusters = colAxis.mins;
    var rowClusters = rowAxis.mins;

    var colWidths = [];
    for (var c = 0; c < colClusters.length; c++) {
      colWidths.push(colAxis.maxes[c] - colClusters[c]);
    }

    var rowHeights = [];
    for (var r = 0; r < rowClusters.length; r++) {
      rowHeights.push(rowAxis.maxes[r] - rowClusters[r]);
    }

    createGridLines(container, colClusters, rowClusters, colWidths, rowHeights, cs.x, cs.y, cs.width, cs.height);
  }
}

// Scan all pages for PropStar (GusPropstar) doc wrappers — NOT our own autodocs.
// PropStar signatures: ❖ frame with GRID layout, or frames containing GROUP "labels" / FRAME "instances"
async function scanFileForDocs() {
  var pages = figma.root.children;
  var results = [];
  figma.notify('Scanning ' + pages.length + ' page(s) for PropStar docs — this could take a while…', { timeout: 60000 });

  function isPropStarWrapper(frame) {
    if (frame.type !== 'FRAME') return false;
    // Our own wrappers have pluginData — skip them
    if (frame.getPluginData('sourceComponentSetId') || frame.getPluginData('standaloneDoc') === 'true') return false;

    // ❖ frame with grid layout = PropStar
    if (frame.name.startsWith('❖') && frame.layoutMode === 'GRID') return 'grid';

    // Frame containing GROUP "labels" or FRAME "instances" = PropStar artifacts
    for (var i = 0; i < frame.children.length; i++) {
      var child = frame.children[i];
      if (child.type === 'GROUP' && child.name.toLowerCase() === 'labels') return 'labels';
      if (child.type === 'FRAME' && child.name.toLowerCase() === 'instances') return 'instances';
    }
    return false;
  }

  for (var p = 0; p < pages.length; p++) {
    var page = pages[p];
    await page.loadAsync();
    var frames = page.findAll(function(n) { return n.type === 'FRAME'; });
    var pageHits = 0;
    for (var f = 0; f < frames.length; f++) {
      var frame = frames[f];
      var signal = isPropStarWrapper(frame);
      if (signal) {
        pageHits++;
        results.push({
          page: page.name,
          pageId: page.id,
          name: frame.name || '(unnamed)',
          id: frame.id,
          signal: signal,
          x: Math.round(frame.x),
          y: Math.round(frame.y),
          w: Math.round(frame.width),
          h: Math.round(frame.height)
        });
      }
    }
    console.log('[Scan] Page ' + (p + 1) + '/' + pages.length + ' "' + page.name + '" — ' + pageHits + ' PropStar doc(s)');
  }
  console.log('[Scan] Total: ' + results.length + ' PropStar doc(s) across ' + pages.length + ' page(s)');
  for (var r = 0; r < results.length; r++) {
    var d = results[r];
    console.log('  ' + (r + 1) + '. "' + d.name + '" on page "' + d.page + '" — signal: ' + d.signal +
      ' — pos: (' + d.x + ', ' + d.y + ') size: ' + d.w + '×' + d.h);
  }
  figma.notify('Found ' + results.length + ' PropStar doc(s)');
  figma.ui.postMessage({ type: 'scanDocsResults', results: results, pageCount: pages.length });
}

function debugGridCoords() {
  var cs = getComponentSet();
  if (!cs) {
    console.log('[Grid Debug] No component set selected');
    return;
  }

  var wrapper = findExistingWrapper(cs);
  var variants = cs.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });

  // Alignment-aware clustering
  var colAxis = clusterVariantAxis(variants, 'x', 'width');
  var rowAxis = clusterVariantAxis(variants, 'y', 'height');
  var colClusters = colAxis.mins;
  var rowClusters = rowAxis.mins;

  var colWidths = [];
  for (var c = 0; c < colClusters.length; c++) {
    colWidths.push(colAxis.maxes[c] - colClusters[c]);
  }

  var rowHeights = [];
  for (var r = 0; r < rowClusters.length; r++) {
    rowHeights.push(rowAxis.maxes[r] - rowClusters[r]);
  }

  console.log('=== GRID DEBUG ===');
  console.log('Component set:', cs.name);
  console.log('CS position in wrapper: x=' + cs.x + ', y=' + cs.y);
  console.log('CS size: w=' + cs.width + ', h=' + cs.height);
  console.log('Wrapper:', wrapper ? ('x=' + wrapper.x + ', y=' + wrapper.y + ', w=' + wrapper.width + ', h=' + wrapper.height) : 'none');
  console.log('Variants:', variants.length);
  console.log('Column alignment mode:', colAxis.mode);
  console.log('Row alignment mode:', rowAxis.mode);
  console.log('---');
  console.log('Column refs (' + colAxis.clusters.length + ', ' + colAxis.mode + '):', JSON.stringify(colAxis.clusters));
  console.log('Column mins (left edges):', JSON.stringify(colClusters));
  console.log('Column maxes (right edges):', JSON.stringify(colAxis.maxes));
  console.log('Column widths:', JSON.stringify(colWidths));
  console.log('Row refs (' + rowAxis.clusters.length + ', ' + rowAxis.mode + '):', JSON.stringify(rowAxis.clusters));
  console.log('Row mins (top edges):', JSON.stringify(rowClusters));
  console.log('Row maxes (bottom edges):', JSON.stringify(rowAxis.maxes));
  console.log('Row heights:', JSON.stringify(rowHeights));
  console.log('---');

  // Horizontal line positions
  console.log('Horizontal lines (' + (rowClusters.length - 1) + '):');
  for (var ri = 1; ri < rowClusters.length; ri++) {
    var prevBottom = rowClusters[ri - 1] + rowHeights[ri - 1];
    var midY = (prevBottom + rowClusters[ri]) / 2;
    var gap = rowClusters[ri] - prevBottom;
    console.log('  H-line ' + ri + ': prevBottom=' + prevBottom + ' nextTop=' + rowClusters[ri] + ' gap=' + gap + ' → midY=' + midY);
  }

  // Vertical line positions
  console.log('Vertical lines (' + (colClusters.length - 1) + '):');
  for (var ci = 1; ci < colClusters.length; ci++) {
    var prevRight = colClusters[ci - 1] + colWidths[ci - 1];
    var midX = (prevRight + colClusters[ci]) / 2;
    var gapX = colClusters[ci] - prevRight;
    console.log('  V-line ' + ci + ': prevRight=' + prevRight + ' nextLeft=' + colClusters[ci] + ' gap=' + gapX + ' → midX=' + midX);
  }

  // Sample variant positions (first 10)
  console.log('---');
  console.log('Sample variants (first 10):');
  for (var si = 0; si < Math.min(10, variants.length); si++) {
    var v = variants[si];
    console.log('  "' + v.name + '" x=' + v.x + ' y=' + v.y + ' w=' + v.width + ' h=' + v.height + ' centerX=' + (v.x + v.width / 2) + ' centerY=' + (v.y + v.height / 2));
  }
  console.log('=== END GRID DEBUG ===');
}

function debugLabelCoords() {
  var cs = getComponentSet();
  if (!cs) {
    console.log('[Label Debug] No component set selected');
    return;
  }

  var wrapper = findExistingWrapper(cs);
  if (!wrapper) {
    console.log('[Label Debug] No wrapper found — generate docs first');
    return;
  }

  var variants = cs.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });

  // Re-analyze grid to get column/row info
  var colAxis = clusterVariantAxis(variants, 'x', 'width');
  var rowAxis = clusterVariantAxis(variants, 'y', 'height');
  var colClusters = colAxis.mins;
  var rowClusters = rowAxis.mins;

  var colWidths = [];
  for (var c = 0; c < colClusters.length; c++) {
    colWidths.push(colAxis.maxes[c] - colClusters[c]);
  }
  var rowHeights = [];
  for (var r = 0; r < rowClusters.length; r++) {
    rowHeights.push(rowAxis.maxes[r] - rowClusters[r]);
  }

  console.log('=== LABEL DEBUG ===');
  console.log('Wrapper: x=' + wrapper.x + ', y=' + wrapper.y + ', w=' + wrapper.width + ', h=' + wrapper.height);
  console.log('CS in wrapper: x=' + cs.x + ', y=' + cs.y + ', w=' + cs.width + ', h=' + cs.height);
  console.log('---');

  // Column analysis: variant positions vs cell centers
  console.log('Column analysis (' + colClusters.length + ' cols):');
  for (var ci = 0; ci < colClusters.length; ci++) {
    var varLeft = colClusters[ci];
    var varW = colWidths[ci];
    var varCenterX = varLeft + varW / 2;

    // Compute visual cell boundaries (midpoints between adjacent column edges, or CS edge)
    var cellLeft, cellRight;
    if (ci === 0) {
      cellLeft = 0;
    } else {
      var prevRight = colClusters[ci - 1] + colWidths[ci - 1];
      cellLeft = (prevRight + colClusters[ci]) / 2;
    }
    if (ci === colClusters.length - 1) {
      cellRight = cs.width;
    } else {
      var nextLeft = colClusters[ci + 1];
      var curRight = colClusters[ci] + colWidths[ci];
      cellRight = (curRight + nextLeft) / 2;
    }
    var cellCenterX = (cellLeft + cellRight) / 2;
    var offset = varCenterX - cellCenterX;

    console.log('  Col ' + ci + ': varLeft=' + varLeft + ' varW=' + varW + ' varCenter=' + varCenterX +
      ' | cell=[' + cellLeft.toFixed(1) + ',' + cellRight.toFixed(1) + '] cellCenter=' + cellCenterX.toFixed(1) +
      ' | offset=' + offset.toFixed(1) + (Math.abs(offset) > 1 ? ' ⚠ OFF-CENTER' : ' ✓'));
  }

  console.log('---');

  // Row analysis
  console.log('Row analysis (' + rowClusters.length + ' rows):');
  for (var ri = 0; ri < rowClusters.length; ri++) {
    var varTop = rowClusters[ri];
    var varH = rowHeights[ri];
    var varCenterY = varTop + varH / 2;

    var cellTop, cellBottom;
    if (ri === 0) {
      cellTop = 0;
    } else {
      var prevBottom = rowClusters[ri - 1] + rowHeights[ri - 1];
      cellTop = (prevBottom + rowClusters[ri]) / 2;
    }
    if (ri === rowClusters.length - 1) {
      cellBottom = cs.height;
    } else {
      var nextTop = rowClusters[ri + 1];
      var curBottom = rowClusters[ri] + rowHeights[ri];
      cellBottom = (curBottom + nextTop) / 2;
    }
    var cellCenterY = (cellTop + cellBottom) / 2;
    var offsetY = varCenterY - cellCenterY;

    console.log('  Row ' + ri + ': varTop=' + varTop + ' varH=' + varH + ' varCenter=' + varCenterY +
      ' | cell=[' + cellTop.toFixed(1) + ',' + cellBottom.toFixed(1) + '] cellCenter=' + cellCenterY.toFixed(1) +
      ' | offset=' + offsetY.toFixed(1) + (Math.abs(offsetY) > 1 ? ' ⚠ OFF-CENTER' : ' ✓'));
  }

  console.log('---');

  // Find the container with labels (Documentation frame or wrapper itself)
  var labelContainer = wrapper;
  for (var dfi = 0; dfi < wrapper.children.length; dfi++) {
    if (wrapper.children[dfi].type === 'FRAME' && wrapper.children[dfi].name === 'Documentation') {
      labelContainer = wrapper.children[dfi];
      break;
    }
  }

  // All label frames
  var labelFrames = labelContainer.children.filter(function(child) { return child.name === 'Label'; });
  console.log('Label frames (' + labelFrames.length + '):');
  for (var li = 0; li < labelFrames.length; li++) {
    var lf = labelFrames[li];
    var textChild = lf.children.find(function(ch) { return ch.type === 'TEXT'; });
    var textContent = textChild ? textChild.characters : '?';
    var textX = textChild ? textChild.x.toFixed(1) : '?';
    var textW = textChild ? textChild.width.toFixed(1) : '?';
    console.log('  "' + textContent + '" frame: x=' + lf.x + ' y=' + lf.y + ' w=' + lf.width + ' h=' + lf.height +
      ' | text: localX=' + textX + ' textW=' + textW);
  }

  console.log('=== END LABEL DEBUG ===');
}

function debugBooleanGrid() {
  var cs = getComponentSet();
  if (!cs) {
    console.log('[Boolean Grid Debug] No component set selected');
    return;
  }

  var wrapper = findExistingWrapper(cs);
  if (!wrapper) {
    console.log('[Boolean Grid Debug] No wrapper found — generate docs first');
    return;
  }

  // Find the Documentation frame or use wrapper directly
  var container = wrapper;
  for (var dfi = 0; dfi < wrapper.children.length; dfi++) {
    if (wrapper.children[dfi].type === 'FRAME' && wrapper.children[dfi].name === 'Documentation') {
      container = wrapper.children[dfi];
      break;
    }
  }

  console.log('=== BOOLEAN GRID DEBUG ===');
  console.log('Wrapper: x=' + wrapper.x + ', y=' + wrapper.y + ', w=' + wrapper.width + ', h=' + wrapper.height);
  console.log('CS: x=' + cs.x + ', y=' + cs.y + ', w=' + cs.width + ', h=' + cs.height);
  console.log('CS strokes:', JSON.stringify(cs.strokes));
  console.log('CS strokeWeight:', cs.strokeWeight, 'dashPattern:', JSON.stringify(cs.dashPattern));

  var savedStrokes = cs.getPluginData('originalStrokes');
  console.log('Saved original strokes:', savedStrokes || '(none)');
  console.log('---');

  // Boolean Grid Border
  var borderFrames = container.children.filter(function(child) { return child.name === 'Boolean Grid Border'; });
  if (borderFrames.length > 0) {
    for (var bi = 0; bi < borderFrames.length; bi++) {
      var bf = borderFrames[bi];
      console.log('Boolean Grid Border:');
      console.log('  position: x=' + bf.x + ', y=' + bf.y);
      console.log('  size: w=' + bf.width + ', h=' + bf.height);
      console.log('  strokes:', JSON.stringify(bf.strokes));
      console.log('  strokeWeight:', bf.strokeWeight, 'dashPattern:', JSON.stringify(bf.dashPattern));
      console.log('  strokeAlign:', bf.strokeAlign, 'cornerRadius:', bf.cornerRadius);
      console.log('  locked:', bf.locked);
      console.log('  fills:', JSON.stringify(bf.fills));
      console.log('  clipsContent:', bf.clipsContent);
      // Check coverage
      console.log('  covers CS: x=' + bf.x + '==' + cs.x + '? ' + (bf.x === cs.x) +
        ' y=' + bf.y + '==' + cs.y + '? ' + (bf.y === cs.y) +
        ' w=' + bf.width + '==' + cs.width + '? ' + (bf.width === cs.width));
      var borderBottom = bf.y + bf.height;
      var wrapperBottom = wrapper.height;
      console.log('  border bottom: ' + borderBottom + ' wrapper bottom: ' + wrapperBottom + ' match: ' + (Math.abs(borderBottom - wrapperBottom) < 1));
    }
  } else {
    console.log('Boolean Grid Border: NOT FOUND');
  }

  console.log('---');

  // Boolean Tints
  var tints = container.children.filter(function(child) { return child.name === 'Boolean Tint'; });
  if (tints.length > 0) {
    for (var ti = 0; ti < tints.length; ti++) {
      var t = tints[ti];
      console.log('Boolean Tint ' + ti + ':');
      console.log('  position: x=' + t.x + ', y=' + t.y);
      console.log('  size: w=' + t.width + ', h=' + t.height);
      console.log('  fills:', JSON.stringify(t.fills));
      console.log('  locked:', t.locked);
      var tintBottom = t.y + t.height;
      console.log('  tint bottom: ' + tintBottom);
    }
  } else {
    console.log('Boolean Tints: NOT FOUND');
  }

  console.log('---');

  // Instances
  var instances = container.children.filter(function(child) { return child.type === 'INSTANCE'; });
  if (instances.length > 0) {
    console.log('Boolean instances (' + instances.length + '):');
    for (var ii = 0; ii < instances.length; ii++) {
      var inst = instances[ii];
      console.log('  ' + inst.name + ': x=' + inst.x + ', y=' + inst.y + ', w=' + inst.width + ', h=' + inst.height +
        ' bottom=' + (inst.y + inst.height));
    }
  } else {
    console.log('Boolean instances: NOT FOUND');
  }

  console.log('---');

  // Divider lines
  var lines = container.children.filter(function(child) { return child.type === 'LINE'; });
  if (lines.length > 0) {
    console.log('Lines (' + lines.length + '):');
    for (var li = 0; li < lines.length; li++) {
      var line = lines[li];
      console.log('  Line ' + li + ': x=' + line.x + ', y=' + line.y + ', w=' + line.width + ', h=' + line.height +
        ' rotation=' + line.rotation);
    }
  }

  // Grid frames
  var grids = container.children.filter(function(child) { return child.name === 'Grid'; });
  if (grids.length > 0) {
    console.log('Grid frames (' + grids.length + '):');
    for (var gi = 0; gi < grids.length; gi++) {
      var g = grids[gi];
      console.log('  Grid ' + gi + ': x=' + g.x + ', y=' + g.y + ', w=' + g.width + ', h=' + g.height);
    }
  }

  console.log('=== END BOOLEAN GRID DEBUG ===');
}

function removeOldLabels(wrapper) {
  const toRemove = [];

  // Find CS node for stroke restoration
  var csNode = null;
  for (const child of wrapper.children) {
    if (child.type === 'COMPONENT_SET' || child.type === 'COMPONENT') {
      csNode = child;
      break;
    }
  }

  // Restore CS strokes if they were saved
  if (csNode) {
    var savedStrokes = csNode.getPluginData('originalStrokes');
    if (savedStrokes) {
      try {
        var strokeData = JSON.parse(savedStrokes);
        csNode.strokes = strokeData.strokes;
        csNode.strokeWeight = strokeData.strokeWeight;
        csNode.dashPattern = strokeData.dashPattern;
        csNode.strokeAlign = strokeData.strokeAlign;
        if (strokeData.cornerRadius !== undefined) csNode.cornerRadius = strokeData.cornerRadius;
      } catch (e) {
        console.log('[removeOldLabels] Error restoring strokes:', e.message);
      }
    }
    var savedFills = csNode.getPluginData('originalFills');
    if (savedFills) {
      try {
        csNode.fills = JSON.parse(savedFills);
      } catch (e) {
        console.log('[removeOldLabels] Error restoring fills:', e.message);
      }
    }
  }

  // Check for Documentation frame (new structure) — remove it entirely
  // Also check for individual elements (legacy structure)
  for (const child of wrapper.children) {
    if (child.type === 'FRAME' && child.name === 'Documentation') {
      toRemove.push(child);
    }
    if (child.type === 'FRAME' && (child.name === 'Label' || child.name === 'Bracket' || child.name === 'Grid' || child.name === 'Boolean Visibility' || child.name === 'Nested Instances' || child.name === 'Property Combinations' || child.name === 'Boolean Grid' || child.name === 'Boolean Grid Border')) {
      toRemove.push(child);
    }
    if (child.type === 'RECTANGLE' && child.name === 'Boolean Tint') {
      toRemove.push(child);
    }
    if (child.type === 'TEXT' && child.name === 'Title') {
      toRemove.push(child);
    }
    if (child.type === 'FRAME' && child.name === 'Description') {
      toRemove.push(child);
    }
    if (child.type === 'TEXT' && child.name === 'Documentation Link') {
      toRemove.push(child);
    }
    if (child.type === 'INSTANCE') {
      toRemove.push(child);
    }
    if (child.type === 'LINE') {
      toRemove.push(child);
    }
  }
  for (const node of toRemove) {
    node.remove();
  }
  return toRemove.length;
}

// ─── Boolean Visibility Section ──────────────────────────────────────────────

// Compute the horizontal gap between columns from variant positions in a component set
function computeVariantColumnGap(variants) {
  if (variants.length < 2) return 40;
  var sorted = variants.slice().sort(function(a, b) { return a.y - b.y || a.x - b.x; });
  var i = 0;
  while (i < sorted.length) {
    var rowY = sorted[i].y;
    var row = [];
    while (i < sorted.length && Math.abs(sorted[i].y - rowY) <= CLUSTER_TOLERANCE) {
      row.push(sorted[i]);
      i++;
    }
    if (row.length >= 2) {
      row.sort(function(a, b) { return a.x - b.x; });
      var gap = row[1].x - (row[0].x + row[0].width);
      if (gap > 0) return Math.round(gap);
    }
  }
  return 40;
}

function computeVariantRowGap(variants) {
  if (variants.length < 2) return 40;
  var sorted = variants.slice().sort(function(a, b) { return a.x - b.x || a.y - b.y; });
  var i = 0;
  while (i < sorted.length) {
    var colX = sorted[i].x;
    var col = [];
    while (i < sorted.length && Math.abs(sorted[i].x - colX) <= CLUSTER_TOLERANCE) {
      col.push(sorted[i]);
      i++;
    }
    if (col.length >= 2) {
      col.sort(function(a, b) { return a.y - b.y; });
      var gap = col[1].y - (col[0].y + col[0].height);
      if (gap > 0) return Math.round(gap);
    }
  }
  return 40;
}

// Generate all non-empty subsets of an array, sorted by size (singles first, then pairs, etc.)
function getNonEmptySubsets(arr) {
  var result = [];
  var n = arr.length;
  for (var mask = 1; mask < (1 << n); mask++) {
    var subset = [];
    for (var i = 0; i < n; i++) {
      if (mask & (1 << i)) subset.push(arr[i]);
    }
    result.push(subset);
  }
  result.sort(function(a, b) { return a.length - b.length; });
  return result;
}

// Build boolean groups from properties and combination mode.
// labelFormat: 'onoff' uses On/Off, anything else uses True/False.
function buildBooleanGroups(boolProps, combinationMode, labelFormat) {
  var subsets;
  if (combinationMode === 'all') {
    subsets = getNonEmptySubsets(boolProps);
  } else if (combinationMode === 'combined' && boolProps.length > 1) {
    subsets = boolProps.map(function(p) { return [p]; });
    subsets.push(boolProps.slice());
  } else {
    subsets = boolProps.map(function(p) { return [p]; });
  }
  var trueLabel = labelFormat === 'onoff' ? 'On' : 'True';
  var falseLabel = labelFormat === 'onoff' ? 'Off' : 'False';
  var groups = [];
  for (var si = 0; si < subsets.length; si++) {
    var subset = subsets[si];
    var labelParts = [];
    var propsToSet = {};
    for (var pi = 0; pi < subset.length; pi++) {
      var nonDefault = !subset[pi].defaultValue;
      labelParts.push(formatLabel(subset[pi].name, nonDefault ? trueLabel : falseLabel));
      propsToSet[subset[pi].key] = nonDefault;
    }
    groups.push({ label: labelParts.join(' + '), props: propsToSet });
  }
  return groups;
}

async function createBooleanVisibilitySection(sourceNode, combinationMode, enabledPropNames, displayMode) {
  var boolProps = getBooleanComponentProperties(sourceNode);
  if (enabledPropNames) {
    boolProps = boolProps.filter(function(p) { return enabledPropNames.indexOf(p.name) !== -1; });
  }
  if (boolProps.length === 0) return null;

  var standalone = isStandaloneComponent(sourceNode);
  var variants = standalone ? [sourceNode] : sourceNode.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
  if (variants.length === 0) return null;

  var columns = buildBooleanGroups(boolProps, combinationMode, 'onoff');

  if (DEBUG) { console.log('[Boolean Visibility] Found', boolProps.length, 'boolean props, mode:', combinationMode, ', display:', displayMode, ', subsets:', columns.length, ', variants:', variants.length); }

  var section = figma.createFrame();
  section.name = 'Boolean Visibility';
  section.fills = [];
  section.clipsContent = false;

  var COMBO_GAP = 24;
  var LABEL_INSTANCE_GAP = 6;

  var csWidth = sourceNode.width;
  var csHeight = standalone ? variants[0].height : sourceNode.height;

  if (displayMode === 'grid') {
    // ─── Grid mode: columns side by side ───
    // First column is "Default" (no overrides), then one column per subset

    var allColumns = [{ label: 'Default', props: null }].concat(columns);
    var colX = 0;
    var labelHeight = 0;

    // First pass: create labels, measure max label height
    var labelNodes = [];
    for (var ci = 0; ci < allColumns.length; ci++) {
      var colLabel = createTextNode(allColumns[ci].label, LABEL_FONT_SIZE);
      labelNodes.push(colLabel);
      if (colLabel.height > labelHeight) labelHeight = colLabel.height;
    }
    labelHeight += LABEL_INSTANCE_GAP;

    // Second pass: position columns
    for (var ci2 = 0; ci2 < allColumns.length; ci2++) {
      var col = allColumns[ci2];

      // Position label
      labelNodes[ci2].x = colX;
      labelNodes[ci2].y = 0;
      section.appendChild(labelNodes[ci2]);

      // Create instances
      for (var vi = 0; vi < variants.length; vi++) {
        try {
          var variant = variants[vi];
          var instance = variant.createInstance();
          if (col.props) instance.setProperties(col.props);

          instance.x = colX + (standalone ? 0 : variant.x);
          instance.y = labelHeight + (standalone ? 0 : variant.y);
          section.appendChild(instance);
        } catch (e) {
          console.log('[Boolean Visibility Grid] Error creating instance:', e.message);
        }
      }

      colX += csWidth + COMBO_GAP;
    }

    var totalWidth = colX > COMBO_GAP ? colX - COMBO_GAP : 10;
    section.resize(totalWidth, labelHeight + csHeight);
  } else if (_autoLayoutExtras) {
    // ─── List mode: vertical stack with auto layout ───

    section.layoutMode = 'VERTICAL';
    section.itemSpacing = COMBO_GAP;
    section.counterAxisSizingMode = 'AUTO';
    section.primaryAxisSizingMode = 'AUTO';

    var INSTANCE_GAP = standalone ? 40 : computeVariantRowGap(variants);
    var colGap = standalone ? 0 : computeVariantColumnGap(variants);

    for (var si2 = 0; si2 < columns.length; si2++) {
      var col2 = columns[si2];

      if (DEBUG) { console.log('[Boolean Visibility] Subset ' + si2 + ': "' + col2.label + '"'); }

      // Combo frame: label + instance rows
      var comboFrame = figma.createFrame();
      comboFrame.name = col2.label;
      comboFrame.fills = [];
      comboFrame.layoutMode = 'VERTICAL';
      comboFrame.itemSpacing = LABEL_INSTANCE_GAP;
      comboFrame.counterAxisSizingMode = 'AUTO';
      comboFrame.primaryAxisSizingMode = 'AUTO';

      var label = createTextNode(col2.label, LABEL_FONT_SIZE);
      comboFrame.appendChild(label);

      // Create instances, collect with original positions
      var comboInstances = [];
      for (var vi2 = 0; vi2 < variants.length; vi2++) {
        try {
          var variant2 = variants[vi2];
          var instance2 = variant2.createInstance();
          instance2.setProperties(col2.props);
          comboInstances.push({ inst: instance2, origX: standalone ? 0 : variant2.x, origY: standalone ? 0 : variant2.y });
        } catch (e) {
          console.log('[Boolean Visibility] Error creating instance from "' + variants[vi2].name + '":', e.message);
        }
      }
      comboInstances.sort(function(a, b) { return a.origY - b.origY || a.origX - b.origX; });

      // Rows container
      var rowsContainer = figma.createFrame();
      rowsContainer.name = 'Instances';
      rowsContainer.fills = [];
      rowsContainer.layoutMode = 'VERTICAL';
      rowsContainer.itemSpacing = INSTANCE_GAP;
      rowsContainer.counterAxisSizingMode = 'AUTO';
      rowsContainer.primaryAxisSizingMode = 'AUTO';

      var ri = 0;
      while (ri < comboInstances.length) {
        var rowY = comboInstances[ri].origY;
        var rowItems = [];
        while (ri < comboInstances.length && Math.abs(comboInstances[ri].origY - rowY) <= CLUSTER_TOLERANCE) {
          rowItems.push(comboInstances[ri]);
          ri++;
        }

        if (rowItems.length === 1) {
          rowsContainer.appendChild(rowItems[0].inst);
        } else {
          rowItems.sort(function(a, b) { return a.origX - b.origX; });
          var rowFrame = figma.createFrame();
          rowFrame.name = 'Row';
          rowFrame.fills = [];
          rowFrame.layoutMode = 'HORIZONTAL';
          rowFrame.itemSpacing = colGap;
          rowFrame.counterAxisSizingMode = 'AUTO';
          rowFrame.primaryAxisSizingMode = 'AUTO';
          for (var rj = 0; rj < rowItems.length; rj++) {
            rowFrame.appendChild(rowItems[rj].inst);
          }
          rowsContainer.appendChild(rowFrame);
        }
      }

      comboFrame.appendChild(rowsContainer);
      section.appendChild(comboFrame);
    }
  } else {
    // ─── List mode: vertical stack (manual positioning) ───

    var yOffset = 0;
    var maxSectionWidth = 0;

    for (var si2 = 0; si2 < columns.length; si2++) {
      var col2 = columns[si2];

      if (DEBUG) { console.log('[Boolean Visibility] Subset ' + si2 + ': "' + col2.label + '"'); }

      var label = createTextNode(col2.label, LABEL_FONT_SIZE);
      label.x = 0;
      label.y = yOffset;
      section.appendChild(label);
      yOffset += label.height + LABEL_INSTANCE_GAP;

      var comboInstances = [];
      for (var vi2 = 0; vi2 < variants.length; vi2++) {
        try {
          var variant2 = variants[vi2];
          var instance2 = variant2.createInstance();
          instance2.setProperties(col2.props);
          instance2.x = standalone ? 0 : variant2.x;
          section.appendChild(instance2);
          comboInstances.push({ inst: instance2, origY: standalone ? 0 : variant2.y });
        } catch (e) {
          console.log('[Boolean Visibility] Error creating instance from "' + variants[vi2].name + '":', e.message);
        }
      }
      comboInstances.sort(function(a, b) { return a.origY - b.origY; });

      var comboHeight = 0;
      var INSTANCE_GAP = standalone ? 40 : computeVariantRowGap(variants);
      var ri = 0;
      while (ri < comboInstances.length) {
        var rowY = comboInstances[ri].origY;
        var rowMaxHeight = 0;
        while (ri < comboInstances.length && Math.abs(comboInstances[ri].origY - rowY) <= CLUSTER_TOLERANCE) {
          comboInstances[ri].inst.y = yOffset + comboHeight;
          if (comboInstances[ri].inst.height > rowMaxHeight) rowMaxHeight = comboInstances[ri].inst.height;
          ri++;
        }
        comboHeight += rowMaxHeight;
        if (ri < comboInstances.length) comboHeight += INSTANCE_GAP;
      }

      if (csWidth > maxSectionWidth) maxSectionWidth = csWidth;
      yOffset += (comboHeight > 0 ? comboHeight : csHeight) + COMBO_GAP;
    }

    var finalHeight = yOffset > COMBO_GAP ? yOffset - COMBO_GAP : 10;
    section.resize(Math.max(maxSectionWidth, 100), finalHeight);
  }

  if (DEBUG) { console.log('[Boolean Visibility] Section size:', section.width + 'x' + section.height); }
  return section;
}

// ─── Scoped Boolean Visibility Section ────────────────────────────────────────

function filterVariantsByProperty(variants, propName, propValue) {
  return variants.filter(function(v) {
    var parts = v.name.split(',');
    for (var i = 0; i < parts.length; i++) {
      var kv = parts[i].trim().split('=');
      if (kv[0].trim() === propName && kv[1].trim() === propValue) return true;
    }
    return false;
  });
}

function filterVariantsByProperties(variants, propValues) {
  return variants.filter(function(v) {
    var parts = v.name.split(',');
    var props = {};
    for (var i = 0; i < parts.length; i++) {
      var kv = parts[i].trim().split('=');
      props[kv[0].trim()] = kv[1].trim();
    }
    for (var key in propValues) {
      if (props[key] !== propValues[key]) return false;
    }
    return true;
  });
}

async function createScopedBooleanSection(sourceNode, scopeData, combinationMode, displayMode) {
  var allBoolProps = getBooleanComponentProperties(sourceNode);
  if (allBoolProps.length === 0) return null;

  var standalone = isStandaloneComponent(sourceNode);
  var allVariants = standalone ? [sourceNode] : sourceNode.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
  if (allVariants.length === 0) return null;

  if (DEBUG) {
    console.log('[Scoped Boolean] scope properties:', (scopeData.properties || [scopeData.property]).join(', '), ', groups:', scopeData.groups.length, ', combinationMode:', combinationMode);
  }

  var section = figma.createFrame();
  section.name = 'Boolean Visibility';
  section.fills = [];
  section.clipsContent = false;

  var GROUP_GAP = 32;
  var COMBO_GAP = 24;
  var LABEL_INSTANCE_GAP = 6;
  var INSTANCE_GAP = standalone ? 40 : computeVariantRowGap(allVariants);

  if (_autoLayoutExtras) {
    section.layoutMode = 'VERTICAL';
    section.itemSpacing = GROUP_GAP;
    section.counterAxisSizingMode = 'AUTO';
    section.primaryAxisSizingMode = 'AUTO';

    var hasContent = false;

    for (var gi = 0; gi < scopeData.groups.length; gi++) {
      var group = scopeData.groups[gi];
      var variants = standalone ? allVariants : (group.values
        ? filterVariantsByProperties(allVariants, group.values)
        : filterVariantsByProperty(allVariants, scopeData.property, group.value));
      if (variants.length === 0) continue;
      var boolProps = allBoolProps.filter(function(p) { return group.booleans.indexOf(p.name) !== -1; });
      if (boolProps.length === 0) continue;
      hasContent = true;

      var colGap = standalone ? 0 : computeVariantColumnGap(variants);
      var minX = Infinity, minY = Infinity;
      for (var vi = 0; vi < variants.length; vi++) {
        if (variants[vi].x < minX) minX = variants[vi].x;
        if (variants[vi].y < minY) minY = variants[vi].y;
      }

      var groupFrame = figma.createFrame();
      groupFrame.name = group.values ? Object.values(group.values).join(', ') : group.value;
      groupFrame.fills = [];
      groupFrame.layoutMode = 'VERTICAL';
      groupFrame.itemSpacing = LABEL_INSTANCE_GAP;
      groupFrame.counterAxisSizingMode = 'AUTO';
      groupFrame.primaryAxisSizingMode = 'AUTO';

      var headerText = group.values ? Object.values(group.values).join(', ') : group.value;
      var groupHeader = createTextNode(headerText, LABEL_FONT_SIZE);
      groupHeader.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.5 }];
      groupFrame.appendChild(groupHeader);

      var columns = buildBooleanGroups(boolProps, combinationMode, 'onoff');

      var combosContainer = figma.createFrame();
      combosContainer.name = 'Combinations';
      combosContainer.fills = [];
      combosContainer.layoutMode = 'VERTICAL';
      combosContainer.itemSpacing = COMBO_GAP;
      combosContainer.counterAxisSizingMode = 'AUTO';
      combosContainer.primaryAxisSizingMode = 'AUTO';

      for (var ci = 0; ci < columns.length; ci++) {
        var col = columns[ci];
        var comboFrame = figma.createFrame();
        comboFrame.name = col.label;
        comboFrame.fills = [];
        comboFrame.layoutMode = 'VERTICAL';
        comboFrame.itemSpacing = LABEL_INSTANCE_GAP;
        comboFrame.counterAxisSizingMode = 'AUTO';
        comboFrame.primaryAxisSizingMode = 'AUTO';

        var label = createTextNode(col.label, LABEL_FONT_SIZE);
        comboFrame.appendChild(label);

        var scopeInstances = [];
        for (var vi3 = 0; vi3 < variants.length; vi3++) {
          try {
            var variant = variants[vi3];
            var instance = variant.createInstance();
            instance.setProperties(col.props);
            scopeInstances.push({ inst: instance, origX: standalone ? 0 : (variant.x - minX), origY: standalone ? 0 : (variant.y - minY) });
          } catch (e) {
            console.log('[Scoped Boolean] Error creating instance:', e.message);
          }
        }
        scopeInstances.sort(function(a, b) { return a.origY - b.origY || a.origX - b.origX; });

        var rowsContainer = figma.createFrame();
        rowsContainer.name = 'Instances';
        rowsContainer.fills = [];
        rowsContainer.layoutMode = 'VERTICAL';
        rowsContainer.itemSpacing = INSTANCE_GAP;
        rowsContainer.counterAxisSizingMode = 'AUTO';
        rowsContainer.primaryAxisSizingMode = 'AUTO';

        var sri = 0;
        while (sri < scopeInstances.length) {
          var sRowY = scopeInstances[sri].origY;
          var rowItems = [];
          while (sri < scopeInstances.length && Math.abs(scopeInstances[sri].origY - sRowY) <= CLUSTER_TOLERANCE) {
            rowItems.push(scopeInstances[sri]);
            sri++;
          }
          if (rowItems.length === 1) {
            rowsContainer.appendChild(rowItems[0].inst);
          } else {
            rowItems.sort(function(a, b) { return a.origX - b.origX; });
            var rowFrame = figma.createFrame();
            rowFrame.name = 'Row';
            rowFrame.fills = [];
            rowFrame.layoutMode = 'HORIZONTAL';
            rowFrame.itemSpacing = colGap;
            rowFrame.counterAxisSizingMode = 'AUTO';
            rowFrame.primaryAxisSizingMode = 'AUTO';
            for (var rj = 0; rj < rowItems.length; rj++) {
              rowFrame.appendChild(rowItems[rj].inst);
            }
            rowsContainer.appendChild(rowFrame);
          }
        }

        comboFrame.appendChild(rowsContainer);
        combosContainer.appendChild(comboFrame);
      }

      groupFrame.appendChild(combosContainer);
      section.appendChild(groupFrame);
    }

    if (!hasContent) return null;
  } else {
    // ─── Manual positioning (original behavior) ───
    var yOffset = 0;
    var maxSectionWidth = 0;

    for (var gi = 0; gi < scopeData.groups.length; gi++) {
      var group = scopeData.groups[gi];
      var variants = standalone ? allVariants : (group.values
        ? filterVariantsByProperties(allVariants, group.values)
        : filterVariantsByProperty(allVariants, scopeData.property, group.value));
      if (variants.length === 0) continue;
      var boolProps = allBoolProps.filter(function(p) { return group.booleans.indexOf(p.name) !== -1; });
      if (boolProps.length === 0) continue;

      var minX = Infinity, minY = Infinity;
      for (var vi = 0; vi < variants.length; vi++) {
        if (variants[vi].x < minX) minX = variants[vi].x;
        if (variants[vi].y < minY) minY = variants[vi].y;
      }
      var groupWidth = 0, groupHeight = 0;
      for (var vi2 = 0; vi2 < variants.length; vi2++) {
        var extX = (variants[vi2].x - minX) + variants[vi2].width;
        var extY = (variants[vi2].y - minY) + variants[vi2].height;
        if (extX > groupWidth) groupWidth = extX;
        if (extY > groupHeight) groupHeight = extY;
      }

      var headerText = group.values ? Object.values(group.values).join(', ') : group.value;
      var groupHeader = createTextNode(headerText, LABEL_FONT_SIZE);
      groupHeader.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.5 }];
      groupHeader.x = 0;
      groupHeader.y = yOffset;
      section.appendChild(groupHeader);
      yOffset += groupHeader.height + LABEL_INSTANCE_GAP;

      var columns = buildBooleanGroups(boolProps, combinationMode, 'onoff');

      for (var ci = 0; ci < columns.length; ci++) {
        var col = columns[ci];
        var label = createTextNode(col.label, LABEL_FONT_SIZE);
        label.x = 0;
        label.y = yOffset;
        section.appendChild(label);
        yOffset += label.height + LABEL_INSTANCE_GAP;

        var scopeInstances = [];
        for (var vi3 = 0; vi3 < variants.length; vi3++) {
          try {
            var variant = variants[vi3];
            var instance = variant.createInstance();
            instance.setProperties(col.props);
            instance.x = standalone ? 0 : (variant.x - minX);
            section.appendChild(instance);
            scopeInstances.push({ inst: instance, origY: standalone ? 0 : (variant.y - minY) });
          } catch (e) {
            console.log('[Scoped Boolean] Error creating instance:', e.message);
          }
        }
        scopeInstances.sort(function(a, b) { return a.origY - b.origY; });

        var scopeComboH = 0;
        var sri = 0;
        while (sri < scopeInstances.length) {
          var sRowY = scopeInstances[sri].origY;
          var sRowMaxH = 0;
          while (sri < scopeInstances.length && Math.abs(scopeInstances[sri].origY - sRowY) <= CLUSTER_TOLERANCE) {
            scopeInstances[sri].inst.y = yOffset + scopeComboH;
            if (scopeInstances[sri].inst.height > sRowMaxH) sRowMaxH = scopeInstances[sri].inst.height;
            sri++;
          }
          scopeComboH += sRowMaxH;
          if (sri < scopeInstances.length) scopeComboH += INSTANCE_GAP;
        }

        if (groupWidth > maxSectionWidth) maxSectionWidth = groupWidth;
        yOffset += (scopeComboH > 0 ? scopeComboH : groupHeight) + COMBO_GAP;
      }

      if (gi < scopeData.groups.length - 1) {
        yOffset += GROUP_GAP - COMBO_GAP;
      }
    }

    if (yOffset === 0) return null;
    var finalHeight = yOffset > COMBO_GAP ? yOffset - COMBO_GAP : 10;
    section.resize(Math.max(maxSectionWidth, 100), finalHeight);
  }

  if (DEBUG) { console.log('[Scoped Boolean] Section size:', section.width + 'x' + section.height); }
  return section;
}

// ─── Nested Instances Section ────────────────────────────────────────────────

async function createNestedInstancesSection(sourceNode, mode, enabledNames) {
  var standalone = isStandaloneComponent(sourceNode);
  var nestedSets = standalone
    ? await findNestedComponentSetsForComponent(sourceNode)
    : await findNestedComponentSets(sourceNode);
  if (enabledNames) {
    nestedSets = nestedSets.filter(function(ns) { return enabledNames.indexOf(ns.componentSetName) !== -1; });
  }
  if (nestedSets.length === 0) return null;

  var variants = standalone ? [sourceNode] : sourceNode.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
  if (variants.length === 0) return null;

  if (DEBUG) {
    console.log('[Nested Instances] mode:', mode, ', nested sets:', nestedSets.length, ', standalone:', standalone);
  }

  var section = figma.createFrame();
  section.name = 'Nested Instances';
  section.fills = [];
  section.clipsContent = false;

  var COMBO_GAP = 24;
  var LABEL_INSTANCE_GAP = 6;
  var PROP_GAP = 32;

  // Pre-compute boolean property overrides for parent instances
  var boolProps = getBooleanComponentProperties(sourceNode);
  var boolOverrides = {};
  for (var bp = 0; bp < boolProps.length; bp++) {
    boolOverrides[boolProps[bp].key] = true;
  }
  var hasBoolOverrides = boolProps.length > 0;

  if (_autoLayoutExtras) {
    section.layoutMode = 'VERTICAL';
    section.itemSpacing = PROP_GAP;
    section.counterAxisSizingMode = 'AUTO';
    section.primaryAxisSizingMode = 'AUTO';

    for (var ni = 0; ni < nestedSets.length; ni++) {
      var nested = nestedSets[ni];

      var setFrame = figma.createFrame();
      setFrame.name = nested.componentSetName;
      setFrame.fills = [];
      setFrame.layoutMode = 'VERTICAL';
      setFrame.itemSpacing = LABEL_INSTANCE_GAP;
      setFrame.counterAxisSizingMode = 'AUTO';
      setFrame.primaryAxisSizingMode = 'AUTO';

      var csHeader = createTextNode(nested.componentSetName, LABEL_FONT_SIZE);
      csHeader.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.5 }];
      setFrame.appendChild(csHeader);

      var nestedGroupProps = nested.variantGroupProperties;
      var defaultProps = nested.defaultVariantProperties;
      var propNames = Object.keys(nestedGroupProps);

      if (DEBUG) {
        console.log('[Nested Instances] "' + nested.componentSetName + '": ' + propNames.length + ' properties, defaults:', JSON.stringify(defaultProps));
      }

      var propsContainer = figma.createFrame();
      propsContainer.name = 'Properties';
      propsContainer.fills = [];
      propsContainer.layoutMode = 'VERTICAL';
      propsContainer.itemSpacing = PROP_GAP;
      propsContainer.counterAxisSizingMode = 'AUTO';
      propsContainer.primaryAxisSizingMode = 'AUTO';

      for (var pi = 0; pi < propNames.length; pi++) {
        var propName = propNames[pi];
        var propValues = nestedGroupProps[propName];
        var defaultValue = defaultProps[propName];

        var propFrame = figma.createFrame();
        propFrame.name = propName;
        propFrame.fills = [];
        propFrame.layoutMode = 'VERTICAL';
        propFrame.itemSpacing = LABEL_INSTANCE_GAP;
        propFrame.counterAxisSizingMode = 'AUTO';
        propFrame.primaryAxisSizingMode = 'AUTO';

        var propHeader = createTextNode(propName, LABEL_FONT_SIZE);
        propHeader.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.35 }];
        propFrame.appendChild(propHeader);

        var valuesContainer = figma.createFrame();
        valuesContainer.name = 'Values';
        valuesContainer.fills = [];
        valuesContainer.layoutMode = 'VERTICAL';
        valuesContainer.itemSpacing = COMBO_GAP;
        valuesContainer.counterAxisSizingMode = 'AUTO';
        valuesContainer.primaryAxisSizingMode = 'AUTO';

        for (var vi2 = 0; vi2 < propValues.length; vi2++) {
          var value = propValues[vi2];
          if (value === defaultValue) continue;

          var targetProps = {};
          for (var key in defaultProps) {
            if (defaultProps.hasOwnProperty(key)) targetProps[key] = defaultProps[key];
          }
          targetProps[propName] = value;

          var swapTarget = findVariantByProperties(nested.variants, targetProps);
          if (!swapTarget) continue;

          var targetComponent = await figma.getNodeByIdAsync(swapTarget.id);
          if (!targetComponent) continue;

          var displayLabel = formatLabel(propName, value);
          var valueFrame = figma.createFrame();
          valueFrame.name = displayLabel;
          valueFrame.fills = [];
          valueFrame.layoutMode = 'VERTICAL';
          valueFrame.itemSpacing = LABEL_INSTANCE_GAP;
          valueFrame.counterAxisSizingMode = 'AUTO';
          valueFrame.primaryAxisSizingMode = 'AUTO';

          var label = createTextNode(displayLabel, LABEL_FONT_SIZE);
          valueFrame.appendChild(label);

          if (mode === 'representative') {
            var baseIdx = (nested.parentVariantIndices && nested.parentVariantIndices.length > 0)
              ? nested.parentVariantIndices[0] : 0;
            try {
              var instance = variants[baseIdx].createInstance();
              if (hasBoolOverrides) instance.setProperties(boolOverrides);
              var nestedInst = findNestedInstanceByName(instance, nested.instanceName);
              if (nestedInst) nestedInst.swapComponent(targetComponent);
              valueFrame.appendChild(instance);
            } catch (e) {
              console.log('[Nested Instances] Error creating representative instance:', e.message);
            }
          } else {
            var fullIndices = [];
            if (nested.parentVariantIndices && nested.parentVariantIndices.length > 0) {
              fullIndices = nested.parentVariantIndices.slice();
            } else {
              for (var ai = 0; ai < variants.length; ai++) fullIndices.push(ai);
            }

            var instancesContainer = figma.createFrame();
            instancesContainer.name = 'Instances';
            instancesContainer.fills = [];
            instancesContainer.layoutMode = 'VERTICAL';
            instancesContainer.itemSpacing = COMBO_GAP;
            instancesContainer.counterAxisSizingMode = 'AUTO';
            instancesContainer.primaryAxisSizingMode = 'AUTO';

            for (var fi2 = 0; fi2 < fullIndices.length; fi2++) {
              try {
                var variant = variants[fullIndices[fi2]];
                var inst = variant.createInstance();
                if (hasBoolOverrides) inst.setProperties(boolOverrides);
                var nestedChild = findNestedInstanceByName(inst, nested.instanceName);
                if (nestedChild) nestedChild.swapComponent(targetComponent);
                instancesContainer.appendChild(inst);
              } catch (e) {
                console.log('[Nested Instances] Error creating instance from "' + variants[fullIndices[fi2]].name + '":', e.message);
              }
            }

            valueFrame.appendChild(instancesContainer);
          }

          valuesContainer.appendChild(valueFrame);
        }

        propFrame.appendChild(valuesContainer);
        propsContainer.appendChild(propFrame);
      }

      setFrame.appendChild(propsContainer);
      section.appendChild(setFrame);
    }

    if (section.children.length === 0) return null;
  } else {
    // ─── Manual positioning (original behavior) ───
    var csWidth = sourceNode.width;
    var csHeight = standalone ? variants[0].height : sourceNode.height;
    var yOffset = 0;
    var maxSectionWidth = 0;

    for (var ni = 0; ni < nestedSets.length; ni++) {
      var nested = nestedSets[ni];

      var csHeader = createTextNode(nested.componentSetName, LABEL_FONT_SIZE);
      csHeader.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.5 }];
      csHeader.x = 0;
      csHeader.y = yOffset;
      section.appendChild(csHeader);
      yOffset += csHeader.height + LABEL_INSTANCE_GAP;

      var nestedGroupProps = nested.variantGroupProperties;
      var defaultProps = nested.defaultVariantProperties;
      var propNames = Object.keys(nestedGroupProps);

      if (DEBUG) {
        console.log('[Nested Instances] "' + nested.componentSetName + '": ' + propNames.length + ' properties, defaults:', JSON.stringify(defaultProps));
      }

      for (var pi = 0; pi < propNames.length; pi++) {
        var propName = propNames[pi];
        var propValues = nestedGroupProps[propName];
        var defaultValue = defaultProps[propName];

        var propHeader = createTextNode(propName, LABEL_FONT_SIZE);
        propHeader.fills = [{ type: 'SOLID', color: DOC_COLOR, opacity: 0.35 }];
        propHeader.x = 0;
        propHeader.y = yOffset;
        section.appendChild(propHeader);
        yOffset += propHeader.height + LABEL_INSTANCE_GAP;

        for (var vi2 = 0; vi2 < propValues.length; vi2++) {
          var value = propValues[vi2];
          if (value === defaultValue) continue;

          var targetProps = {};
          for (var key in defaultProps) {
            if (defaultProps.hasOwnProperty(key)) targetProps[key] = defaultProps[key];
          }
          targetProps[propName] = value;

          var swapTarget = findVariantByProperties(nested.variants, targetProps);
          if (!swapTarget) continue;

          var displayLabel = formatLabel(propName, value);
          var label = createTextNode(displayLabel, LABEL_FONT_SIZE);
          label.x = 0;
          label.y = yOffset;
          section.appendChild(label);
          yOffset += label.height + LABEL_INSTANCE_GAP;

          var targetComponent = await figma.getNodeByIdAsync(swapTarget.id);
          if (!targetComponent) continue;

          if (mode === 'representative') {
            var indicesToUse = [];
            if (nested.parentVariantIndices && nested.parentVariantIndices.length > 0) {
              indicesToUse.push(nested.parentVariantIndices[0]);
            } else {
              indicesToUse.push(0);
            }

            for (var bvi = 0; bvi < indicesToUse.length; bvi++) {
              try {
                var baseVariantIdx = indicesToUse[bvi];
                var instance = variants[baseVariantIdx].createInstance();
                if (hasBoolOverrides) instance.setProperties(boolOverrides);
                var nestedInst = findNestedInstanceByName(instance, nested.instanceName);
                if (nestedInst) nestedInst.swapComponent(targetComponent);
                instance.x = 0;
                instance.y = yOffset;
                section.appendChild(instance);
                if (instance.width > maxSectionWidth) maxSectionWidth = instance.width;
                yOffset += instance.height + COMBO_GAP;
              } catch (e) {
                console.log('[Nested Instances] Error creating representative instance:', e.message);
              }
            }
          } else {
            var fullIndices = [];
            if (nested.parentVariantIndices && nested.parentVariantIndices.length > 0) {
              fullIndices = nested.parentVariantIndices.slice();
            } else {
              for (var ai = 0; ai < variants.length; ai++) fullIndices.push(ai);
            }
            var comboY = yOffset;
            for (var fi2 = 0; fi2 < fullIndices.length; fi2++) {
              try {
                var variant = variants[fullIndices[fi2]];
                var inst = variant.createInstance();
                if (hasBoolOverrides) inst.setProperties(boolOverrides);
                var nestedChild = findNestedInstanceByName(inst, nested.instanceName);
                if (nestedChild) nestedChild.swapComponent(targetComponent);
                inst.x = standalone ? 0 : variant.x;
                inst.y = comboY;
                section.appendChild(inst);
                if (inst.width > maxSectionWidth) maxSectionWidth = inst.width;
                comboY += inst.height + COMBO_GAP;
              } catch (e) {
                console.log('[Nested Instances] Error creating instance from "' + variants[fullIndices[fi2]].name + '":', e.message);
              }
            }
            yOffset = comboY;
          }
        }

        if (pi < propNames.length - 1) {
          yOffset += PROP_GAP - COMBO_GAP;
        }
      }

      if (ni < nestedSets.length - 1) {
        yOffset += PROP_GAP - COMBO_GAP;
      }
    }

    var finalHeight = yOffset > COMBO_GAP ? yOffset - COMBO_GAP : 10;
    section.resize(Math.max(maxSectionWidth, 100), finalHeight);
  }

  if (DEBUG) { console.log('[Nested Instances] Section size:', section.width + 'x' + section.height); }
  return section;
}

// Find a nested INSTANCE node by layer name (recursive)
function findNestedInstanceByName(parent, name) {
  if (!('children' in parent)) return null;
  for (var i = 0; i < parent.children.length; i++) {
    var child = parent.children[i];
    if (child.type === 'INSTANCE' && child.name === name) return child;
    var found = findNestedInstanceByName(child, name);
    if (found) return found;
  }
  return null;
}

// ─── Main Generation (Step 6) ────────────────────────────────────────────────

async function generateForStandaloneComponent(component, options) {
  var t0 = Date.now();
  if (DEBUG) { console.log('[Generate Standalone] Options:', JSON.stringify(options)); }

  // Save settings on the component for later restoration
  saveComponentSettings(component, options);

  // Set doc color
  if (options.color) {
    DOC_COLOR = hexToRgb(options.color);
  } else {
    DOC_COLOR = { r: DEFAULT_COLOR.r, g: DEFAULT_COLOR.g, b: DEFAULT_COLOR.b };
  }

  _hidePropertyNames = options.hidePropertyNames || false;
  _autoLayoutExtras = options.autoLayoutExtras || false;

  if (!_fontLoaded && _fontLoadPromise) {
    figma.ui.postMessage({ type: 'status', message: 'Loading fonts...' });
    await _fontLoadPromise;
  }
  if (!_fontLoaded) throw new Error('Inter font not available. Try restarting Figma.');
  if (options.fontFamily) await loadLabelFont(options.fontFamily);
  initMeasureNode();

  // Create or reuse wrapper
  var wrapper = findExistingWrapper(component);
  var originalParent = component.parent;
  var originalX = component.x;
  var originalY = component.y;

  if (wrapper) {
    removeOldLabels(wrapper);
  } else {
    wrapper = figma.createFrame();
    wrapper.name = '❖ ' + component.name + (options.docLabel ? ' — ' + options.docLabel : ' — Obra Autodocs');
    wrapper.setPluginData('sourceComponentSetId', component.id);
    wrapper.fills = [];
    wrapper.clipsContent = false;

    var compIndex = originalParent.children.indexOf(component);
    originalParent.insertChild(compIndex, wrapper);
    wrapper.x = originalX;
    wrapper.y = originalY;

    wrapper.appendChild(component);
  }

  // ─── Boolean grid mode for standalone ───

  var standaloneBoolGrid = null;
  var sBoolLabelWidth = 0;
  var sBoolLabelHeight = 0;
  var SBOOL_GROUP_GAP = 24; // updated below for single-boolean cases

  if (options.showBooleanVisibility && (options.booleanDisplayMode || 'list') === 'grid') {
    var sBoolProps = getBooleanComponentProperties(component);
    if (options.enabledBooleanProps) {
      sBoolProps = sBoolProps.filter(function(p) { return options.enabledBooleanProps.indexOf(p.name) !== -1; });
    }
    if (sBoolProps.length > 0) {
      var sBoolAxis = determineBooleanAxis(component.width, component.height);
      var sBoolGroups = buildBooleanGroups(sBoolProps, options.booleanCombination || 'individual', 'truefalse');
      standaloneBoolGrid = {
        axis: sBoolAxis,
        defaultLabel: 'Base',
        groups: sBoolGroups
      };
      // Single-boolean: no gap (tint sits flush against divider line), multi-boolean: full gap for double-line dividers
      SBOOL_GROUP_GAP = sBoolProps.length > 1 ? 24 : 0;

      // Calculate boolean label dimensions
      if (sBoolAxis === 'row') {
        var sMaxBoolTextWidth = measureTextWidth(standaloneBoolGrid.defaultLabel, LABEL_FONT_SIZE);
        for (var sbli = 0; sbli < sBoolGroups.length; sbli++) {
          var sblw = measureTextWidth(sBoolGroups[sbli].label, LABEL_FONT_SIZE);
          if (sblw > sMaxBoolTextWidth) sMaxBoolTextWidth = sblw;
        }
        sBoolLabelWidth = Math.ceil(sMaxBoolTextWidth + LABEL_GAP * 2);
      } else {
        sBoolLabelHeight = SIMPLE_LABEL_ROW_HEIGHT;
      }
    }
  }

  // Reserve space for title + description + doc link
  var sTitleOffset = 0;
  var sDescOffset = 0;
  var sDocLinkOffset = 0;
  var sHeaderDescFrame = null;
  var sHeaderDocLinkText = null;
  if (options.showTitle) {
    sTitleOffset = TITLE_FONT_SIZE + TITLE_GAP;
  }
  var sDescText = component.descriptionMarkdown || component.description || '';
  if (options.showDescription && sDescText.trim()) {
    sHeaderDescFrame = createDescriptionFrame(sDescText, Math.max(component.width, 240));
    sDescOffset = sHeaderDescFrame.height + DESC_GAP;
  }
  var sDocLinkUri = getDocLink(component);
  if (options.showDocLink && sDocLinkUri) {
    sHeaderDocLinkText = createDocLinkText(sDocLinkUri);
    sDocLinkOffset = DOC_LINK_FONT_SIZE + DOC_LINK_GAP;
  }
  var sHeaderOffset = sTitleOffset + sDescOffset + sDocLinkOffset;
  sBoolLabelHeight += sHeaderOffset;

  // Position component (offset by boolean label area)
  component.constraints = { horizontal: 'MIN', vertical: 'MIN' };
  component.x = sBoolLabelWidth;
  component.y = sBoolLabelHeight;

  // Calculate expanded dimensions
  var sExpandedWidth = component.width;
  var sExpandedHeight = component.height;
  if (standaloneBoolGrid) {
    var sNumNonDefault = standaloneBoolGrid.groups.length;
    if (standaloneBoolGrid.axis === 'row') {
      sExpandedHeight = component.height + sNumNonDefault * (component.height + SBOOL_GROUP_GAP);
    } else {
      sExpandedWidth = component.width + sNumNonDefault * (component.width + SBOOL_GROUP_GAP);
    }
  }

  wrapper.resize(sBoolLabelWidth + sExpandedWidth, sBoolLabelHeight + sExpandedHeight);

  // ─── Header: title + description (standalone component) ───

  var scHeaderY = 0;
  if (options.showTitle && sTitleOffset > 0) {
    var scTitleText = createTitleText(component.name);
    scTitleText.x = sBoolLabelWidth;
    scTitleText.y = scHeaderY;
    wrapper.appendChild(scTitleText);
    scHeaderY += sTitleOffset;
  }
  if (sHeaderDescFrame) {
    sHeaderDescFrame.x = sBoolLabelWidth;
    sHeaderDescFrame.y = scHeaderY;
    wrapper.appendChild(sHeaderDescFrame);
    scHeaderY += sDescOffset;
  }
  if (sHeaderDocLinkText) {
    sHeaderDocLinkText.x = sBoolLabelWidth;
    sHeaderDocLinkText.y = scHeaderY;
    wrapper.appendChild(sHeaderDocLinkText);
  }

  // Create boolean grid instances + tints + labels
  if (standaloneBoolGrid) {
    for (var sbgi = 0; sbgi < standaloneBoolGrid.groups.length; sbgi++) {
      var sbGroup = standaloneBoolGrid.groups[sbgi];
      var sbGroupIndex = sbgi + 1;
      var sbOffsetX, sbOffsetY;

      if (standaloneBoolGrid.axis === 'row') {
        sbOffsetX = sBoolLabelWidth;
        sbOffsetY = sBoolLabelHeight + sbGroupIndex * (component.height + SBOOL_GROUP_GAP);
      } else {
        sbOffsetX = sBoolLabelWidth + sbGroupIndex * (component.width + SBOOL_GROUP_GAP);
        sbOffsetY = sBoolLabelHeight;
      }

      var sTint = createBackgroundTint(sbOffsetX, sbOffsetY, component.width, component.height);
      wrapper.appendChild(sTint);

      var sInstance = component.createInstance();
      sInstance.setProperties(sbGroup.props);
      sInstance.x = sbOffsetX;
      sInstance.y = sbOffsetY;
      wrapper.appendChild(sInstance);
    }

    // Boolean labels
    if (standaloneBoolGrid.axis === 'row') {
      var sDefLabel = createRowLabel(standaloneBoolGrid.defaultLabel, 0, sBoolLabelHeight, sBoolLabelWidth, component.height, false);
      wrapper.appendChild(sDefLabel);
      for (var sbli = 0; sbli < standaloneBoolGrid.groups.length; sbli++) {
        var sbLabelY = sBoolLabelHeight + (sbli + 1) * (component.height + SBOOL_GROUP_GAP);
        var sbLabel = createRowLabel(standaloneBoolGrid.groups[sbli].label, 0, sbLabelY, sBoolLabelWidth, component.height, false);
        wrapper.appendChild(sbLabel);
      }
    } else {
      var sDefLabel = createColLabel(standaloneBoolGrid.defaultLabel, sBoolLabelWidth, 0, component.width, sBoolLabelHeight, false);
      wrapper.appendChild(sDefLabel);
      for (var sbli = 0; sbli < standaloneBoolGrid.groups.length; sbli++) {
        var sbLabelX = sBoolLabelWidth + (sbli + 1) * (component.width + SBOOL_GROUP_GAP);
        var sbLabel = createColLabel(standaloneBoolGrid.groups[sbli].label, sbLabelX, 0, component.width, sBoolLabelHeight, false);
        wrapper.appendChild(sbLabel);
      }
    }
  }

  // Build extra sections (boolean list mode + nested)
  var extraSections = [];

  if (options.showBooleanVisibility && !standaloneBoolGrid) {
    figma.ui.postMessage({ type: 'status', message: 'Creating boolean visibility examples...' });
    var boolSection;
    if (options.booleanScope) {
      boolSection = await createScopedBooleanSection(component, options.booleanScope, options.booleanCombination || 'individual', options.booleanDisplayMode || 'list');
    } else {
      boolSection = await createBooleanVisibilitySection(component, options.booleanCombination || 'individual', options.enabledBooleanProps, options.booleanDisplayMode || 'list');
    }
    if (boolSection) extraSections.push(boolSection);
  }

  if (options.showNestedInstances) {
    figma.ui.postMessage({ type: 'status', message: 'Creating nested instance examples...' });
    var nestedSection = await createNestedInstancesSection(component, options.nestedInstancesMode || 'representative', options.enabledNestedInstances);
    if (nestedSection) extraSections.push(nestedSection);
  }

  if (extraSections.length > 0) {
    var EXTRAS_GAP = 24;
    var extrasContainer = figma.createFrame();
    extrasContainer.name = 'Property Combinations';
    extrasContainer.fills = [];
    extrasContainer.clipsContent = false;
    extrasContainer.layoutMode = 'VERTICAL';
    extrasContainer.itemSpacing = EXTRAS_GAP;
    extrasContainer.counterAxisSizingMode = 'AUTO';
    extrasContainer.primaryAxisSizingMode = 'AUTO';

    for (var i = 0; i < extraSections.length; i++) {
      extrasContainer.appendChild(extraSections[i]);
    }

    if (_autoLayoutExtras) {
      extrasContainer.paddingLeft = sBoolLabelWidth;
    } else {
      extrasContainer.x = sBoolLabelWidth;
      extrasContainer.y = sBoolLabelHeight + sExpandedHeight + EXTRAS_GAP;
    }
    wrapper.appendChild(extrasContainer);

    if (!_autoLayoutExtras) {
      wrapper.resize(
        Math.max(wrapper.width, sBoolLabelWidth + extrasContainer.width),
        extrasContainer.y + extrasContainer.height
      );
    }
  }

  if (_autoLayoutExtras) {
    // ─── Wrap in Documentation frame + make wrapper auto layout ───

    var scDocsFrame = figma.createFrame();
    scDocsFrame.name = 'Documentation';
    scDocsFrame.fills = [];
    scDocsFrame.clipsContent = false;
    scDocsFrame.locked = false;
    scDocsFrame.resize(sBoolLabelWidth + sExpandedWidth, sBoolLabelHeight + sExpandedHeight);
    wrapper.appendChild(scDocsFrame);

    var scDocsChildren = [];
    for (var scdi = 0; scdi < wrapper.children.length; scdi++) {
      var scChild = wrapper.children[scdi];
      if (scChild !== scDocsFrame && scChild.name !== 'Property Combinations') {
        scDocsChildren.push(scChild);
      }
    }
    for (var scdi2 = 0; scdi2 < scDocsChildren.length; scdi2++) {
      scDocsChildren[scdi2].constraints = { horizontal: 'MIN', vertical: 'MIN' };
      scDocsFrame.appendChild(scDocsChildren[scdi2]);
    }

    var scExtrasInWrapper = null;
    for (var scewi = 0; scewi < wrapper.children.length; scewi++) {
      if (wrapper.children[scewi].name === 'Property Combinations') {
        scExtrasInWrapper = wrapper.children[scewi];
        break;
      }
    }
    if (scExtrasInWrapper) {
      wrapper.appendChild(scExtrasInWrapper);
    }

    wrapper.layoutMode = 'VERTICAL';
    wrapper.primaryAxisSizingMode = 'AUTO';
    wrapper.counterAxisSizingMode = 'AUTO';
    wrapper.itemSpacing = 24;
    wrapper.paddingTop = 0;
    wrapper.paddingBottom = 0;
    wrapper.paddingLeft = 0;
    wrapper.paddingRight = 0;
  }

  disposeMeasureNode();

  var totalMs = Date.now() - t0;
  var totalSec = (totalMs / 1000).toFixed(2);
  console.log('[Perf] === Standalone generation time: ' + totalMs + 'ms (' + totalSec + 's) ===');

  figma.currentPage.selection = [wrapper];
  figma.viewport.scrollAndZoomIntoView([wrapper]);

  var doneMsg = 'Generated docs in ' + totalSec + 's.';
  figma.notify(doneMsg);
  figma.ui.postMessage({ type: 'done', message: doneMsg });
}

async function generate(options = {}) {
  var t0 = Date.now();
  if (DEBUG) { console.log('[Generate] Options received:', JSON.stringify(options)); }
  const cs = getComponentSet();
  if (!cs) {
    var standalone = getStandaloneComponent();
    if (standalone) {
      return generateForStandaloneComponent(standalone, options);
    }
    figma.ui.postMessage({ type: 'error', message: 'Please select a component set or component.' });
    return;
  }

  // Variable modes no longer force standalone — they attach to the regular wrapper

  // Save settings on the component set for later restoration
  saveComponentSettings(cs, options);

  // Set doc color from options
  if (options.color) {
    DOC_COLOR = hexToRgb(options.color);
  } else {
    DOC_COLOR = { r: DEFAULT_COLOR.r, g: DEFAULT_COLOR.g, b: DEFAULT_COLOR.b };
  }

  _hidePropertyNames = options.hidePropertyNames || false;
  _autoLayoutExtras = options.autoLayoutExtras || false;

  if (!_fontLoaded && _fontLoadPromise) {
    figma.ui.postMessage({ type: 'status', message: 'Loading fonts...' });
    await _fontLoadPromise;
  }
  if (!_fontLoaded) throw new Error('Inter font not available. Try restarting Figma.');
  if (options.fontFamily) await loadLabelFont(options.fontFamily);
  console.log('[Perf] Font loaded:', (Date.now() - t0) + 'ms');

  // Init reusable text measurement node
  initMeasureNode();

  // Get enum properties
  const enumProps = getEnumProperties(cs, options);
  if (DEBUG) { console.log('[Generate] properties:', enumProps.map(function(p) { return p.name; })); }

  var layout, grid, gridByRow, gridByCol, colClusters, rowClusters, colAxisProps, rowAxisProps, variants;

  if (enumProps.length === 0) {
    // No enum properties at all — wrap without labels
    variants = cs.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
    grid = [];
    gridByRow = {};
    gridByCol = {};
    colClusters = [];
    rowClusters = [];
    colAxisProps = [];
    rowAxisProps = [];
  } else {
    figma.ui.postMessage({ type: 'status', message: 'Analyzing layout...' });

    // Detect variants with absolute-positioned overflow (e.g. dropdown menus, tooltips)
    // before layout analysis so effectiveSize() uses visual sizes for clustering
    var overflowInfo = detectComponentSetOverflow(cs);
    if (overflowInfo.hasOverflow) {
      console.log('[Overflow] Accommodating absolute-positioned overflow in grid layout.');
      console.log('[Overflow] Affected: ' + overflowInfo.affectedVariants.length + ' variant(s), max overflow: bottom +' + Math.round(overflowInfo.maxOverflowBottom) + 'px, right +' + Math.round(overflowInfo.maxOverflowRight) + 'px');
      // Build per-variant visual size map: frame size + overflow
      _visualSizeOverrides = {};
      var ovVariants = cs.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
      for (var ovi = 0; ovi < ovVariants.length; ovi++) {
        var ov = detectVariantOverflow(ovVariants[ovi]);
        _visualSizeOverrides[ovVariants[ovi].id] = {
          width: ovVariants[ovi].width + Math.max(0, ov.overflowRight),
          height: ovVariants[ovi].height + Math.max(0, ov.overflowBottom)
        };
      }
      figma.notify('Overflow detected: grid adjusted for absolute-positioned elements. Organize tab may not work for this component.', { timeout: 6000 });
    }

    // Analyze layout (uses visual sizes via effectiveSize() when overflow is present)
    var tLayout = Date.now();
    layout = analyzeLayout(cs, enumProps, options.gridAlignment || 'auto', options.allowSpanning || false);
    if (DEBUG) { console.log('[Perf] Layout analysis:', (Date.now() - tLayout) + 'ms'); }

    // Clear visual size overrides after layout analysis
    _visualSizeOverrides = null;

    grid = layout.grid;
    gridByRow = layout.gridByRow;
    gridByCol = layout.gridByCol;
    colClusters = layout.colClusters;
    rowClusters = layout.rowClusters;
    colAxisProps = layout.colAxisProps;
    rowAxisProps = layout.rowAxisProps;
    variants = layout.variants;
  }

  var singleTypeGrid = layout ? layout.singleTypeGrid : false;
  var singleTypeProp = layout ? layout.singleTypeProp : null;
  if (singleTypeGrid) {
    console.log('[Generate] Single-type grid mode: per-item labels for "' + singleTypeProp.name + '"');
    figma.notify('Single-property grid detected — labels will appear below each item.', { timeout: 4000 });
  }

  // ─── Disable auto-layout on CS before any repositioning ───
  // Component sets often have auto-layout (HORIZONTAL with wrap). Any resize or
  // position change would trigger Figma to reflow children. We disable it here,
  // saving and restoring positions since setting layoutMode='NONE' can flatten
  // a wrapped layout (moving children to a single row).
  if (cs.layoutMode && cs.layoutMode !== 'NONE') {
    if (DEBUG) { console.log('[Generate] Disabling CS auto-layout (was: ' + cs.layoutMode + ')'); }
    var _savedPos = {};
    for (var _si = 0; _si < variants.length; _si++) {
      _savedPos[variants[_si].id] = { x: variants[_si].x, y: variants[_si].y };
    }
    var _savedW = cs.width;
    var _savedH = cs.height;
    cs.layoutMode = 'NONE';
    // Restore positions and size that Figma changed during the layoutMode switch
    cs.resize(_savedW, _savedH);
    for (var _ri = 0; _ri < variants.length; _ri++) {
      var _sp = _savedPos[variants[_ri].id];
      variants[_ri].x = _sp.x;
      variants[_ri].y = _sp.y;
    }
  }

  figma.ui.postMessage({ type: 'status', message: 'Calculating dimensions...' });
  var tDims = Date.now();

  // ─── Boolean grid mode detection ───

  var booleanGridInfo = null;
  var BOOL_GROUP_GAP = 24; // updated below for single-boolean cases

  if (options.showBooleanVisibility && (options.booleanDisplayMode || 'list') === 'grid') {
    var boolProps = getBooleanComponentProperties(cs);
    if (options.enabledBooleanProps) {
      boolProps = boolProps.filter(function(p) { return options.enabledBooleanProps.indexOf(p.name) !== -1; });
    }
    if (boolProps.length > 0) {
      var boolAxis = determineBooleanAxis(cs.width, cs.height);
      var boolGroups = buildBooleanGroups(boolProps, options.booleanCombination || 'individual', 'truefalse');
      booleanGridInfo = {
        axis: boolAxis,
        defaultLabel: 'Base',
        groups: boolGroups,
        numGroups: boolGroups.length + 1,
        boolPropCount: boolProps.length
      };
      // Single-boolean: no gap (tint sits flush against divider line), multi-boolean: full gap for double-line dividers
      BOOL_GROUP_GAP = boolProps.length > 1 ? 24 : 0;
      if (DEBUG) { console.log('[Boolean Grid] axis:', boolAxis, 'groups:', boolGroups.length, 'default:', booleanGridInfo.defaultLabel, 'gap:', BOOL_GROUP_GAP); }
    }
  }

  // ─── Calculate label column widths (for row-axis properties) ───

  const rowLabelWidths = [];
  for (let i = 0; i < rowAxisProps.length; i++) {
    const prop = rowAxisProps[i];
    const hasBracket = rowAxisProps.length > 1 && i < rowAxisProps.length - 1;

    let maxTextWidth = 0;
    for (const val of prop.values) {
      const displayText = formatLabel(prop.name, val);
      const w = measureTextWidth(displayText, LABEL_FONT_SIZE);
      if (w > maxTextWidth) maxTextWidth = w;
    }

    const labelWidth = hasBracket
      ? Math.ceil(maxTextWidth + LABEL_GAP + BRACKET_THICKNESS + LABEL_GAP)
      : Math.ceil(maxTextWidth + LABEL_GAP * 2);
    rowLabelWidths.push(labelWidth);
  }

  const totalLabelWidth = rowLabelWidths.reduce((sum, w) => sum + w, 0);

  // ─── Calculate label row heights (for column-axis properties) ───

  const colLabelHeights = [];
  for (let i = 0; i < colAxisProps.length; i++) {
    const hasBracket = colAxisProps.length > 1 && i < colAxisProps.length - 1;
    colLabelHeights.push(hasBracket ? BRACKET_LABEL_ROW_HEIGHT : SIMPLE_LABEL_ROW_HEIGHT);
  }

  const totalLabelHeight = colLabelHeights.reduce((sum, h) => sum + h, 0);

  // ─── Boolean grid label dimensions ───

  var boolLabelWidth = 0;
  var boolLabelHeight = 0;

  if (booleanGridInfo) {
    if (booleanGridInfo.axis === 'row') {
      var hasBoolBracket = rowAxisProps.length > 0;
      var maxBoolTextWidth = measureTextWidth(booleanGridInfo.defaultLabel, LABEL_FONT_SIZE);
      for (var bli = 0; bli < booleanGridInfo.groups.length; bli++) {
        var blw = measureTextWidth(booleanGridInfo.groups[bli].label, LABEL_FONT_SIZE);
        if (blw > maxBoolTextWidth) maxBoolTextWidth = blw;
      }
      boolLabelWidth = hasBoolBracket
        ? Math.ceil(maxBoolTextWidth + LABEL_GAP + BRACKET_THICKNESS + LABEL_GAP)
        : Math.ceil(maxBoolTextWidth + LABEL_GAP * 2);
    } else {
      var hasBoolBracket = colAxisProps.length > 0;
      boolLabelHeight = hasBoolBracket ? BRACKET_LABEL_ROW_HEIGHT : SIMPLE_LABEL_ROW_HEIGHT;
    }
  }

  var adjustedTotalLabelWidth = totalLabelWidth + boolLabelWidth;
  var adjustedTotalLabelHeight = totalLabelHeight + boolLabelHeight;

  // Reserve space above labels for title, description, and doc link
  var titleOffset = 0;
  var descOffset = 0;
  var docLinkOffset = 0;
  var headerDescFrame = null;
  var headerDocLinkText = null;
  if (options.showTitle) {
    titleOffset = TITLE_FONT_SIZE + TITLE_GAP;
  }
  var descText = cs.descriptionMarkdown || cs.description || '';
  if (options.showDescription && descText.trim()) {
    console.log('[Description RAW]', JSON.stringify(descText));
    headerDescFrame = createDescriptionFrame(descText, Math.max(cs.width, 240));
    descOffset = headerDescFrame.height + DESC_GAP;
  }
  var docLinkUri = getDocLink(cs);
  if (options.showDocLink && docLinkUri) {
    headerDocLinkText = createDocLinkText(docLinkUri);
    docLinkOffset = DOC_LINK_FONT_SIZE + DOC_LINK_GAP;
  }
  var headerOffset = titleOffset + descOffset + docLinkOffset;
  adjustedTotalLabelHeight += headerOffset;

  // ─── Calculate column widths and row heights from bounding boxes ───

  var colWidths = [];
  if (layout && layout.colMaxes) {
    for (var c = 0; c < colClusters.length; c++) {
      colWidths.push(layout.colMaxes[c] - colClusters[c]);
    }
  }

  var rowHeights = [];
  if (layout && layout.rowMaxes) {
    for (var r = 0; r < rowClusters.length; r++) {
      rowHeights.push(layout.rowMaxes[r] - rowClusters[r]);
    }
  }

  if (DEBUG) {
    console.log('[Step 3] Column widths (variant sizes):', colWidths);
    console.log('[Step 3] Row heights (variant sizes):', rowHeights);
  }

  // ═══════════════════════════════════════════════════════════════════════════════
  // VARIABLE MODES — separate action, generates mode columns in its own frame
  // ═══════════════════════════════════════════════════════════════════════════════

  // Variable modes handled in Step 9b below (uses .groups format)

  // ═══════════════════════════════════════════════════════════════════════════════
  // STANDALONE DOC MODE — generate docs as a separate frame, CS stays untouched
  // ═══════════════════════════════════════════════════════════════════════════════

  if (options.standaloneDoc) {
    figma.ui.postMessage({ type: 'status', message: 'Creating standalone docs...' });

    // Build position map for instances (variant ID → position within grid)
    var saPositions = {};
    for (var sapi = 0; sapi < grid.length; sapi++) {
      saPositions[grid[sapi].node.id] = { x: grid[sapi].node.x, y: grid[sapi].node.y };
    }

    // ─── Uniform cell grid (compute positions for instances, don't modify CS) ───

    if (colAxisProps.length > 0 && colClusters.length > 1) {
      var saInnerColProp = colAxisProps[colAxisProps.length - 1];
      var saColLabelTextWidths = [];
      for (var sac = 0; sac < colClusters.length; sac++) {
        var saInCol = gridByCol[sac] || [];
        if (saInCol.length === 0) { saColLabelTextWidths.push(0); continue; }
        var saValue = saInCol[0].props[saInnerColProp.name];
        var saDisplayText = formatLabel(saInnerColProp.name, saValue);
        saColLabelTextWidths.push(measureTextWidth(saDisplayText, LABEL_FONT_SIZE) + LABEL_GAP * 2);
      }
      var saMaxLabelW = Math.max.apply(null, saColLabelTextWidths);
      var saMaxVariantW = Math.max.apply(null, colWidths);
      if (saMaxLabelW > saMaxVariantW) {
        var saOrigCC = colClusters.slice();
        var saOrigCW = colWidths.slice();
        var saFirstCenter = saOrigCC[0] + saOrigCW[0] / 2;
        var saLastCenter = saOrigCC[saOrigCC.length - 1] + saOrigCW[saOrigCW.length - 1] / 2;
        var saOrigPitch = (saLastCenter - saFirstCenter) / (colClusters.length - 1);
        var saCellW = Math.max(saOrigPitch, saMaxLabelW);
        for (var sac2 = 0; sac2 < colClusters.length; sac2++) {
          var saIdealCenter = saCellW / 2 + sac2 * saCellW;
          var saVarHalfW = saOrigCW[sac2] / 2;
          var saIdealLeft = saIdealCenter - saVarHalfW;
          var saInCol2 = gridByCol[sac2] || [];
          for (var saci = 0; saci < saInCol2.length; saci++) {
            saPositions[saInCol2[saci].node.id].x = saIdealLeft + (saInCol2[saci].node.x - saOrigCC[sac2]);
          }
          colClusters[sac2] = sac2 * saCellW;
          colWidths[sac2] = saCellW;
        }
      }
    }

    if (rowAxisProps.length > 0 && rowClusters.length > 1) {
      var saMinRowLabelH = LABEL_FONT_SIZE + LABEL_GAP * 2;
      var saMaxVariantH = Math.max.apply(null, rowHeights);
      if (saMinRowLabelH > saMaxVariantH) {
        var saOrigRC = rowClusters.slice();
        var saOrigRH = rowHeights.slice();
        var saFirstRCenter = saOrigRC[0] + saOrigRH[0] / 2;
        var saLastRCenter = saOrigRC[saOrigRC.length - 1] + saOrigRH[saOrigRH.length - 1] / 2;
        var saOrigRPitch = (saLastRCenter - saFirstRCenter) / (rowClusters.length - 1);
        var saCellH = Math.max(saOrigRPitch, saMinRowLabelH);
        for (var sar = 0; sar < rowClusters.length; sar++) {
          var saRIdealCenter = saCellH / 2 + sar * saCellH;
          var saVarHalfH = saOrigRH[sar] / 2;
          var saRIdealTop = saRIdealCenter - saVarHalfH;
          var saInRow = gridByRow[sar] || [];
          for (var sari = 0; sari < saInRow.length; sari++) {
            saPositions[saInRow[sari].node.id].y = saRIdealTop + (saInRow[sari].node.y - saOrigRC[sar]);
          }
          rowClusters[sar] = sar * saCellH;
          rowHeights[sar] = saCellH;
        }
      }
    }

    // ─── Single-type grid: build uniform grid with per-item label space ───

    var stGridInfo = null; // shared across standalone + variable modes
    if (singleTypeGrid) {
      stGridInfo = buildSingleTypeUniformGrid(
        grid, gridByCol, colClusters, colWidths, rowClusters, rowHeights, singleTypeProp,
        function(g, x, y) { saPositions[g.node.id].x = x; saPositions[g.node.id].y = y; }
      );
    }

    // Compute grid dimensions from clusters
    var saGridW, saGridH;
    if (singleTypeGrid) {
      var saStLastRowBottom = rowClusters[rowClusters.length - 1] + rowHeights[rowHeights.length - 1];
      saGridW = colClusters[colClusters.length - 1] + colWidths[colClusters.length - 1] + stGridInfo.pad;
      saGridH = saStLastRowBottom + stGridInfo.pad;
    } else {
      saGridW = colClusters.length > 0 ? (colClusters[colClusters.length - 1] + colWidths[colWidths.length - 1]) : cs.width;
      saGridH = rowClusters.length > 0 ? (rowClusters[rowClusters.length - 1] + rowHeights[rowHeights.length - 1]) : cs.height;
    }

    // ─── Create or reuse standalone wrapper ───

    var saWrapper = findStandaloneWrapper(cs);
    if (saWrapper) {
      removeOldLabels(saWrapper);
    } else {
      saWrapper = figma.createFrame();
      saWrapper.name = '❖ ' + cs.name + (options.docLabel ? ' — ' + options.docLabel : ' — Obra Autodocs');
      saWrapper.setPluginData('standaloneDoc', 'true');
      saWrapper.setPluginData('sourceComponentSetId', cs.id);
      saWrapper.fills = [];
      saWrapper.clipsContent = false;
    }
    // Place wrapper to the right of the parent frame (outside of it, on the page/grandparent)
    // If the CS lives directly on the page, place next to the CS itself
    var saParentFrame = cs.parent;
    var saIsOnPage = saParentFrame.type === 'PAGE';
    var saRefFrame = saIsOnPage ? cs : saParentFrame;
    var saPlacementParent = saIsOnPage ? saParentFrame : (saParentFrame.parent || saParentFrame);
    if (saWrapper.parent !== saPlacementParent) {
      saPlacementParent.appendChild(saWrapper);
    }
    saWrapper.x = saRefFrame.x + saRefFrame.width + 48;
    saWrapper.y = saRefFrame.y;

    // ─── Header: title + description (standalone doc) ───

    var saHeaderY = 0;
    if (options.showTitle && titleOffset > 0) {
      var saTitleText = createTitleText(cs.name);
      saTitleText.x = adjustedTotalLabelWidth;
      saTitleText.y = saHeaderY;
      saWrapper.appendChild(saTitleText);
      saHeaderY += titleOffset;
    }
    if (headerDescFrame) {
      headerDescFrame.x = adjustedTotalLabelWidth;
      headerDescFrame.y = saHeaderY;
      saWrapper.appendChild(headerDescFrame);
      saHeaderY += descOffset;
    }
    if (headerDocLinkText) {
      headerDocLinkText.x = adjustedTotalLabelWidth;
      headerDocLinkText.y = saHeaderY;
      saWrapper.appendChild(headerDocLinkText);
    }

    // ─── Create variant instances in grid ───

    var saExpandedW = saGridW;
    var saExpandedH = saGridH;

    for (var sagi = 0; sagi < grid.length; sagi++) {
      var saG = grid[sagi];
      var saInst = saG.node.createInstance();
      var saPos = saPositions[saG.node.id];
      saInst.x = adjustedTotalLabelWidth + saPos.x;
      saInst.y = adjustedTotalLabelHeight + saPos.y;
      saWrapper.appendChild(saInst);
    }

    // ─── Per-item labels for single-type grid ───

    if (singleTypeGrid) {
      createSingleTypeGridLabels(saWrapper, grid, singleTypeProp, colClusters, colWidths,
        rowClusters, stGridInfo.maxVarH, adjustedTotalLabelWidth, adjustedTotalLabelHeight);
    }

    // ─── Dashed border around grid area ───

    var saBorder = figma.createFrame();
    saBorder.name = 'Grid Border';
    saBorder.fills = [];
    saBorder.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
    saBorder.dashPattern = [4, 4];
    saBorder.x = adjustedTotalLabelWidth;
    saBorder.y = adjustedTotalLabelHeight;
    saBorder.resize(saGridW, saGridH);
    saWrapper.appendChild(saBorder);

    // ─── Boolean grid (standalone) ───

    var saBoolGroupHeights = [];
    var saBoolGroupWidths = [];
    var saBoolGroupData = [];

    if (booleanGridInfo) {
      var saGridPadTop = rowClusters.length > 0 ? rowClusters[0] : 0;
      var saGridPadLeft = colClusters.length > 0 ? colClusters[0] : 0;
      var saGridPadBottom = saGridH - (rowClusters.length > 0 ? (rowClusters[rowClusters.length - 1] + rowHeights[rowHeights.length - 1]) : saGridH);
      var saGridPadRight = saGridW - (colClusters.length > 0 ? (colClusters[colClusters.length - 1] + colWidths[colWidths.length - 1]) : saGridW);

      var saBoolTotalCombs = booleanGridInfo.groups.length * grid.length;
      var saUseBoolAccuracy = options.booleanImprovedAccuracy && saBoolTotalCombs <= BOOL_ACCURACY_MAX_COMBINATIONS;

      for (var sabgi = 0; sabgi < booleanGridInfo.groups.length; sabgi++) {
        var saBGroup = booleanGridInfo.groups[sabgi];
        var saBResult = createBooleanInstanceGroup(cs, saBGroup.props, 0, 0);
        var saAdjRow = null, saAdjCol = null;

        if (saUseBoolAccuracy) {
          if (rowClusters.length > 0) {
            var saDRC = [], saDRH = [], saDCurY = saGridPadTop;
            for (var sadri = 0; sadri < rowClusters.length; sadri++) {
              var saDMaxH = rowHeights[sadri];
              for (var sadii = 0; sadii < saBResult.instances.length; sadii++) {
                var _saiy = saBResult.instances[sadii].y;
                if (_saiy >= rowClusters[sadri] - CLUSTER_TOLERANCE && _saiy <= rowClusters[sadri] + rowHeights[sadri] + CLUSTER_TOLERANCE) {
                  if (saBResult.instances[sadii].height > saDMaxH) saDMaxH = saBResult.instances[sadii].height;
                }
              }
              saDRC.push(saDCurY);
              saDRH.push(saDMaxH);
              if (sadri < rowClusters.length - 1) {
                saDCurY += saDMaxH + (rowClusters[sadri + 1] - (rowClusters[sadri] + rowHeights[sadri]));
              }
            }
            saAdjRow = { clusters: saDRC, heights: saDRH };
          }
          if (colClusters.length > 0) {
            var saDCC = [], saDCW = [], saDCurX = saGridPadLeft;
            for (var sadci = 0; sadci < colClusters.length; sadci++) {
              var saDMaxW = colWidths[sadci];
              for (var sadci2 = 0; sadci2 < saBResult.instances.length; sadci2++) {
                var _saix = saBResult.instances[sadci2].x;
                if (_saix >= colClusters[sadci] - CLUSTER_TOLERANCE && _saix <= colClusters[sadci] + colWidths[sadci] + CLUSTER_TOLERANCE) {
                  if (saBResult.instances[sadci2].width > saDMaxW) saDMaxW = saBResult.instances[sadci2].width;
                }
              }
              saDCC.push(saDCurX);
              saDCW.push(saDMaxW);
              if (sadci < colClusters.length - 1) {
                saDCurX += saDMaxW + (colClusters[sadci + 1] - (colClusters[sadci] + colWidths[sadci]));
              }
            }
            saAdjCol = { clusters: saDCC, widths: saDCW };
          }
        }

        // Reposition instances to centered positions
        if (saAdjRow || saAdjCol) {
          for (var sarii = 0; sarii < saBResult.instances.length; sarii++) {
            var saRI = saBResult.instances[sarii];
            if (saAdjRow) {
              for (var sarmi = 0; sarmi < rowClusters.length; sarmi++) {
                if (saRI.y >= rowClusters[sarmi] - CLUSTER_TOLERANCE && saRI.y <= rowClusters[sarmi] + rowHeights[sarmi] + CLUSTER_TOLERANCE) {
                  saRI.y = saAdjRow.clusters[sarmi] + (saAdjRow.heights[sarmi] - saRI.height) / 2;
                  break;
                }
              }
            }
            if (saAdjCol) {
              for (var sacmi = 0; sacmi < colClusters.length; sacmi++) {
                if (saRI.x >= colClusters[sacmi] - CLUSTER_TOLERANCE && saRI.x <= colClusters[sacmi] + colWidths[sacmi] + CLUSTER_TOLERANCE) {
                  saRI.x = saAdjCol.clusters[sacmi] + (saAdjCol.widths[sacmi] - saRI.width) / 2;
                  break;
                }
              }
            }
          }
        }

        var saGroupH, saGroupW;
        if (saAdjRow) {
          var saLRB = saAdjRow.clusters[saAdjRow.clusters.length - 1] + saAdjRow.heights[saAdjRow.heights.length - 1];
          saGroupH = Math.max(saGridH, saLRB + saGridPadBottom);
        } else {
          saGroupH = Math.max(saGridH, saBResult.height + saGridPadBottom);
        }
        if (saAdjCol) {
          var saLCR = saAdjCol.clusters[saAdjCol.clusters.length - 1] + saAdjCol.widths[saAdjCol.widths.length - 1];
          saGroupW = Math.max(saGridW, saLCR + saGridPadRight);
        } else {
          saGroupW = Math.max(saGridW, saBResult.width + saGridPadRight);
        }

        saBoolGroupHeights.push(saGroupH);
        saBoolGroupWidths.push(saGroupW);
        saBoolGroupData.push({ instances: saBResult.instances, width: saBResult.width, height: saBResult.height, adjRowData: saAdjRow, adjColData: saAdjCol });
      }

      // Calculate expanded dimensions
      if (booleanGridInfo.axis === 'row') {
        var saTotalGH = 0;
        for (var sabhi = 0; sabhi < saBoolGroupHeights.length; sabhi++) {
          saTotalGH += saBoolGroupHeights[sabhi] + BOOL_GROUP_GAP;
        }
        saExpandedH = saGridH + saTotalGH;
      } else {
        var saTotalGW = 0;
        for (var sabwi = 0; sabwi < saBoolGroupWidths.length; sabwi++) {
          saTotalGW += saBoolGroupWidths[sabwi] + BOOL_GROUP_GAP;
        }
        saExpandedW = saGridW + saTotalGW;
      }

      // Update border to cover expanded grid
      if (booleanGridInfo.axis === 'row') {
        saBorder.resize(saGridW, saExpandedH);
      } else {
        saBorder.resize(saExpandedW, saGridH);
      }
    }

    // Resize wrapper
    saWrapper.resize(
      adjustedTotalLabelWidth + saExpandedW,
      adjustedTotalLabelHeight + saExpandedH
    );

    // ─── Position boolean grid instances, tints, and labels ───

    if (booleanGridInfo) {
      var saBCumOffset = 0;
      for (var sabgi2 = 0; sabgi2 < booleanGridInfo.groups.length; sabgi2++) {
        var saBGH = saBoolGroupHeights[sabgi2];
        var saBGW = saBoolGroupWidths[sabgi2];
        var saBGOffX, saBGOffY;
        if (booleanGridInfo.axis === 'row') {
          saBGOffX = adjustedTotalLabelWidth;
          saBGOffY = adjustedTotalLabelHeight + saGridH + saBCumOffset + BOOL_GROUP_GAP;
          saBCumOffset += saBGH + BOOL_GROUP_GAP;
        } else {
          saBGOffX = adjustedTotalLabelWidth + saGridW + saBCumOffset + BOOL_GROUP_GAP;
          saBGOffY = adjustedTotalLabelHeight;
          saBCumOffset += saBGW + BOOL_GROUP_GAP;
        }
        saBoolGroupData[sabgi2].offsetX = saBGOffX;
        saBoolGroupData[sabgi2].offsetY = saBGOffY;

        // Background tint
        var saTintW = booleanGridInfo.axis === 'row' ? saGridW : saBGW;
        var saTintH = booleanGridInfo.axis === 'row' ? saBGH : saGridH;
        var saTint = createBackgroundTint(saBGOffX, saBGOffY, saTintW, saTintH);
        saWrapper.appendChild(saTint);

        // Reposition boolean instances
        var saBInsts = saBoolGroupData[sabgi2].instances;
        for (var sabii = 0; sabii < saBInsts.length; sabii++) {
          saBInsts[sabii].x += saBGOffX;
          saBInsts[sabii].y += saBGOffY;
          saWrapper.appendChild(saBInsts[sabii]);
        }
      }

      // Boolean axis labels
      if (booleanGridInfo.axis === 'row') {
        var saHBB = rowAxisProps.length > 0;
        saWrapper.appendChild(createRowLabel(booleanGridInfo.defaultLabel, 0, adjustedTotalLabelHeight, boolLabelWidth, saGridH, saHBB));
        for (var sabli = 0; sabli < booleanGridInfo.groups.length; sabli++) {
          saWrapper.appendChild(createRowLabel(booleanGridInfo.groups[sabli].label, 0, saBoolGroupData[sabli].offsetY, boolLabelWidth, saBoolGroupHeights[sabli], saHBB));
        }
      } else {
        var saHBB = colAxisProps.length > 0;
        saWrapper.appendChild(createColLabel(booleanGridInfo.defaultLabel, adjustedTotalLabelWidth, headerOffset, saGridW, boolLabelHeight, saHBB));
        for (var sabli = 0; sabli < booleanGridInfo.groups.length; sabli++) {
          saWrapper.appendChild(createColLabel(booleanGridInfo.groups[sabli].label, saBoolGroupData[sabli].offsetX, headerOffset, saBoolGroupWidths[sabli], boolLabelHeight, saHBB));
        }
      }
    }

    // ─── Row labels (left side) ───

    var saLabelXOff = boolLabelWidth;
    for (var sarpi = 0; sarpi < rowAxisProps.length; sarpi++) {
      var saRProp = rowAxisProps[sarpi];
      var saRColW = rowLabelWidths[sarpi];
      var saRIsInner = rowAxisProps.length > 1 && sarpi === rowAxisProps.length - 1;
      var saRHasBracket = !saRIsInner && rowAxisProps.length > 1;
      var saRGroups = getRowGroups(saRProp.name, grid, gridByRow, rowClusters, rowHeights);
      for (var sarvi = 0; sarvi < saRGroups.length; sarvi++) {
        var saRG = saRGroups[sarvi];
        saWrapper.appendChild(createRowLabel(
          formatLabel(saRProp.name, saRG.value), saLabelXOff,
          adjustedTotalLabelHeight + saRG.startY, saRColW,
          saRG.endY - saRG.startY, saRHasBracket
        ));
      }
      saLabelXOff += saRColW;
    }

    // ─── Column labels (top) ───

    var saLabelYOff = boolLabelHeight + headerOffset;
    for (var sacpi = 0; sacpi < colAxisProps.length; sacpi++) {
      var saCProp = colAxisProps[sacpi];
      var saCRowH = colLabelHeights[sacpi];
      var saCIsInner = colAxisProps.length > 1 && sacpi === colAxisProps.length - 1;
      var saCHasBracket = !saCIsInner && colAxisProps.length > 1;
      var saCGroups = getColGroups(saCProp.name, grid, gridByCol, colClusters, colWidths);
      for (var sacvi = 0; sacvi < saCGroups.length; sacvi++) {
        var saCG = saCGroups[sacvi];
        saWrapper.appendChild(createColLabel(
          formatLabel(saCProp.name, saCG.value),
          adjustedTotalLabelWidth + saCG.startX, saLabelYOff,
          saCG.endX - saCG.startX, saCRowH, saCHasBracket
        ));
      }
      saLabelYOff += saCRowH;
    }

    // ─── Repeat labels for boolean groups ───

    if (booleanGridInfo) {
      for (var sabrgi = 0; sabrgi < booleanGridInfo.groups.length; sabrgi++) {
        var sabrAdjRC = (saBoolGroupData[sabrgi].adjRowData && saBoolGroupData[sabrgi].adjRowData.clusters) || rowClusters;
        var sabrAdjRH = (saBoolGroupData[sabrgi].adjRowData && saBoolGroupData[sabrgi].adjRowData.heights) || rowHeights;
        var sabrAdjCC = (saBoolGroupData[sabrgi].adjColData && saBoolGroupData[sabrgi].adjColData.clusters) || colClusters;
        var sabrAdjCW = (saBoolGroupData[sabrgi].adjColData && saBoolGroupData[sabrgi].adjColData.widths) || colWidths;
        if (booleanGridInfo.axis === 'row') {
          var sabrLXO = boolLabelWidth;
          for (var sabrpi = 0; sabrpi < rowAxisProps.length; sabrpi++) {
            var sabrP = rowAxisProps[sabrpi];
            var sabrCW = rowLabelWidths[sabrpi];
            var sabrIsInner = rowAxisProps.length > 1 && sabrpi === rowAxisProps.length - 1;
            var sabrHB = !sabrIsInner && rowAxisProps.length > 1;
            var sabrVGs = getRowGroups(sabrP.name, grid, gridByRow, sabrAdjRC, sabrAdjRH);
            for (var sabrvi = 0; sabrvi < sabrVGs.length; sabrvi++) {
              var sabrVG = sabrVGs[sabrvi];
              saWrapper.appendChild(createRowLabel(
                formatLabel(sabrP.name, sabrVG.value), sabrLXO,
                saBoolGroupData[sabrgi].offsetY + sabrVG.startY,
                sabrCW, sabrVG.endY - sabrVG.startY, sabrHB
              ));
            }
            sabrLXO += sabrCW;
          }
        } else {
          var sabrLYO = boolLabelHeight + headerOffset;
          for (var sabrpi = 0; sabrpi < colAxisProps.length; sabrpi++) {
            var sabrP = colAxisProps[sabrpi];
            var sabrRH = colLabelHeights[sabrpi];
            var sabrIsInner = colAxisProps.length > 1 && sabrpi === colAxisProps.length - 1;
            var sabrHB = !sabrIsInner && colAxisProps.length > 1;
            var sabrVGs = getColGroups(sabrP.name, grid, gridByCol, sabrAdjCC, sabrAdjCW);
            for (var sabrvi = 0; sabrvi < sabrVGs.length; sabrvi++) {
              var sabrVG = sabrVGs[sabrvi];
              saWrapper.appendChild(createColLabel(
                formatLabel(sabrP.name, sabrVG.value),
                saBoolGroupData[sabrgi].offsetX + sabrVG.startX, sabrLYO,
                sabrVG.endX - sabrVG.startX, sabrRH, sabrHB
              ));
            }
            sabrLYO += sabrRH;
          }
        }
      }
    }

    // ─── Grid lines ───

    if (options.showGrid) {
      createGridLines(saWrapper, colClusters, rowClusters, colWidths, rowHeights,
        adjustedTotalLabelWidth, adjustedTotalLabelHeight, saGridW, saGridH,
        singleTypeGrid ? grid : null);

      if (booleanGridInfo) {
        for (var saggi = 0; saggi < booleanGridInfo.groups.length; saggi++) {
          var sagRC = (saBoolGroupData[saggi].adjRowData && saBoolGroupData[saggi].adjRowData.clusters) || rowClusters;
          var sagRH = (saBoolGroupData[saggi].adjRowData && saBoolGroupData[saggi].adjRowData.heights) || rowHeights;
          var sagCC = (saBoolGroupData[saggi].adjColData && saBoolGroupData[saggi].adjColData.clusters) || colClusters;
          var sagCW = (saBoolGroupData[saggi].adjColData && saBoolGroupData[saggi].adjColData.widths) || colWidths;
          if (booleanGridInfo.axis === 'row') {
            createGridLines(saWrapper, sagCC, sagRC, sagCW, sagRH,
              saBoolGroupData[saggi].offsetX, saBoolGroupData[saggi].offsetY, saGridW, saBoolGroupHeights[saggi]);
          } else {
            createGridLines(saWrapper, sagCC, sagRC, sagCW, sagRH,
              saBoolGroupData[saggi].offsetX, saBoolGroupData[saggi].offsetY, saBoolGroupWidths[saggi], saGridH);
          }
        }

        // Dividers between boolean groups
        var saUseDblLine = booleanGridInfo.boolPropCount > 1;
        for (var sadgi = 0; sadgi < booleanGridInfo.groups.length; sadgi++) {
          var saDivL1 = figma.createLine();
          saDivL1.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
          saDivL1.strokeWeight = GRID_STROKE_WEIGHT;
          saDivL1.dashPattern = GRID_DASH_PATTERN;
          if (booleanGridInfo.axis === 'row') {
            var saPrevBot = sadgi === 0 ? (adjustedTotalLabelHeight + saGridH) : (saBoolGroupData[sadgi - 1].offsetY + saBoolGroupHeights[sadgi - 1]);
            saDivL1.resize(saGridW, 0);
            saDivL1.x = adjustedTotalLabelWidth;
            saDivL1.y = saPrevBot;
          } else {
            var saPrevRight = sadgi === 0 ? (adjustedTotalLabelWidth + saGridW) : (saBoolGroupData[sadgi - 1].offsetX + saBoolGroupWidths[sadgi - 1]);
            saDivL1.rotation = -90;
            saDivL1.resize(saGridH, 0);
            saDivL1.x = saPrevRight;
            saDivL1.y = adjustedTotalLabelHeight;
          }
          saWrapper.appendChild(saDivL1);
          if (saUseDblLine) {
            var saDivL2 = figma.createLine();
            saDivL2.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
            saDivL2.strokeWeight = GRID_STROKE_WEIGHT;
            saDivL2.dashPattern = GRID_DASH_PATTERN;
            if (booleanGridInfo.axis === 'row') {
              saDivL2.resize(saGridW, 0);
              saDivL2.x = adjustedTotalLabelWidth;
              saDivL2.y = saBoolGroupData[sadgi].offsetY;
            } else {
              saDivL2.rotation = -90;
              saDivL2.resize(saGridH, 0);
              saDivL2.x = saBoolGroupData[sadgi].offsetX;
              saDivL2.y = adjustedTotalLabelHeight;
            }
            saWrapper.appendChild(saDivL2);
          }
        }
      }
    }

    // ─── Extra sections ───

    var saExtraSections = [];
    if (options.showBooleanVisibility && !booleanGridInfo) {
      figma.ui.postMessage({ type: 'status', message: 'Creating boolean visibility examples...' });
      var saBoolSection;
      if (options.booleanScope) {
        saBoolSection = await createScopedBooleanSection(cs, options.booleanScope, options.booleanCombination || 'individual', options.booleanDisplayMode || 'list');
      } else {
        saBoolSection = await createBooleanVisibilitySection(cs, options.booleanCombination || 'individual', options.enabledBooleanProps, options.booleanDisplayMode || 'list');
      }
      if (saBoolSection) saExtraSections.push(saBoolSection);
    }
    if (options.showNestedInstances) {
      figma.ui.postMessage({ type: 'status', message: 'Creating nested instance examples...' });
      var saNestedSection = await createNestedInstancesSection(cs, options.nestedInstancesMode || 'representative', options.enabledNestedInstances);
      if (saNestedSection) saExtraSections.push(saNestedSection);
    }
    if (saExtraSections.length > 0) {
      var SA_EXTRAS_GAP = 24;
      var saExtras = figma.createFrame();
      saExtras.name = 'Property Combinations';
      saExtras.fills = [];
      saExtras.clipsContent = false;
      saExtras.layoutMode = 'VERTICAL';
      saExtras.itemSpacing = SA_EXTRAS_GAP;
      saExtras.counterAxisSizingMode = 'AUTO';
      saExtras.primaryAxisSizingMode = 'AUTO';
      for (var saei = 0; saei < saExtraSections.length; saei++) {
        saExtras.appendChild(saExtraSections[saei]);
      }
      if (_autoLayoutExtras) {
        saExtras.paddingLeft = adjustedTotalLabelWidth;
      } else {
        saExtras.x = adjustedTotalLabelWidth;
        saExtras.y = adjustedTotalLabelHeight + saExpandedH + SA_EXTRAS_GAP;
      }
      saWrapper.appendChild(saExtras);
      if (!_autoLayoutExtras) {
        saWrapper.resize(
          Math.max(saWrapper.width, adjustedTotalLabelWidth + saExtras.width),
          saExtras.y + saExtras.height
        );
      }
    }

    // ─── Variable Modes (experimental) ───
    // Each selected mode creates an extension of the grid to the right.
    // The extension is a frame with the variable mode set — Figma propagates
    // the mode to all child instances inside that frame.

    var vmModesContainer = null;

    if (options.variableModes && options.variableModes.groups && options.variableModes.groups.length > 0) {
      figma.ui.postMessage({ type: 'status', message: 'Generating variable modes...' });

      var vmCollId = options.variableModes.collectionId;
      var vmCollection = await figma.variables.getVariableCollectionByIdAsync(vmCollId);
      var vmModes = options.variableModes.modes;
      console.log('[VarModes] Collection resolved:', vmCollection ? vmCollection.name : 'NOT FOUND');
      var VM_GAP = 48;
      var VM_LABEL_HEIGHT = SIMPLE_LABEL_ROW_HEIGHT;

      console.log('[VarModes] collectionId:', vmCollId);
      console.log('[VarModes] modes:', JSON.stringify(vmModes));
      console.log('[VarModes] grid:', grid.length, 'variants, saExpandedW:', saExpandedW, 'saExpandedH:', saExpandedH);
      console.log('[VarModes] colClusters:', JSON.stringify(colClusters), 'colWidths:', JSON.stringify(colWidths));

      // Try to find a background color variable from the collection for dark mode support
      var vmBgVariable = null;
      try {
        var vmVarIds = vmCollection.variableIds;
        for (var vmvi = 0; vmvi < vmVarIds.length; vmvi++) {
          var vmVar = await figma.variables.getVariableByIdAsync(vmVarIds[vmvi]);
          if (vmVar && vmVar.resolvedType === 'COLOR') {
            var vmVarName = vmVar.name.toLowerCase();
            if (vmVarName.indexOf('background') !== -1 || vmVarName.indexOf('bg') !== -1 || vmVarName.indexOf('surface') !== -1 || vmVarName.indexOf('body') !== -1) {
              vmBgVariable = vmVar;
              console.log('[VarModes] Found background variable:', vmVar.name, vmVar.id);
              break;
            }
          }
        }
      } catch (vmBgErr) {
        console.warn('[VarModes] Could not search for background variable:', vmBgErr.message);
      }

      // Create a "Modes" container — will be a sibling of the Documentation frame
      vmModesContainer = figma.createFrame();
      vmModesContainer.name = 'Modes';
      vmModesContainer.fills = [];
      vmModesContainer.clipsContent = false;

      // Create one extension per selected (non-default) mode
      var vmCumOffsetX = 0;

      for (var vmmi = 0; vmmi < vmModes.length; vmmi++) {
        var vmMode = vmModes[vmmi];
        var vmOffX = vmCumOffsetX;

        // Per-mode wrapper frame
        var vmModeWrapper = figma.createFrame();
        vmModeWrapper.name = 'Mode: ' + vmMode.name;
        vmModeWrapper.fills = [];
        vmModeWrapper.clipsContent = false;

        // Mode label (above the grid)
        vmModeWrapper.appendChild(createColLabel(vmMode.name, adjustedTotalLabelWidth, boolLabelHeight, saExpandedW, VM_LABEL_HEIGHT, false));

        // Mode frame — instances inside inherit this variable mode
        var vmFrame = figma.createFrame();
        vmFrame.name = vmMode.name;
        vmFrame.clipsContent = false;
        vmFrame.x = adjustedTotalLabelWidth;
        vmFrame.y = adjustedTotalLabelHeight;
        vmFrame.resize(saExpandedW, saExpandedH);

        // Set background fill — use variable if found, otherwise transparent
        if (vmBgVariable) {
          vmFrame.fills = [figma.variables.setBoundVariableForPaint({ type: 'SOLID', color: { r: 1, g: 1, b: 1 } }, 'color', vmBgVariable)];
        } else {
          vmFrame.fills = [];
        }

        try {
          vmFrame.setExplicitVariableModeForCollection(vmCollection, vmMode.modeId);
          console.log('[VarModes] Mode "' + vmMode.name + '" set. explicitVariableModes:', JSON.stringify(vmFrame.explicitVariableModes));
        } catch (vmErr) {
          console.error('[VarModes] Failed to set mode "' + vmMode.name + '":', vmErr.message);
        }

        // Create instances inside the mode frame (positions relative to grid origin)
        for (var vmgi = 0; vmgi < grid.length; vmgi++) {
          var vmInst = grid[vmgi].node.createInstance();
          var vmPos = saPositions[grid[vmgi].node.id];
          vmInst.x = vmPos.x;
          vmInst.y = vmPos.y;
          vmFrame.appendChild(vmInst);
        }
        vmModeWrapper.appendChild(vmFrame);

        // Dashed border (on top of the mode frame)
        var vmBorder = figma.createFrame();
        vmBorder.name = 'Grid Border';
        vmBorder.fills = [];
        vmBorder.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
        vmBorder.dashPattern = [4, 4];
        vmBorder.x = adjustedTotalLabelWidth;
        vmBorder.y = adjustedTotalLabelHeight;
        vmBorder.resize(saExpandedW, saExpandedH);
        vmModeWrapper.appendChild(vmBorder);

        // Grid lines
        if (options.showGrid) {
          createGridLines(vmModeWrapper, colClusters, rowClusters, colWidths, rowHeights,
            adjustedTotalLabelWidth, adjustedTotalLabelHeight, saExpandedW, saExpandedH,
            singleTypeGrid ? grid : null);
        }

        // Repeat row labels for this mode
        for (var vmrpi = 0; vmrpi < rowAxisProps.length; vmrpi++) {
          var vmRProp = rowAxisProps[vmrpi];
          var vmRColW = rowLabelWidths[vmrpi];
          var vmRIsInner = rowAxisProps.length > 1 && vmrpi === rowAxisProps.length - 1;
          var vmRHasBracket = !vmRIsInner && rowAxisProps.length > 1;
          var vmRGroups = getRowGroups(vmRProp.name, grid, gridByRow, rowClusters, rowHeights);
          var vmRLXO = 0;
          for (var vmrpj = 0; vmrpj < vmrpi; vmrpj++) vmRLXO += rowLabelWidths[vmrpj];
          for (var vmrvi = 0; vmrvi < vmRGroups.length; vmrvi++) {
            var vmRG = vmRGroups[vmrvi];
            vmModeWrapper.appendChild(createRowLabel(
              formatLabel(vmRProp.name, vmRG.value),
              vmRLXO, adjustedTotalLabelHeight + vmRG.startY,
              vmRColW, vmRG.endY - vmRG.startY, vmRHasBracket
            ));
          }
        }

        // Repeat column labels for this mode
        var vmLYO = boolLabelHeight + VM_LABEL_HEIGHT;
        for (var vmcpi = 0; vmcpi < colAxisProps.length; vmcpi++) {
          var vmCProp = colAxisProps[vmcpi];
          var vmCRowH = colLabelHeights[vmcpi];
          var vmCIsInner = colAxisProps.length > 1 && vmcpi === colAxisProps.length - 1;
          var vmCHasBracket = !vmCIsInner && colAxisProps.length > 1;
          var vmCGroups = getColGroups(vmCProp.name, grid, gridByCol, colClusters, colWidths);
          for (var vmcvi = 0; vmcvi < vmCGroups.length; vmcvi++) {
            var vmCG = vmCGroups[vmcvi];
            vmModeWrapper.appendChild(createColLabel(
              formatLabel(vmCProp.name, vmCG.value),
              adjustedTotalLabelWidth + vmCG.startX, vmLYO,
              vmCG.endX - vmCG.startX, vmCRowH, vmCHasBracket
            ));
          }
          vmLYO += vmCRowH;
        }

        // Per-item labels for single-type grid (variable modes)
        if (singleTypeGrid && stGridInfo) {
          createSingleTypeGridLabels(vmModeWrapper, grid, singleTypeProp, colClusters, colWidths,
            rowClusters, stGridInfo.maxVarH, adjustedTotalLabelWidth, adjustedTotalLabelHeight);
        }

        // Size the per-mode wrapper to fit its content
        var vmModeW = adjustedTotalLabelWidth + saExpandedW;
        var vmModeH = adjustedTotalLabelHeight + saExpandedH;
        vmModeWrapper.resize(vmModeW, vmModeH);

        // Position inside the Modes container
        vmModeWrapper.x = vmCumOffsetX;
        vmModeWrapper.y = 0;
        vmModesContainer.appendChild(vmModeWrapper);

        vmCumOffsetX += vmModeW + VM_GAP;
      }

      // Size the Modes container
      vmModesContainer.resize(vmCumOffsetX > 0 ? vmCumOffsetX - VM_GAP : 1, vmModesContainer.children.length > 0 ? vmModesContainer.children[0].height : 1);
    }

    // ─── Wrap in Documentation frame ───

    var saDocsFrame = figma.createFrame();
    saDocsFrame.name = 'Documentation';
    saDocsFrame.fills = [];
    saDocsFrame.clipsContent = false;
    saDocsFrame.locked = false;

    // Collect wrapper children into Documentation
    var saDocsChildren = [];
    for (var sadi = 0; sadi < saWrapper.children.length; sadi++) {
      if (_autoLayoutExtras && saWrapper.children[sadi].name === 'Property Combinations') continue;
      saDocsChildren.push(saWrapper.children[sadi]);
    }

    // Size Documentation frame
    var saDocsW = adjustedTotalLabelWidth + saExpandedW;
    var saDocsH = adjustedTotalLabelHeight + saExpandedH;
    if (!_autoLayoutExtras && saWrapper.height > saDocsH) saDocsH = saWrapper.height;
    saDocsFrame.resize(saDocsW, saDocsH);

    saWrapper.appendChild(saDocsFrame);
    for (var sadi2 = 0; sadi2 < saDocsChildren.length; sadi2++) {
      saDocsChildren[sadi2].constraints = { horizontal: 'MIN', vertical: 'MIN' };
      saDocsFrame.appendChild(saDocsChildren[sadi2]);
    }

    // When auto layout, ensure Property Combinations is ordered after docsFrame
    if (_autoLayoutExtras) {
      var saExtrasInWrapper = null;
      for (var saewi = 0; saewi < saWrapper.children.length; saewi++) {
        if (saWrapper.children[saewi].name === 'Property Combinations') {
          saExtrasInWrapper = saWrapper.children[saewi];
          break;
        }
      }
      if (saExtrasInWrapper) {
        saWrapper.appendChild(saExtrasInWrapper);
      }
    }

    // ─── Make wrapper auto-layout vertical, hug height ───

    var saHasVarModes = vmModesContainer !== null;
    if (_autoLayoutExtras || saHasVarModes) {
      saWrapper.layoutMode = 'VERTICAL';
      saWrapper.primaryAxisSizingMode = 'AUTO';
      saWrapper.counterAxisSizingMode = 'AUTO';
      saWrapper.itemSpacing = _autoLayoutExtras ? 24 : 48;
      saWrapper.paddingTop = 0;
      saWrapper.paddingBottom = 0;
      saWrapper.paddingLeft = 0;
      saWrapper.paddingRight = 0;
    }

    // Add Modes container if variable modes were generated
    if (vmModesContainer) {
      saWrapper.appendChild(vmModesContainer);
    }

    // ─── Cleanup ───

    disposeMeasureNode();

    var saTotalMs = Date.now() - t0;
    var saTotalSec = (saTotalMs / 1000).toFixed(2);
    console.log('[Perf] === TOTAL standalone generation time: ' + saTotalMs + 'ms (' + saTotalSec + 's) ===');

    figma.currentPage.selection = [saWrapper];
    figma.viewport.scrollAndZoomIntoView([saWrapper]);

    var saDoneMsg = 'Generated standalone docs in ' + saTotalSec + 's.';
    figma.notify(saDoneMsg);
    figma.ui.postMessage({ type: 'done', message: saDoneMsg });
    return;
  }

  // ─── Ensure columns form a clean uniform grid when labels are wider than variants ───
  // When variants are small (e.g. 16px radio buttons), labels like "State: Error Focus"
  // are much wider. In that case, reposition variants into a uniform cell grid so that:
  // 1. Each cell is wide enough for the widest label
  // 2. Variants are centered in their cells (including edge cells)
  // 3. colClusters/colWidths represent cell extents (for label & grid line positioning)

  if (colAxisProps.length > 0 && colClusters.length > 1) {
    const innerColProp = colAxisProps[colAxisProps.length - 1];

    // Measure text width for each column's innermost label
    const colLabelTextWidths = [];
    for (let c = 0; c < colClusters.length; c++) {
      const inCol = gridByCol[c] || [];
      if (inCol.length === 0) {
        colLabelTextWidths.push(0);
        continue;
      }
      const value = inCol[0].props[innerColProp.name];
      const displayText = formatLabel(innerColProp.name, value);
      const textWidth = measureTextWidth(displayText, LABEL_FONT_SIZE);
      colLabelTextWidths.push(textWidth + LABEL_GAP * 2);
    }

    const maxLabelWidth = Math.max(...colLabelTextWidths);
    const maxVariantWidth = Math.max(...colWidths);

    // Only reposition when labels are wider than the widest variant
    if (maxLabelWidth > maxVariantWidth) {
      const originalColClusters = colClusters.slice();
      const originalColWidths = colWidths.slice();

      // Compute original pitch (center-to-center distance, averaged)
      const firstCenter = originalColClusters[0] + originalColWidths[0] / 2;
      const lastCenter = originalColClusters[colClusters.length - 1] + originalColWidths[colClusters.length - 1] / 2;
      const originalPitch = (lastCenter - firstCenter) / (colClusters.length - 1);

      // Uniform cell width = max of original pitch and widest label
      const cellWidth = Math.max(originalPitch, maxLabelWidth);

      // Reposition all variants to be centered in uniform cells
      for (let c = 0; c < colClusters.length; c++) {
        const idealCenter = cellWidth / 2 + c * cellWidth;
        const variantHalfWidth = originalColWidths[c] / 2;
        const idealLeft = idealCenter - variantHalfWidth;

        const inCol = gridByCol[c] || [];
        for (const item of inCol) {
          item.node.x = idealLeft + (item.node.x - originalColClusters[c]);
        }
      }

      // Update colClusters and colWidths to represent cell extents
      for (let c = 0; c < colClusters.length; c++) {
        colClusters[c] = c * cellWidth;
        colWidths[c] = cellWidth;
      }

      // Resize component set to fit uniform grid
      cs.resize(cellWidth * colClusters.length, cs.height);
    }
  }

  // ─── Same for rows: uniform grid when labels are taller than variants ───

  if (rowAxisProps.length > 0 && rowClusters.length > 1) {
    const innerRowProp = rowAxisProps[rowAxisProps.length - 1];

    // Row labels need enough height for text
    const minRowLabelHeight = LABEL_FONT_SIZE + LABEL_GAP * 2;
    const maxVariantHeight = Math.max(...rowHeights);

    if (minRowLabelHeight > maxVariantHeight) {
      const originalRowClusters = rowClusters.slice();
      const originalRowHeights = rowHeights.slice();

      const firstCenter = originalRowClusters[0] + originalRowHeights[0] / 2;
      const lastCenter = originalRowClusters[rowClusters.length - 1] + originalRowHeights[rowClusters.length - 1] / 2;
      const originalPitch = (lastCenter - firstCenter) / (rowClusters.length - 1);

      const cellHeight = Math.max(originalPitch, minRowLabelHeight);

      for (let r = 0; r < rowClusters.length; r++) {
        const idealCenter = cellHeight / 2 + r * cellHeight;
        const variantHalfHeight = originalRowHeights[r] / 2;
        const idealTop = idealCenter - variantHalfHeight;

        const inRow = gridByRow[r] || [];
        for (const item of inRow) {
          item.node.y = idealTop + (item.node.y - originalRowClusters[r]);
        }
      }

      for (let r = 0; r < rowClusters.length; r++) {
        rowClusters[r] = r * cellHeight;
        rowHeights[r] = cellHeight;
      }

      cs.resize(cs.width, cellHeight * rowClusters.length);
    }
  }

  // ─── Single-type grid (regular mode): build a uniform grid with per-item label space ───

  var rgStGridInfo = null;
  if (singleTypeGrid) {
    rgStGridInfo = buildSingleTypeUniformGrid(
      grid, gridByCol, colClusters, colWidths, rowClusters, rowHeights, singleTypeProp,
      function(g, x, y) { g.node.x = x; g.node.y = y; }
    );

    // Resize CS to fit uniform grid
    var stNewCSW = rgStGridInfo.pad + colClusters.length * rgStGridInfo.cellW + (colClusters.length - 1) * rgStGridInfo.pad + rgStGridInfo.pad;
    var stLastRowBottom = rowClusters[rowClusters.length - 1] + rgStGridInfo.cellH;
    var stNewCSH = stLastRowBottom + rgStGridInfo.pad;
    cs.resize(stNewCSW, stNewCSH);
  }

  console.log('[Perf] Dimensions + uniform grid:', (Date.now() - tDims) + 'ms');

  figma.ui.postMessage({ type: 'status', message: 'Creating wrapper...' });

  // ─── 6a: Create or reuse wrapper frame ───

  let wrapper = findExistingWrapper(cs);
  const csOriginalParent = cs.parent;
  const csOriginalX = cs.x;
  const csOriginalY = cs.y;

  if (wrapper) {
    removeOldLabels(wrapper);
  } else {
    wrapper = figma.createFrame();
    wrapper.name = '❖ ' + cs.name + (options.docLabel ? ' — ' + options.docLabel : ' — Obra Autodocs');
    wrapper.setPluginData('sourceComponentSetId', cs.id);
    wrapper.fills = [];
    wrapper.clipsContent = false;

    // Insert wrapper at the component set's position in its parent
    const csIndex = csOriginalParent.children.indexOf(cs);
    csOriginalParent.insertChild(csIndex, wrapper);
    wrapper.x = csOriginalX;
    wrapper.y = csOriginalY;

    // Move component set into wrapper
    wrapper.appendChild(cs);
  }

  // Position component set within wrapper (offset by label area)
  // Pin CS to top-left so wrapper.resize() doesn't shift it via constraints
  cs.constraints = { horizontal: 'MIN', vertical: 'MIN' };

  cs.x = adjustedTotalLabelWidth;
  cs.y = adjustedTotalLabelHeight;

  // Set component set stroke to match doc color with dashes
  cs.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
  cs.dashPattern = [4, 4];

  // Save and clear CS fills so grid lines are visible through the component set
  cs.setPluginData('originalFills', JSON.stringify(cs.fills));
  cs.fills = [];

  // ─── Header: title + description (above all labels and grid) ───

  var headerY = 0;
  if (options.showTitle && titleOffset > 0) {
    var titleText = createTitleText(cs.name);
    titleText.x = adjustedTotalLabelWidth;
    titleText.y = headerY;
    wrapper.appendChild(titleText);
    headerY += titleOffset;
  }
  if (headerDescFrame) {
    headerDescFrame.x = adjustedTotalLabelWidth;
    headerDescFrame.y = headerY;
    wrapper.appendChild(headerDescFrame);
    headerY += descOffset;
  }
  if (headerDocLinkText) {
    headerDocLinkText.x = adjustedTotalLabelWidth;
    headerDocLinkText.y = headerY;
    wrapper.appendChild(headerDocLinkText);
  }

  // ─── Boolean grid: create instances, measure actual sizes, then position everything ───

  var expandedGridWidth = cs.width;
  var expandedGridHeight = cs.height;
  var boolGroupHeights = [];  // actual height of each non-default group
  var boolGroupWidths = [];   // actual width of each non-default group
  var boolGroupData = [];     // { instances, width, height } for each group

  if (booleanGridInfo) {
    // Compute CS padding from actual variant extents — instances that grow
    // beyond original variant sizes need the same padding maintained
    var csVariants = cs.children.filter(function(c) { return c.type === 'COMPONENT' && c.visible !== false; });
    var csMaxExtentY = 0;
    var csMaxExtentX = 0;
    for (var cvi = 0; cvi < csVariants.length; cvi++) {
      var cvExtY = csVariants[cvi].y + csVariants[cvi].height;
      var cvExtX = csVariants[cvi].x + csVariants[cvi].width;
      if (cvExtY > csMaxExtentY) csMaxExtentY = cvExtY;
      if (cvExtX > csMaxExtentX) csMaxExtentX = cvExtX;
    }
    var csBottomPadding = cs.height - csMaxExtentY;
    var csRightPadding = cs.width - csMaxExtentX;

    // Compute CS layout metrics for adaptive grid
    var csPaddingTop = rowClusters.length > 0 ? rowClusters[0] : 0;
    var csPaddingLeft = colClusters.length > 0 ? colClusters[0] : 0;

    // ─── Create real instances and compute group dimensions ───
    // Auto-disable improved accuracy when combinations exceed threshold (protects against slow regeneration)
    var boolTotalCombinations = booleanGridInfo.groups.length * csVariants.length;
    var useBoolAccuracy = options.booleanImprovedAccuracy && boolTotalCombinations <= BOOL_ACCURACY_MAX_COMBINATIONS;
    if (options.booleanImprovedAccuracy && !useBoolAccuracy) {
      console.log('[Boolean Grid] Improved accuracy auto-disabled: ' + boolTotalCombinations + ' combinations exceeds ' + BOOL_ACCURACY_MAX_COMBINATIONS + ' limit');
    }
    var tBoolCreate = Date.now();

    // Pass 1: Create instances and measure actual sizes. Do NOT reposition yet —
    // we need all groups measured before we can compute the global (uniform) adj.
    for (var bgi = 0; bgi < booleanGridInfo.groups.length; bgi++) {
      var bGroup = booleanGridInfo.groups[bgi];
      var result = createBooleanInstanceGroup(cs, bGroup.props, 0, 0);

      var adjRowData = null;
      var adjColData = null;

      // Direct measurement: compute per-group adjusted layout from real instance sizes
      if (useBoolAccuracy) {
        if (rowClusters.length > 0) {
          var dAdjRC = [];
          var dAdjRH = [];
          var dCurY = csPaddingTop;
          for (var dari = 0; dari < rowClusters.length; dari++) {
            var dMaxH = rowHeights[dari];
            for (var daii = 0; daii < result.instances.length; daii++) {
              var _diy = result.instances[daii].y;
              if (_diy >= rowClusters[dari] - CLUSTER_TOLERANCE && _diy <= rowClusters[dari] + rowHeights[dari] + CLUSTER_TOLERANCE) {
                if (result.instances[daii].height > dMaxH) dMaxH = result.instances[daii].height;
              }
            }
            dAdjRC.push(dCurY);
            dAdjRH.push(dMaxH);
            if (dari < rowClusters.length - 1) {
              dCurY += dMaxH + (rowClusters[dari + 1] - (rowClusters[dari] + rowHeights[dari]));
            }
          }
          adjRowData = { clusters: dAdjRC, heights: dAdjRH };
        }
        if (colClusters.length > 0) {
          var dAdjCC = [];
          var dAdjCW = [];
          var dCurX = csPaddingLeft;
          for (var daci = 0; daci < colClusters.length; daci++) {
            var dMaxW = colWidths[daci];
            for (var daii2 = 0; daii2 < result.instances.length; daii2++) {
              var _dix = result.instances[daii2].x;
              if (_dix >= colClusters[daci] - CLUSTER_TOLERANCE && _dix <= colClusters[daci] + colWidths[daci] + CLUSTER_TOLERANCE) {
                if (result.instances[daii2].width > dMaxW) dMaxW = result.instances[daii2].width;
              }
            }
            dAdjCC.push(dCurX);
            dAdjCW.push(dMaxW);
            if (daci < colClusters.length - 1) {
              dCurX += dMaxW + (colClusters[daci + 1] - (colClusters[daci] + colWidths[daci]));
            }
          }
          adjColData = { clusters: dAdjCC, widths: dAdjCW };
        }
      }

      // Store measurement results only (no repositioning yet)
      boolGroupData.push({ instances: result.instances, width: result.width, height: result.height, adjRowData: adjRowData, adjColData: adjColData });
    }

    // Compute global adj: take the max row height / col width across all boolean groups.
    // For column-axis booleans, all groups and the base CS are side-by-side and must share
    // the same row positions. For row-axis booleans they share column positions.
    var globalAdjRowData = null;
    var globalAdjColData = null;
    if (useBoolAccuracy) {
      if (booleanGridInfo.axis === 'col' && rowClusters.length > 0) {
        var gRH = rowHeights.slice();
        for (var gri = 0; gri < boolGroupData.length; gri++) {
          if (boolGroupData[gri].adjRowData) {
            for (var grii = 0; grii < gRH.length; grii++) {
              if (boolGroupData[gri].adjRowData.heights[grii] > gRH[grii]) gRH[grii] = boolGroupData[gri].adjRowData.heights[grii];
            }
          }
        }
        var rowsGrew = gRH.some(function(h, i) { return h > rowHeights[i]; });
        if (rowsGrew) {
          var gRC = [];
          var gCurY = csPaddingTop;
          for (var gri2 = 0; gri2 < rowClusters.length; gri2++) {
            gRC.push(gCurY);
            if (gri2 < rowClusters.length - 1) {
              gCurY += gRH[gri2] + (rowClusters[gri2 + 1] - (rowClusters[gri2] + rowHeights[gri2]));
            }
          }
          globalAdjRowData = { clusters: gRC, heights: gRH };
        }
      }
      if (booleanGridInfo.axis === 'row' && colClusters.length > 0) {
        var gCW = colWidths.slice();
        for (var gci = 0; gci < boolGroupData.length; gci++) {
          if (boolGroupData[gci].adjColData) {
            for (var gcii = 0; gcii < gCW.length; gcii++) {
              if (boolGroupData[gci].adjColData.widths[gcii] > gCW[gcii]) gCW[gcii] = boolGroupData[gci].adjColData.widths[gcii];
            }
          }
        }
        var colsGrew = gCW.some(function(w, i) { return w > colWidths[i]; });
        if (colsGrew) {
          var gCC = [];
          var gCurX = csPaddingLeft;
          for (var gcii2 = 0; gcii2 < colClusters.length; gcii2++) {
            gCC.push(gCurX);
            if (gcii2 < colClusters.length - 1) {
              gCurX += gCW[gcii2] + (colClusters[gcii2 + 1] - (colClusters[gcii2] + colWidths[gcii2]));
            }
          }
          globalAdjColData = { clusters: gCC, widths: gCW };
        }
      }
    }

    // Pass 2: Apply the global (uniform) adj to every boolean group, then reposition their
    // instances and compute each group's final dimensions.
    for (var bgi2a = 0; bgi2a < boolGroupData.length; bgi2a++) {
      var bgData = boolGroupData[bgi2a];
      var effAdjRow = globalAdjRowData || bgData.adjRowData;
      var effAdjCol = globalAdjColData || bgData.adjColData;

      // Reposition instances using the effective (globally-uniform) adj data
      if (effAdjRow || effAdjCol) {
        for (var rii = 0; rii < bgData.instances.length; rii++) {
          var inst = bgData.instances[rii];
          if (effAdjRow) {
            for (var rmi = 0; rmi < rowClusters.length; rmi++) {
              if (inst.y >= rowClusters[rmi] - CLUSTER_TOLERANCE && inst.y <= rowClusters[rmi] + rowHeights[rmi] + CLUSTER_TOLERANCE) {
                inst.y = effAdjRow.clusters[rmi];
                break;
              }
            }
          }
          if (effAdjCol) {
            for (var cmi = 0; cmi < colClusters.length; cmi++) {
              if (inst.x >= colClusters[cmi] - CLUSTER_TOLERANCE && inst.x <= colClusters[cmi] + colWidths[cmi] + CLUSTER_TOLERANCE) {
                inst.x = effAdjCol.clusters[cmi];
                break;
              }
            }
          }
        }
      }

      // Store effective adj so downstream label/grid-line code uses the uniform values
      bgData.adjRowData = effAdjRow;
      bgData.adjColData = effAdjCol;

      // Compute group dimensions using the effective adj data
      var groupHeight, groupWidth;
      if (effAdjRow) {
        var lastRowBottom = effAdjRow.clusters[effAdjRow.clusters.length - 1] + effAdjRow.heights[effAdjRow.heights.length - 1];
        groupHeight = Math.max(cs.height, lastRowBottom + csBottomPadding);
      } else {
        groupHeight = Math.max(cs.height, bgData.height + csBottomPadding);
      }
      if (effAdjCol) {
        var lastColRight = effAdjCol.clusters[effAdjCol.clusters.length - 1] + effAdjCol.widths[effAdjCol.widths.length - 1];
        groupWidth = Math.max(cs.width, lastColRight + csRightPadding);
      } else {
        groupWidth = Math.max(cs.width, bgData.width + csRightPadding);
      }

      boolGroupHeights.push(groupHeight);
      boolGroupWidths.push(groupWidth);
    }

    // Apply global adj to the base CS so all groups share a uniform grid.
    // For column-axis booleans: expand row heights → reposition CS variants vertically and resize CS.
    // For row-axis booleans: expand col widths → reposition CS variants horizontally and resize CS.
    if (globalAdjRowData) {
      for (var grv = 0; grv < csVariants.length; grv++) {
        var grvY = csVariants[grv].y;
        for (var grvr = 0; grvr < rowClusters.length; grvr++) {
          if (grvY >= rowClusters[grvr] - CLUSTER_TOLERANCE && grvY <= rowClusters[grvr] + rowHeights[grvr] + CLUSTER_TOLERANCE) {
            csVariants[grv].y = globalAdjRowData.clusters[grvr];
            break;
          }
        }
      }
      var newCSHeight = globalAdjRowData.clusters[globalAdjRowData.clusters.length - 1] + globalAdjRowData.heights[globalAdjRowData.heights.length - 1] + csBottomPadding;
      cs.resize(cs.width, newCSHeight);
      // Update rowClusters/rowHeights in-place so all downstream code (labels, grid lines) uses the new positions
      for (var gru = 0; gru < rowClusters.length; gru++) {
        rowClusters[gru] = globalAdjRowData.clusters[gru];
        rowHeights[gru] = globalAdjRowData.heights[gru];
      }
      expandedGridHeight = newCSHeight;
    }
    if (globalAdjColData) {
      for (var gcv = 0; gcv < csVariants.length; gcv++) {
        var gcvX = csVariants[gcv].x;
        for (var gcvc = 0; gcvc < colClusters.length; gcvc++) {
          if (gcvX >= colClusters[gcvc] - CLUSTER_TOLERANCE && gcvX <= colClusters[gcvc] + colWidths[gcvc] + CLUSTER_TOLERANCE) {
            csVariants[gcv].x = globalAdjColData.clusters[gcvc];
            break;
          }
        }
      }
      var newCSWidth = globalAdjColData.clusters[globalAdjColData.clusters.length - 1] + globalAdjColData.widths[globalAdjColData.widths.length - 1] + csRightPadding;
      cs.resize(newCSWidth, cs.height);
      // Update colClusters/colWidths in-place so all downstream code uses the new positions
      for (var gcu = 0; gcu < colClusters.length; gcu++) {
        colClusters[gcu] = globalAdjColData.clusters[gcu];
        colWidths[gcu] = globalAdjColData.widths[gcu];
      }
      expandedGridWidth = newCSWidth;
    }

    console.log('⏱ Boolean grid (instances + accuracy):', (Date.now() - tBoolCreate) + 'ms');

    // Calculate expanded dimensions using actual group sizes
    if (booleanGridInfo.axis === 'row') {
      var totalGroupHeight = 0;
      for (var bhi = 0; bhi < boolGroupHeights.length; bhi++) {
        totalGroupHeight += boolGroupHeights[bhi] + BOOL_GROUP_GAP;
      }
      expandedGridHeight = cs.height + totalGroupHeight;
    } else {
      var totalGroupWidth = 0;
      for (var bwi = 0; bwi < boolGroupWidths.length; bwi++) {
        totalGroupWidth += boolGroupWidths[bwi] + BOOL_GROUP_GAP;
      }
      expandedGridWidth = cs.width + totalGroupWidth;
    }
  }

  // Resize wrapper to fit labels + expanded grid
  wrapper.resize(
    adjustedTotalLabelWidth + expandedGridWidth,
    adjustedTotalLabelHeight + expandedGridHeight
  );

  figma.ui.postMessage({ type: 'status', message: 'Creating labels...' });
  var tLabels = Date.now();

  // ─── Boolean grid: position instances, tints, and labels ───

  if (booleanGridInfo) {
    // Calculate cumulative Y/X offsets for each group
    var bCumulativeOffset = 0;

    for (var bgi2 = 0; bgi2 < booleanGridInfo.groups.length; bgi2++) {
      var bGroupIndex = bgi2 + 1;
      var bGroupOffsetX, bGroupOffsetY;
      var bGroupH = boolGroupHeights[bgi2];
      var bGroupW = boolGroupWidths[bgi2];

      if (booleanGridInfo.axis === 'row') {
        bGroupOffsetX = adjustedTotalLabelWidth;
        bGroupOffsetY = adjustedTotalLabelHeight + cs.height + bCumulativeOffset + BOOL_GROUP_GAP;
        bCumulativeOffset += bGroupH + BOOL_GROUP_GAP;
      } else {
        bGroupOffsetX = adjustedTotalLabelWidth + cs.width + bCumulativeOffset + BOOL_GROUP_GAP;
        bGroupOffsetY = adjustedTotalLabelHeight;
        bCumulativeOffset += bGroupW + BOOL_GROUP_GAP;
      }

      // Store offset for label/grid use
      boolGroupData[bgi2].offsetX = bGroupOffsetX;
      boolGroupData[bgi2].offsetY = bGroupOffsetY;

      // Background tint — covers exactly the group content area
      var tintW = booleanGridInfo.axis === 'row' ? cs.width : bGroupW;
      var tintH = booleanGridInfo.axis === 'row' ? bGroupH : cs.height;
      var tint = createBackgroundTint(bGroupOffsetX, bGroupOffsetY, tintW, tintH);
      wrapper.appendChild(tint);
      boolGroupData[bgi2].tint = tint;

      // Reposition instances from (0,0) to actual offset
      var bInstances = boolGroupData[bgi2].instances;
      for (var bii = 0; bii < bInstances.length; bii++) {
        bInstances[bii].x += bGroupOffsetX;
        bInstances[bii].y += bGroupOffsetY;
        wrapper.appendChild(bInstances[bii]);
      }
    }

    // Boolean axis labels
    if (booleanGridInfo.axis === 'row') {
      var hasBoolBracket = rowAxisProps.length > 0;
      // Default group label
      var bDefaultLabel = createRowLabel(
        booleanGridInfo.defaultLabel, 0, adjustedTotalLabelHeight,
        boolLabelWidth, cs.height, hasBoolBracket
      );
      wrapper.appendChild(bDefaultLabel);
      // Non-default group labels
      for (var bgli = 0; bgli < booleanGridInfo.groups.length; bgli++) {
        var bLabel = createRowLabel(
          booleanGridInfo.groups[bgli].label, 0, boolGroupData[bgli].offsetY,
          boolLabelWidth, boolGroupHeights[bgli], hasBoolBracket
        );
        wrapper.appendChild(bLabel);
      }
    } else {
      var hasBoolBracket = colAxisProps.length > 0;
      // Default group label
      var bDefaultLabel = createColLabel(
        booleanGridInfo.defaultLabel, adjustedTotalLabelWidth, headerOffset,
        cs.width, boolLabelHeight, hasBoolBracket
      );
      wrapper.appendChild(bDefaultLabel);
      // Non-default group labels
      for (var bgli = 0; bgli < booleanGridInfo.groups.length; bgli++) {
        var bLabel = createColLabel(
          booleanGridInfo.groups[bgli].label, boolGroupData[bgli].offsetX, headerOffset,
          boolGroupWidths[bgli], boolLabelHeight, hasBoolBracket
        );
        wrapper.appendChild(bLabel);
      }
    }

    // ─── Boolean Grid Border: unified stroke around entire grid area ───
    // Save CS strokes, create a border frame with the same style, remove CS strokes
    var csStrokeData = {
      strokes: JSON.parse(JSON.stringify(cs.strokes)),
      strokeWeight: cs.strokeWeight,
      dashPattern: cs.dashPattern.slice(),
      strokeAlign: cs.strokeAlign,
      cornerRadius: cs.cornerRadius
    };
    cs.setPluginData('originalStrokes', JSON.stringify(csStrokeData));

    var borderFrame = figma.createFrame();
    borderFrame.name = 'Boolean Grid Border';
    borderFrame.fills = [];
    borderFrame.clipsContent = false;
    borderFrame.strokes = csStrokeData.strokes;
    borderFrame.strokeWeight = csStrokeData.strokeWeight;
    borderFrame.dashPattern = csStrokeData.dashPattern;
    borderFrame.strokeAlign = csStrokeData.strokeAlign;
    borderFrame.cornerRadius = csStrokeData.cornerRadius;
    borderFrame.locked = true;
    borderFrame.x = cs.x;
    borderFrame.y = cs.y;
    if (booleanGridInfo.axis === 'row') {
      borderFrame.resize(cs.width, expandedGridHeight);
    } else {
      borderFrame.resize(expandedGridWidth, cs.height);
    }
    wrapper.appendChild(borderFrame);

    // Remove strokes from CS so the border frame is the only visible border
    cs.strokes = [];
  }

  // ─── 6b: Create row labels (left side) ───
  // Row labels are positioned relative to the wrapper.
  // Their y-coordinates need to account for the adjusted label height offset (column label area above).

  let labelXOffset = boolLabelWidth;
  for (let pi = 0; pi < rowAxisProps.length; pi++) {
    const prop = rowAxisProps[pi];
    const colWidth = rowLabelWidths[pi];
    const isInner = rowAxisProps.length > 1 && pi === rowAxisProps.length - 1;
    const hasBracket = !isInner && rowAxisProps.length > 1;

    // Group rows by this property's value (using original grid positions)
    const valueGroups = getRowGroups(prop.name, grid, gridByRow, rowClusters, rowHeights);

    for (const group of valueGroups) {
      // Offset y by totalLabelHeight to sit below column labels
      const groupY = adjustedTotalLabelHeight + group.startY;
      const groupHeight = group.endY - group.startY;

      const displayValue = formatLabel(prop.name, group.value);

      const label = createRowLabel(
        displayValue,
        labelXOffset,
        groupY,
        colWidth,
        groupHeight,
        hasBracket
      );
      wrapper.appendChild(label);
    }

    labelXOffset += colWidth;
  }

  // ─── 6c: Create column labels (top) ───
  // Column labels are positioned relative to the wrapper.
  // Their x-coordinates need to account for the totalLabelWidth offset (row label area to the left).

  let labelYOffset = boolLabelHeight + headerOffset;
  for (let pi = 0; pi < colAxisProps.length; pi++) {
    const prop = colAxisProps[pi];
    const rowHeight = colLabelHeights[pi];
    const isInner = colAxisProps.length > 1 && pi === colAxisProps.length - 1;
    const hasBracket = !isInner && colAxisProps.length > 1;

    // Group columns by this property's value (using original grid positions)
    const valueGroups = getColGroups(prop.name, grid, gridByCol, colClusters, colWidths);

    for (const group of valueGroups) {
      // Offset x by totalLabelWidth to sit to the right of row labels
      const groupX = adjustedTotalLabelWidth + group.startX;
      const groupWidth = group.endX - group.startX;

      const displayValue = formatLabel(prop.name, group.value);

      const label = createColLabel(
        displayValue,
        groupX,
        labelYOffset,
        groupWidth,
        rowHeight,
        hasBracket
      );
      wrapper.appendChild(label);
    }

    labelYOffset += rowHeight;
  }

  // ─── Per-item labels for single-type grid (regular mode) ───

  if (singleTypeGrid && rgStGridInfo) {
    createSingleTypeGridLabels(wrapper, grid, singleTypeProp, colClusters, colWidths,
      rowClusters, rgStGridInfo.maxVarH, adjustedTotalLabelWidth, adjustedTotalLabelHeight);
  }

  // ─── Repeat inner variant labels for non-default boolean groups ───

  if (booleanGridInfo) {
    for (var brgi = 0; brgi < booleanGridInfo.groups.length; brgi++) {
      var brAdjRC = (boolGroupData[brgi].adjRowData && boolGroupData[brgi].adjRowData.clusters) || rowClusters;
      var brAdjRH = (boolGroupData[brgi].adjRowData && boolGroupData[brgi].adjRowData.heights) || rowHeights;
      var brAdjCC = (boolGroupData[brgi].adjColData && boolGroupData[brgi].adjColData.clusters) || colClusters;
      var brAdjCW = (boolGroupData[brgi].adjColData && boolGroupData[brgi].adjColData.widths) || colWidths;
      if (booleanGridInfo.axis === 'row') {
        // Repeat row labels (left side) for this boolean group
        var brLabelXOffset = boolLabelWidth;
        for (var brpi = 0; brpi < rowAxisProps.length; brpi++) {
          var brProp = rowAxisProps[brpi];
          var brColWidth = rowLabelWidths[brpi];
          var brIsInner = rowAxisProps.length > 1 && brpi === rowAxisProps.length - 1;
          var brHasBracket = !brIsInner && rowAxisProps.length > 1;
          var brValueGroups = getRowGroups(brProp.name, grid, gridByRow, brAdjRC, brAdjRH);
          for (var brvi = 0; brvi < brValueGroups.length; brvi++) {
            var brVGroup = brValueGroups[brvi];
            var brGroupY = boolGroupData[brgi].offsetY + brVGroup.startY;
            var brGroupHeight = brVGroup.endY - brVGroup.startY;
            var brDisplayValue = formatLabel(brProp.name, brVGroup.value);
            var brLabel = createRowLabel(brDisplayValue, brLabelXOffset, brGroupY, brColWidth, brGroupHeight, brHasBracket);
            wrapper.appendChild(brLabel);
          }
          brLabelXOffset += brColWidth;
        }
      } else {
        // Repeat column labels (top) for this boolean group
        var brLabelYOffset = boolLabelHeight + headerOffset;
        for (var brpi = 0; brpi < colAxisProps.length; brpi++) {
          var brProp = colAxisProps[brpi];
          var brRowHeight = colLabelHeights[brpi];
          var brIsInner = colAxisProps.length > 1 && brpi === colAxisProps.length - 1;
          var brHasBracket = !brIsInner && colAxisProps.length > 1;
          var brValueGroups = getColGroups(brProp.name, grid, gridByCol, brAdjCC, brAdjCW);
          for (var brvi = 0; brvi < brValueGroups.length; brvi++) {
            var brVGroup = brValueGroups[brvi];
            var brGroupX = boolGroupData[brgi].offsetX + brVGroup.startX;
            var brGroupWidth = brVGroup.endX - brVGroup.startX;
            var brDisplayValue = formatLabel(brProp.name, brVGroup.value);
            var brLabel = createColLabel(brDisplayValue, brGroupX, brLabelYOffset, brGroupWidth, brRowHeight, brHasBracket);
            wrapper.appendChild(brLabel);
          }
          brLabelYOffset += brRowHeight;
        }
      }
    }
  }

  console.log('[Perf] Label creation:', (Date.now() - tLabels) + 'ms');

  // ─── Grid lines (if selected) ───

  if (options.showGrid) {
    createGridLines(wrapper, colClusters, rowClusters, colWidths, rowHeights, cs.x, cs.y, cs.width, cs.height,
      singleTypeGrid ? grid : null);
    // Additional grid lines for boolean groups (using actual group dimensions)
    if (booleanGridInfo) {
      for (var ggi = 0; ggi < booleanGridInfo.groups.length; ggi++) {
        var gRC = (boolGroupData[ggi].adjRowData && boolGroupData[ggi].adjRowData.clusters) || rowClusters;
        var gRH = (boolGroupData[ggi].adjRowData && boolGroupData[ggi].adjRowData.heights) || rowHeights;
        var gCC = (boolGroupData[ggi].adjColData && boolGroupData[ggi].adjColData.clusters) || colClusters;
        var gCW = (boolGroupData[ggi].adjColData && boolGroupData[ggi].adjColData.widths) || colWidths;
        if (booleanGridInfo.axis === 'row') {
          createGridLines(wrapper, gCC, gRC, gCW, gRH,
            boolGroupData[ggi].offsetX, boolGroupData[ggi].offsetY, cs.width, boolGroupHeights[ggi]);
        } else {
          createGridLines(wrapper, gCC, gRC, gCW, gRH,
            boolGroupData[ggi].offsetX, boolGroupData[ggi].offsetY, boolGroupWidths[ggi], cs.height);
        }
      }
      // Dividers between boolean groups
      // Double-line dividers only for nested boolean groups (multiple bool props);
      // single-line divider for simple single-boolean components
      var useDoubleLine = booleanGridInfo.boolPropCount > 1;
      for (var dgi = 0; dgi < booleanGridInfo.groups.length; dgi++) {
        var divLine1 = figma.createLine();
        divLine1.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
        divLine1.strokeWeight = GRID_STROKE_WEIGHT;
        divLine1.dashPattern = GRID_DASH_PATTERN;
        if (booleanGridInfo.axis === 'row') {
          var prevBottom = dgi === 0 ? (cs.y + cs.height) : (boolGroupData[dgi - 1].offsetY + boolGroupHeights[dgi - 1]);
          divLine1.resize(cs.width, 0);
          divLine1.x = cs.x;
          divLine1.y = prevBottom;
        } else {
          var prevRight = dgi === 0 ? (cs.x + cs.width) : (boolGroupData[dgi - 1].offsetX + boolGroupWidths[dgi - 1]);
          divLine1.rotation = -90;
          divLine1.resize(cs.height, 0);
          divLine1.x = prevRight;
          divLine1.y = cs.y;
        }
        wrapper.appendChild(divLine1);
        if (useDoubleLine) {
          var divLine2 = figma.createLine();
          divLine2.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
          divLine2.strokeWeight = GRID_STROKE_WEIGHT;
          divLine2.dashPattern = GRID_DASH_PATTERN;
          if (booleanGridInfo.axis === 'row') {
            divLine2.resize(cs.width, 0);
            divLine2.x = cs.x;
            divLine2.y = boolGroupData[dgi].offsetY;
          } else {
            divLine2.rotation = -90;
            divLine2.resize(cs.height, 0);
            divLine2.x = boolGroupData[dgi].offsetX;
            divLine2.y = cs.y;
          }
          wrapper.appendChild(divLine2);
        }
      }
    }
  }

  // ─── Extra sections (Boolean visibility + Nested instances) ───

  var extraSections = [];

  if (options.showBooleanVisibility && !booleanGridInfo) {
    figma.ui.postMessage({ type: 'status', message: 'Creating boolean visibility examples...' });
    var boolSection;
    if (options.booleanScope) {
      boolSection = await createScopedBooleanSection(cs, options.booleanScope, options.booleanCombination || 'individual', options.booleanDisplayMode || 'list');
    } else {
      boolSection = await createBooleanVisibilitySection(cs, options.booleanCombination || 'individual', options.enabledBooleanProps, options.booleanDisplayMode || 'list');
    }
    if (boolSection) extraSections.push(boolSection);
  }

  if (options.showNestedInstances) {
    figma.ui.postMessage({ type: 'status', message: 'Creating nested instance examples...' });
    var nestedSection = await createNestedInstancesSection(cs, options.nestedInstancesMode || 'representative', options.enabledNestedInstances);
    if (nestedSection) extraSections.push(nestedSection);
  }

  if (extraSections.length > 0) {
    var EXTRAS_GAP = 24;
    var extrasContainer = figma.createFrame();
    extrasContainer.name = 'Property Combinations';
    extrasContainer.fills = [];
    extrasContainer.clipsContent = false;
    extrasContainer.layoutMode = 'VERTICAL';
    extrasContainer.itemSpacing = EXTRAS_GAP;
    extrasContainer.counterAxisSizingMode = 'AUTO';
    extrasContainer.primaryAxisSizingMode = 'AUTO';

    for (var ei = 0; ei < extraSections.length; ei++) {
      extrasContainer.appendChild(extraSections[ei]);
    }

    if (_autoLayoutExtras) {
      extrasContainer.paddingLeft = adjustedTotalLabelWidth;
    } else {
      extrasContainer.x = adjustedTotalLabelWidth;
      extrasContainer.y = adjustedTotalLabelHeight + expandedGridHeight + EXTRAS_GAP;
    }
    wrapper.appendChild(extrasContainer);

    if (!_autoLayoutExtras) {
      wrapper.resize(
        Math.max(wrapper.width, adjustedTotalLabelWidth + extrasContainer.width),
        extrasContainer.y + extrasContainer.height
      );
    }
    if (DEBUG) { console.log('[Generate] Extras container added (' + extraSections.length + ' sections)'); }
  }

  // Description is now placed at the top (headerDescFrame) alongside the title

  // ─── Step 9: Wrap generated elements in Documentation frame ───

  var docsFrame = figma.createFrame();
  docsFrame.name = 'Documentation';
  docsFrame.fills = [];
  docsFrame.clipsContent = false;
  docsFrame.locked = false;
  docsFrame.constraints = { horizontal: 'MIN', vertical: 'MIN' };
  docsFrame.x = 0;
  docsFrame.y = 0;
  // Size to base grid area when auto layout (extras are a sibling), or full wrapper when manual
  if (_autoLayoutExtras) {
    docsFrame.resize(adjustedTotalLabelWidth + expandedGridWidth, adjustedTotalLabelHeight + expandedGridHeight);
  } else {
    docsFrame.resize(wrapper.width, wrapper.height);
  }
  wrapper.appendChild(docsFrame);

  // Move generated elements into docs frame (keep CS, docsFrame itself, and Property Combinations out when auto layout)
  var docsChildren = [];
  for (var di = 0; di < wrapper.children.length; di++) {
    var dChild = wrapper.children[di];
    if (dChild.type !== 'COMPONENT_SET' && dChild.type !== 'COMPONENT' && dChild !== docsFrame) {
      if (_autoLayoutExtras && dChild.name === 'Property Combinations') continue;
      docsChildren.push(dChild);
    }
  }
  for (var di2 = 0; di2 < docsChildren.length; di2++) {
    docsChildren[di2].constraints = { horizontal: 'MIN', vertical: 'MIN' };
    docsFrame.appendChild(docsChildren[di2]);
  }

  // When auto layout, ensure Property Combinations is ordered after docsFrame
  if (_autoLayoutExtras) {
    var extrasInWrapper = null;
    for (var ewi = 0; ewi < wrapper.children.length; ewi++) {
      if (wrapper.children[ewi].name === 'Property Combinations') {
        extrasInWrapper = wrapper.children[ewi];
        break;
      }
    }
    if (extrasInWrapper) {
      wrapper.appendChild(extrasInWrapper);
    }
  }

  // ─── Step 9b: Variable modes ───

  if (options.variableModes && options.variableModes.groups && options.variableModes.groups.length > 0) {
    figma.ui.postMessage({ type: 'status', message: 'Generating variable modes...' });

    var vmGroups = options.variableModes.groups;
    var VM_GAP = 48;
    var VM_LABEL_HEIGHT = SIMPLE_LABEL_ROW_HEIGHT;

    // Resolve all collections referenced by any group
    var vmCollectionCache = {};
    for (var vmci = 0; vmci < vmGroups.length; vmci++) {
      for (var vmcj = 0; vmcj < vmGroups[vmci].collections.length; vmcj++) {
        var vmCid = vmGroups[vmci].collections[vmcj].collectionId;
        if (!vmCollectionCache[vmCid]) {
          vmCollectionCache[vmCid] = await figma.variables.getVariableCollectionByIdAsync(vmCid);
        }
      }
    }

    // Try to find a background color variable for dark mode support (search all collections)
    var vmBgVariable = null;
    try {
      for (var vmBgCid in vmCollectionCache) {
        if (vmBgVariable) break;
        var vmBgColl = vmCollectionCache[vmBgCid];
        if (!vmBgColl) continue;
        var vmVarIds = vmBgColl.variableIds;
        for (var vmvi = 0; vmvi < vmVarIds.length; vmvi++) {
          var vmVar = await figma.variables.getVariableByIdAsync(vmVarIds[vmvi]);
          if (vmVar && vmVar.resolvedType === 'COLOR') {
            var vmVarName = vmVar.name.toLowerCase();
            if (vmVarName.indexOf('background') !== -1 || vmVarName.indexOf('bg') !== -1 || vmVarName.indexOf('surface') !== -1 || vmVarName.indexOf('body') !== -1) {
              vmBgVariable = vmVar;
              break;
            }
          }
        }
      }
    } catch (vmBgErr) {
      console.warn('[VarModes] Could not search for background variable:', vmBgErr.message);
    }

    var vmModesContainer = figma.createFrame();
    vmModesContainer.name = 'Modes';
    vmModesContainer.fills = [];
    vmModesContainer.clipsContent = false;
    vmModesContainer.layoutMode = 'VERTICAL';
    vmModesContainer.primaryAxisSizingMode = 'AUTO';
    vmModesContainer.counterAxisSizingMode = 'AUTO';
    vmModesContainer.itemSpacing = VM_GAP;

    var vmGridW = expandedGridWidth;
    var vmGridH = cs.height;

    for (var vmmi = 0; vmmi < vmGroups.length; vmmi++) {
      var vmGroup = vmGroups[vmmi];

      // Per-mode wrapper — include collection names for traceability
      var vmModeWrapper = figma.createFrame();
      var vmCollNames = [];
      for (var vmni = 0; vmni < vmGroup.collections.length; vmni++) {
        var vmNC = vmCollectionCache[vmGroup.collections[vmni].collectionId];
        if (vmNC) vmCollNames.push(vmNC.name);
      }
      vmModeWrapper.name = 'Mode: ' + vmGroup.name + (vmCollNames.length > 0 ? ' (' + vmCollNames.join(', ') + ')' : '');
      vmModeWrapper.fills = [];
      vmModeWrapper.clipsContent = false;

      // Mode label (sits at top of mode wrapper)
      vmModeWrapper.appendChild(createColLabel(vmGroup.name, adjustedTotalLabelWidth, boolLabelHeight, vmGridW, VM_LABEL_HEIGHT, false));

      // The grid area starts after mode label + column labels
      var vmGridY = VM_LABEL_HEIGHT + adjustedTotalLabelHeight;

      // Mode frame with variable modes set from all collections in this group
      var vmFrame = figma.createFrame();
      vmFrame.name = vmGroup.name;
      vmFrame.clipsContent = false;
      vmFrame.x = adjustedTotalLabelWidth;
      vmFrame.y = vmGridY;
      vmFrame.resize(vmGridW, vmGridH);

      if (vmBgVariable) {
        vmFrame.fills = [figma.variables.setBoundVariableForPaint({ type: 'SOLID', color: { r: 1, g: 1, b: 1 } }, 'color', vmBgVariable)];
      } else {
        vmFrame.fills = [];
      }

      // Apply all collection modes for this group
      for (var vmgci = 0; vmgci < vmGroup.collections.length; vmgci++) {
        var vmGC = vmGroup.collections[vmgci];
        var vmGColl = vmCollectionCache[vmGC.collectionId];
        if (!vmGColl) continue;
        try {
          vmFrame.setExplicitVariableModeForCollection(vmGColl, vmGC.modeId);
        } catch (vmErr) {
          console.error('[VarModes] Failed to set mode "' + vmGroup.name + '" for collection "' + vmGColl.name + '":', vmErr.message);
        }
      }

      // Create instances at variant positions
      for (var vmgi = 0; vmgi < grid.length; vmgi++) {
        var vmInst = grid[vmgi].node.createInstance();
        vmInst.x = grid[vmgi].node.x;
        vmInst.y = grid[vmgi].node.y;
        vmFrame.appendChild(vmInst);
      }
      vmModeWrapper.appendChild(vmFrame);

      // Dashed border
      var vmBorder = figma.createFrame();
      vmBorder.name = 'Grid Border';
      vmBorder.fills = [];
      vmBorder.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
      vmBorder.dashPattern = [4, 4];
      vmBorder.x = adjustedTotalLabelWidth;
      vmBorder.y = vmGridY;
      vmBorder.resize(vmGridW, vmGridH);
      vmModeWrapper.appendChild(vmBorder);

      // Grid lines
      if (options.showGrid) {
        createGridLines(vmModeWrapper, colClusters, rowClusters, colWidths, rowHeights,
          adjustedTotalLabelWidth, vmGridY, vmGridW, vmGridH,
          singleTypeGrid ? grid : null);
      }

      // Row labels
      for (var vmrpi = 0; vmrpi < rowAxisProps.length; vmrpi++) {
        var vmRProp = rowAxisProps[vmrpi];
        var vmRColW = rowLabelWidths[vmrpi];
        var vmRIsInner = rowAxisProps.length > 1 && vmrpi === rowAxisProps.length - 1;
        var vmRHasBracket = !vmRIsInner && rowAxisProps.length > 1;
        var vmRGroups = getRowGroups(vmRProp.name, grid, gridByRow, rowClusters, rowHeights);
        var vmRLXO = 0;
        for (var vmrpj = 0; vmrpj < vmrpi; vmrpj++) vmRLXO += rowLabelWidths[vmrpj];
        for (var vmrvi = 0; vmrvi < vmRGroups.length; vmrvi++) {
          var vmRG = vmRGroups[vmrvi];
          vmModeWrapper.appendChild(createRowLabel(
            formatLabel(vmRProp.name, vmRG.value),
            vmRLXO, vmGridY + vmRG.startY,
            vmRColW, vmRG.endY - vmRG.startY, vmRHasBracket
          ));
        }
      }

      // Column labels (after mode label, before grid)
      var vmLYO = boolLabelHeight + VM_LABEL_HEIGHT;
      for (var vmcpi = 0; vmcpi < colAxisProps.length; vmcpi++) {
        var vmCProp = colAxisProps[vmcpi];
        var vmCRowH = colLabelHeights[vmcpi];
        var vmCIsInner = colAxisProps.length > 1 && vmcpi === colAxisProps.length - 1;
        var vmCHasBracket = !vmCIsInner && colAxisProps.length > 1;
        var vmCGroups = getColGroups(vmCProp.name, grid, gridByCol, colClusters, colWidths);
        for (var vmcvi = 0; vmcvi < vmCGroups.length; vmcvi++) {
          var vmCG = vmCGroups[vmcvi];
          vmModeWrapper.appendChild(createColLabel(
            formatLabel(vmCProp.name, vmCG.value),
            adjustedTotalLabelWidth + vmCG.startX, vmLYO,
            vmCG.endX - vmCG.startX, vmCRowH, vmCHasBracket
          ));
        }
        vmLYO += vmCRowH;
      }

      // Per-item labels for single-type grid (regular mode variable modes)
      if (singleTypeGrid && rgStGridInfo) {
        createSingleTypeGridLabels(vmModeWrapper, grid, singleTypeProp, colClusters, colWidths,
          rowClusters, rgStGridInfo.maxVarH, adjustedTotalLabelWidth, vmGridY);
      }

      // Size the per-mode wrapper (mode label + col labels + grid)
      var vmModeW = adjustedTotalLabelWidth + vmGridW;
      var vmModeH = vmGridY + vmGridH;
      vmModeWrapper.resize(vmModeW, vmModeH);

      vmModesContainer.appendChild(vmModeWrapper);
    }

    wrapper.appendChild(vmModesContainer);
  }

  // ─── Make wrapper auto-layout vertical (when auto layout extras enabled or variable modes present) ───

  var hasVarModes = options.variableModes && options.variableModes.groups && options.variableModes.groups.length > 0;
  if (_autoLayoutExtras || hasVarModes) {
    wrapper.layoutMode = 'VERTICAL';
    wrapper.primaryAxisSizingMode = 'AUTO';
    wrapper.counterAxisSizingMode = 'AUTO';
    wrapper.itemSpacing = _autoLayoutExtras ? 24 : 48;
    wrapper.paddingTop = 0;
    wrapper.paddingBottom = 0;
    wrapper.paddingLeft = 0;
    wrapper.paddingRight = 0;

    cs.layoutPositioning = 'ABSOLUTE';
    cs.x = adjustedTotalLabelWidth;
    cs.y = adjustedTotalLabelHeight;
  }

  // Ensure CS is above Documentation in the layer panel (last child = top)
  wrapper.appendChild(cs);

  // ─── Step 10: Cleanup ───

  disposeMeasureNode();

  var totalMs = Date.now() - t0;
  var totalSec = (totalMs / 1000).toFixed(2);
  console.log('[Perf] === TOTAL generation time: ' + totalMs + 'ms (' + totalSec + 's) ===');

  figma.currentPage.selection = [wrapper];
  figma.viewport.scrollAndZoomIntoView([wrapper]);

  var doneMsg = 'Generated docs in ' + totalSec + 's.';
  figma.notify(doneMsg);
  figma.ui.postMessage({
    type: 'done',
    message: doneMsg
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// VARIABLE MODES — separate frame with mode columns, placed next to regular docs
// ═══════════════════════════════════════════════════════════════════════════════

async function generateVariableModes(cs, options, grid, gridByRow, gridByCol,
    colClusters, rowClusters, colWidths, rowHeights,
    colAxisProps, rowAxisProps, colLabelHeights, rowLabelWidths,
    boolLabelHeight, adjustedTotalLabelWidth, adjustedTotalLabelHeight, t0) {

  figma.ui.postMessage({ type: 'status', message: 'Generating variable modes...' });

  var vmCollId = options.variableModes.collectionId;
  var vmCollection = await figma.variables.getVariableCollectionByIdAsync(vmCollId);
  var vmModes = options.variableModes.modes;
  var VM_GAP = 48;
  var VM_MODE_GAP = 24; // gap between mode columns

  // For variable modes, we only care about the base grid labels (no boolean label offsets)
  var vmColLabelHeight = colLabelHeights.reduce(function(s, h) { return s + h; }, 0);
  var vmRowLabelWidth = rowLabelWidths.reduce(function(s, w) { return s + w; }, 0);

  // Set doc color
  if (options.color) {
    DOC_COLOR = hexToRgb(options.color);
  } else {
    DOC_COLOR = { r: DEFAULT_COLOR.r, g: DEFAULT_COLOR.g, b: DEFAULT_COLOR.b };
  }

  // Pre-fetch all other collections for cross-collection mode matching
  var vmOtherCollections = [];
  try {
    var vmAllColls = await figma.variables.getLocalVariableCollectionsAsync();
    for (var vmaci = 0; vmaci < vmAllColls.length; vmaci++) {
      if (vmAllColls[vmaci].id !== vmCollId) {
        vmOtherCollections.push(vmAllColls[vmaci]);
      }
    }
  } catch (vmFetchErr) {
    console.warn('[VarModes] Could not fetch collections for mode matching:', vmFetchErr.message);
  }

  // Try to find a background color variable — search primary collection first, then others
  var vmBgVariable = null;
  var vmBgSearchColls = [vmCollection].concat(vmOtherCollections);
  for (var vmBgSi = 0; vmBgSi < vmBgSearchColls.length && !vmBgVariable; vmBgSi++) {
    try {
      var vmBgColl = vmBgSearchColls[vmBgSi];
      var vmVarIds = vmBgColl.variableIds;
      for (var vmvi = 0; vmvi < vmVarIds.length; vmvi++) {
        var vmVar = await figma.variables.getVariableByIdAsync(vmVarIds[vmvi]);
        if (vmVar && vmVar.resolvedType === 'COLOR') {
          var vmVarName = vmVar.name.toLowerCase();
          if (vmVarName.indexOf('background') !== -1 || vmVarName.indexOf('bg') !== -1 || vmVarName.indexOf('surface') !== -1 || vmVarName.indexOf('body') !== -1) {
            vmBgVariable = vmVar;
            console.log('[VarModes] Found background variable:', vmVar.name, 'in "' + vmBgColl.name + '"');
            break;
          }
        }
      }
    } catch (vmBgErr) {
      console.warn('[VarModes] Could not search for background variable:', vmBgErr.message);
    }
  }

  // Find or create the variable modes wrapper
  var vmWrapper = findVariableModesWrapper(cs);
  if (vmWrapper) {
    removeOldLabels(vmWrapper);
  } else {
    vmWrapper = figma.createFrame();
    vmWrapper.name = '❖ ' + cs.name + ' — Variable Modes';
    vmWrapper.setPluginData('variableModesDoc', 'true');
    vmWrapper.setPluginData('sourceComponentSetId', cs.id);
    vmWrapper.fills = [];
    vmWrapper.clipsContent = false;
  }

  // Place to the right of the regular docs wrapper (or CS parent if no wrapper)
  var vmRefFrame = findExistingWrapper(cs) || cs.parent;
  var vmPlacementParent = vmRefFrame.parent || vmRefFrame;
  if (vmWrapper.parent !== vmPlacementParent) {
    vmPlacementParent.appendChild(vmWrapper);
  }
  vmWrapper.x = vmRefFrame.x + vmRefFrame.width + VM_GAP;
  vmWrapper.y = vmRefFrame.y;

  // Use CS dimensions — instances are positioned at their original CS-relative coords
  var vmGridW = cs.width;
  var vmGridH = cs.height;

  // Create one mode block per selected mode, stacked vertically
  var vmCumOffsetY = 0;
  var vmMaxW = 0; // track widest mode block for wrapper sizing

  for (var vmmi = 0; vmmi < vmModes.length; vmmi++) {
    var vmMode = vmModes[vmmi];
    var vmBlockY = vmCumOffsetY;

    // Mode name title
    var vmModeLabel = figma.createText();
    vmModeLabel.fontName = { family: _labelFontFamily, style: 'Regular' };
    vmModeLabel.fontSize = LABEL_FONT_SIZE;
    vmModeLabel.characters = vmMode.name;
    vmModeLabel.fills = [{ type: 'SOLID', color: DOC_COLOR }];
    vmModeLabel.x = vmRowLabelWidth;
    vmModeLabel.y = vmBlockY;
    vmWrapper.appendChild(vmModeLabel);
    vmBlockY += SIMPLE_LABEL_ROW_HEIGHT;

    // Column labels for this mode
    var vmLYO = vmBlockY;
    for (var vmcpi = 0; vmcpi < colAxisProps.length; vmcpi++) {
      var vmCProp = colAxisProps[vmcpi];
      var vmCRowH = colLabelHeights[vmcpi];
      var vmCIsInner = colAxisProps.length > 1 && vmcpi === colAxisProps.length - 1;
      var vmCHasBracket = !vmCIsInner && colAxisProps.length > 1;
      var vmCGroups = getColGroups(vmCProp.name, grid, gridByCol, colClusters, colWidths);
      for (var vmcvi = 0; vmcvi < vmCGroups.length; vmcvi++) {
        var vmCG = vmCGroups[vmcvi];
        vmWrapper.appendChild(createColLabel(
          formatLabel(vmCProp.name, vmCG.value),
          vmRowLabelWidth + vmCG.startX, vmLYO,
          vmCG.endX - vmCG.startX, vmCRowH, vmCHasBracket
        ));
      }
      vmLYO += vmCRowH;
    }

    var vmGridTopY = vmLYO; // where the grid starts for this mode block

    // Row labels for this mode block
    var vmRLXO = 0;
    for (var vmrpi2 = 0; vmrpi2 < rowAxisProps.length; vmrpi2++) {
      var vmRP = rowAxisProps[vmrpi2];
      var vmRII = rowAxisProps.length > 1 && vmrpi2 === rowAxisProps.length - 1;
      var vmRHB = !vmRII && rowAxisProps.length > 1;
      var vmRCW = rowLabelWidths[vmrpi2] || 100;
      var vmRGS = getRowGroups(vmRP.name, grid, gridByRow, rowClusters, rowHeights);
      for (var vmrvi2 = 0; vmrvi2 < vmRGS.length; vmrvi2++) {
        var vmRG2 = vmRGS[vmrvi2];
        vmWrapper.appendChild(createRowLabel(
          formatLabel(vmRP.name, vmRG2.value),
          vmRLXO, vmRG2.startY + vmGridTopY,
          vmRCW, vmRG2.endY - vmRG2.startY, vmRHB
        ));
      }
      vmRLXO += vmRCW;
    }

    // Mode frame — instances inside inherit this variable mode
    var vmFrame = figma.createFrame();
    vmFrame.name = 'Mode: ' + vmMode.name;
    vmFrame.clipsContent = true;
    vmFrame.x = vmRowLabelWidth;
    vmFrame.y = vmGridTopY;
    vmFrame.resize(vmGridW, vmGridH);
    vmWrapper.appendChild(vmFrame);

    // Set background fill using variable (resolves per mode)
    if (vmBgVariable) {
      vmFrame.fills = [figma.variables.setBoundVariableForPaint({ type: 'SOLID', color: { r: 1, g: 1, b: 1 } }, 'color', vmBgVariable)];
    } else {
      vmFrame.fills = [];
    }

    // Set primary collection mode
    try {
      vmFrame.setExplicitVariableModeForCollection(vmCollection, vmMode.modeId);
    } catch (vmErr) {
      console.error('[VarModes] Failed to set mode "' + vmMode.name + '":', vmErr.message);
    }

    // Also apply matching modes from other collections
    for (var vmci = 0; vmci < vmOtherCollections.length; vmci++) {
      var vmOtherColl = vmOtherCollections[vmci];
      for (var vmomi = 0; vmomi < vmOtherColl.modes.length; vmomi++) {
        if (vmOtherColl.modes[vmomi].name === vmMode.name) {
          try {
            vmFrame.setExplicitVariableModeForCollection(vmOtherColl, vmOtherColl.modes[vmomi].modeId);
          } catch (vmOtherErr) {
            // Collection may not be used by these instances — skip silently
          }
          break;
        }
      }
    }

    // Create instances inside the mode frame
    for (var vmgi = 0; vmgi < grid.length; vmgi++) {
      var vmInst = grid[vmgi].node.createInstance();
      vmInst.x = grid[vmgi].node.x;
      vmInst.y = grid[vmgi].node.y;
      vmFrame.appendChild(vmInst);
    }

    // Dashed border
    var vmBorder = figma.createFrame();
    vmBorder.name = 'Grid Border';
    vmBorder.fills = [];
    vmBorder.strokes = [{ type: 'SOLID', color: DOC_COLOR }];
    vmBorder.dashPattern = [4, 4];
    vmBorder.x = vmRowLabelWidth;
    vmBorder.y = vmGridTopY;
    vmBorder.resize(vmGridW, vmGridH);
    vmWrapper.appendChild(vmBorder);

    // Grid lines
    if (options.showGrid) {
      createGridLines(vmWrapper, colClusters, rowClusters, colWidths, rowHeights,
        vmRowLabelWidth, vmGridTopY, vmGridW, vmGridH);
    }

    var vmBlockW = vmRowLabelWidth + vmGridW;
    if (vmBlockW > vmMaxW) vmMaxW = vmBlockW;
    vmCumOffsetY = vmGridTopY + vmGridH + VM_GAP;
  }

  // Resize wrapper to fit all stacked mode blocks
  var vmTotalW = vmMaxW;
  var vmTotalH = vmCumOffsetY - VM_GAP;
  vmWrapper.resize(Math.max(vmTotalW, 1), Math.max(vmTotalH, 1));

  // Wrap contents in a Documentation frame
  var vmDocsFrame = figma.createFrame();
  vmDocsFrame.name = 'Documentation';
  vmDocsFrame.fills = [];
  vmDocsFrame.clipsContent = false;
  vmDocsFrame.locked = false;
  vmDocsFrame.constraints = { horizontal: 'MIN', vertical: 'MIN' };
  vmDocsFrame.x = 0;
  vmDocsFrame.y = 0;
  vmDocsFrame.resize(vmWrapper.width, vmWrapper.height);
  vmWrapper.appendChild(vmDocsFrame);

  var vmDocsChildren = [];
  for (var vmdi = 0; vmdi < vmWrapper.children.length; vmdi++) {
    var vmDChild = vmWrapper.children[vmdi];
    if (vmDChild !== vmDocsFrame) vmDocsChildren.push(vmDChild);
  }
  for (var vmdi2 = 0; vmdi2 < vmDocsChildren.length; vmdi2++) {
    vmDocsChildren[vmdi2].constraints = { horizontal: 'MIN', vertical: 'MIN' };
    vmDocsFrame.appendChild(vmDocsChildren[vmdi2]);
  }

  // Done
  var vmTotalMs = Date.now() - t0;
  var vmTotalSec = (vmTotalMs / 1000).toFixed(2);
  console.log('[VarModes] Generated ' + vmModes.length + ' mode column(s) in ' + vmTotalSec + 's');

  figma.currentPage.selection = [vmWrapper];
  figma.viewport.scrollAndZoomIntoView([vmWrapper]);

  var vmDoneMsg = 'Generated variable modes docs in ' + vmTotalSec + 's.';
  figma.notify(vmDoneMsg);
  figma.ui.postMessage({ type: 'done', message: vmDoneMsg });
}

// ─── Grouping Helpers ────────────────────────────────────────────────────────

function getRowGroups(propName, grid, gridByRow, rowClusters, rowHeights) {
  // Find contiguous row ranges that share the same property value
  const groups = [];
  let currentValue = null;
  let startRow = 0;

  for (let r = 0; r < rowClusters.length; r++) {
    // Get the value of this property for any variant in this row
    const inRow = gridByRow[r] || [];
    if (inRow.length === 0) continue;
    const value = inRow[0].props[propName];

    if (value !== currentValue) {
      if (currentValue !== null) {
        groups.push({
          value: currentValue,
          startRow,
          endRow: r - 1,
          startY: rowClusters[startRow],
          endY: rowClusters[r - 1] + rowHeights[r - 1],
        });
      }
      currentValue = value;
      startRow = r;
    }
  }

  // Push last group
  if (currentValue !== null) {
    const lastRow = rowClusters.length - 1;
    groups.push({
      value: currentValue,
      startRow,
      endRow: lastRow,
      startY: rowClusters[startRow],
      endY: rowClusters[lastRow] + rowHeights[lastRow],
    });
  }

  return groups;
}

function getColGroups(propName, grid, gridByCol, colClusters, colWidths) {
  const groups = [];
  let currentValue = null;
  let startCol = 0;

  for (let c = 0; c < colClusters.length; c++) {
    const inCol = gridByCol[c] || [];
    if (inCol.length === 0) continue;
    const value = inCol[0].props[propName];

    if (value !== currentValue) {
      if (currentValue !== null) {
        groups.push({
          value: currentValue,
          startCol,
          endCol: c - 1,
          startX: colClusters[startCol],
          endX: colClusters[c - 1] + colWidths[c - 1],
        });
      }
      currentValue = value;
      startCol = c;
    }
  }

  if (currentValue !== null) {
    const lastCol = colClusters.length - 1;
    groups.push({
      value: currentValue,
      startCol,
      endCol: lastCol,
      startX: colClusters[startCol],
      endX: colClusters[lastCol] + colWidths[lastCol],
    });
  }

  return groups;
}

// ─── Groupify ─────────────────────────────────────────────────────────────────

function getGroupifyComponentSetData(preferredAxis) {
  var selection = figma.currentPage.selection;
  if (selection.length !== 1) {
    return { error: "Please select exactly one component set." };
  }

  var node = selection[0];

  // If a variant is selected, go up to the component set
  if (node.type === "COMPONENT" && node.parent && node.parent.type === "COMPONENT_SET") {
    node = node.parent;
  }

  if (node.type !== "COMPONENT_SET") {
    return { error: "Please select a component set (a component with variants)." };
  }

  // Extract properties and their values from variant names (visible variants only)
  var properties = {};
  var propertyOrder = [];

  for (var i = 0; i < node.children.length; i++) {
    var variant = node.children[i];
    if (variant.type !== 'COMPONENT' || variant.visible === false) continue;
    var pairs = variant.name.split(",");
    for (var j = 0; j < pairs.length; j++) {
      var eq = pairs[j].indexOf("=");
      if (eq !== -1) {
        var propName = pairs[j].substring(0, eq).trim();
        var propValue = pairs[j].substring(eq + 1).trim();
        if (!properties[propName]) {
          properties[propName] = [];
          propertyOrder.push(propName);
        }
        if (properties[propName].indexOf(propValue) === -1) {
          properties[propName].push(propValue);
        }
      }
    }
  }

  var propsArray = [];
  for (var i = 0; i < propertyOrder.length; i++) {
    propsArray.push({
      name: propertyOrder[i],
      values: properties[propertyOrder[i]]
    });
  }

  // Infer current layout config from canvas positions
  var inferred = groupifyInferConfigFromCanvas(node, propsArray, preferredAxis);

  return { nodeId: node.id, properties: propsArray, inferred: inferred };
}

// Optimize property directions to minimize entirely empty rows/columns.
// Tries all 2^N-2 non-degenerate ROW/COLUMN assignments and picks the one
// with the fewest empty rows + empty columns. On tie, prefers the assignment
// closest to the position-based inference (fewest direction changes).
// preferredAxis: 'x' (more columns), 'y' (more rows), or null/undefined (auto)
function groupifyOptimizeDirections(propsArray, variantKeys, currentConfigs, preferredAxis) {
  var n = propsArray.length;
  if (n < 2) return null; // need at least 2 props for a grid
  if (n > 8) return null; // skip brute-force for large property counts (2^N combinatorial explosion)

  var bestScore = Infinity;
  var bestAxisCount = -1; // higher = more combos on preferred axis
  var bestChanges = Infinity;
  var bestMap = null;

  // Bitmask: bit i = 0 → ROW, bit i = 1 → COLUMN
  // Skip 0 (all ROW) and (2^n - 1) (all COLUMN)
  var limit = (1 << n) - 1;
  for (var mask = 1; mask < limit; mask++) {
    var rowP = [];
    var colP = [];
    for (var i = 0; i < n; i++) {
      var prop = { name: propsArray[i].name, values: propsArray[i].values };
      if (mask & (1 << i)) {
        colP.push(prop);
      } else {
        rowP.push(prop);
      }
    }

    var rowCombos = groupifyBuildCombinations(rowP);
    var colCombos = groupifyBuildCombinations(colP);

    // Count empty rows
    var emptyRows = 0;
    for (var r = 0; r < rowCombos.length; r++) {
      var found = false;
      for (var c = 0; c < colCombos.length && !found; c++) {
        var combo = groupifyMergeObjects(rowCombos[r], colCombos[c]);
        if (variantKeys[groupifyComboToKey(combo)]) found = true;
      }
      if (!found) emptyRows++;
    }

    // Count empty columns
    var emptyCols = 0;
    for (var c = 0; c < colCombos.length; c++) {
      var found = false;
      for (var r = 0; r < rowCombos.length && !found; r++) {
        var combo = groupifyMergeObjects(rowCombos[r], colCombos[c]);
        if (variantKeys[groupifyComboToKey(combo)]) found = true;
      }
      if (!found) emptyCols++;
    }

    var score = emptyRows + emptyCols;

    // Preferred axis tiebreaker: maximize combos on preferred axis
    var axisCount = 0;
    if (preferredAxis === 'x') axisCount = colCombos.length;
    else if (preferredAxis === 'y') axisCount = rowCombos.length;

    // Fallback tiebreaker: fewest direction changes from position-based inference
    var changes = 0;
    if (currentConfigs) {
      for (var i = 0; i < n; i++) {
        var newDir = (mask & (1 << i)) ? "COLUMN" : "ROW";
        if (currentConfigs[i].direction !== newDir) changes++;
      }
    }

    var isBetter = false;
    if (score < bestScore) {
      isBetter = true;
    } else if (score === bestScore) {
      if (preferredAxis && axisCount > bestAxisCount) {
        isBetter = true;
      } else if ((!preferredAxis || axisCount === bestAxisCount) && changes < bestChanges) {
        isBetter = true;
      }
    }

    if (isBetter) {
      bestScore = score;
      bestAxisCount = axisCount;
      bestChanges = changes;
      bestMap = {};
      for (var i = 0; i < n; i++) {
        bestMap[propsArray[i].name] = (mask & (1 << i)) ? "COLUMN" : "ROW";
      }
    }
  }

  // Only return if we improved over the position-based inference,
  // or if a preferred axis was specified (always use optimizer result)
  if (!preferredAxis && currentConfigs) {
    var currentScore = groupifyScoreDirections(propsArray, variantKeys, currentConfigs);
    if (bestScore >= currentScore) return null;
  }

  return bestMap;
}

// Score the current direction assignment (count empty rows + empty cols)
function groupifyScoreDirections(propsArray, variantKeys, configs) {
  var rowP = [], colP = [];
  for (var i = 0; i < configs.length; i++) {
    var prop = { name: propsArray[i].name, values: propsArray[i].values };
    if (configs[i].direction === "COLUMN") colP.push(prop);
    else rowP.push(prop);
  }

  var rowCombos = groupifyBuildCombinations(rowP);
  var colCombos = groupifyBuildCombinations(colP);
  // Degenerate case: all props on one axis → treat as 1-row or 1-col grid
  if (rowCombos.length === 0) rowCombos = [{}];
  if (colCombos.length === 0) colCombos = [{}];
  var score = 0;

  for (var r = 0; r < rowCombos.length; r++) {
    var found = false;
    for (var c = 0; c < colCombos.length && !found; c++) {
      if (variantKeys[groupifyComboToKey(groupifyMergeObjects(rowCombos[r], colCombos[c]))]) found = true;
    }
    if (!found) score++;
  }
  for (var c = 0; c < colCombos.length; c++) {
    var found = false;
    for (var r = 0; r < rowCombos.length && !found; r++) {
      if (variantKeys[groupifyComboToKey(groupifyMergeObjects(rowCombos[r], colCombos[c]))]) found = true;
    }
    if (!found) score++;
  }
  return score;
}

function groupifyInferConfigFromCanvas(node, propsArray, preferredAxis) {
  var children = node.children;
  if (children.length === 0) return null;

  var variants = [];
  for (var i = 0; i < children.length; i++) {
    var v = children[i];
    if (v.type !== 'COMPONENT' || v.visible === false) continue;
    var props = {};
    var pairs = v.name.split(",");
    for (var j = 0; j < pairs.length; j++) {
      var eq = pairs[j].indexOf("=");
      if (eq !== -1) {
        props[pairs[j].substring(0, eq).trim()] = pairs[j].substring(eq + 1).trim();
      }
    }
    variants.push({ props: props, x: v.x, y: v.y, w: v.width, h: v.height });
  }

  var propConfigs = [];
  for (var p = 0; p < propsArray.length; p++) {
    var propName = propsArray[p].name;
    var values = propsArray[p].values;

    var sumX = {}, sumY = {}, cnt = {};
    for (var v = 0; v < values.length; v++) {
      sumX[values[v]] = 0;
      sumY[values[v]] = 0;
      cnt[values[v]] = 0;
    }
    for (var i = 0; i < variants.length; i++) {
      var val = variants[i].props[propName];
      if (val && cnt[val] !== undefined) {
        sumX[val] += variants[i].x + variants[i].w / 2;
        sumY[val] += variants[i].y + variants[i].h / 2;
        cnt[val]++;
      }
    }

    var centroids = {};
    for (var v = 0; v < values.length; v++) {
      var c = cnt[values[v]];
      centroids[values[v]] = {
        x: c > 0 ? sumX[values[v]] / c : 0,
        y: c > 0 ? sumY[values[v]] / c : 0
      };
    }

    var minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    for (var v = 0; v < values.length; v++) {
      var cx = centroids[values[v]].x;
      var cy = centroids[values[v]].y;
      if (cx < minX) minX = cx;
      if (cx > maxX) maxX = cx;
      if (cy < minY) minY = cy;
      if (cy > maxY) maxY = cy;
    }
    var rangeX = maxX - minX;
    var rangeY = maxY - minY;

    var direction, sortedValues;
    if (rangeX < 1 && rangeY < 1) {
      direction = "COLUMN";
      sortedValues = values.slice();
    } else if (rangeY > rangeX) {
      direction = "ROW";
      sortedValues = values.slice().sort(function (a, b) {
        return centroids[a].y - centroids[b].y;
      });
    } else {
      direction = "COLUMN";
      sortedValues = values.slice().sort(function (a, b) {
        return centroids[a].x - centroids[b].x;
      });
    }

    propConfigs.push({
      name: propName,
      values: sortedValues,
      direction: direction,
      spacing: 48
    });
  }

  // Optimize directions to avoid entirely empty rows/columns
  var variantKeys = {};
  for (var i = 0; i < variants.length; i++) {
    var keyParts = [];
    for (var pn in variants[i].props) keyParts.push(pn + "=" + variants[i].props[pn]);
    keyParts.sort();
    variantKeys[keyParts.join(", ")] = true;
  }
  var optimized = groupifyOptimizeDirections(propsArray, variantKeys, propConfigs, preferredAxis);
  if (optimized) {
    for (var p = 0; p < propConfigs.length; p++) {
      propConfigs[p].direction = optimized[propConfigs[p].name];
    }
  }

  var yMap = {}, xMap = {};
  for (var i = 0; i < variants.length; i++) {
    var yk = Math.round(variants[i].y);
    var xk = Math.round(variants[i].x);
    if (!yMap[yk]) yMap[yk] = { pos: variants[i].y, size: 0 };
    yMap[yk].size = Math.max(yMap[yk].size, variants[i].h);
    if (!xMap[xk]) xMap[xk] = { pos: variants[i].x, size: 0 };
    xMap[xk].size = Math.max(xMap[xk].size, variants[i].w);
  }

  var yRows = groupifyObjToSortedArray(yMap);
  var xCols = groupifyObjToSortedArray(xMap);

  var yGaps = groupifyComputeGaps(yRows);
  var xGaps = groupifyComputeGaps(xCols);

  var rowProps = [], colProps = [];
  for (var p = 0; p < propConfigs.length; p++) {
    if (propConfigs[p].direction === "ROW") rowProps.push(propConfigs[p]);
    else colProps.push(propConfigs[p]);
  }
  groupifyAssignSpacings(yGaps, rowProps);
  groupifyAssignSpacings(xGaps, colProps);

  var minVx = Infinity, minVy = Infinity, maxVx = -Infinity, maxVy = -Infinity;
  for (var i = 0; i < variants.length; i++) {
    if (variants[i].x < minVx) minVx = variants[i].x;
    if (variants[i].y < minVy) minVy = variants[i].y;
    var r = variants[i].x + variants[i].w;
    var b = variants[i].y + variants[i].h;
    if (r > maxVx) maxVx = r;
    if (b > maxVy) maxVy = b;
  }

  var padding = {
    top: Math.max(0, Math.round(minVy)),
    right: Math.max(0, Math.round(node.width - maxVx)),
    bottom: Math.max(0, Math.round(node.height - maxVy)),
    left: Math.max(0, Math.round(minVx))
  };

  var resultProps = [];
  for (var p = 0; p < propConfigs.length; p++) {
    resultProps.push({
      name: propConfigs[p].name,
      values: propConfigs[p].values,
      direction: propConfigs[p].direction,
      spacing: propConfigs[p].spacing
    });
  }

  return {
    padding: padding,
    sectionDirection: "GRID",
    properties: resultProps
  };
}

function groupifyObjToSortedArray(map) {
  var arr = [];
  for (var k in map) arr.push(map[k]);
  arr.sort(function (a, b) { return a.pos - b.pos; });
  return arr;
}

function groupifyComputeGaps(entries) {
  var gaps = [];
  for (var i = 1; i < entries.length; i++) {
    var gap = Math.round(entries[i].pos - entries[i - 1].pos - entries[i - 1].size);
    if (gap > 0) gaps.push(gap);
  }
  return gaps;
}

function groupifyAssignSpacings(gaps, axisProps) {
  if (axisProps.length === 0 || gaps.length === 0) return;

  var sorted = gaps.slice().sort(function (a, b) { return a - b; });
  var clusters = [{ total: sorted[0], count: 1, rep: sorted[0] }];
  for (var i = 1; i < sorted.length; i++) {
    var last = clusters[clusters.length - 1];
    if (sorted[i] - last.rep <= 2) {
      last.total += sorted[i];
      last.count++;
      last.rep = Math.round(last.total / last.count);
    } else {
      clusters.push({ total: sorted[i], count: 1, rep: sorted[i] });
    }
  }

  clusters.sort(function (a, b) { return b.count - a.count; });

  var propsSorted = axisProps.slice().sort(function (a, b) {
    return b.values.length - a.values.length;
  });

  for (var i = 0; i < propsSorted.length && i < clusters.length; i++) {
    propsSorted[i].spacing = clusters[i].rep;
  }
}

function sendGroupifySelectionData(preferredAxis) {
  var _tg0 = Date.now();
  var data = getGroupifyComponentSetData(preferredAxis);
  if (DEBUG) console.log('[Perf] sendGroupifySelectionData: ' + (Date.now() - _tg0) + 'ms');
  if (data.error) {
    figma.ui.postMessage({ type: "groupify-error", message: data.error });
  } else {
    figma.ui.postMessage({ type: "groupify-init", data: data });
  }
}

// Quick align: re-infer directions with preferred axis, apply with uniform spacing/padding
async function groupifyQuickAlign(msg) {
  var data = getGroupifyComponentSetData(msg.preferredAxis);
  if (data.error) {
    figma.ui.postMessage({ type: "groupify-error", message: data.error });
    return;
  }

  var sp = msg.spacing || 40;
  var properties = data.inferred.properties;
  for (var i = 0; i < properties.length; i++) {
    properties[i].spacing = sp;
  }

  var config = {
    nodeId: data.nodeId,
    sectionDirection: data.inferred.sectionDirection,
    padding: { top: sp, right: sp, bottom: sp, left: sp },
    properties: properties
  };

  await applyGroupifyLayout(config);

  // Re-send init so UI reflects the new directions, padding, and spacing
  sendGroupifySelectionData(msg.preferredAxis);
}

async function applyGroupifyLayout(config) {
  var node = await figma.getNodeByIdAsync(config.nodeId);
  if (!node || node.type !== "COMPONENT_SET") {
    figma.ui.postMessage({ type: "groupify-error", message: "Component set not found." });
    return;
  }

  // Disable auto-layout before positioning (prevents reflow on resize)
  if (node.layoutMode !== 'NONE') {
    node.layoutMode = 'NONE';
  }

  var pad = config.padding;
  var properties = config.properties;
  var sectionDirection = config.sectionDirection || "GRID";

  var sectionProps = [];
  var rowProps = [];
  var colProps = [];
  for (var i = 0; i < properties.length; i++) {
    if (properties[i].direction === "SECTION") {
      sectionProps.push(properties[i]);
    } else if (properties[i].direction === "ROW") {
      rowProps.push(properties[i]);
    } else {
      colProps.push(properties[i]);
    }
  }

  var sectionRowProps, sectionColProps;
  if (sectionProps.length === 0) {
    sectionRowProps = [];
    sectionColProps = [];
  } else if (sectionProps.length === 1 || sectionDirection !== "GRID") {
    if (sectionDirection === "VERTICAL") {
      sectionRowProps = sectionProps;
      sectionColProps = [];
    } else {
      sectionRowProps = [];
      sectionColProps = sectionProps;
    }
  } else {
    sectionRowProps = sectionProps.slice(0, 1);
    sectionColProps = sectionProps.slice(1);
  }

  var sectionRowCombos = groupifyBuildCombinations(sectionRowProps);
  var sectionColCombos = groupifyBuildCombinations(sectionColProps);
  var rowCombos = groupifyBuildCombinations(rowProps);
  var colCombos = groupifyBuildCombinations(colProps);
  if (sectionRowCombos.length === 0) sectionRowCombos = [{}];
  if (sectionColCombos.length === 0) sectionColCombos = [{}];
  if (rowCombos.length === 0) rowCombos = [{}];
  if (colCombos.length === 0) colCombos = [{}];

  var variantMap = {};
  for (var i = 0; i < node.children.length; i++) {
    var child = node.children[i];
    if (child.type !== 'COMPONENT' || child.visible === false) continue;
    var key = groupifyVariantToKey(child.name);
    variantMap[key] = child;
  }

  var grids = [];
  for (var sr = 0; sr < sectionRowCombos.length; sr++) {
    grids[sr] = [];
    for (var sc = 0; sc < sectionColCombos.length; sc++) {
      var sectionCombo = groupifyMergeObjects(sectionRowCombos[sr], sectionColCombos[sc]);
      var grid = [];
      for (var r = 0; r < rowCombos.length; r++) {
        grid[r] = [];
        for (var c = 0; c < colCombos.length; c++) {
          var fullCombo = groupifyMergeObjects(groupifyMergeObjects(sectionCombo, rowCombos[r]), colCombos[c]);
          grid[r][c] = variantMap[groupifyComboToKey(fullCombo)] || null;
        }
      }
      grids[sr][sc] = grid;
    }
  }

  var colWidths = [];
  var rowHeights = [];
  for (var c = 0; c < colCombos.length; c++) {
    colWidths[c] = 0;
    for (var sr = 0; sr < sectionRowCombos.length; sr++) {
      for (var sc = 0; sc < sectionColCombos.length; sc++) {
        for (var r = 0; r < rowCombos.length; r++) {
          var v = grids[sr][sc][r][c];
          if (v) colWidths[c] = Math.max(colWidths[c], Math.max(v.width, 1));
        }
      }
    }
  }
  for (var r = 0; r < rowCombos.length; r++) {
    rowHeights[r] = 0;
    for (var sr = 0; sr < sectionRowCombos.length; sr++) {
      for (var sc = 0; sc < sectionColCombos.length; sc++) {
        for (var c = 0; c < colCombos.length; c++) {
          var v = grids[sr][sc][r][c];
          if (v) rowHeights[r] = Math.max(rowHeights[r], Math.max(v.height, 1));
        }
      }
    }
  }

  // Fill empty row/column dimensions with fallback from populated ones
  var maxRowH = 0, maxColW = 0;
  for (var r = 0; r < rowHeights.length; r++) {
    if (rowHeights[r] > maxRowH) maxRowH = rowHeights[r];
  }
  for (var c = 0; c < colWidths.length; c++) {
    if (colWidths[c] > maxColW) maxColW = colWidths[c];
  }
  for (var r = 0; r < rowHeights.length; r++) {
    if (rowHeights[r] === 0) rowHeights[r] = maxRowH;
  }
  for (var c = 0; c < colWidths.length; c++) {
    if (colWidths[c] === 0) colWidths[c] = maxColW;
  }

  var rowSpacings = groupifyCalcSpacings(rowCombos, rowProps);
  var colSpacings = groupifyCalcSpacings(colCombos, colProps);

  var innerWidth = 0;
  for (var c = 0; c < colWidths.length; c++) {
    innerWidth += colWidths[c];
    if (c > 0) innerWidth += colSpacings[c - 1];
  }
  var innerHeight = 0;
  for (var r = 0; r < rowHeights.length; r++) {
    innerHeight += rowHeights[r];
    if (r > 0) innerHeight += rowSpacings[r - 1];
  }

  var sectionRowSpacings = groupifyCalcSpacings(sectionRowCombos, sectionRowProps);
  var sectionColSpacings = groupifyCalcSpacings(sectionColCombos, sectionColProps);

  var emptyCells = [];
  var sy = pad.top;
  for (var sr = 0; sr < sectionRowCombos.length; sr++) {
    if (sr > 0) sy += sectionRowSpacings[sr - 1];
    var sx = pad.left;
    for (var sc = 0; sc < sectionColCombos.length; sc++) {
      if (sc > 0) sx += sectionColSpacings[sc - 1];

      var iy = sy;
      for (var r = 0; r < rowCombos.length; r++) {
        if (r > 0) iy += rowSpacings[r - 1];
        var ix = sx;
        for (var c = 0; c < colCombos.length; c++) {
          if (c > 0) ix += colSpacings[c - 1];
          var v = grids[sr][sc][r][c];
          if (v) {
            v.x = ix;
            v.y = iy;
          } else {
            emptyCells.push({ x: ix, y: iy, w: colWidths[c], h: rowHeights[r] });
          }
          ix += colWidths[c];
        }
        iy += rowHeights[r];
      }

      sx += innerWidth;
    }
    sy += innerHeight;
  }

  var totalWidth = pad.left + pad.right;
  for (var sc = 0; sc < sectionColCombos.length; sc++) {
    totalWidth += innerWidth;
    if (sc > 0) totalWidth += sectionColSpacings[sc - 1];
  }
  var totalHeight = pad.top + pad.bottom;
  for (var sr = 0; sr < sectionRowCombos.length; sr++) {
    totalHeight += innerHeight;
    if (sr > 0) totalHeight += sectionRowSpacings[sr - 1];
  }

  node.resizeWithoutConstraints(Math.max(1, totalWidth), Math.max(1, totalHeight));

  var numSections = sectionRowCombos.length * sectionColCombos.length;
  var sectionInfo = numSections > 1
    ? " (" + sectionRowCombos.length + "\u00d7" + sectionColCombos.length + " sections)"
    : "";
  var emptyInfo = emptyCells.length > 0
    ? " (" + emptyCells.length + " empty)"
    : "";
  figma.ui.postMessage({
    type: "groupify-success",
    message: "Layout applied! " + rowCombos.length + " rows \u00d7 " + colCombos.length + " cols" + sectionInfo + emptyInfo + "."
  });
}

function groupifyBuildCombinations(props) {
  if (props.length === 0) return [];
  var result = [{}];
  for (var i = 0; i < props.length; i++) {
    var next = [];
    var values = props[i].values;
    for (var j = 0; j < result.length; j++) {
      for (var k = 0; k < values.length; k++) {
        var combo = {};
        for (var key in result[j]) combo[key] = result[j][key];
        combo[props[i].name] = values[k];
        next.push(combo);
      }
    }
    result = next;
  }
  return result;
}

function groupifyVariantToKey(name) {
  var pairs = name.split(",");
  var sorted = [];
  for (var i = 0; i < pairs.length; i++) {
    var eq = pairs[i].indexOf("=");
    if (eq !== -1) {
      sorted.push(pairs[i].substring(0, eq).trim() + "=" + pairs[i].substring(eq + 1).trim());
    }
  }
  sorted.sort();
  return sorted.join(", ");
}

function groupifyComboToKey(combo) {
  var pairs = [];
  for (var key in combo) pairs.push(key + "=" + combo[key]);
  pairs.sort();
  return pairs.join(", ");
}

function groupifyMergeObjects(a, b) {
  var r = {};
  for (var k in a) r[k] = a[k];
  for (var k in b) r[k] = b[k];
  return r;
}

function groupifyCalcSpacings(combos, props) {
  var spacings = [];
  for (var i = 1; i < combos.length; i++) {
    var spacing = 0;
    for (var p = 0; p < props.length; p++) {
      if (combos[i][props[p].name] !== combos[i - 1][props[p].name]) {
        spacing = props[p].spacing;
        break;
      }
    }
    spacings.push(spacing);
  }
  return spacings;
}
