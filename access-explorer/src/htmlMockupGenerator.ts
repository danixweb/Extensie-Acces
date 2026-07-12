import { DesignControl, DesignSection, ParsedDesign } from './formDesignParser';

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function px(n: number): number {
  return Math.round(n * 10) / 10;
}

function controlStyle(c: DesignControl): string {
  const parts = [
    'position:absolute',
    `left:${px(c.left)}px`,
    `top:${px(c.top)}px`,
    `width:${px(Math.max(c.width, 4))}px`,
    `height:${px(Math.max(c.height, 4))}px`,
    'box-sizing:border-box',
    'overflow:hidden',
  ];
  if (c.fontName) {
    parts.push(`font-family:'${c.fontName.replace(/'/g, '')}'`);
  }
  if (c.fontSize) {
    parts.push(`font-size:${c.fontSize}pt`);
  }
  if (c.fontWeight) {
    parts.push(`font-weight:${c.fontWeight}`);
  }
  if (c.foreColor) {
    parts.push(`color:${c.foreColor}`);
  }
  if (c.backColor) {
    parts.push(`background-color:${c.backColor}`);
  }
  if (c.borderStyle === 'solid') {
    parts.push('border:1px solid #888');
  }
  return parts.join(';');
}

function renderControl(c: DesignControl): string {
  const style = controlStyle(c);
  const label = escapeHtml(c.caption ?? c.name);
  switch (c.type) {
    case 'Label':
      return `<div class="ctl ctl-label" style="${style}">${label}</div>`;
    case 'CommandButton':
      return `<div class="ctl ctl-button" style="${style}">${label}</div>`;
    case 'TextBox':
    case 'ComboBox':
    case 'ListBox': {
      const text = c.controlSource ? `[${escapeHtml(c.controlSource)}]` : '';
      const cls = c.type === 'ComboBox' ? 'ctl-combobox' : c.type === 'ListBox' ? 'ctl-listbox' : 'ctl-textbox';
      return `<div class="ctl ${cls}" style="${style}">${text}</div>`;
    }
    case 'CheckBox':
    case 'OptionButton':
    case 'ToggleButton':
      return `<div class="ctl ctl-check" style="${style}"><span class="ctl-check-box"></span>${label}</div>`;
    case 'OptionGroup':
      return `<fieldset class="ctl ctl-optiongroup" style="${style}"><legend>${label}</legend></fieldset>`;
    case 'Line':
      return `<div class="ctl ctl-line" style="${style}"></div>`;
    case 'Rectangle':
      return `<div class="ctl ctl-rectangle" style="${style}"></div>`;
    case 'Image':
    case 'BoundObjectFrame':
    case 'UnboundObjectFrame':
      // No access to the actual embedded image binary through this text export —
      // an explicit placeholder, not a broken/missing image, per the confirmed design decision.
      return `<div class="ctl ctl-image" style="${style}">🖼 ${label}</div>`;
    case 'Subform':
    case 'Subreport':
    case 'SubReport': {
      const child = c.sourceObject?.replace(/^(Form|Report)\./, '') ?? c.name;
      return `<div class="ctl ctl-subform" style="${style}">⬚ ${escapeHtml(child)} (subformular/subraport neexpandat)</div>`;
    }
    case 'Tab':
      // Approximate only: page strip shown, individual Page contents not parsed in this first pass.
      return `<div class="ctl ctl-tab" style="${style}"><div class="ctl-tab-strip">${label}</div></div>`;
    default:
      return `<div class="ctl ctl-generic" style="${style}" title="${escapeHtml(c.type)}">${label}</div>`;
  }
}

function renderSection(s: DesignSection): string {
  const style = [
    'position:relative',
    'width:100%',
    `min-height:${px(Math.max(s.height, 2))}px`,
    s.backColor ? `background-color:${s.backColor}` : '',
  ]
    .filter(Boolean)
    .join(';');
  const controls = s.controls.map(renderControl).join('\n    ');
  return `  <div class="section" style="${style}">
    <div class="section-label">${escapeHtml(s.name)}</div>
    ${controls}
  </div>`;
}

/** Produces a self-contained HTML document (inline CSS, no external resources). */
export function generateHtml(design: ParsedDesign, objectName: string): string {
  const sections = design.sections.map(renderSection).join('\n');
  const title = escapeHtml(design.caption || objectName);
  const kind = design.objectType === 'Form' ? 'Formular' : 'Raport';
  return `<!doctype html>
<html lang="ro">
<head>
<meta charset="utf-8">
<title>${title} — mockup</title>
<style>
  body { margin: 0; padding: 24px; background: #ddd; font-family: 'Segoe UI', Tahoma, sans-serif; }
  .page-title { font: 600 13px 'Segoe UI', sans-serif; margin-bottom: 10px; color: #333; }
  .container { position: relative; width: ${px(Math.max(design.width, 100))}px; background: #f0f0f0; border: 1px solid #999; box-shadow: 0 1px 6px rgba(0,0,0,.25); }
  .section { border-bottom: 1px dashed #bbb; }
  .section:last-child { border-bottom: none; }
  .section-label { position: absolute; left: 2px; top: 2px; font-size: 9px; color: #999; z-index: 1000; pointer-events: none; }
  .ctl { font-size: 8pt; font-family: 'Segoe UI', sans-serif; color: #000; }
  .ctl-label { display: flex; align-items: center; }
  .ctl-button { display: flex; align-items: center; justify-content: center; background: #e1e1e1; border: 1px solid #adadad; border-radius: 2px; text-align: center; }
  .ctl-textbox, .ctl-combobox, .ctl-listbox { background: #fff; border: 1px solid #7a7a7a; display: flex; align-items: center; padding: 0 2px; color: #555; }
  .ctl-check { display: flex; align-items: center; gap: 4px; }
  .ctl-check-box { width: 11px; height: 11px; border: 1px solid #555; display: inline-block; flex: none; }
  .ctl-optiongroup { border: 1px solid #888; margin: 0; }
  .ctl-line { border-top: 1px solid #000; height: 0 !important; }
  .ctl-rectangle { border: 1px solid #555; background: transparent; }
  .ctl-image { border: 1px dashed #999; display: flex; align-items: center; justify-content: center; color: #777; background: #f7f7f7; text-align: center; }
  .ctl-subform { border: 1px dashed #4a6fa5; display: flex; align-items: center; justify-content: center; color: #4a6fa5; background: #eef3fa; text-align: center; }
  .ctl-tab { border: 1px solid #888; background: #f4f4f4; }
  .ctl-tab-strip { padding: 2px 6px; border-bottom: 1px solid #888; display: inline-block; background: #e8e8e8; }
  .ctl-generic { border: 1px dotted #aaa; }
</style>
</head>
<body>
  <div class="page-title">${title} (${kind}) — mockup generat automat din design-ul Access, aproximativ</div>
  <div class="container">
${sections}
  </div>
</body>
</html>
`;
}
