import React, { useRef, useState, useEffect } from 'react';
import {
  Layout, Typography, Button, Input, Tooltip, Breadcrumb,
  Select, message, Tag, Badge, Modal, Form, Switch,
} from 'antd';
import {
  HomeOutlined, ReloadOutlined, ArrowLeftOutlined, ArrowRightOutlined,
  VideoCameraOutlined, StopOutlined, LockOutlined,
  AimOutlined, FileTextOutlined, CheckSquareOutlined,
  DeleteOutlined, CloseOutlined, MenuFoldOutlined, MenuUnfoldOutlined,
  KeyOutlined, SettingOutlined, UserOutlined, EyeInvisibleOutlined, EyeTwoTone,
  CameraOutlined,
} from '@ant-design/icons';
import { Link, useNavigate } from 'react-router-dom';
import ScreenshotAnnotator from '../../components/ScreenshotAnnotator';
import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  ImageRun, HeadingLevel, TextRun, WidthType, ShadingType,
  AlignmentType,
} from 'docx';

const { Text } = Typography;
const { Option } = Select;

const REDWOOD = '#C74634';

interface CapturedField {
  fieldName: string;
  action: string;
  value: string;
  description: string;
}

interface Step {
  id: string;
  type: 'click' | 'input' | 'navigate' | 'snapshot';
  description: string;
  fieldName: string;   // label / button text / element name
  action: string;      // Click | Enter | Select | Check | Navigate
  value: string;       // entered value, selected option, or button text
  url: string;
  pageTitle: string;
  timestamp: number;
  screenshot?: string;
  fields?: CapturedField[];  // consolidated fields from Capture Fields / Capture Tab
}

const FUSION_URLS = [
  { label: 'Oracle Fusion Home',          value: '' },
  { label: 'Payables — Manage Invoices',  value: '' },
  { label: 'Payables — Manage Payments',  value: '' },
  { label: 'General Ledger',              value: '' },
  { label: 'Suppliers',                   value: '' },
];

// Injected into the webview to capture user interactions
const INJECT_SCRIPT = `
(function() {
  if (window.__reactErpTracking) return;
  window.__reactErpTracking = true;
  window.__reactErpSteps = window.__reactErpSteps || [];

  // ── Label resolution ──────────────────────────────────────────────────────
  function getDirectLabel(el) {
    if (!el || !el.getAttribute) return '';
    // aria-label is most reliable
    var al = el.getAttribute('aria-label'); if (al) return al.trim();
    // aria-labelledby
    var alby = el.getAttribute('aria-labelledby');
    if (alby) {
      var parts = alby.split(' ');
      var texts = [];
      for (var k = 0; k < parts.length; k++) {
        var ref = document.getElementById(parts[k]);
        if (ref) texts.push(ref.textContent.trim());
      }
      if (texts.length) return texts.join(' ');
    }
    // explicit <label for="id">
    if (el.id) {
      var lbl = document.querySelector('label[for="' + el.id + '"]');
      if (lbl) return lbl.textContent.trim().replace(/:$/, '').trim();
    }
    // title / placeholder / name
    return (el.getAttribute('title') || el.getAttribute('placeholder') || el.name || '').trim();
  }

  function findLabel(el) {
    var direct = getDirectLabel(el);
    if (direct) return direct;

    // Walk up DOM — Oracle ADF puts label in adjacent cell (previousElementSibling)
    var node = el.parentElement;
    for (var i = 0; i < 8 && node; i++) {
      // Check previous sibling for label text (ADF panel form layout)
      var prev = node.previousElementSibling;
      if (prev) {
        var lblEl = prev.querySelector('label') || prev.querySelector('[role="label"]');
        if (lblEl) {
          var t = lblEl.textContent.trim().replace(/:$/, '').trim();
          if (t && t.length < 80) return t;
        }
        // Sometimes the sibling itself is the label cell
        var pt = prev.textContent.trim().replace(/:$/, '').trim();
        if (pt && pt.length < 60 && !pt.includes('\\n') && !/^[0-9,.$%]+$/.test(pt)) return pt;
      }
      // <label> anywhere in current ancestor
      var labels = node.querySelectorAll('label');
      for (var j = 0; j < labels.length; j++) {
        var lt = labels[j].textContent.trim().replace(/:$/, '').trim();
        if (lt && lt.length < 80) return lt;
      }
      node = node.parentElement;
    }
    return '';
  }

  // ── Field info extraction ─────────────────────────────────────────────────
  function getFieldInfo(el) {
    if (!el || !el.tagName) return null;
    var tag  = el.tagName.toLowerCase();
    var type = (el.getAttribute('type') || '').toLowerCase();
    var role = (el.getAttribute('role') || '').toLowerCase();
    var fieldName, action, value;

    if (tag === 'select') {
      fieldName = findLabel(el) || 'Dropdown';
      action    = 'Select';
      value     = el.options && el.selectedIndex >= 0 ? (el.options[el.selectedIndex].text || el.value) : el.value;
    } else if (type === 'checkbox') {
      fieldName = findLabel(el) || 'Checkbox';
      action    = el.checked ? 'Check' : 'Uncheck';
      value     = el.checked ? 'Yes' : 'No';
    } else if (type === 'radio') {
      fieldName = findLabel(el) || 'Radio';
      action    = 'Select';
      value     = el.value || '';
    } else if (tag === 'input' || tag === 'textarea' || role === 'textbox' || role === 'combobox' || role === 'spinbutton') {
      fieldName = findLabel(el) || el.getAttribute('placeholder') || 'Field';
      action    = 'Enter';
      value     = (el.value || el.textContent || '').trim().slice(0, 100);
    } else if (tag === 'button' || type === 'button' || type === 'submit') {
      var btnText = (el.innerText || el.textContent || '').trim().slice(0, 80);
      fieldName   = el.getAttribute('aria-label') || btnText || 'Button';
      action      = 'Click';
      value       = fieldName;
    } else if (tag === 'a') {
      var linkText = (el.innerText || el.textContent || '').trim().slice(0, 80);
      fieldName    = el.getAttribute('aria-label') || linkText || el.getAttribute('title') || 'Link';
      action       = 'Click';
      value        = fieldName;
    } else {
      // Generic — get the most meaningful visible text
      var elText = (el.getAttribute('aria-label') || el.getAttribute('title') ||
                   (el.innerText || el.textContent || '').trim()).slice(0, 80);
      if (!elText || elText.length < 2) return null; // skip meaningless elements
      fieldName = elText;
      action    = 'Click';
      value     = elText;
    }
    return { fieldName: fieldName, action: action, value: value };
  }

  // ── Click listener — skip plain input/textarea/select (captured via change) ─
  document.addEventListener('click', function(e) {
    var el = e.target;
    if (!el || !el.tagName) return;
    var tag  = el.tagName.toLowerCase();
    var type = (el.getAttribute('type') || '').toLowerCase();
    var role = (el.getAttribute('role') || '').toLowerCase();

    // Skip focusable field elements — their data entry is captured via 'change'
    if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
    if (role === 'textbox' || role === 'combobox' || role === 'spinbutton') return;
    if (type === 'text' || type === 'number' || type === 'date' || type === 'email') return;

    var info = getFieldInfo(el);
    if (!info) return;
    window.__reactErpSteps.push({
      type: 'click', fieldName: info.fieldName, action: info.action, value: info.value,
      description: info.action + ': ' + info.fieldName,
      url: location.href, pageTitle: document.title, timestamp: Date.now()
    });
  }, true);

  // ── Change listener — fires when field value is committed ─────────────────
  document.addEventListener('change', function(e) {
    var el = e.target;
    var info = getFieldInfo(el);
    if (!info) return;
    window.__reactErpSteps.push({
      type: 'input', fieldName: info.fieldName, action: info.action, value: info.value,
      description: info.action + ' "' + info.value + '" in: ' + info.fieldName,
      url: location.href, pageTitle: document.title, timestamp: Date.now()
    });
  }, true);

  // Also capture blur on text inputs (ADF doesn't always fire 'change')
  document.addEventListener('blur', function(e) {
    var el = e.target;
    if (!el || !el.tagName) return;
    var tag  = el.tagName.toLowerCase();
    var role = (el.getAttribute('role') || '').toLowerCase();
    if (tag !== 'input' && tag !== 'textarea' && role !== 'textbox' && role !== 'combobox') return;
    var val = (el.value || el.textContent || '').trim();
    if (!val) return;
    var info = getFieldInfo(el);
    if (!info || !info.fieldName) return;
    // Avoid double-capture if 'change' already pushed this
    var last = window.__reactErpSteps[window.__reactErpSteps.length - 1];
    if (last && last.fieldName === info.fieldName && last.value === val) return;
    window.__reactErpSteps.push({
      type: 'input', fieldName: info.fieldName, action: 'Enter', value: val,
      description: 'Enter "' + val + '" in: ' + info.fieldName,
      url: location.href, pageTitle: document.title, timestamp: Date.now()
    });
  }, true);

  console.log('[ReactERP] Step tracking active');

  // ── Snapshot: scan all filled inputs AND read-only ADF display fields ────────
  window.__reactErpCaptureFields = function() {
    var results = [];
    var seen = {};

    // ── 1. Standard form inputs / ARIA widgets ──────────────────────────────
    var els = Array.from(document.querySelectorAll(
      'input, select, textarea, [role="textbox"], [role="combobox"], [role="spinbutton"]'
    ));
    els.forEach(function(el) {
      var tag  = el.tagName.toLowerCase();
      var type = (el.getAttribute('type') || 'text').toLowerCase();
      if (type === 'hidden' || type === 'submit' || type === 'button' ||
          type === 'image'  || type === 'reset'  || type === 'file') return;

      var value = '';
      if (tag === 'select') {
        value = el.options && el.selectedIndex >= 0 ? (el.options[el.selectedIndex].text || el.value) : el.value;
      } else if (type === 'checkbox' || type === 'radio') {
        if (!el.checked) return;
        value = findLabel(el) || (el.checked ? 'Yes' : 'No');
      } else {
        value = (el.value || el.textContent || '').trim();
      }
      if (!value || value.length < 1) return;

      var label = findLabel(el);
      if (!label || label.length < 2) return;
      if (seen[label]) return;
      seen[label] = true;

      var action = tag === 'select' ? 'Select' : (type === 'checkbox' || type === 'radio') ? 'Check' : 'Enter';
      results.push({
        type: 'input',
        fieldName: label,
        action: action,
        value: value,
        description: action + ' "' + value + '" in: ' + label,
        url: location.href,
        pageTitle: document.title,
        timestamp: Date.now()
      });
    });

    // ── 2. Oracle ADF read-only display fields ──────────────────────────────
    // ADF renders read-only fields as label text + adjacent output text (spans/divs/tds).
    // Strategy: find label elements, then look for the value in the next sibling cell/element.
    var labelEls = Array.from(document.querySelectorAll(
      'label, [class*="af_panelLabelAndMessage_label"], [class*="AFPanelFormLayoutLabel"]'
    ));
    labelEls.forEach(function(lbl) {
      var labelText = lbl.textContent.trim().replace(/:$/, '').trim();
      if (!labelText || labelText.length < 2 || labelText.length > 80) return;
      if (seen[labelText]) return;

      // Try to find the value in the adjacent element
      // Pattern 1: <td>label</td><td>value</td>
      var parentCell = lbl.closest('td, th');
      var valueEl = parentCell ? parentCell.nextElementSibling : null;
      // Pattern 2: label is directly followed by sibling span/div
      if (!valueEl) valueEl = lbl.nextElementSibling;
      // Pattern 3: parent div → next sibling div (ADF panelFormLayout)
      if (!valueEl) {
        var p = lbl.parentElement;
        if (p) valueEl = p.nextElementSibling;
      }

      if (!valueEl) return;

      // Get text, skip if empty, too short, or contains nested inputs (handled above)
      var val = (valueEl.textContent || '').trim();
      // Strip internal whitespace runs
      val = val.replace(/\\s+/g, ' ').trim();
      if (!val || val.length < 1 || val.length > 300) return;
      // Skip if the element itself contains an input (already captured above)
      if (valueEl.querySelector('input, select, textarea')) return;
      // Skip noise values
      if (/^[0-9]+$/.test(val) && val.length < 2) return;

      seen[labelText] = true;
      results.push({
        type: 'input',
        fieldName: labelText,
        action: 'Display',
        value: val,
        description: 'Display "' + val + '" — ' + labelText,
        url: location.href,
        pageTitle: document.title,
        timestamp: Date.now()
      });
    });

    return JSON.stringify(results);
  };
})();
true;
`;

const escapeHtml = (s: string) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const downloadHtml = (content: string, filename: string) => {
  const blob = new Blob([content], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

const generateUserManual = (steps: Step[]): string => {
  const date = new Date().toLocaleString();

  // Group steps by screen (pageTitle)
  const screens: { title: string; steps: (Step & { globalIdx: number })[] }[] = [];
  let globalIdx = 0;
  for (const s of steps) {
    const title = s.pageTitle || 'Oracle Fusion';
    let screen = screens.find(sc => sc.title === title);
    if (!screen) { screen = { title, steps: [] }; screens.push(screen); }
    screen.steps.push({ ...s, globalIdx: ++globalIdx });
  }

  const typeLabel: Record<string, string> = { click: 'Click', input: 'Input', navigate: 'Navigate' };
  const typeColor: Record<string, string> = { click: '#1565c0', input: '#2e7d32', navigate: '#e65100' };
  const typeBg:    Record<string, string> = { click: '#e3f2fd', input: '#e8f5e9', navigate: '#fff3e0' };

  const toc = screens.map((sc, i) =>
    `<li><a href="#screen-${i}">${escapeHtml(sc.title)}</a> <span style="color:#999">(${sc.steps.length} step${sc.steps.length !== 1 ? 's' : ''})</span></li>`
  ).join('\n');

  const sectionsHtml = screens.map((sc, si) => {
    // Navigate step screenshot (shows the screen on arrival)
    const navStep = sc.steps.find(s => s.type === 'navigate' && s.screenshot);
    const navScreenShot = navStep?.screenshot || '';

    // Split non-navigate steps into groups separated by 'snapshot' steps.
    // Each snapshot step starts a new group and carries that group's screenshot.
    const nonNavSteps = sc.steps.filter(s => s.type !== 'navigate');

    type Group = { screenshot: string; label: string; steps: (Step & { globalIdx: number })[]; fields?: CapturedField[] };
    const groups: Group[] = [];
    let current: Group = { screenshot: '', label: '', steps: [] };

    for (const step of nonNavSteps) {
      if (step.type === 'snapshot') {
        // Push whatever we've accumulated so far (even if empty — preserves order)
        groups.push(current);
        current = { screenshot: step.screenshot || '', label: step.fieldName || '', steps: [], fields: step.fields };
      } else {
        current.steps.push(step);
      }
    }
    groups.push(current);

    // If no snapshot steps were captured, use the nav screenshot for the single group
    if (groups.length === 1 && !groups[0].screenshot && navScreenShot) {
      groups[0].screenshot = navScreenShot;
    }

    const renderGroup = (g: Group, gIdx: number) => {
      const shotHtml = g.screenshot
        ? `<div style="margin:${gIdx === 0 ? '14px' : '24px'} 0 16px;">
             <img src="${g.screenshot}" alt="${escapeHtml(sc.title)} screenshot"
               style="width:100%;border:1px solid #ddd;border-radius:6px;display:block;box-shadow:0 2px 8px rgba(0,0,0,.08);" />
             <div style="font-size:11px;color:#aaa;margin-top:4px;text-align:right;">
               ${g.label ? `Tab: ${escapeHtml(g.label)}` : `Screenshot: ${escapeHtml(sc.title)}`}
             </div>
           </div>`
        : '';

      // Consolidated fields from a "Capture Fields" or "Capture Tab" snapshot step
      if (!g.steps.length && g.fields?.length) {
        const fieldRows = g.fields.map((f, fidx) => {
          const fa    = f.action || 'Enter';
          const fColor = fa === 'Display' ? '#e65100' : (fa === 'Enter' || fa === 'Select' || fa === 'Check') ? '#2e7d32' : '#1565c0';
          const fBg    = fa === 'Display' ? '#fff3e0' : (fa === 'Enter' || fa === 'Select' || fa === 'Check') ? '#e8f5e9' : '#e3f2fd';
          const badge  = `<span style="display:inline-block;padding:1px 7px;border-radius:3px;font-size:10px;font-weight:700;background:${fBg};color:${fColor}">${fa}</span>`;
          const explanation = fa === 'Display'
            ? `<strong>${escapeHtml(f.fieldName)}</strong>: ${escapeHtml(f.value)}`
            : fa === 'Enter'
            ? `In the <strong>${escapeHtml(f.fieldName)}</strong> field, enter <strong>${escapeHtml(f.value)}</strong>.`
            : fa === 'Select'
            ? `In the <strong>${escapeHtml(f.fieldName)}</strong> field, select <strong>${escapeHtml(f.value)}</strong>.`
            : `Click the <strong>${escapeHtml(f.fieldName)}</strong>.`;
          return `
          <tr>
            <td style="text-align:center;font-weight:700;font-size:15px;color:#444;white-space:nowrap;">${fidx + 1}</td>
            <td>${badge}</td>
            <td style="font-weight:600;color:#222;">${escapeHtml(f.fieldName)}</td>
            <td style="color:#555;">${f.value ? escapeHtml(f.value) : '—'}</td>
            <td style="line-height:1.5;">${explanation}</td>
          </tr>`;
        }).join('\n');
        const fieldTableHtml = `
        <table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;margin-bottom:8px;">
          <thead>
            <tr style="background:#f5f5f5;">
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:40px;">#</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:72px;">Type</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:170px;">Field / Element</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:140px;">Value</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;">Description</th>
            </tr>
          </thead>
          <tbody>${fieldRows}</tbody>
        </table>`;
        return shotHtml + fieldTableHtml;
      }

      if (!g.steps.length) return shotHtml;

      const rows = g.steps.map(s => {
        const explanation = s.type === 'input'
          ? `In the <strong>${escapeHtml(s.fieldName || 'field')}</strong> field, enter <strong>${escapeHtml(s.value || '')}</strong>.`
          : `Click the <strong>${escapeHtml(s.fieldName || s.description)}</strong>${s.value && s.value !== s.fieldName ? ' — <em>' + escapeHtml(s.value) + '</em>' : ''}.`;
        const badge = `<span style="display:inline-block;padding:1px 7px;border-radius:3px;font-size:10px;font-weight:700;background:${typeBg[s.type] || '#eee'};color:${typeColor[s.type] || '#333'}">${typeLabel[s.type] || s.type}</span>`;
        return `
        <tr>
          <td style="text-align:center;font-weight:700;font-size:15px;color:#444;white-space:nowrap;">${s.globalIdx}</td>
          <td>${badge}</td>
          <td style="font-weight:600;color:#222;">${escapeHtml(s.fieldName || s.description)}</td>
          <td style="color:#555;">${s.value ? escapeHtml(s.value) : '—'}</td>
          <td style="line-height:1.5;">${explanation}</td>
        </tr>`;
      }).join('\n');

      const tableHtml = `
        <table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;margin-bottom:8px;">
          <thead>
            <tr style="background:#f5f5f5;">
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:40px;">#</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:72px;">Type</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:170px;">Field / Element</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;width:140px;">Value Entered</th>
              <th style="padding:9px 12px;border-bottom:2px solid #ddd;">Step Explanation</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>`;

      return shotHtml + tableHtml;
    };

    const bodyHtml = groups.map((g, gi) => renderGroup(g, gi)).join('\n');
    const hasAnyFields = groups.some(g => g.steps.length > 0);
    const content = hasAnyFields
      ? bodyHtml
      : (navScreenShot
          ? `<div style="margin:14px 0 20px;"><img src="${navScreenShot}" style="width:100%;border:1px solid #ddd;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.08);" /></div>` : '')
        + '<p style="color:#aaa;font-size:12px;font-style:italic;">No field interactions recorded on this screen.</p>';

    return `
    <section id="screen-${si}" style="margin-bottom:50px;page-break-inside:avoid;">
      <h2 style="color:${REDWOOD};border-bottom:2px solid ${REDWOOD};padding-bottom:6px;margin-top:36px;">
        ${escapeHtml(sc.title)}
      </h2>
      ${content}
    </section>`;
  }).join('\n');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>User Manual — Oracle Fusion</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; max-width: 1000px; margin: 0 auto; padding: 30px; color: #333; }
  h1 { color: ${REDWOOD}; margin-bottom: 4px; }
  .meta { color: #888; font-size: 13px; margin-bottom: 20px; }
  .toc { background:#f9f9f9; border:1px solid #e0e0e0; border-radius:8px; padding:16px 24px; margin-bottom:32px; }
  .toc h3 { margin:0 0 10px; color:#444; font-size:14px; }
  .toc ol { margin:0; padding-left:20px; }
  .toc li { margin:4px 0; font-size:13px; }
  .toc a { color:${REDWOOD}; text-decoration:none; }
  .toc a:hover { text-decoration:underline; }
  tbody tr:hover td { background:#fafafa; }
  td, th { padding:9px 12px; border-bottom:1px solid #eee; vertical-align:top; text-align:left; }
  @media print {
    body { max-width:100%; padding:15px; }
    section { break-before:auto; }
    tr { break-inside:avoid; }
  }
</style>
</head>
<body>
<h1>&#128214; User Manual — Oracle Fusion</h1>
<div class="meta">Generated: ${date} &nbsp;|&nbsp; Total steps: ${steps.length} &nbsp;|&nbsp; Screens: ${screens.length}</div>

<div class="toc">
  <h3>&#128196; Table of Contents</h3>
  <ol>${toc}</ol>
</div>

${sectionsHtml}
</body>
</html>`;
};

const generateUATScript = (steps: Step[]): string => {
  const date = new Date().toLocaleString();
  const rows = steps.map((s, i) => `
    <tr>
      <td style="text-align:center;font-weight:600;">${i + 1}</td>
      <td>${escapeHtml(s.pageTitle || 'Oracle Fusion')}</td>
      <td>${escapeHtml(s.description)}</td>
      <td style="color:#555;font-style:italic;">&nbsp;</td>
      <td style="white-space:nowrap;">
        <label style="cursor:pointer;margin-right:10px;">
          <input type="radio" name="res${i}" value="pass" style="accent-color:#2e7d32;"> <span style="color:#2e7d32;font-weight:600;">Pass</span>
        </label>
        <label style="cursor:pointer;">
          <input type="radio" name="res${i}" value="fail" style="accent-color:#c62828;"> <span style="color:#c62828;font-weight:600;">Fail</span>
        </label>
      </td>
      <td><input type="text" placeholder="Add comments…" style="width:100%;border:1px solid #ccc;border-radius:4px;padding:4px 6px;font-size:12px;box-sizing:border-box;"></td>
    </tr>`).join('\n');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>UAT Script — Oracle Fusion</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; padding: 30px; color: #333; }
  h1 { color: ${REDWOOD}; }
  .meta { color: #888; font-size: 13px; margin-bottom: 20px; }
  .tester-row { display:flex; gap:30px; margin-bottom:20px; }
  .tester-row label { font-size:13px; font-weight:600; }
  .tester-row input { border:none; border-bottom:1px solid #999; width:180px; padding:2px 4px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  thead tr { background: ${REDWOOD}; color: #fff; }
  th { padding: 10px 12px; text-align: left; font-weight: 600; }
  td { padding: 9px 12px; border-bottom: 1px solid #e0e0e0; vertical-align: top; }
  tr:nth-child(even) td { background: #f8f8f8; }
  .no-print { display: block; }
  @media print {
    .no-print { display: none; }
    table { font-size: 11px; }
    body { padding: 10px; }
  }
</style>
</head>
<body>
<h1>&#9989; UAT Test Script — Oracle Fusion</h1>
<div class="meta">Generated: ${date} &nbsp;|&nbsp; Total test steps: ${steps.length}</div>
<div class="tester-row">
  <div><label>Tester Name: </label><input type="text"></div>
  <div><label>Test Date: </label><input type="text"></div>
  <div><label>Version/Release: </label><input type="text"></div>
  <div><label>Overall Status: </label><input type="text"></div>
</div>
<table>
  <thead>
    <tr>
      <th style="width:42px;">#</th>
      <th style="width:140px;">Screen</th>
      <th>Action / Step Description</th>
      <th style="width:160px;">Expected Result</th>
      <th style="width:120px;">Result</th>
      <th style="width:150px;">Comments</th>
    </tr>
  </thead>
  <tbody>
    ${rows}
  </tbody>
</table>
<br>
<button class="no-print" onclick="window.print()" style="background:${REDWOOD};color:#fff;border:none;padding:10px 24px;border-radius:6px;font-size:14px;cursor:pointer;">&#128438; Print / Save as PDF</button>
</body>
</html>`;
};

// ── Word export helpers ───────────────────────────────────────────────────────
async function loadImgForWord(dataUrl: string): Promise<{ data: Uint8Array; width: number; height: number } | null> {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      try {
        const base64 = dataUrl.split(',')[1];
        const binary = atob(base64);
        const data = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) data[i] = binary.charCodeAt(i);
        const maxW = 580;
        const scale = img.width > maxW ? maxW / img.width : 1;
        resolve({ data, width: Math.round(img.width * scale), height: Math.round(img.height * scale) });
      } catch { resolve(null); }
    };
    img.onerror = () => resolve(null);
    img.src = dataUrl;
  });
}

function groupByScreen(steps: Step[]) {
  const screens: { title: string; steps: Step[] }[] = [];
  for (const s of steps) {
    const title = s.pageTitle || 'Oracle Fusion';
    let sc = screens.find(x => x.title === title);
    if (!sc) { sc = { title, steps: [] }; screens.push(sc); }
    sc.steps.push(s);
  }
  return screens;
}

async function buildWordManual(steps: Step[]): Promise<Blob> {
  const screens = groupByScreen(steps);
  const docChildren: any[] = [
    new Paragraph({ text: 'User Manual — Oracle Fusion', heading: HeadingLevel.TITLE }),
    new Paragraph({
      children: [new TextRun({ text: `Generated: ${new Date().toLocaleString()}  |  Steps: ${steps.length}`, color: '888888', size: 20 })],
    }),
    new Paragraph({ text: '' }),
  ];

  let globalIdx = 0;
  for (const screen of screens) {
    docChildren.push(new Paragraph({
      text: screen.title,
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 320, after: 160 },
    }));

    const shotStep = screen.steps.find(s => s.screenshot);
    if (shotStep?.screenshot) {
      const img = await loadImgForWord(shotStep.screenshot);
      if (img && img.data.length > 0) {
        docChildren.push(new Paragraph({
          children: [new ImageRun({ type: 'png', data: img.data, transformation: { width: img.width, height: img.height } })],
          spacing: { after: 160 },
        }));
      }
    }

    const actionSteps = screen.steps.filter(s => s.type !== 'navigate' && s.type !== 'snapshot');
    if (actionSteps.length > 0) {
      const headerRow = new TableRow({
        tableHeader: true,
        children: ['#', 'Type', 'Field / Element', 'Value', 'Description'].map(txt =>
          new TableCell({
            shading: { type: ShadingType.SOLID, fill: 'C74634' },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: txt, bold: true, color: 'FFFFFF', size: 20 })] })],
          })
        ),
      });
      const dataRows = actionSteps.map(s => {
        globalIdx++;
        return new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(globalIdx), size: 20 })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.type, size: 20 })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.fieldName || '', size: 20 })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.value || '', size: 20 })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.description || '', size: 20 })] })] }),
          ],
        });
      });
      docChildren.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] }));
    }
    docChildren.push(new Paragraph({ text: '' }));
  }

  const doc = new Document({ sections: [{ children: docChildren }] });
  return Packer.toBlob(doc);
}

async function buildWordUAT(steps: Step[]): Promise<Blob> {
  const headerRow = new TableRow({
    tableHeader: true,
    children: ['#', 'Screen', 'Step Description', 'Expected Result', 'Result', 'Comments'].map(txt =>
      new TableCell({
        shading: { type: ShadingType.SOLID, fill: 'C74634' },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: txt, bold: true, color: 'FFFFFF', size: 20 })] })],
      })
    ),
  });
  const dataRows = steps.map((s, i) =>
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(i + 1), size: 20 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.pageTitle || 'Oracle Fusion', size: 20 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.description || '', size: 20 })] })] }),
        new TableCell({ children: [new Paragraph({ text: '' })] }),
        new TableCell({ children: [new Paragraph({ text: '' })] }),
        new TableCell({ children: [new Paragraph({ text: '' })] }),
      ],
    })
  );

  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ text: 'UAT Test Script — Oracle Fusion', heading: HeadingLevel.TITLE }),
        new Paragraph({ children: [new TextRun({ text: `Generated: ${new Date().toLocaleString()}  |  Steps: ${steps.length}`, color: '888888', size: 20 })] }),
        new Paragraph({ text: '' }),
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] }),
      ],
    }],
  });
  return Packer.toBlob(doc);
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

const isElectron = () => !!(window as any).electronAPI?.isElectron;
const formatTime = (s: number) =>
  `${Math.floor(s / 60).toString().padStart(2, '0')}:${(s % 60).toString().padStart(2, '0')}`;

const STEP_COLORS: Record<string, string> = { click: 'blue', input: 'green', navigate: 'orange', snapshot: 'purple' };

const OracleFusion: React.FC = () => {
  const navigate = useNavigate();
  const webviewRef = useRef<any>(null);
  const [url, setUrl] = useState('https://iacney-test.fa.ocs.oraclecloud.com/');
  const [inputUrl, setInputUrl] = useState('https://iacney-test.fa.ocs.oraclecloud.com/');
  const [canGoBack, setCanGoBack] = useState(false);
  const [canGoFwd, setCanGoFwd] = useState(false);
  const [loading, setLoading] = useState(false);

  // --- Step tracking ---
  const [tracking, setTracking] = useState(false);
  const [steps, setSteps] = useState<Step[]>([]);
  const [showPanel, setShowPanel] = useState(false);
  const trackingRef = useRef(false);
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const screenshotIntervalRef = useRef<ReturnType<typeof setInterval> | null>(null);
  // Tracks the pageTitle of the screen currently being interacted with
  const currentPageTitleRef = useRef<string>('');
  // Rolling screenshot updated every 2s while tracking — applied to steps on navigate away
  const lastScreenshotRef = useRef<string>('');

  // --- Auto screenshot toggle ---
  const [autoShot, setAutoShot] = useState(true);
  const autoShotRef = useRef(true);

  // --- Annotation ---
  const [annotatingStepId, setAnnotatingStepId] = useState<string | null>(null);

  // --- Preview ---
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewUrl, setPreviewUrl] = useState<string>('');

  // --- Step drag-and-drop ---
  const [dragIdx, setDragIdx]         = useState<number | null>(null);
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null);

  // --- Saved credentials ---
  const [credsModalOpen, setCredsModalOpen] = useState(false);
  const [hasSavedCreds, setHasSavedCreds] = useState(false);
  const [credsForm] = Form.useForm();

  // --- Screen recording ---
  const [recording, setRecording] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const mediaRecorder = useRef<MediaRecorder | null>(null);
  const chunks = useRef<Blob[]>([]);
  const streamRef = useRef<MediaStream | null>(null);
  const elapsedRef = useRef(0);

  // Inject tracking script and add navigate step on each page load
  const injectTracking = async (wv: any) => {
    try {
      await wv.executeJavaScript(INJECT_SCRIPT);
    } catch (_) {}
  };

  useEffect(() => {
    const wv = webviewRef.current;
    if (!wv || !isElectron()) return;

    // Helper: capture webview screenshot as data URL
    const captureShot = async (): Promise<string> => {
      try {
        const shot = await wv.capturePage();
        return shot?.resize?.({ width: 900 })?.toDataURL?.() || shot?.toDataURL?.() || '';
      } catch (_) { return ''; }
    };

    // BEFORE leaving a screen: capture it while fields are still filled in.
    const onWillNavigate = async () => {
      if (!trackingRef.current || !autoShotRef.current) return;
      const leavingTitle = currentPageTitleRef.current;
      const dataUrl = await captureShot();
      if (dataUrl && leavingTitle) {
        setSteps(prev => prev.map(s =>
          !s.screenshot && s.pageTitle === leavingTitle ? { ...s, screenshot: dataUrl } : s
        ));
      }
    };

    // AFTER arriving at a new screen: add navigate step + capture screenshot of new screen.
    const onLoad = async () => {
      setLoading(false);
      const currentUrl = wv.getURL ? wv.getURL() : url;
      setInputUrl(currentUrl);
      setCanGoBack(wv.canGoBack?.() ?? false);
      setCanGoFwd(wv.canGoForward?.() ?? false);

      if (!trackingRef.current) return;

      await injectTracking(wv);
      try {
        const title = await wv.executeJavaScript('document.title');
        const navUrl = wv.getURL?.() || currentUrl;
        const navId = Date.now() + Math.random() + '';
        currentPageTitleRef.current = title || navUrl;

        // Add navigate step (screenshot filled in shortly after page settles)
        setSteps(prev => [...prev, {
          id: navId,
          type: 'navigate',
          fieldName: title || navUrl,
          action: 'Navigate',
          value: '',
          description: `Navigate to: ${title || navUrl}`,
          url: navUrl,
          pageTitle: title || navUrl,
          timestamp: Date.now(),
        }]);

        // Capture screenshot of the new screen once it has fully rendered
        if (autoShotRef.current) {
          setTimeout(async () => {
            const dataUrl = await captureShot();
            if (dataUrl) {
              setSteps(prev => prev.map(s => s.id === navId ? { ...s, screenshot: dataUrl } : s));
            }
          }, 1000);
        }
      } catch (_) {}
    };

    const onStart = () => setLoading(true);
    const onFail  = () => setLoading(false);

    wv.addEventListener('will-navigate',    onWillNavigate);
    wv.addEventListener('did-finish-load',  onLoad);
    wv.addEventListener('did-start-loading', onStart);
    wv.addEventListener('did-fail-load',    onFail);
    return () => {
      wv.removeEventListener('will-navigate',    onWillNavigate);
      wv.removeEventListener('did-finish-load',  onLoad);
      wv.removeEventListener('did-start-loading', onStart);
      wv.removeEventListener('did-fail-load',    onFail);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Cleanup on unmount
  useEffect(() => () => {
    if (timerRef.current) clearInterval(timerRef.current);
    if (pollRef.current) clearInterval(pollRef.current);
    if (screenshotIntervalRef.current) clearInterval(screenshotIntervalRef.current);
    streamRef.current?.getTracks().forEach(t => t.stop());
  }, []);

  const navigate_to = (dest: string) => {
    const wv = webviewRef.current;
    setUrl(dest);
    setInputUrl(dest);
    if (wv) wv.src = dest;
  };

  const handleUrlInput = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      let dest = inputUrl.trim();
      if (dest && !dest.startsWith('http')) dest = 'https://' + dest;
      navigate_to(dest);
    }
  };

  // ---- Step Tracking ----
  const startTracking = async () => {
    const wv = webviewRef.current;
    if (!wv) { message.warning('WebView not ready'); return; }
    setTracking(true);
    setSteps([]);
    setShowPanel(true);
    trackingRef.current = true;
    lastScreenshotRef.current = '';
    // Seed current page title
    try { currentPageTitleRef.current = await wv.executeJavaScript('document.title'); } catch (_) {}
    await injectTracking(wv);

    // ── Rolling screenshot every 2s (auto mode only) ──
    if (autoShotRef.current) {
      screenshotIntervalRef.current = setInterval(async () => {
        if (!trackingRef.current || !autoShotRef.current) return;
        try {
          const shot = await wv.capturePage();
          const url = shot?.resize?.({ width: 960 })?.toDataURL?.() || shot?.toDataURL?.() || '';
          if (url) lastScreenshotRef.current = url;
        } catch (_) {}
      }, 2000);
    }

    // ── Poll every 800ms: collect steps + detect SPA navigation by title change ──
    pollRef.current = setInterval(async () => {
      if (!trackingRef.current) return;
      try {
        const pollData = await wv.executeJavaScript(
          '({steps:(window.__reactErpSteps||[]).splice(0), title:document.title, url:location.href})'
        );
        const { steps: rawSteps, title: liveTitle, url: liveUrl } = pollData as any;

        // ── SPA navigation detected (Oracle Fusion ADF navigates without full reload) ──
        if (liveTitle && liveTitle !== currentPageTitleRef.current) {
          const oldTitle = currentPageTitleRef.current;
          const navId = Date.now() + Math.random() + '';

          // 1. Apply last rolling screenshot to all steps from the OLD screen
          if (autoShotRef.current && lastScreenshotRef.current && oldTitle) {
            const snap = lastScreenshotRef.current;
            setSteps(prev => prev.map(s =>
              !s.screenshot && s.pageTitle === oldTitle ? { ...s, screenshot: snap } : s
            ));
          }

          currentPageTitleRef.current = liveTitle;
          lastScreenshotRef.current = '';

          // 2. Add navigate step
          setSteps(prev => [...prev, {
            id: navId,
            type: 'navigate' as const,
            fieldName: liveTitle,
            action: 'Navigate',
            value: '',
            description: `Navigate to: ${liveTitle}`,
            url: liveUrl,
            pageTitle: liveTitle,
            timestamp: Date.now(),
          }]);

          // 3. Re-inject tracker into new SPA page
          await injectTracking(wv);

          // 4. Capture screenshot of new screen once settled (1s) — auto mode only
          if (autoShotRef.current) setTimeout(async () => {
            try {
              const shot = await wv.capturePage();
              const dataUrl = shot?.resize?.({ width: 960 })?.toDataURL?.() || shot?.toDataURL?.() || '';
              if (dataUrl) {
                lastScreenshotRef.current = dataUrl;
                setSteps(prev => prev.map(s => s.id === navId ? { ...s, screenshot: dataUrl } : s));
              }
            } catch (_) {}
          }, 1000);
        }

        // ── Collect field/click steps ──
        if (rawSteps?.length) {
          const enriched: Step[] = rawSteps.map((s: any) => ({
            fieldName: s.fieldName || s.description || '',
            action: s.action || s.type || '',
            value: s.value || '',
            ...s,
            id: Date.now() + Math.random() + '',
          }));
          setSteps(prev => [...prev, ...enriched]);
        }
      } catch (_) {}
    }, 800);

    message.success('Step tracking started — perform your Oracle Fusion workflow');
  };

  const stopTracking = async () => {
    setTracking(false);
    trackingRef.current = false;
    if (pollRef.current) { clearInterval(pollRef.current); pollRef.current = null; }
    if (screenshotIntervalRef.current) { clearInterval(screenshotIntervalRef.current); screenshotIntervalRef.current = null; }
    // Apply last rolling screenshot to any steps still missing one
    if (lastScreenshotRef.current) {
      const snap = lastScreenshotRef.current;
      const title = currentPageTitleRef.current;
      setSteps(prev => prev.map(s =>
        !s.screenshot && s.pageTitle === title ? { ...s, screenshot: snap } : s
      ));
    }
    message.info(`Tracking stopped — ${steps.length} steps captured`);
  };

  // ---- Saved Credentials ----
  // Check on mount whether credentials are already stored
  useEffect(() => {
    (window as any).electronAPI?.getFusionCredentials?.().then((creds: any) => {
      setHasSavedCreds(!!creds?.username);
      if (creds?.username) credsForm.setFieldsValue({ username: creds.username, password: creds.password });
    });
  }, []);

  const openCredsModal = async () => {
    // Pre-fill form with saved values
    const creds = await (window as any).electronAPI?.getFusionCredentials?.();
    if (creds) credsForm.setFieldsValue({ username: creds.username, password: creds.password });
    setCredsModalOpen(true);
  };

  const handleSaveCreds = async () => {
    const values = await credsForm.validateFields();
    const result = await (window as any).electronAPI?.saveFusionCredentials?.(values.username, values.password);
    if (result?.success) {
      setHasSavedCreds(true);
      setCredsModalOpen(false);
      message.success('Credentials saved — use the key button to auto-login');
    } else {
      message.error('Failed to save credentials');
    }
  };

  const handleClearCreds = async () => {
    await (window as any).electronAPI?.clearFusionCredentials?.();
    credsForm.resetFields();
    setHasSavedCreds(false);
    message.info('Credentials cleared');
  };

  // Inject credentials into the Oracle Fusion login page
  const handleAutoLogin = async () => {
    const wv = webviewRef.current;
    if (!wv) { message.warning('WebView not ready'); return; }
    const creds = await (window as any).electronAPI?.getFusionCredentials?.();
    if (!creds) { message.warning('No saved credentials — click the settings icon to set up'); return; }

    const script = `
      (function() {
        function fillField(el, val) {
          el.focus();
          // Clear first
          el.value = '';
          el.dispatchEvent(new Event('input', {bubbles:true}));
          // Set via native setter (bypasses React/ADF synthetic events)
          try {
            var proto = el.tagName === 'INPUT' ? window.HTMLInputElement.prototype : window.HTMLTextAreaElement.prototype;
            var setter = Object.getOwnPropertyDescriptor(proto, 'value');
            if (setter && setter.set) setter.set.call(el, val);
            else el.value = val;
          } catch(e) { el.value = val; }
          // Fire all events ADF/OAM listens to
          el.dispatchEvent(new Event('input',  {bubbles:true, cancelable:true}));
          el.dispatchEvent(new Event('change', {bubbles:true, cancelable:true}));
          el.dispatchEvent(new KeyboardEvent('keydown', {bubbles:true, cancelable:true, key:'Tab'}));
          el.dispatchEvent(new KeyboardEvent('keyup',   {bubbles:true, cancelable:true, key:'Tab'}));
          el.blur();
        }

        var allInputs = Array.from(document.querySelectorAll('input'));

        // Username: prefer named/id'd fields, then first text-type
        var userField =
          document.querySelector('input[name="userid"]') ||
          document.querySelector('input[id="userid"]') ||
          document.querySelector('input[autocomplete="username"]') ||
          document.querySelector('input[name="username"]') ||
          document.querySelector('input[id="username"]') ||
          allInputs.find(function(el) {
            var t = (el.type || 'text').toLowerCase();
            return t !== 'password' && t !== 'hidden' && t !== 'submit' &&
                   t !== 'button'   && t !== 'checkbox' && t !== 'radio' && t !== 'file';
          });

        // Password: first password-type
        var passField =
          document.querySelector('input[type="password"]') ||
          document.querySelector('input[name="password"]');

        if (!userField && !passField) {
          return 'no-inputs | found: ' + allInputs.map(function(i){
            return (i.type||'text') + '|' + i.name + '|' + i.id + '|' + i.placeholder;
          }).join(' :: ');
        }

        if (userField) fillField(userField, ${JSON.stringify(creds.username)});
        if (passField) fillField(passField, ${JSON.stringify(creds.password)});

        // Sign In button
        var allBtns = Array.from(document.querySelectorAll('button, input[type="submit"], a[role="button"]'));
        var loginBtn = allBtns.find(function(b) {
          var t = (b.textContent || b.value || b.getAttribute('aria-label') || '').toLowerCase().trim();
          return t === 'sign in' || t === 'login' || t === 'log in' || t === 'submit' || t === 'ok' || t === 'next';
        }) || document.querySelector('input[type="submit"]') || document.querySelector('#btnActive');

        if (loginBtn) { loginBtn.click(); return 'ok-submitted'; }
        return 'ok-filled';
      })();
    `;

    try {
      const result = await wv.executeJavaScript(script);
      if (typeof result === 'string' && result.startsWith('ok-submitted')) {
        message.success('Credentials filled and Sign In clicked');
      } else if (typeof result === 'string' && result.startsWith('ok-filled')) {
        message.success('Credentials filled — click Sign In to continue');
      } else if (typeof result === 'string' && result.startsWith('no-inputs')) {
        message.warning(`No login fields found on this page. Detected: ${result.replace('no-inputs | found: ', '')}`, 10);
      } else {
        message.warning('Result: ' + String(result));
      }
    } catch (e: any) {
      message.error('Auto-login error: ' + e.message);
    }
  };

  // ---- Screen Recording ----
  const startRecording = async () => {
    if (!isElectron()) { message.warning('Recording only works in the desktop app'); return; }
    try {
      const sources = await (window as any).electronAPI.getScreenSources();
      const src = sources.find((s: any) => s.name.includes('Screen') || s.name.includes('Entire')) || sources[0];
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: false,
        video: { mandatory: { chromeMediaSource: 'desktop', chromeMediaSourceId: src.id, minWidth: 1280, maxWidth: 1920, minHeight: 720, maxHeight: 1080 } } as any,
      });
      streamRef.current = stream;
      chunks.current = [];
      const rec = new MediaRecorder(stream, { mimeType: 'video/webm;codecs=vp8' });
      rec.ondataavailable = (e) => { if (e.data.size > 0) chunks.current.push(e.data); };
      rec.start(1000);
      mediaRecorder.current = rec;
      elapsedRef.current = 0;
      setElapsed(0);
      setRecording(true);
      timerRef.current = setInterval(() => { elapsedRef.current++; setElapsed(elapsedRef.current); }, 1000);
      message.success('Recording started');
    } catch (e: any) {
      message.error('Could not start recording: ' + e.message);
    }
  };

  const stopRecording = () => {
    if (!mediaRecorder.current) return;
    mediaRecorder.current.stop();
    streamRef.current?.getTracks().forEach(t => t.stop());
    if (timerRef.current) clearInterval(timerRef.current);
    setRecording(false);
    mediaRecorder.current.onstop = async () => {
      const blob = new Blob(chunks.current, { type: 'video/webm' });
      const buf = await blob.arrayBuffer();
      const now = new Date();
      const stamp = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}`;
      const result = await (window as any).electronAPI.saveRecording(
        Array.from(new Uint8Array(buf)),
        { title: `Oracle Fusion ${stamp}`, description: 'Recorded from Oracle Fusion WebView', category: 'Oracle Fusion', defaultName: `OracleFusion_${stamp}.webm`, duration: elapsedRef.current }
      );
      if (result?.success) {
        message.success(<span>Saved! <a onClick={() => navigate('/training')} style={{ textDecoration: 'underline', cursor: 'pointer' }}>View in Training Library</a></span>, 6);
      }
    };
  };

  // ---- Capture Fields Snapshot (ONE consolidated step with all fields) ----
  const handleCaptureFields = async () => {
    const wv = webviewRef.current;
    if (!wv) return;
    try {
      await injectTracking(wv);
      const raw = await wv.executeJavaScript('window.__reactErpCaptureFields ? window.__reactErpCaptureFields() : "[]"');
      const captured: any[] = JSON.parse(raw || '[]');
      if (!captured.length) { message.warning('No fields found — try Capture Tab to save a screenshot instead'); return; }

      const currentTitle = currentPageTitleRef.current ||
        await wv.executeJavaScript('document.title');

      // Capture screenshot of current state
      const shot = await wv.capturePage();
      const dataUrl = shot?.resize?.({ width: 960 })?.toDataURL?.() || shot?.toDataURL?.() || '';
      if (dataUrl) lastScreenshotRef.current = dataUrl;

      const now = Date.now();
      // ONE consolidated snapshot step carrying all fields
      const snapshotStep: Step = {
        id: now + 'fields' + Math.random(),
        type: 'snapshot',
        fieldName: currentTitle,
        action: 'Snapshot',
        value: `${captured.length} fields`,
        description: `Fields captured: ${currentTitle}`,
        url: wv.getURL?.() || '',
        pageTitle: currentTitle,
        timestamp: now,
        screenshot: dataUrl || undefined,
        fields: captured.map((s: any) => ({
          fieldName: s.fieldName || '',
          action: s.action || '',
          value: s.value || '',
          description: s.description || '',
        })),
      };

      // APPEND — never remove existing steps or screenshots
      setSteps(prev => {
        // Apply screenshot to most recent navigate step for this screen if it has none
        const withShot = dataUrl
          ? prev.map(s =>
              s.pageTitle === currentTitle && s.type === 'navigate' && !s.screenshot
                ? { ...s, screenshot: dataUrl } : s
            )
          : prev;
        return [...withShot, snapshotStep];
      });

      message.success(`Captured ${captured.length} fields from "${currentTitle}"`);
    } catch (e: any) {
      message.error('Capture failed: ' + e.message);
    }
  };

  // ---- Capture Screenshot + Fields (multi-tab: appends, never replaces) ----
  const handleCaptureTab = async () => {
    const wv = webviewRef.current;
    if (!wv) return;
    try {
      // Always inject the script first so __reactErpCaptureFields is available
      await injectTracking(wv);

      const raw = await wv.executeJavaScript('window.__reactErpCaptureFields ? window.__reactErpCaptureFields() : "[]"');
      const captured: any[] = JSON.parse(raw || '[]');

      const currentTitle = currentPageTitleRef.current ||
        await wv.executeJavaScript('document.title');

      // Always capture screenshot — even if no fields found
      const shot = await wv.capturePage();
      const dataUrl = shot?.resize?.({ width: 960 })?.toDataURL?.() || shot?.toDataURL?.() || '';

      const now = Date.now();
      // ONE snapshot step carrying screenshot + all fields
      const snapshotStep: Step = {
        id: now + 'snap' + Math.random(),
        type: 'snapshot',
        fieldName: currentTitle,
        action: 'Snapshot',
        value: captured.length ? `${captured.length} fields` : '',
        description: `Screenshot: ${currentTitle}`,
        url: wv.getURL?.() || '',
        pageTitle: currentTitle,
        timestamp: now,
        screenshot: dataUrl,
        fields: captured.length
          ? captured.map((s: any) => ({
              fieldName: s.fieldName || '',
              action: s.action || '',
              value: s.value || '',
              description: s.description || '',
            }))
          : undefined,
      };

      // Always APPEND — never replace — so every tab is preserved independently
      setSteps(prev => [...prev, snapshotStep]);

      if (captured.length > 0) {
        message.success(`Captured screenshot + ${captured.length} fields from "${currentTitle}"`);
      } else {
        message.info(`Screenshot captured from "${currentTitle}" (no fillable fields detected — read-only view)`);
      }
    } catch (e: any) {
      message.error('Capture failed: ' + e.message);
    }
  };

  // ---- Document Generation ----
  const handleGenerateManual = () => {
    if (!steps.length) { message.warning('No steps recorded yet'); return; }
    const html = generateUserManual(steps);
    const ts = new Date().toISOString().slice(0, 10);
    downloadHtml(html, `UserManual_OracleFusion_${ts}.html`);
    message.success('User Manual downloaded');
  };

  const handleGenerateUAT = () => {
    if (!steps.length) { message.warning('No steps recorded yet'); return; }
    const html = generateUATScript(steps);
    const ts = new Date().toISOString().slice(0, 10);
    downloadHtml(html, `UATScript_OracleFusion_${ts}.html`);
    message.success('UAT Script downloaded');
  };

  const handleGenerateWordManual = async () => {
    if (!steps.length) { message.warning('No steps recorded yet'); return; }
    const key = 'word-manual';
    message.loading({ content: 'Building Word document…', key });
    try {
      const blob = await buildWordManual(steps);
      downloadBlob(blob, `UserManual_OracleFusion_${new Date().toISOString().slice(0, 10)}.docx`);
      message.success({ content: 'Word document downloaded', key });
    } catch {
      message.error({ content: 'Failed to generate Word document', key });
    }
  };

  const handleGenerateWordUAT = async () => {
    if (!steps.length) { message.warning('No steps recorded yet'); return; }
    const key = 'word-uat';
    message.loading({ content: 'Building Word document…', key });
    try {
      const blob = await buildWordUAT(steps);
      downloadBlob(blob, `UATScript_OracleFusion_${new Date().toISOString().slice(0, 10)}.docx`);
      message.success({ content: 'Word document downloaded', key });
    } catch {
      message.error({ content: 'Failed to generate Word document', key });
    }
  };

  // ---- Preview (Blob URL avoids srcDoc size limit with many screenshots) ----
  const handleOpenPreview = () => {
    if (!steps.length) return;
    const html = generateUserManual(steps);
    const blob = new Blob([html], { type: 'text/html' });
    const blobUrl = URL.createObjectURL(blob);
    setPreviewUrl(blobUrl);
    setPreviewOpen(true);
  };

  const handleClosePreview = () => {
    setPreviewOpen(false);
    if (previewUrl) { URL.revokeObjectURL(previewUrl); setPreviewUrl(''); }
  };

  // ---- Manual screenshot capture (used when autoShot=false) ----
  const handleManualCapture = async () => {
    const wv = webviewRef.current;
    if (!wv) { message.warning('WebView not ready'); return; }
    try {
      const shot = await wv.capturePage();
      const dataUrl = shot?.resize?.({ width: 960 })?.toDataURL?.() || shot?.toDataURL?.() || '';
      if (!dataUrl) { message.warning('Could not capture screenshot'); return; }
      const title = currentPageTitleRef.current;
      setSteps(prev => prev.map(s =>
        s.pageTitle === title && !s.screenshot ? { ...s, screenshot: dataUrl } : s
      ));
      lastScreenshotRef.current = dataUrl;
      message.success('Screenshot captured');
    } catch (e: any) { message.error('Capture failed: ' + e.message); }
  };

  // ---- Insert screenshot between steps ----
  const handleInsertScreenshot = async (afterIndex: number) => {
    const wv = webviewRef.current;
    if (!wv) { message.warning('WebView not ready'); return; }
    try {
      const shot = await wv.capturePage();
      const dataUrl = shot?.resize?.({ width: 960 })?.toDataURL?.() || shot?.toDataURL?.() || '';
      if (!dataUrl) { message.warning('Could not capture screenshot'); return; }
      const title = await wv.executeJavaScript('document.title').catch(() => currentPageTitleRef.current);
      const newStep: Step = {
        id: Date.now() + Math.random() + '',
        type: 'snapshot',
        fieldName: title || 'Screenshot',
        action: 'Snapshot',
        value: '',
        description: `Screenshot: ${title || 'Oracle Fusion'}`,
        url: wv.getURL?.() || '',
        pageTitle: title || currentPageTitleRef.current,
        timestamp: Date.now(),
        screenshot: dataUrl,
      };
      setSteps(prev => {
        const arr = [...prev];
        arr.splice(afterIndex + 1, 0, newStep);
        return arr;
      });
      message.success('Screenshot inserted');
    } catch (e: any) { message.error('Failed: ' + e.message); }
  };

  // ---- Drag-and-drop step reordering ----
  const handleDragStart = (i: number) => setDragIdx(i);
  const handleDragOver  = (e: React.DragEvent, i: number) => { e.preventDefault(); setDragOverIdx(i); };
  const handleDrop      = (i: number) => {
    if (dragIdx === null || dragIdx === i) { setDragIdx(null); setDragOverIdx(null); return; }
    setSteps(prev => {
      const arr = [...prev];
      const [moved] = arr.splice(dragIdx, 1);
      arr.splice(i, 0, moved);
      return arr;
    });
    setDragIdx(null);
    setDragOverIdx(null);
  };

  return (
    <>
    <Layout style={{ height: 'calc(100vh - 64px)', display: 'flex', flexDirection: 'column', background: '#1a1a2e', overflow: 'hidden' }}>
      <style>{`
        @keyframes pulse-rec { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.35;transform:scale(1.35)} }
        @keyframes pulse-trk { 0%,100%{opacity:1} 50%{opacity:.4} }
        .wv-scroll::-webkit-scrollbar{width:6px} .wv-scroll::-webkit-scrollbar-thumb{background:#555;border-radius:3px}
      `}</style>

      {/* Breadcrumb */}
      <div style={{ padding: '8px 20px', background: '#fff', borderBottom: '1px solid #e5e5e5', flexShrink: 0 }}>
        <Breadcrumb items={[
          { title: <Link to="/home"><HomeOutlined /> Home</Link> },
          { title: 'Oracle Fusion' },
        ]} />
      </div>

      {/* Browser Toolbar */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '7px 10px', background: '#2b2b3b', flexShrink: 0, flexWrap: 'nowrap' }}>
        <Tooltip title="Back">
          <Button size="small" icon={<ArrowLeftOutlined />} disabled={!canGoBack}
            onClick={() => webviewRef.current?.goBack?.()}
            style={{ background: '#444', border: 'none', color: '#fff' }} />
        </Tooltip>
        <Tooltip title="Forward">
          <Button size="small" icon={<ArrowRightOutlined />} disabled={!canGoFwd}
            onClick={() => webviewRef.current?.goForward?.()}
            style={{ background: '#444', border: 'none', color: '#fff' }} />
        </Tooltip>
        <Tooltip title="Reload">
          <Button size="small" icon={<ReloadOutlined spin={loading} />}
            onClick={() => webviewRef.current?.reload?.()}
            style={{ background: '#444', border: 'none', color: '#fff' }} />
        </Tooltip>

        <Select size="small" value={undefined} placeholder="Quick links" onChange={navigate_to} style={{ width: 170 }}>
          {FUSION_URLS.map(f => <Option key={f.value} value={f.value}>{f.label}</Option>)}
        </Select>

        <Input
          size="small"
          value={inputUrl}
          onChange={e => setInputUrl(e.target.value)}
          onKeyDown={handleUrlInput}
          prefix={<LockOutlined style={{ color: '#4caf50', fontSize: 11 }} />}
          style={{ flex: 1, background: '#3a3a4a', border: '1px solid #555', color: '#fff', minWidth: 100 }}
        />

        {/* Track Steps */}
        {tracking ? (
          <>
            <Tooltip title="Scan fields and replace noisy clicks with clean data (replaces existing captures for this screen)">
              <Button size="small" icon={<FileTextOutlined />} onClick={handleCaptureFields}
                style={{ background: '#1b5e20', border: 'none', color: '#fff', fontWeight: 600 }}>
                Capture Fields
              </Button>
            </Tooltip>
            <Tooltip title="Capture screenshot + fields and APPEND — use this when switching between tab pages on the same screen">
              <Button size="small" icon={<CameraOutlined />} onClick={handleCaptureTab}
                style={{ background: '#4a148c', border: 'none', color: '#fff', fontWeight: 600 }}>
                Capture Tab
              </Button>
            </Tooltip>
            <Badge count={steps.length} size="small" offset={[-4, 0]}>
              <Button size="small" onClick={stopTracking}
                style={{ background: '#7b2d00', border: 'none', color: '#fff', display: 'flex', alignItems: 'center', gap: 4 }}>
                <span style={{ width: 7, height: 7, borderRadius: '50%', background: '#ff6b35', display: 'inline-block', animation: 'pulse-trk 1s infinite' }} />
                Stop Tracking
              </Button>
            </Badge>
          </>
        ) : (
          <Tooltip title="Track navigation to generate User Manual / UAT Script">
            <Button size="small" icon={<AimOutlined />} onClick={() => { autoShotRef.current = autoShot; startTracking(); }}
              style={{ background: '#1565c0', border: 'none', color: '#fff' }}>
              Track Steps
            </Button>
          </Tooltip>
        )}

        {/* Auto screenshot toggle */}
        <Tooltip title={autoShot ? 'Auto screenshot ON — screenshots captured automatically' : 'Auto screenshot OFF — click 📷 to capture manually'} placement="bottom">
          <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
            <span style={{ fontSize: 10, color: '#aaa' }}>Auto 📷</span>
            <Switch
              size="small"
              checked={autoShot}
              onChange={v => { setAutoShot(v); autoShotRef.current = v; }}
              style={{ background: autoShot ? '#1565c0' : '#555' }}
            />
          </div>
        </Tooltip>

        {/* Manual capture button — visible when autoShot is off and tracking */}
        {tracking && !autoShot && (
          <Tooltip title="Capture screenshot now for current steps" placement="bottom">
            <Button size="small" icon={<CameraOutlined />} onClick={handleManualCapture}
              style={{ background: '#c77700', border: 'none', color: '#fff', fontWeight: 600 }}>
              📷 Capture
            </Button>
          </Tooltip>
        )}

        {/* Screen Record */}
        {recording ? (
          <Button size="small" onClick={stopRecording}
            style={{ background: '#333', border: 'none', color: '#fff', display: 'flex', alignItems: 'center', gap: 4 }}>
            <span style={{ width: 7, height: 7, borderRadius: '50%', background: '#ff4444', display: 'inline-block', animation: 'pulse-rec 1s infinite' }} />
            <StopOutlined />
            <span style={{ fontWeight: 600 }}>{formatTime(elapsed)}</span>
          </Button>
        ) : (
          <Tooltip title="Record screen">
            <Button size="small" icon={<VideoCameraOutlined />} onClick={startRecording}
              style={{ background: REDWOOD, border: 'none', color: '#fff' }} />
          </Tooltip>
        )}

        <Tooltip title={showPanel ? 'Hide steps panel' : 'Show steps panel'}>
          <Button size="small"
            icon={showPanel ? <MenuUnfoldOutlined /> : <MenuFoldOutlined />}
            onClick={() => setShowPanel(v => !v)}
            style={{ background: '#444', border: 'none', color: steps.length ? '#ffd54f' : '#fff' }} />
        </Tooltip>

        {/* Auto-login button */}
        <Tooltip title={hasSavedCreds ? 'Auto-fill login credentials' : 'No credentials saved yet — click ⚙ to set up'}>
          <Button size="small" icon={<KeyOutlined />} onClick={handleAutoLogin}
            style={{ background: hasSavedCreds ? '#5b3a8c' : '#444', border: 'none', color: hasSavedCreds ? '#fff' : '#666' }} />
        </Tooltip>

        {/* Credentials settings */}
        <Tooltip title="Login credentials settings">
          <Button size="small" icon={<SettingOutlined />} onClick={openCredsModal}
            style={{ background: '#444', border: 'none', color: hasSavedCreds ? '#4caf50' : '#aaa' }} />
        </Tooltip>

        <Tooltip title="Open Training Library">
          <Button size="small" onClick={() => navigate('/training')}
            style={{ background: '#1D7B4D', border: 'none', color: '#fff', fontSize: 11 }}>
            Training
          </Button>
        </Tooltip>
      </div>

      {/* Main Area: WebView + Steps Panel */}
      <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>

        {/* WebView */}
        {isElectron() ? (
          <webview
            ref={webviewRef}
            src={url}
            // @ts-ignore
            disablewebsecurity="true"
            // @ts-ignore
            allowpopups="true"
            style={{ flex: 1, minWidth: 0 }}
          />
        ) : (
          <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: 12, color: '#fff' }}>
            <Text style={{ color: '#fff', fontSize: 16 }}>Oracle Fusion WebView is only available in the desktop app.</Text>
            <Text style={{ color: '#aaa' }}>Please use the installed ReactERP desktop application.</Text>
          </div>
        )}

        {/* Steps Panel */}
        {showPanel && (
          <div style={{
            width: 340, background: '#1e1e2e', borderLeft: '1px solid #333',
            display: 'flex', flexDirection: 'column', flexShrink: 0,
          }}>
            {/* Panel Header */}
            <div style={{ padding: '10px 14px', borderBottom: '1px solid #333', display: 'flex', alignItems: 'center', gap: 8 }}>
              <AimOutlined style={{ color: tracking ? '#ff6b35' : '#888', fontSize: 15 }} />
              <Text style={{ color: '#fff', fontWeight: 600, flex: 1 }}>
                Steps {steps.length > 0 && <Tag color="blue" style={{ marginLeft: 4 }}>{steps.length}</Tag>}
              </Text>
              {tracking && (
                <span style={{ fontSize: 11, color: '#ff6b35', animation: 'pulse-trk 1s infinite' }}>● LIVE</span>
              )}
              <Tooltip title="Close panel">
                <CloseOutlined style={{ color: '#888', cursor: 'pointer' }} onClick={() => setShowPanel(false)} />
              </Tooltip>
            </div>

            {/* Steps List */}
            <div className="wv-scroll" style={{ flex: 1, overflowY: 'auto', padding: '4px 0' }}>
              {steps.length === 0 ? (
                <div style={{ padding: 30, textAlign: 'center' }}>
                  <AimOutlined style={{ fontSize: 32, color: '#555', display: 'block', marginBottom: 10 }} />
                  <Text style={{ color: '#666', fontSize: 13 }}>
                    {tracking ? 'Perform actions in Oracle Fusion…' : 'Click "Track Steps" to start capturing'}
                  </Text>
                </div>
              ) : (
                steps.map((s, i) => (
                  <React.Fragment key={s.id}>
                    {/* Step card */}
                    <div
                      draggable
                      onDragStart={() => handleDragStart(i)}
                      onDragOver={e => handleDragOver(e, i)}
                      onDrop={() => handleDrop(i)}
                      onDragEnd={() => { setDragIdx(null); setDragOverIdx(null); }}
                      style={{
                        padding: '7px 10px 7px 8px',
                        borderBottom: '1px solid #222',
                        background: dragOverIdx === i ? '#2a2a4a' : 'transparent',
                        opacity: dragIdx === i ? 0.4 : 1,
                        cursor: 'grab',
                      }}
                    >
                      <div style={{ display: 'flex', alignItems: 'flex-start', gap: 6 }}>
                        {/* Drag handle + step number */}
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2, flexShrink: 0, marginTop: 2 }}>
                          <span style={{ fontSize: 9, color: '#555', lineHeight: 1 }}>⠿</span>
                          <span style={{
                            width: 18, height: 18, borderRadius: '50%', background: '#333',
                            display: 'flex', alignItems: 'center', justifyContent: 'center',
                            fontSize: 9, color: '#aaa',
                          }}>{i + 1}</span>
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 4, marginBottom: 3 }}>
                            <Tag color={STEP_COLORS[s.type]} style={{ fontSize: 10, margin: 0 }}>{s.type}</Tag>
                            <span style={{ flex: 1 }} />
                            <Tooltip title="Delete this step" placement="left">
                              <span
                                onClick={() => setSteps(prev => prev.filter(x => x.id !== s.id))}
                                style={{ color: '#555', cursor: 'pointer', fontSize: 13, lineHeight: 1, padding: '0 2px' }}
                                onMouseEnter={e => (e.currentTarget.style.color = '#ff6b35')}
                                onMouseLeave={e => (e.currentTarget.style.color = '#555')}
                              >×</span>
                            </Tooltip>
                          </div>
                          <div style={{ fontSize: 11, color: '#e0e0e0', wordBreak: 'break-word', lineHeight: 1.4 }}>{s.description}</div>
                          {s.fields?.length ? (
                            <div style={{ fontSize: 10, color: '#a78bfa', marginTop: 3 }}>
                              📋 {s.fields.length} field{s.fields.length !== 1 ? 's' : ''} captured
                            </div>
                          ) : null}
                          {s.pageTitle && <div style={{ fontSize: 10, color: '#555', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{s.pageTitle}</div>}
                          {s.screenshot && (
                            <Tooltip title="Click to annotate" placement="left">
                              <div style={{ position: 'relative', marginTop: 5, cursor: 'pointer' }} onClick={() => setAnnotatingStepId(s.id)}>
                                <img src={s.screenshot} alt={`step ${i + 1}`}
                                  style={{ width: '100%', borderRadius: 4, border: '1px solid #333', display: 'block' }} />
                                <div style={{
                                  position: 'absolute', top: 4, right: 4,
                                  background: 'rgba(0,0,0,0.55)', borderRadius: 3,
                                  padding: '1px 6px', fontSize: 10, color: '#fff', pointerEvents: 'none',
                                }}>✏ annotate</div>
                              </div>
                            </Tooltip>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Insert screenshot button between steps */}
                    <Tooltip title="Insert a screenshot here" placement="right">
                      <div
                        onClick={() => handleInsertScreenshot(i)}
                        style={{
                          height: 18, display: 'flex', alignItems: 'center', justifyContent: 'center',
                          cursor: 'pointer', opacity: 0.3, fontSize: 10, color: '#7cb9e8',
                          borderBottom: '1px dashed #333',
                        }}
                        onMouseEnter={e => (e.currentTarget.style.opacity = '1')}
                        onMouseLeave={e => (e.currentTarget.style.opacity = '0.3')}
                      >
                        + 📷 insert screenshot here
                      </div>
                    </Tooltip>
                  </React.Fragment>
                ))
              )}
            </div>

            {/* Action Buttons */}
            <div style={{ padding: '12px 14px', borderTop: '1px solid #333', display: 'flex', flexDirection: 'column', gap: 6 }}>
              {/* User Manual row */}
              <div style={{ display: 'flex', gap: 4 }}>
                <Button
                  style={{ flex: 1, background: steps.length ? '#1565c0' : '#333', border: 'none', color: '#fff', fontSize: 12 }}
                  icon={<FileTextOutlined />}
                  onClick={handleGenerateManual}
                  disabled={!steps.length}
                >
                  User Manual
                </Button>
                <Tooltip title="Download as Word (.docx)">
                  <Button
                    style={{ background: steps.length ? '#1e4080' : '#333', border: 'none', color: '#fff', fontSize: 11 }}
                    onClick={handleGenerateWordManual}
                    disabled={!steps.length}
                  >
                    W↓
                  </Button>
                </Tooltip>
              </div>

              {/* UAT Script row */}
              <div style={{ display: 'flex', gap: 4 }}>
                <Button
                  style={{ flex: 1, background: steps.length ? '#2e7d32' : '#333', border: 'none', color: '#fff', fontSize: 12 }}
                  icon={<CheckSquareOutlined />}
                  onClick={handleGenerateUAT}
                  disabled={!steps.length}
                >
                  UAT Script
                </Button>
                <Tooltip title="Download as Word (.docx)">
                  <Button
                    style={{ background: steps.length ? '#1b5e20' : '#333', border: 'none', color: '#fff', fontSize: 11 }}
                    onClick={handleGenerateWordUAT}
                    disabled={!steps.length}
                  >
                    W↓
                  </Button>
                </Tooltip>
              </div>

              {/* Preview + Clear row */}
              <div style={{ display: 'flex', gap: 4 }}>
                <Button
                  style={{ flex: 1, background: steps.length ? '#6a1b9a' : '#333', border: 'none', color: '#fff', fontSize: 12 }}
                  onClick={handleOpenPreview}
                  disabled={!steps.length}
                >
                  👁 Preview
                </Button>
                <Button
                  icon={<DeleteOutlined />}
                  onClick={() => { setSteps([]); message.info('Steps cleared'); }}
                  disabled={!steps.length}
                  style={{ background: '#333', border: 'none', color: steps.length ? '#ff6b35' : '#555' }}
                />
              </div>
            </div>
          </div>
        )}
      </div>
    </Layout>

    {/* ── User Manual Preview Modal ── */}
    <Modal
      open={previewOpen}
      onCancel={handleClosePreview}
      width="92vw"
      style={{ top: 16 }}
      title="User Manual Preview"
      footer={
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
          <Button onClick={handleGenerateManual} icon={<FileTextOutlined />}>Download HTML</Button>
          <Button onClick={handleGenerateWordManual} style={{ background: '#1565c0', color: '#fff', border: 'none' }}>Download Word</Button>
          <Button onClick={handleClosePreview}>Close</Button>
        </div>
      }
    >
      <iframe
        src={previewUrl}
        style={{ width: '100%', height: 'calc(90vh - 130px)', border: 'none', borderRadius: 4 }}
        title="User Manual Preview"
      />
    </Modal>

    {/* ── Credentials Setup Modal ── */}
    <Modal
      title={<span><LockOutlined style={{ color: '#5b3a8c', marginRight: 8 }} />Oracle Fusion Login Credentials</span>}
      open={credsModalOpen}
      onCancel={() => setCredsModalOpen(false)}
      footer={null}
      width={420}
    >
      <p style={{ color: '#666', fontSize: 13, marginBottom: 20 }}>
        Credentials are encrypted using your OS keychain (Windows DPAPI / Mac Keychain) and stored locally.
        They are never sent to any server.
      </p>
      <Form form={credsForm} layout="vertical">
        <Form.Item
          name="username"
          label="Oracle Fusion Username"
          rules={[{ required: true, message: 'Please enter your username' }]}
        >
          <Input prefix={<UserOutlined style={{ color: '#aaa' }} />} placeholder="e.g. john.smith@company.com" />
        </Form.Item>
        <Form.Item
          name="password"
          label="Oracle Fusion Password"
          rules={[{ required: true, message: 'Please enter your password' }]}
        >
          <Input.Password
            placeholder="Your Oracle Fusion password"
            iconRender={visible => (visible ? <EyeTwoTone /> : <EyeInvisibleOutlined />)}
          />
        </Form.Item>
        <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
          <Button type="primary" onClick={handleSaveCreds} style={{ flex: 1, background: '#5b3a8c', borderColor: '#5b3a8c' }}>
            <LockOutlined /> Save Credentials
          </Button>
          {hasSavedCreds && (
            <Button danger onClick={handleClearCreds}>
              Clear
            </Button>
          )}
        </div>
        {hasSavedCreds && (
          <div style={{ marginTop: 12, padding: '8px 12px', background: '#f0fff4', borderRadius: 6, border: '1px solid #b7eb8f', fontSize: 12, color: '#389e0d' }}>
            ✓ Credentials saved — click the <KeyOutlined /> button in the toolbar to auto-fill the login page
          </div>
        )}
      </Form>
    </Modal>
    {annotatingStepId && (() => {
      const step = steps.find(s => s.id === annotatingStepId);
      if (!step?.screenshot) return null;
      return (
        <ScreenshotAnnotator
          screenshot={step.screenshot}
          onSave={(annotated) => {
            setSteps(prev => prev.map(s =>
              s.id === annotatingStepId ? { ...s, screenshot: annotated } : s
            ));
            setAnnotatingStepId(null);
          }}
          onClose={() => setAnnotatingStepId(null)}
        />
      );
    })()}
    </>
  );
};

export default OracleFusion;
