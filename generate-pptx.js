#!/usr/bin/env node
/**
 * PSU Darpan — PPTX Generator
 * Produces 4 vector PPTX files (Short Dark/Light, Full Dark/Light)
 * Uses pptxgenjs: text + shapes, NOT screenshots.
 * Run: node generate-pptx.js
 */

const PptxGenJS = require('pptxgenjs');
const path      = require('path');

// ── Theme palettes ──────────────────────────────────────────────────────────
const THEMES = {
  dark: {
    bg:        '0a0f1e',
    surface:   '111827',
    surface2:  '1e293b',
    border:    '1e3a5f',
    text:      'f0f4ff',
    muted:     '94a3b8',
    faint:     '64748b',
    amber:     'f59e0b',
    amberBg:   '431407',
    amberText: 'fbbf24',
    blueBg:    '0c2d54',
    blueText:  '60a5fa',
    redBg:     '450a0a',
    redText:   'f87171',
    greenBg:   '052e16',
    greenText: '4ade80',
    logo:      'f8fafc',
    tagRed:    { bg: '450a0a', text: 'f87171' },
    tagBlue:   { bg: '0c2d54', text: '93c5fd' },
    tagGreen:  { bg: '052e16', text: '4ade80' },
    tagAmber:  { bg: '431407', text: 'fbbf24' },
  },
  light: {
    bg:        'f1f5f9',
    surface:   'ffffff',
    surface2:  'f8fafc',
    border:    'cbd5e1',
    text:      '0f172a',
    muted:     '475569',
    faint:     '64748b',
    amber:     'd97706',
    amberBg:   'fef3c7',
    amberText: '92400e',
    blueBg:    'dbeafe',
    blueText:  '1d4ed8',
    redBg:     'fee2e2',
    redText:   'b91c1c',
    greenBg:   'dcfce7',
    greenText: '15803d',
    logo:      '0f172a',
    tagRed:    { bg: 'fee2e2', text: 'b91c1c' },
    tagBlue:   { bg: 'dbeafe', text: '1d4ed8' },
    tagGreen:  { bg: 'dcfce7', text: '15803d' },
    tagAmber:  { bg: 'fef3c7', text: '92400e' },
  },
};

// ── Layout constants (inches, LAYOUT_WIDE = 13.33 x 7.5) ───────────────────
const W  = 13.33;
const H  = 7.5;
const PX = 0.45; // horizontal padding
const PY = 0.35; // vertical padding (top)
const CW = W - PX * 2; // content width

// ── Helpers ─────────────────────────────────────────────────────────────────

function newPptx(title) {
  const pptx   = new PptxGenJS();
  pptx.layout  = 'LAYOUT_WIDE';
  pptx.title   = title;
  pptx.author  = 'PSU Darpan';
  pptx.company = 'PSU Darpan';
  return pptx;
}

/** Solid background rect filling the whole slide */
function slideBg(slide, t) {
  slide.addShape(slide.ShapeType ? slide.ShapeType.rect : 'rect', {
    x: 0, y: 0, w: W, h: H,
    fill: { color: t.bg },
    line: { type: 'none' },
  });
}

/** Top-bar: "PSU Darpan" logo left, slide label right */
function addTopBar(slide, t, label) {
  // Logo
  slide.addText([
    { text: 'PSU ', options: { color: t.logo } },
    { text: 'Darpan', options: { color: t.amber } },
  ], { x: PX, y: PY, w: 3, h: 0.3, fontSize: 13, bold: true });

  // Label
  if (label) {
    slide.addText(label, {
      x: PX, y: PY, w: CW, h: 0.3,
      fontSize: 8, color: t.faint, align: 'right',
    });
  }
}

/** Tag pill */
function addTag(slide, t, text, color, x, y) {
  const scheme = color === 'red'   ? t.tagRed
               : color === 'blue'  ? t.tagBlue
               : color === 'green' ? t.tagGreen
               : t.tagAmber;
  slide.addText(text.toUpperCase(), {
    x, y, w: 2.4, h: 0.22,
    fontSize: 7, bold: true, color: scheme.text,
    fill: { color: scheme.bg },
    align: 'center',
    charSpacing: 1.5,
    shape: 'roundRect', rectRadius: 0.04,
  });
}

/** Section heading */
function addTitle(slide, t, text, x, y, w, opts = {}) {
  slide.addText(text, {
    x, y, w, h: 0.7,
    fontSize: opts.fontSize || 22,
    bold: true,
    color: t.logo,
    wrap: true,
    ...opts,
  });
}

/** Body text */
function addBody(slide, t, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: opts.fontSize || 9.5,
    color: opts.color || t.muted,
    wrap: true,
    valign: 'top',
    ...opts,
  });
}

/** Card box */
function addCard(slide, t, x, y, w, h, title, body, titleColor) {
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: t.surface },
    line: { color: t.border, width: 0.5 },
    rectRadius: 0.1,
  });
  slide.addText(title, {
    x: x + 0.12, y: y + 0.1, w: w - 0.24, h: 0.22,
    fontSize: 9, bold: true, color: titleColor || t.amber, wrap: true,
  });
  if (body) {
    slide.addText(body, {
      x: x + 0.12, y: y + 0.34, w: w - 0.24, h: h - 0.44,
      fontSize: 8, color: t.muted, wrap: true, valign: 'top',
    });
  }
}

/** Stat box */
function addStat(slide, t, x, y, w, num, label, src, numColor) {
  slide.addText(num, {
    x, y: y + 0.05, w, h: 0.38,
    fontSize: 20, bold: true,
    color: numColor || t.amber,
    align: 'center',
  });
  slide.addText(label, {
    x, y: y + 0.46, w, h: 0.26,
    fontSize: 7.5, color: t.muted,
    align: 'center', wrap: true,
  });
  if (src) {
    slide.addText(src, {
      x, y: y + 0.74, w, h: 0.16,
      fontSize: 6, color: t.faint,
      align: 'center',
    });
  }
}

/** Bullet list */
function addBullets(slide, t, items, x, y, w, h) {
  const lines = items.map(item => ({ text: '• ' + item, options: { color: t.muted } }));
  slide.addText(lines, {
    x, y, w, h,
    fontSize: 8.5,
    wrap: true,
    valign: 'top',
    paraSpaceAfter: 3,
  });
}

/** Confidential footer */
function addFooter(slide, t) {
  slide.addText('CONFIDENTIAL · NDA  —  PSU Darpan © 2026', {
    x: 0, y: H - 0.22, w: W, h: 0.22,
    fontSize: 6.5, color: t.faint,
    align: 'center',
  });
}

// ── Short Pitch — 5 slides ───────────────────────────────────────────────────

function buildShortDeck(themeName) {
  const t    = THEMES[themeName];
  const pptx = newPptx('PSU Darpan — Short Pitch');

  // ── S1: Cover ──────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '01 of 05 · India\'s PSU intelligence app');

    s.addText([
      { text: 'PSU ', options: { color: t.logo } },
      { text: 'Darpan', options: { color: t.amber } },
    ], { x: PX, y: 0.95, w: CW, h: 0.85, fontSize: 48, bold: true });

    addBody(s, t,
      'India has ~390 PSUs and lakhs of public sector employees. The information they need most reaches them last — distorted, dense, or buried under general news.',
      PX, 1.85, 7, 0.45, { fontSize: 10 });

    // 4 stats
    const stats = [
      ['~390', 'PSUs in India', 'DPE, 2024'],
      ['11L+', 'Permanent CPSE employees', 'PIB / Public Enterprises Survey'],
      ['18%', 'India GDP from public sector', ''],
      ['Lakhs', 'GATE aspirants targeting PSUs annually', ''],
    ];
    const sw = CW / 4 - 0.1;
    stats.forEach(([num, label, src], i) => {
      const x = PX + i * (sw + 0.13);
      s.addShape('roundRect', { x, y: 2.4, w: sw, h: 1.05, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.08 });
      addStat(s, t, x, 2.4, sw, num, label, src || null);
    });

    // info box
    s.addShape('roundRect', { x: PX, y: 3.6, w: CW, h: 0.75, fill: { color: t.blueBg }, line: { type: 'none' }, rectRadius: 0.1 });
    addBody(s, t,
      'PSU Darpan is a native mobile app built exclusively for this audience — personalized, AI-powered, multi-format, and free. The modern experience PSU professionals never had.',
      PX + 0.2, 3.7, CW - 0.4, 0.55, { fontSize: 9, color: t.blueText });

    addFooter(s, t);
  }

  // ── S2: Problem ─────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '02 of 05 · Why we are building this');

    addTitle(s, t, 'Three structural failures. One product that fixes all three.', PX, 0.72, CW, { fontSize: 20 });
    addBody(s, t,
      'The problem isn\'t that PSU news doesn\'t exist. It\'s that the entire information chain was built for government audiences — not for the professional whose career it affects most.',
      PX, 1.35, CW, 0.3, { fontSize: 9 });

    // Column headers
    s.addText('Problem', { x: PX + 0.35, y: 1.72, w: 5.4, h: 0.2, fontSize: 8, bold: true, color: t.redText, align: 'center' });
    s.addText('Our Answer', { x: PX + 6.0, y: 1.72, w: 5.5, h: 0.2, fontSize: 8, bold: true, color: t.greenText, align: 'center' });

    const rows = [
      ['1', 'Production — Unreadable', 'Press releases & PDFs in bureaucratic language. No one translates them for the professional reader.', 'AI rewrites at source', 'LLM auto-summarizes into plain language, assigns category, maps to PSU/sector before it hits the feed.'],
      ['2', 'Distribution — Wrong order', 'News reaches journalists first. Employees get a distorted WhatsApp forward last — too late to act on.', 'Direct push to the employee', 'Native app with push notifications delivers verified news to the PSU professional before WhatsApp does.'],
      ['3', 'Personalization — One-size-fits-all', 'An ONGC engineer and a BHEL officer have near-zero news overlap. Every platform treats them identically.', 'PSU-native feed from day one', 'Select your PSU → feed is immediately personalized by org, sector & 13 structured categories. No warm-up.'],
    ];

    rows.forEach(([num, probTitle, probBody, solTitle, solBody], i) => {
      const ry = 2.0 + i * 1.55;
      // num circle
      s.addShape('ellipse', { x: PX, y: ry + 0.2, w: 0.28, h: 0.28, fill: { color: t.surface2 }, line: { color: t.border, width: 0.5 } });
      s.addText(num, { x: PX, y: ry + 0.2, w: 0.28, h: 0.28, fontSize: 9, bold: true, color: t.amber, align: 'center', valign: 'middle' });
      // problem card
      addCard(s, t, PX + 0.35, ry, 5.4, 1.4, probTitle, probBody, t.redText);
      // solution card
      addCard(s, t, PX + 5.9, ry, 5.5, 1.4, solTitle, solBody, t.greenText);
    });

    addFooter(s, t);
  }

  // ── S3: Product ─────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '03 of 05 · What we\'ve built');

    addTitle(s, t, 'Every format. Every PSU update. One app.', PX, 0.72, CW, { fontSize: 20 });
    addBody(s, t, 'Inshorts × LinkedIn × Moneycontrol — exclusively for India\'s public sector, as a native mobile app.', PX, 1.28, CW, 0.25, { fontSize: 9 });

    const features = [
      ['Personalized Feed', 'Ranked by PSU, sector & topics. Relevant from first open.'],
      ['Shorts', 'Reel-style swipeable cards. Like, share, bookmark inline.'],
      ['Stories', 'Instagram-style visual stories with cube transition & auto-progress.'],
      ['Breaking Ticker', 'Live editor-curated marquee for urgent news.'],
      ['Jobs Page', 'All PSU recruitment & HR pay news in one filtered view.'],
      ['PSU Browser', '~390 PSU profiles with sector, ministry & all articles.'],
      ['Employee Directory', 'OTP-gated verified network — search by name, plant, role.'],
      ['Search', 'Real-time across articles, PSUs, and sectors.'],
      ['AI Helper', 'Natural language queries ("NTPC Q3 profit?") — coming soon.'],
    ];

    const cols = 3;
    const fw   = (CW - 0.2) / cols;
    const fh   = 1.25;
    features.forEach(([title, body], i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      addCard(s, t, PX + col * (fw + 0.1), 1.6 + row * (fh + 0.08), fw, fh, title, body, t.blueText);
    });

    addFooter(s, t);
  }

  // ── S4: Pipeline ─────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '04 of 05 · How we build it');

    addTitle(s, t, 'AI speed. Human accuracy. Built to scale.', PX, 0.72, CW, { fontSize: 20 });

    // Pipeline steps
    const steps = [
      ['🤖', 'AI Scraper', 'RSS from PIB & PSU sites — LLM summarizes, classifies, maps to PSU/sector'],
      ['✍️', 'Writer Network', 'Writers submit articles. Editorial review queue. Admin approves before live.'],
      ['⚡', 'Admin Direct', 'Breaking news live instantly — no queue, no delay.'],
    ];
    const pw = (CW - 0.6) / 3;
    steps.forEach(([icon, title, body], i) => {
      const x = PX + i * (pw + 0.3);
      s.addShape('roundRect', { x, y: 1.3, w: pw, h: 1.9, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.1 });
      s.addText(icon, { x, y: 1.4, w: pw, h: 0.4, fontSize: 20, align: 'center' });
      s.addText(title, { x: x + 0.1, y: 1.84, w: pw - 0.2, h: 0.26, fontSize: 10, bold: true, color: t.amber, align: 'center' });
      addBody(s, t, body, x + 0.1, 2.14, pw - 0.2, 0.9, { fontSize: 8.5, align: 'center' });
      if (i < 2) {
        s.addText('→', { x: x + pw + 0.02, y: 2.0, w: 0.26, h: 0.3, fontSize: 14, color: t.faint, align: 'center' });
      }
    });

    // Two columns: Tech Stack & Trust Built In
    const col2w = CW / 2 - 0.15;
    addBullets(s, t,
      ['Next.js 15 + TypeScript · Tailwind CSS v4', 'Firebase Phone OTP authentication', 'Node.js REST API + PostgreSQL', 'LLM — auto-classify, summarize, map', 'Native mobile app distribution'],
      PX, 3.45, col2w, 1.35);

    addBullets(s, t,
      ['Every AI-scraped article reviewed by humans before publishing', 'OTP-gated employee directory — no fake profiles', 'Role-based access: public · registered · writer · admin', 'Full editorial CMS with audit trail'],
      PX + col2w + 0.3, 3.45, col2w, 1.35);

    s.addText('TECH STACK', { x: PX, y: 3.3, w: col2w, h: 0.18, fontSize: 7, bold: true, color: t.faint, charSpacing: 1 });
    s.addText('TRUST BUILT IN', { x: PX + col2w + 0.3, y: 3.3, w: col2w, h: 0.18, fontSize: 7, bold: true, color: t.faint, charSpacing: 1 });

    addFooter(s, t);
  }

  // ── S5: Traction ─────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '05 of 05 · Traction & positioning');

    addTitle(s, t, 'Live product. Real users. Clear positioning.', PX, 0.72, CW, { fontSize: 20 });

    // 4 stats
    const stats = [
      ['12,400', 'Registered users'],
      ['3,200', 'Daily active users'],
      ['850+', 'Articles published'],
      ['4.8m', 'Avg. session duration'],
    ];
    const sw = CW / 4 - 0.1;
    stats.forEach(([num, label], i) => {
      const x = PX + i * (sw + 0.13);
      s.addShape('roundRect', { x, y: 1.3, w: sw, h: 1.0, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.08 });
      addStat(s, t, x, 1.3, sw, num, label, null, t.greenText);
    });

    // Competitive table
    const rows = [
      ['Existing PSU media sites', 'Outdated UI & tech, no AI, no personalization, not mobile-first'],
      ['ET / Moneycontrol / LinkedIn', 'PSU content is incidental — designed for a general audience'],
      ['WhatsApp / PIB', 'Unverified or unreadable — no personalization, no trust layer'],
      ['PSU Darpan ✦', 'Modern · Personalized · AI-powered · Verified · Multi-format · Built only for PSU professionals'],
    ];

    const th = 0.24;
    s.addShape('rect', { x: PX, y: 2.46, w: 4.0, h: th, fill: { color: t.surface2 }, line: { type: 'none' } });
    s.addShape('rect', { x: PX + 4.0, y: 2.46, w: CW - 4.0, h: th, fill: { color: t.surface2 }, line: { type: 'none' } });
    s.addText('Platform', { x: PX + 0.1, y: 2.46, w: 3.8, h: th, fontSize: 8, bold: true, color: t.text, valign: 'middle' });
    s.addText('What they lack for PSU professionals', { x: PX + 4.1, y: 2.46, w: CW - 4.1, h: th, fontSize: 8, bold: true, color: t.text, valign: 'middle' });

    rows.forEach(([platform, gap], i) => {
      const ry = 2.46 + th + i * 0.7;
      const isOurs = platform.includes('PSU Darpan');
      const rowBg = isOurs ? t.amberBg : (i % 2 === 0 ? t.surface : t.bg);
      s.addShape('rect', { x: PX, y: ry, w: 4.0, h: 0.65, fill: { color: rowBg }, line: { type: 'none' } });
      s.addShape('rect', { x: PX + 4.0, y: ry, w: CW - 4.0, h: 0.65, fill: { color: rowBg }, line: { type: 'none' } });
      s.addText(platform, { x: PX + 0.1, y: ry + 0.05, w: 3.8, h: 0.55, fontSize: isOurs ? 8.5 : 8, bold: isOurs, color: isOurs ? t.amberText : t.text, valign: 'middle', wrap: true });
      s.addText(gap, { x: PX + 4.1, y: ry + 0.05, w: CW - 4.2, h: 0.55, fontSize: 8, color: isOurs ? t.amberText : t.muted, valign: 'middle', wrap: true });
    });

    // Quote
    s.addShape('roundRect', { x: PX, y: 5.58, w: CW, h: 0.65, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.1 });
    addBody(s, t, '"We\'re not competing on coverage. We\'re replacing WhatsApp groups and outdated portals as the primary daily habit of PSU India."', PX + 0.2, 5.68, CW - 0.4, 0.5, { fontSize: 9, color: t.muted, italic: true });

    addFooter(s, t);
  }

  return pptx;
}

// ── Full Pitch Deck — 10 slides ──────────────────────────────────────────────

function buildFullDeck(themeName) {
  const t    = THEMES[themeName];
  const pptx = newPptx('PSU Darpan — Full Pitch Deck');

  // ── F1: Cover ──────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '01 / 10');

    addTag(s, t, "India's PSU Intelligence Platform", 'amber', PX, 0.78);

    s.addText([
      { text: 'PSU ', options: { color: t.logo } },
      { text: 'Darpan', options: { color: t.amber } },
    ], { x: PX, y: 1.1, w: CW, h: 1.1, fontSize: 52, bold: true });

    addBody(s, t, 'India\'s first dedicated intelligence app for public sector professionals', PX, 2.25, CW, 0.3, { fontSize: 11, color: t.muted });

    s.addShape('rect', { x: PX, y: 2.65, w: 1.2, h: 0.025, fill: { color: t.amber }, line: { type: 'none' } });

    addBody(s, t,
      'India has ~390 PSUs and lakhs of public sector employees. The information they need most — about their own organizations — reaches them last, distorted, or not at all. PSU Darpan is built to fix that.',
      PX, 2.82, 7.5, 0.55, { fontSize: 10 });

    // Pills
    const pills = ['Native mobile app', 'AI-powered pipeline', 'Live product', 'Free for users'];
    pills.forEach((p, i) => {
      const pw = 1.65;
      s.addText(p, {
        x: PX + i * (pw + 0.1), y: 3.5, w: pw, h: 0.26,
        fontSize: 8, color: t.blueText,
        fill: { color: t.blueBg },
        align: 'center',
        shape: 'roundRect', rectRadius: 0.04,
      });
    });

    // Big numbers
    const bigs = [['~390', 'PSUs in India', 'DPE, 2024'], ['13', 'Content categories', ''], ['Live', 'Product, today', '']];
    bigs.forEach(([num, label, src], i) => {
      const bw = 2.5;
      const bx = PX + i * (bw + 0.5);
      s.addText(num, { x: bx, y: 4.05, w: bw, h: 0.7, fontSize: 32, bold: true, color: t.amber });
      s.addText(label, { x: bx, y: 4.78, w: bw, h: 0.22, fontSize: 8.5, color: t.muted });
      if (src) s.addText(src, { x: bx, y: 5.02, w: bw, h: 0.18, fontSize: 7, color: t.faint });
    });

    addFooter(s, t);
  }

  // ── F2: Why the Problem Exists ─────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '02 / 10');

    addTag(s, t, 'Why the Problem Exists', 'red', PX, 0.78);
    addTitle(s, t, 'PSU News Was Never Designed for the PSU Employee', PX, 1.08, CW, { fontSize: 19 });
    addBody(s, t,
      'The core issue isn\'t that PSU news doesn\'t exist. It\'s that the entire information chain was built for government audiences — not for the professional whose career it directly affects.',
      PX, 1.72, CW, 0.3, { fontSize: 9 });

    const hw = CW / 2 - 0.15;
    // How produced
    s.addShape('roundRect', { x: PX, y: 2.12, w: hw, h: 2.5, fill: { color: t.surface }, line: { color: t.redBg, width: 1.5 }, rectRadius: 0.1 });
    s.addText('How PSU news is produced', { x: PX + 0.15, y: 2.22, w: hw - 0.3, h: 0.25, fontSize: 9.5, bold: true, color: t.redText });
    addBullets(s, t,
      ['Press releases written in bureaucratic language', 'Ministry circulars formatted as dense PDFs', 'PIB bulletins designed for journalists, not employees', 'No translation layer for the professional consumer'],
      PX + 0.15, 2.5, hw - 0.3, 1.9);

    // How travels
    s.addShape('roundRect', { x: PX + hw + 0.3, y: 2.12, w: hw, h: 2.5, fill: { color: t.surface }, line: { color: t.redBg, width: 1.5 }, rectRadius: 0.1 });
    s.addText('How PSU news travels', { x: PX + hw + 0.45, y: 2.22, w: hw - 0.3, h: 0.25, fontSize: 9.5, bold: true, color: t.redText });
    addBullets(s, t,
      ['Journalist picks it up → publishes for general audience', 'Someone screenshots → shares to WhatsApp group', 'Forward gets distorted 3–4 times along the way', 'Employee gets it last — often wrong, always late'],
      PX + hw + 0.45, 2.5, hw - 0.3, 1.9);

    // Structural gap info box
    s.addShape('roundRect', { x: PX, y: 4.78, w: CW, h: 0.55, fill: { color: t.amberBg }, line: { type: 'none' }, rectRadius: 0.08 });
    addBody(s, t,
      'The structural gap: there is no channel that goes directly from PSU event → relevant employee, in readable language, before WhatsApp distortion sets in.',
      PX + 0.2, 4.88, CW - 0.4, 0.35, { fontSize: 8.5, color: t.amberText, align: 'center' });

    // Quote
    s.addShape('roundRect', { x: PX, y: 5.46, w: CW, h: 0.62, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.08 });
    addBody(s, t,
      '"Employees at SAIL\'s Bhilai plant often learn about DPC results from a WhatsApp forward — not from their own organization\'s communication."',
      PX + 0.2, 5.56, CW - 0.4, 0.42, { fontSize: 8.5, italic: true });

    addFooter(s, t);
  }

  // ── F3: Three Problem Layers ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '03 / 10');

    addTag(s, t, 'The Three Problem Layers', 'red', PX, 0.78);
    addTitle(s, t, 'Three Structural Failures. One Product That Fixes All Three.', PX, 1.08, CW, { fontSize: 19 });

    s.addText('Problem', { x: PX + 0.45, y: 1.68, w: 5.5, h: 0.2, fontSize: 8, bold: true, color: t.redText, align: 'center' });
    s.addText('PSU Darpan\'s Answer', { x: PX + 6.2, y: 1.68, w: 5.5, h: 0.2, fontSize: 8, bold: true, color: t.greenText, align: 'center' });

    const rows = [
      ['1', 'Production — Unreadable content', 'PSU news originates as press releases & PDFs in bureaucratic language. No one translates it for the professional reader.', 'AI rewrites at source', 'LLM auto-summarizes every article into plain language, assigns category, and maps it to the right PSU & sector — before it ever reaches the feed.'],
      ['2', 'Distribution — Wrong channel, wrong order', 'News reaches journalists and investors first. Employees get a distorted WhatsApp forward last — often too late to act on.', 'Direct push to the employee', 'Native app with push notifications delivers verified news directly to the PSU professional — before the WhatsApp chain even starts.'],
      ['3', 'Personalization — One-size-fits-all', 'An ONGC engineer and a BHEL officer have near-zero news overlap. Every existing platform treats them identically.', 'PSU-native feed from day one', 'User selects their PSU. Feed is immediately personalized by organization, sector, and 13 structured categories. No algorithm warm-up needed.'],
    ];

    rows.forEach(([num, probTitle, probBody, solTitle, solBody], i) => {
      const ry = 1.95 + i * 1.65;
      s.addShape('ellipse', { x: PX, y: ry + 0.28, w: 0.3, h: 0.3, fill: { color: t.surface2 }, line: { color: t.border, width: 0.5 } });
      s.addText(num, { x: PX, y: ry + 0.28, w: 0.3, h: 0.3, fontSize: 9, bold: true, color: t.amber, align: 'center', valign: 'middle' });
      addCard(s, t, PX + 0.4, ry, 5.5, 1.52, probTitle, probBody, t.redText);
      addCard(s, t, PX + 6.1, ry, 5.5, 1.52, solTitle, solBody, t.greenText);
    });

    addFooter(s, t);
  }

  // ── F4: Solution ───────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '04 / 10');

    addTag(s, t, 'The Solution', 'blue', PX, 0.78);
    addTitle(s, t, 'Not Just Better UI. A New Trust Infrastructure for PSU India.', PX, 1.08, CW, { fontSize: 19 });
    addBody(s, t,
      'Existing PSU media is present but built for a different era — no personalization, no AI, no multi-format consumption, and no verified community layer. PSU Darpan is the modern rebuild of that stack.',
      PX, 1.72, CW, 0.3, { fontSize: 9 });

    // Centre box
    s.addShape('roundRect', { x: PX, y: 2.1, w: CW, h: 0.65, fill: { color: t.blueBg }, line: { type: 'none' }, rectRadius: 0.1 });
    s.addText([
      { text: 'Inshorts', options: { bold: true, color: t.blueText } },
      { text: ' × ', options: { color: t.muted } },
      { text: 'LinkedIn', options: { bold: true, color: t.blueText } },
      { text: ' × ', options: { color: t.muted } },
      { text: 'Moneycontrol', options: { bold: true, color: t.blueText } },
    ], { x: PX + 0.2, y: 2.18, w: CW - 0.4, h: 0.28, fontSize: 13, align: 'center' });
    addBody(s, t, '— exclusively for India\'s public sector, as a native mobile app', PX + 0.2, 2.48, CW - 0.4, 0.2, { fontSize: 9, align: 'center', color: t.muted });

    // 3 cards
    const cw3 = (CW - 0.2) / 3;
    const cards = [
      ['📰', 'Readable news', 'AI converts bureaucratic source content into clean, scannable summaries across 13 structured categories.'],
      ['🎯', 'Personalized from day one', 'PSU + sector + category selection means the feed is relevant the first time a user opens the app — no algorithm warming.'],
      ['✅', 'Verified trust layer', 'OTP-gated employee directory and editorial review pipeline create a citable, verified source — not a WhatsApp rumour.'],
    ];
    cards.forEach(([icon, title, body], i) => {
      const cx = PX + i * (cw3 + 0.1);
      s.addShape('roundRect', { x: cx, y: 2.9, w: cw3, h: 2.2, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.1 });
      s.addText(icon, { x: cx, y: 3.0, w: cw3, h: 0.35, fontSize: 18, align: 'center' });
      s.addText(title, { x: cx + 0.12, y: 3.38, w: cw3 - 0.24, h: 0.26, fontSize: 9.5, bold: true, color: t.blueText, align: 'center' });
      addBody(s, t, body, cx + 0.12, 3.68, cw3 - 0.24, 1.2, { fontSize: 8.5, align: 'center' });
    });

    // Quote
    s.addShape('roundRect', { x: PX, y: 5.28, w: CW, h: 0.8, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.1 });
    addBody(s, t,
      '"Lakhs of employees make career decisions — deputation applications, transfer requests — based on unverified forwards. PSU Darpan is the verified source they should have always had."',
      PX + 0.2, 5.38, CW - 0.4, 0.6, { fontSize: 9, italic: true });

    addFooter(s, t);
  }

  // ── F5: Product ────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '05 / 10');

    addTag(s, t, 'The Product', 'blue', PX, 0.78);
    addTitle(s, t, 'Every Format. Every Update. One App.', PX, 1.08, CW, { fontSize: 19 });

    const features = [
      ['Personalized Feed', 'Ranked by PSU, sector & topics. Infinite scroll, pull-to-refresh.'],
      ['Shorts', 'Vertical reel-style cards. Swipe, like, share — without leaving the feed.'],
      ['Stories', 'Instagram-style bubbles. Cube transition. Auto-progress. Time-gated.'],
      ['Breaking Ticker', 'Live editor-curated scrolling marquee for urgent news.'],
      ['Jobs Page', 'All PSU recruitment, vacancy & HR pay revision in one filtered view.'],
      ['PSU Browser', '~390 PSU profiles with sector, ministry & all related articles.'],
      ['Employee Directory', 'OTP-gated verified network. Search by name, plant, designation.'],
      ['Search', 'Real-time across articles, PSUs, and sectors.'],
      ['AI Helper', 'Natural language PSU queries. ("NTPC\'s Q3 profit?") — coming soon.'],
    ];

    const cols = 3;
    const fw   = (CW - 0.2) / cols;
    const fh   = 1.08;
    features.forEach(([title, body], i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      addCard(s, t, PX + col * (fw + 0.1), 1.55 + row * (fh + 0.06), fw, fh, title, body, t.blueText);
    });

    // Categories
    s.addText('13 Content Categories', { x: PX, y: 5.0, w: CW, h: 0.2, fontSize: 7.5, bold: true, color: t.faint, charSpacing: 1 });
    const cats = ['Appointments', 'Promotions', 'Jobs', 'Policy', 'Tenders', 'Projects', 'CSR', 'Awards', 'Finance', 'HR & Pay', 'General', 'Breaking', 'Analysis'];
    const catColors = [t.blueText, t.blueText, t.amberText, t.blueText, t.amberText, t.blueText, t.greenText, t.greenText, t.amberText, t.blueText, t.blueText, t.redText, t.blueText];
    const catBgs    = [t.blueBg, t.blueBg, t.amberBg, t.blueBg, t.amberBg, t.blueBg, t.greenBg, t.greenBg, t.amberBg, t.blueBg, t.blueBg, t.redBg, t.blueBg];
    cats.forEach((cat, i) => {
      const cw = 1.4;
      const cx = PX + (i % 7) * (cw + 0.08);
      const cy = 5.26 + Math.floor(i / 7) * 0.32;
      s.addText(cat, {
        x: cx, y: cy, w: cw, h: 0.24,
        fontSize: 7.5, color: catColors[i],
        fill: { color: catBgs[i] },
        align: 'center',
        shape: 'roundRect', rectRadius: 0.04,
      });
    });

    addFooter(s, t);
  }

  // ── F6: Who We Serve ───────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '06 / 10');

    addTag(s, t, 'Who We Serve', 'amber', PX, 0.78);
    addTitle(s, t, 'One Universe. Four Distinct Audiences.', PX, 1.08, CW, { fontSize: 19 });

    // Stats
    const stats = [
      ['~390', 'PSUs in India (central + state)', 'DPE, 2024'],
      ['11L+', 'Permanent CPSE employees', 'PIB / Public Enterprises Survey'],
      ['18%', 'India\'s GDP from public sector', ''],
      ['Lakhs', 'GATE aspirants targeting PSU jobs annually', ''],
    ];
    const sw = CW / 4 - 0.1;
    stats.forEach(([num, label, src], i) => {
      const x = PX + i * (sw + 0.13);
      s.addShape('roundRect', { x, y: 1.6, w: sw, h: 1.0, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.08 });
      addStat(s, t, x, 1.6, sw, num, label, src || null);
    });

    // 4 audience cards
    const aw = CW / 2 - 0.15;
    const audiences = [
      ['PSU Employees', t.amber, 'Officers, engineers, managers in CPSEs and state PSUs. Need appointments, policy, pay revision, tenders — fast and verified.'],
      ['PSU Aspirants', t.blueText, 'Lakhs of GATE-qualified graduates targeting PSU recruitment every year. Obsessively track vacancy notifications, cut-offs, and DPC results. Currently underserved by every platform.'],
      ['PSU Investors & Analysts', t.greenText, 'Retail investors in listed PSUs (ONGC, NTPC, SAIL, Coal India, HAL, BEL) who need financial news, project wins, and leadership changes faster than ET or Moneycontrol surfaces them.'],
      ['Extended Ecosystem', '#a78bfa', 'Retirees, PSU families, government vendors & contractors, policy researchers, ministry stakeholders — all orbiting the same PSU universe and underserved by general news platforms.'],
    ];
    audiences.forEach(([title, color, body], i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      addCard(s, t, PX + col * (aw + 0.3), 2.72 + row * 1.65, aw, 1.52, title, body, color);
    });

    addFooter(s, t);
  }

  // ── F7: Content Pipeline ───────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '07 / 10');

    addTag(s, t, 'How We Build Content', 'blue', PX, 0.78);
    addTitle(s, t, 'AI Speed. Human Accuracy. Three-Stream Pipeline.', PX, 1.08, CW, { fontSize: 19 });
    addBody(s, t,
      'Fresh, categorized, structured PSU content — at scale, around the clock. Every article is readable, categorized, and mapped to the right PSU before it reaches a single user.',
      PX, 1.72, CW, 0.3, { fontSize: 9 });

    // Pipeline
    const steps = [
      ['🤖', 'AI Scraper', 'RSS from PIB, PSU press sites, financial news — LLM auto-summarizes, classifies category, maps to PSU & sector'],
      ['✍️', 'Writer Network', 'Registered writers submit original articles. Editorial review queue. Admin approval before any article goes live.'],
      ['⚡', 'Admin Direct', 'Breaking news published instantly by editors without queue. Real-time urgency when it matters most.'],
    ];
    const pw = (CW - 0.6) / 3;
    steps.forEach(([icon, title, body], i) => {
      const x = PX + i * (pw + 0.3);
      s.addShape('roundRect', { x, y: 2.12, w: pw, h: 1.85, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.1 });
      s.addText(icon, { x, y: 2.22, w: pw, h: 0.35, fontSize: 18, align: 'center' });
      s.addText(title, { x: x + 0.1, y: 2.6, w: pw - 0.2, h: 0.26, fontSize: 10, bold: true, color: t.amber, align: 'center' });
      addBody(s, t, body, x + 0.1, 2.9, pw - 0.2, 0.9, { fontSize: 8.5, align: 'center' });
      if (i < 2) {
        s.addText('→', { x: x + pw + 0.02, y: 2.9, w: 0.26, h: 0.3, fontSize: 14, color: t.faint, align: 'center' });
      }
    });

    // 3 benefit cards
    const bw = (CW - 0.2) / 3;
    const benefits = [
      ['Scale without noise', t.greenText, 'AI handles bulk ingestion — but every AI-scraped article still goes through human review before publishing. No hallucination risk.'],
      ['Always categorized', t.blueText, 'Every article lands in one of 13 categories — so an Appointments story never reaches someone who only follows Finance.'],
      ['Always fresh', t.amber, 'Breaking news live in minutes. Scraped content same day. No PSU story sits unread in an RSS feed overnight.'],
    ];
    benefits.forEach(([title, color, body], i) => {
      addCard(s, t, PX + i * (bw + 0.1), 4.18, bw, 1.9, title, body, color);
    });

    addFooter(s, t);
  }

  // ── F8: What Has Been Built ────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '08 / 10');

    addTag(s, t, 'What Has Been Built', 'green', PX, 0.78);
    addTitle(s, t, 'Not a Concept. A Live, Working Product.', PX, 1.08, CW, { fontSize: 19 });

    // Stats
    const stats = [
      ['12,400', 'Registered users'],
      ['3,200', 'Daily active users'],
      ['850+', 'Articles published'],
      ['4.8m', 'Avg. session duration'],
    ];
    const sw = CW / 4 - 0.1;
    stats.forEach(([num, label], i) => {
      const x = PX + i * (sw + 0.13);
      s.addShape('roundRect', { x, y: 1.6, w: sw, h: 0.95, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.08 });
      addStat(s, t, x, 1.6, sw, num, label, null, t.greenText);
    });

    const col2w = CW / 2 - 0.15;
    s.addText('FULLY LIVE TODAY', { x: PX, y: 2.68, w: col2w, h: 0.18, fontSize: 7, bold: true, color: t.faint, charSpacing: 1 });
    addBullets(s, t,
      ['~390 PSU profiles populated & actively curated', '13 categories with daily editorial coverage', 'Shorts & Stories fully functional', 'Threaded comments, reactions & bookmarks live', 'Employee OTP email verification operational', 'Full dark mode across every screen & component', 'Admin CMS — articles, tickers, stories, user roles'],
      PX, 2.9, col2w, 2.3);

    s.addText('REAL ENGAGEMENT SIGNALS', { x: PX + col2w + 0.3, y: 2.68, w: col2w, h: 0.18, fontSize: 7, bold: true, color: t.faint, charSpacing: 1 });
    addBullets(s, t,
      ['Top categories: Appointments, Jobs & Finance', 'Most followed PSUs: ONGC, NTPC, SAIL, Coal India, HAL', 'Shorts avg. 6 swipes per session', '40% of users return within 48 hours', 'Comments most active on Appointments news'],
      PX + col2w + 0.3, 2.9, col2w, 2.3);

    s.addShape('roundRect', { x: PX, y: 5.36, w: CW, h: 0.42, fill: { color: t.greenBg }, line: { type: 'none' }, rectRadius: 0.08 });
    addBody(s, t, '⚠️  Replace dummy metrics with your live product data before presenting.', PX + 0.2, 5.44, CW - 0.4, 0.26, { fontSize: 8, color: t.greenText, align: 'center' });

    addFooter(s, t);
  }

  // ── F9: Competition ────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '09 / 10');

    addTag(s, t, 'Competitive Landscape', 'red', PX, 0.78);
    addTitle(s, t, 'Others Cover PSUs. None Are Built for Them.', PX, 1.08, CW, { fontSize: 19 });
    addBody(s, t,
      'The gap is not in coverage — it\'s in experience, personalization, trust, and technology. Every existing option was designed for a different primary audience.',
      PX, 1.72, CW, 0.3, { fontSize: 9 });

    // Table
    const th = 0.24;
    const colW = [3.2, 3.0, CW - 6.2];
    const startX = [PX, PX + colW[0], PX + colW[0] + colW[1]];
    const headers = ['Platform', 'Primary audience', 'PSU professional\'s experience'];
    headers.forEach((h, i) => {
      s.addShape('rect', { x: startX[i], y: 2.1, w: colW[i], h: th, fill: { color: t.surface2 }, line: { type: 'none' } });
      s.addText(h, { x: startX[i] + 0.08, y: 2.1, w: colW[i] - 0.08, h: th, fontSize: 8, bold: true, color: t.text, valign: 'middle' });
    });

    const rows = [
      ['Existing PSU media sites', 'General PSU readership', 'Outdated UI, no personalization, no AI, no mobile-first design'],
      ['Economic Times / Moneycontrol', 'Business & finance readers', 'PSU news incidental — buried under Sensex, politics, IPOs'],
      ['LinkedIn', 'Pan-professional network', 'PSU content algorithm-dependent; no structured categories'],
      ['PIB / Press Releases', 'Journalists & government', 'Dense, unreadable, zero personalization or search'],
      ['WhatsApp Groups', 'Peer networks', 'Unverified, distorted, unsearchable, no archive'],
      ['PSU Darpan ✦', 'PSU professionals (only)', 'Personalized · AI-powered · Verified · Multi-format · Native app'],
    ];
    rows.forEach(([p, a, exp], i) => {
      const ry = 2.1 + th + i * 0.6;
      const isOurs = p.includes('PSU Darpan');
      const rowBg = isOurs ? t.amberBg : (i % 2 === 0 ? t.surface : t.bg);
      [p, a, exp].forEach((cell, ci) => {
        s.addShape('rect', { x: startX[ci], y: ry, w: colW[ci], h: 0.55, fill: { color: rowBg }, line: { type: 'none' } });
        s.addText(cell, { x: startX[ci] + 0.08, y: ry + 0.04, w: colW[ci] - 0.1, h: 0.47, fontSize: isOurs ? 8 : 7.5, bold: isOurs, color: isOurs ? t.amberText : t.muted, valign: 'middle', wrap: true });
      });
    });

    // Two positioning cards
    const cw2 = CW / 2 - 0.15;
    addCard(s, t, PX, 6.0, cw2, 1.1, 'Our position', 'We\'re not competing with ET or existing PSU sites on coverage. We\'re replacing WhatsApp groups and outdated portals as the primary daily habit of PSU professionals.', t.amber);
    addCard(s, t, PX + cw2 + 0.3, 6.0, cw2, 1.1, 'Compounding moat', 'Structured taxonomy + verified employee directory + AI pipeline + multi-format UX = a product that gets harder to replicate the more it grows.', t.blueText);

    addFooter(s, t);
  }

  // ── F10: Technology ────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    slideBg(s, t);
    addTopBar(s, t, '10 / 10');

    addTag(s, t, 'How We Build — Technology', 'blue', PX, 0.78);
    addTitle(s, t, 'Built to Scale Across India\'s Entire Public Sector', PX, 1.08, CW, { fontSize: 19 });

    const col2w = CW / 2 - 0.15;
    s.addText('TECH STACK', { x: PX, y: 1.7, w: col2w, h: 0.18, fontSize: 7, bold: true, color: t.faint, charSpacing: 1 });
    addBullets(s, t,
      ['Frontend: Next.js 15 (App Router) + TypeScript', 'Styling: Tailwind CSS v4 with full dark mode', 'Auth: Firebase Phone OTP — zero friction onboarding', 'Backend: Node.js REST API', 'Database: PostgreSQL', 'AI: LLM — auto-classify, summarize, map to PSU/sector', 'Distribution: Native mobile app'],
      PX, 1.92, col2w, 2.5);

    s.addText('DESIGNED TO SCALE', { x: PX + col2w + 0.3, y: 1.7, w: col2w, h: 0.18, fontSize: 7, bold: true, color: t.faint, charSpacing: 1 });
    addBullets(s, t,
      ['PSU Browser ready for all ~390 PSUs + state expansion', 'AI pipeline built for high-volume automated ingestion', 'Role-based access: public · registered · writer · admin', 'Modular backend — new sectors without full rebuild', 'Full admin CMS: articles, stories, tickers, analytics'],
      PX + col2w + 0.3, 1.92, col2w, 2.5);

    // 3 icon cards
    const icw = (CW - 0.2) / 3;
    const icards = [
      ['⚡', 'Performance', 'Next.js image optimization. Fast even on 4G mobile networks across India.'],
      ['🔐', 'Trust & Verification', 'Firebase auth + OTP-gated employee directory + editorial human review. Verified at every layer.'],
      ['🛠️', 'Editorial Control', 'Full CMS: create, approve, reject, manage articles, tickers, stories, and user roles.'],
    ];
    icards.forEach(([icon, title, body], i) => {
      const cx = PX + i * (icw + 0.1);
      s.addShape('roundRect', { x: cx, y: 4.6, w: icw, h: 2.0, fill: { color: t.surface }, line: { color: t.border, width: 0.5 }, rectRadius: 0.1 });
      s.addText(icon, { x: cx, y: 4.7, w: icw, h: 0.4, fontSize: 20, align: 'center' });
      s.addText(title, { x: cx + 0.12, y: 5.14, w: icw - 0.24, h: 0.26, fontSize: 9.5, bold: true, color: t.blueText, align: 'center' });
      addBody(s, t, body, cx + 0.12, 5.44, icw - 0.24, 0.96, { fontSize: 8.5, align: 'center' });
    });

    addFooter(s, t);
  }

  return pptx;
}

// ── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  const out = path.join(__dirname, 'output');
  const fs  = require('fs');
  if (!fs.existsSync(out)) fs.mkdirSync(out);

  const jobs = [
    { fn: buildShortDeck, theme: 'dark',  name: 'PSU-Darpan-Short-Pitch-Dark'  },
    { fn: buildShortDeck, theme: 'light', name: 'PSU-Darpan-Short-Pitch-Light' },
    { fn: buildFullDeck,  theme: 'dark',  name: 'PSU-Darpan-Full-Deck-Dark'    },
    { fn: buildFullDeck,  theme: 'light', name: 'PSU-Darpan-Full-Deck-Light'   },
  ];

  for (const { fn, theme, name } of jobs) {
    process.stdout.write(`Building ${name}.pptx … `);
    const pptx = fn(theme);
    const file = path.join(out, `${name}.pptx`);
    await pptx.writeFile({ fileName: file });
    console.log('done');
  }

  console.log(`\nAll 4 files saved to: ${out}/`);
}

main().catch(err => { console.error(err); process.exit(1); });
