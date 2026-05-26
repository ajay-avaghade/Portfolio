const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D1117',
  hero:   '161B22',
  card:   '1C2333',
  bright: '818CF8',
  accent: '6366F1',
  green:  '34D399',
  orange: 'F97316',
  red:    'EF4444',
  white:  'F0F6FC',
  lgray:  '8B949E',
  border: '30363D',
};

const makeShadow = () => ({ type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.15 });
const W = 10, H = 5.625;

// ── SLIDE 1 — COVER ──────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  for (let i = 0; i < 5; i++) {
    const lx = -0.5 + i * 2.2;
    const lw = Math.min(3.5, W - lx - 0.15);
    if (lx > W || lw <= 0) continue;
    s.addShape(prs.ShapeType.line, { x: lx, y: 0, w: lw, h: H, line: { color: C.bright, width: 0.4 }, rotate: 30, transparency: 88 });
  }
  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.06, h: H, fill: { color: C.bright } });

  s.addShape(prs.ShapeType.roundRect, { x: 0.5, y: 0.35, w: 1.8, h: 0.28, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.06 });
  s.addText('DEEL · EQUITY PRODUCT', { x: 0.5, y: 0.35, w: 1.8, h: 0.28, color: C.bright, fontSize: 7, bold: true, align: 'center', valign: 'middle', charSpacing: 1.5 });

  s.addText('EquityConnect', { x: 0.5, y: 0.8, w: 6, h: 0.9, color: C.white, fontSize: 44, bold: true, fontFace: 'Calibri' });
  s.addText('Universal Equity Data Integration Hub', { x: 0.5, y: 1.65, w: 6.5, h: 0.42, color: C.bright, fontSize: 18, bold: true });
  s.addText('Connect any cap table platform to Deel in days, not months — with real-time sync & drift alerting', { x: 0.5, y: 2.1, w: 6.5, h: 0.35, color: C.lgray, fontSize: 12 });
  s.addText('Presented by Ajay Avaghade · Senior PM, Equity', { x: 0.5, y: 2.6, w: 5, h: 0.28, color: C.lgray, fontSize: 10 });

  s.addShape(prs.ShapeType.roundRect, { x: 7.2, y: 1.6, w: 2.4, h: 1.6, fill: { color: C.hero }, line: { color: C.bright, width: 1 }, rectRadius: 0.1, shadow: makeShadow() });
  s.addText('2 days', { x: 7.2, y: 1.75, w: 2.4, h: 0.55, color: C.bright, fontSize: 28, bold: true, align: 'center', fontFace: 'Courier New' });
  s.addText('new cap table\nintegration setup', { x: 7.2, y: 2.3, w: 2.4, h: 0.55, color: C.lgray, fontSize: 9, align: 'center', valign: 'middle' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityConnect · Confidential', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
  s.addText('ajay-avaghade.github.io/Portfolio', { x: 5.5, y: 5.38, w: 4.1, h: 0.245, color: C.bright, fontSize: 8, align: 'right', valign: 'middle' });
}

// ── SLIDE 2 — THE PROBLEM ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE PROBLEM', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("Deel Needs Equity Data From Dozens of Platforms — Each Speaking a Different Language", { x: 0.4, y: 0.44, w: 9.2, h: 0.65, color: C.white, fontSize: 22, bold: true });

  const cols = [
    { icon: '🔌', stat: '10+', label: 'Cap table platforms in market', sub: 'Carta, Shareworks, Fidelity, Pulley, Capdesk, Shareholder IntelliAgent, and more — each with unique schemas' },
    { icon: '⏳', stat: '3–6 mo', label: 'Custom integration build time', sub: 'Each point-to-point integration is a custom engineering project: schema mapping, webhook setup, reconciliation, testing' },
    { icon: '💔', stat: '40%', label: 'Integrations break annually', sub: 'Upstream schema changes go undetected until payroll items start failing. No systematic drift detection exists today' },
  ];
  cols.forEach((col, i) => {
    const x = 0.4 + i * 3.1;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.2, w: 2.9, h: 2.4, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addText(col.icon, { x, y: 1.32, w: 2.9, h: 0.4, fontSize: 22, align: 'center' });
    s.addText(col.stat, { x, y: 1.7, w: 2.9, h: 0.55, color: C.bright, fontSize: 28, bold: true, align: 'center', fontFace: 'Courier New' });
    s.addText(col.label, { x: x + 0.1, y: 2.25, w: 2.7, h: 0.3, color: C.white, fontSize: 10, bold: true, align: 'center' });
    s.addText(col.sub, { x: x + 0.1, y: 2.55, w: 2.7, h: 0.7, color: C.lgray, fontSize: 8, align: 'center', valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 3.75, w: 9.2, h: 0.65, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('Root cause: Every new cap table platform requires a custom integration — fragile, slow to build, and invisible to maintain. A schema change at Carta or Shareworks can silently corrupt equity data until the next payroll run.', {
    x: 0.6, y: 3.78, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityConnect', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 3 — THE INSIGHT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE INSIGHT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('A Canonical Equity Model + Schema Normalisation Layer Makes Every Platform Pluggable', { x: 0.4, y: 0.44, w: 9.2, h: 0.55, color: '111111', fontSize: 21, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 1.1, w: 4.1, h: 3.4, fill: { color: 'FFF5F5' }, line: { color: 'FECACA', width: 0.8 }, rectRadius: 0.1 });
  s.addText('❌  Point-to-Point (Current)', { x: 0.6, y: 1.22, w: 3.7, h: 0.3, color: 'DC2626', fontSize: 11, bold: true });
  const leftPts = [
    'Each platform requires a bespoke integration: Carta schema ≠ Shareworks schema ≠ Fidelity schema',
    'Changes to upstream platform schema go undetected until payroll fails',
    'New platform onboarding: 3–6 month engineering project each time',
    'No single source of truth — equity data fragmented across platforms',
    'Schema evolution is invisible — no alerting, no auto-resolution',
  ];
  leftPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 0.6, y: 1.6 + i * 0.44, w: 3.7, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  s.addShape(prs.ShapeType.ellipse, { x: 4.55, y: 2.4, w: 0.6, h: 0.45, fill: { color: C.dark }, line: { color: C.dark, width: 0 } });
  s.addText('VS', { x: 4.55, y: 2.4, w: 0.6, h: 0.45, color: C.white, fontSize: 9, bold: true, align: 'center', valign: 'middle' });

  s.addShape(prs.ShapeType.roundRect, { x: 5.2, y: 1.1, w: 4.4, h: 3.4, fill: { color: 'F5F3FF' }, line: { color: 'C4B5FD', width: 0.8 }, rectRadius: 0.1 });
  s.addText('✅  EquityConnect (Proposed)', { x: 5.4, y: 1.22, w: 4.0, h: 0.3, color: '7C3AED', fontSize: 11, bold: true });
  const rightPts = [
    'Single canonical equity data model: employee_id, equity_type, vesting_date, vested_units, fair_market_value, country_iso, withholding_usd',
    'AI schema mapping wizard: new platform field mapping in <2 hours with AI suggestions confirmed by PM',
    'Real-time webhook + schema drift detection: field changes detected immediately, auto-resolution queued',
    'New platform onboarding: 2-day connector setup instead of 3-6 month build',
    'All equity data normalised into one queryable layer for EquityPulse, EquityOnboard, TotalRewards',
  ];
  rightPts.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.4, y: 1.6 + i * 0.44, w: 4.0, h: 0.38, color: '333333', fontSize: 9, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.4, y: 4.62, w: 9.2, h: 0.65, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.08 });
  s.addText("Key insight: All cap table platforms store the same 8–12 core equity fields — they just name them differently. A normalisation layer built once unlocks every platform automatically, and makes schema drift detectable in real time.", {
    x: 0.6, y: 4.65, w: 8.8, h: 0.6, color: C.white, fontSize: 9, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityConnect', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 4 — THE MECHANIC ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('THE MECHANIC', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Connect Any Platform in 3 Steps', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 24, bold: true });

  const steps = [
    { n: '1', title: 'Platform Auth', body: 'OAuth / API key connection to cap table platform. Parte of pre-built connector library (Carta, Shareworks, Fidelity, Pulley). New platforms via generic REST connector.' },
    { n: '2', title: 'AI Schema Mapping', body: 'AI engine maps source fields to Deel canonical equity model (employee_id, equity_type, vesting_date, vested_units, FMV, country, withholding). PM confirms or edits in 2-hour wizard.' },
    { n: '3', title: 'Live Sync + Monitoring', body: 'Real-time webhook fires on equity events. Reconciliation job runs nightly. Schema drift engine compares current schema against baseline — alerts on any field change with AI-suggested resolution.' },
    { n: '4', title: 'Drift Auto-Resolve', body: 'When upstream schema changes (new field, renamed enum, type change), drift engine detects, classifies impact level (critical / warning / info), and suggests mapping fix. PM approves in 1 click.' },
    { n: '5', title: 'EquityPulse Unlock', body: 'Once connected, all vesting events auto-flow into EquityPulse orchestration engine. Zero additional config. New countries and worker types automatically covered as platform data expands.' },
  ];

  steps.forEach((step, i) => {
    const x = 0.38 + i * 1.88;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.05, w: 1.72, h: 3.4, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.ellipse, { x: x + 0.62, y: 1.12, w: 0.46, h: 0.38, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(step.n, { x: x + 0.62, y: 1.12, w: 0.46, h: 0.38, color: C.dark, fontSize: 12, bold: true, align: 'center', valign: 'middle' });
    s.addText(step.title, { x: x + 0.06, y: 1.58, w: 1.6, h: 0.38, color: C.white, fontSize: 10, bold: true, align: 'center' });
    s.addText(step.body, { x: x + 0.08, y: 2.0, w: 1.56, h: 1.9, color: C.lgray, fontSize: 8, align: 'center', valign: 'top' });
    if (i < 4) s.addShape(prs.ShapeType.line, { x: x + 1.72, y: 2.2, w: 0.16, h: 0, line: { color: C.bright, width: 1.2 } });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.38, y: 4.56, w: 9.24, h: 0.64, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('PhonePe proof: Built a self-serve multi-tenant PG integration platform that onboarded 5,000+ B2B merchants from scratch — the same connector + schema normalisation pattern applied at merchant scale. EquityConnect applies this to equity data platforms.', {
    x: 0.56, y: 4.58, w: 9.0, h: 0.6, color: C.white, fontSize: 8.5, valign: 'middle'
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityConnect', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 5 — THE PRODUCT ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('THE PRODUCT', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('4 Key Screens — Connect, Map, Monitor, Resolve', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  const cards = [
    { n: '01', title: 'Integration Hub', body: 'Connected platforms at a glance: Carta (live), Shareworks (alert), Fidelity (in setup). Last sync time, record count, drift status per connector.' },
    { n: '02', title: 'Schema Mapping Wizard', body: '3-step flow: platform auth → AI field mapping (12 of 15 suggested) → confirm. New platform live in under 2 hours. AI suggests mappings based on field name similarity and existing integrations.' },
    { n: '03', title: 'Connection Live', body: 'Confirmation: X records ingested, vesting schedules mapped, employee profiles linked, drift detection active. EquityPulse automatically unlocked.' },
    { n: '04', title: 'Drift Alert Console', body: 'Schema change detected: field name, before/after values, impact scope. AI resolution suggestion (e.g., map new enum PSU → performance_share). 1-click approve and resume sync.' },
  ];

  cards.forEach((card, i) => {
    const x = 0.38 + i * 2.38;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.08, w: 2.2, h: 3.55, fill: { color: 'F9FAFB' }, line: { color: 'E5E7EB', width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.08, w: 2.2, h: 0.06, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(card.n, { x, y: 1.16, w: 2.2, h: 0.35, color: C.bright, fontSize: 20, bold: true, align: 'center', fontFace: 'Courier New' });
    s.addText(card.title, { x: x + 0.1, y: 1.5, w: 2.0, h: 0.34, color: '111111', fontSize: 11, bold: true, align: 'center' });
    s.addText(card.body, { x: x + 0.1, y: 1.88, w: 2.0, h: 2.4, color: '555555', fontSize: 8.5, align: 'left', valign: 'top' });
  });

  s.addText('Interactive prototype: deel-equityconnect-prototype.html', { x: 0.4, y: 4.74, w: 9.2, h: 0.3, color: '777777', fontSize: 9, align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityConnect', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 6 — IMPACT & ROI ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('IMPACT & ROI', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('From 3-Month Integration Projects to 2-Day Connector Setup', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 22, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('ENGINEERING IMPACT', { x: 0.5, y: 1.15, w: 4.2, h: 0.22, color: C.bright, fontSize: 8, bold: true, charSpacing: 1.5 });
  const leftMetrics = [
    { val: '3 mo → 2 days', lbl: 'New cap table integration setup time' },
    { val: '−40%', lbl: 'Annual integration break rate (drift detection)' },
    { val: '1 platform', lbl: 'Canonical data model — no more ad-hoc schemas' },
    { val: '< 2 hrs', lbl: 'Schema mapping wizard — AI-assisted, PM-owned' },
  ];
  leftMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 0.5, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 0.6, y: y + 0.04, w: 4.0, h: 0.3, color: C.bright, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 0.6, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.55, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1 });
  s.addText('BUSINESS IMPACT', { x: 5.25, y: 1.15, w: 4.2, h: 0.22, color: C.orange, fontSize: 8, bold: true, charSpacing: 1.5 });
  const rightMetrics = [
    { val: '10+', lbl: 'Cap table platforms connected via one framework' },
    { val: '$500K+', lbl: 'Engineering cost saved per integration avoided' },
    { val: 'Real-time', lbl: 'Equity data freshness across all platforms' },
    { val: 'Zero', lbl: 'Payroll failures from undetected schema drift' },
  ];
  rightMetrics.forEach((m, i) => {
    const y = 1.45 + i * 0.76;
    s.addShape(prs.ShapeType.roundRect, { x: 5.25, y, w: 4.2, h: 0.6, fill: { color: C.card }, line: { color: C.border, width: 0.4 }, rectRadius: 0.07 });
    s.addText(m.val, { x: 5.35, y: y + 0.04, w: 4.0, h: 0.3, color: C.orange, fontSize: 16, bold: true, fontFace: 'Courier New' });
    s.addText(m.lbl, { x: 5.35, y: y + 0.34, w: 4.0, h: 0.22, color: C.lgray, fontSize: 8.5 });
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Deel · EquityConnect', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: C.lgray, fontSize: 8, valign: 'middle' });
}

// ── SLIDE 7 — PROOF OF WORK ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: 'FFFFFF' };

  s.addText('PROOF OF WORK', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: '777777', fontSize: 8, bold: true, charSpacing: 2 });
  s.addText("I Built This Multi-Tenant Integration Architecture Before.", { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: '111111', fontSize: 22, bold: true });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 1.05, w: 4.55, h: 3.5, fill: { color: C.dark }, line: { color: C.border, width: 0 }, rectRadius: 0.1 });
  s.addText('What I Built at PhonePe', { x: 0.5, y: 1.15, w: 4.2, h: 0.28, color: C.bright, fontSize: 10, bold: true });
  const leftPoints = [
    'Zero-to-one self-serve PG integration platform: multi-tenant connector for 5,000+ B2B merchants (Instant Discount + EMI subvention)',
    'Schema normalisation: mapped each merchant\'s transaction schema to PhonePe\'s canonical model — same pattern as EquityConnect',
    'Rebuilt ops workflow from 2-day manual onboarding to 30-minute self-serve (TAT reduction: 97%)',
    'Compliance + Legal alignment: built consent flows, T&C frameworks, and regulatory checks into the onboarding UX',
    '23% lower CAC and clear pre/post KPI baselines validating each release',
  ];
  leftPoints.forEach((pt, i) => {
    s.addText(`${i + 1}. ${pt}`, { x: 0.5, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: C.white, fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 5.1, y: 1.05, w: 4.55, h: 3.5, fill: { color: 'F5F3FF' }, line: { color: 'C4B5FD', width: 0.8 }, rectRadius: 0.1 });
  s.addText('How It Maps to This Role', { x: 5.25, y: 1.15, w: 4.2, h: 0.28, color: '7C3AED', fontSize: 10, bold: true });
  const rightPoints = [
    'PG integration platform → EquityConnect: same multi-tenant connector architecture, applied to cap table platforms instead of payment gateways',
    'Schema normalisation at PhonePe → Canonical equity data model: identical technical problem — map upstream schema to internal canonical',
    '30-min onboarding TAT → 2-day platform setup: same self-serve wizard pattern',
    'Compliance + T&C flows → Equity platform legal alignment: experienced with regulated domain integration',
    'KPI instrumentation from day 0 → Equity product success metrics pre-defined before build',
  ];
  rightPoints.forEach((pt, i) => {
    s.addText(`• ${pt}`, { x: 5.25, y: 1.5 + i * 0.58, w: 4.2, h: 0.52, color: '333333', fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 4.65, w: 9.3, h: 0.5, fill: { color: C.dark }, line: { color: C.dark, width: 0 }, rectRadius: 0.07 });
  s.addText('"The self-serve PG integration platform I built at PhonePe is EquityConnect — the domain changed, the architecture didn\'t."', { x: 0.5, y: 4.67, w: 9.0, h: 0.46, color: C.bright, fontSize: 9.5, italic: true, valign: 'middle', align: 'center' });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: 'F3F4F6' } });
  s.addText('Deel · EquityConnect', { x: 0.4, y: 5.38, w: 5, h: 0.245, color: '777777', fontSize: 8, valign: 'middle' });
}

// ── SLIDE 8 — ROLLOUT PLAN ──────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.dark };

  s.addText('ROLLOUT PLAN', { x: 0.4, y: 0.22, w: 9, h: 0.22, color: C.lgray, fontSize: 8, bold: true, charSpacing: 2 });
  s.addText('Start With Top 3 Platforms, Build Universal Framework', { x: 0.4, y: 0.44, w: 9.2, h: 0.5, color: C.white, fontSize: 22, bold: true });

  const phases = [
    { ph: 'Phase 1', period: 'Months 1–2', title: 'Core Connectors (Carta + Shareworks)', body: 'Build canonical equity data model. Implement Carta + Shareworks connectors. Launch AI schema mapping wizard. Instrument integration health KPIs.' },
    { ph: 'Phase 2', period: 'Months 3–4', title: 'Fidelity + Drift Engine', body: 'Add Fidelity connector. Build schema drift detection and auto-resolution engine. Launch ops monitoring dashboard. A/B test wizard completion rate.' },
    { ph: 'Phase 3', period: 'Months 5–6', title: 'Universal Framework GA', body: 'Generic REST connector for any platform. Self-serve connector builder for Deel engineering team. Full EquityPulse integration. Connect to TotalRewards layer.' },
  ];

  phases.forEach((phase, i) => {
    const x = 0.35 + i * 3.2;
    s.addShape(prs.ShapeType.roundRect, { x, y: 1.05, w: 3.0, h: 2.85, fill: { color: C.hero }, line: { color: C.border, width: 0.6 }, rectRadius: 0.1, shadow: makeShadow() });
    s.addShape(prs.ShapeType.rect, { x, y: 1.05, w: 3.0, h: 0.05, fill: { color: C.bright }, line: { color: C.bright, width: 0 } });
    s.addText(phase.ph, { x: x + 0.12, y: 1.12, w: 2.76, h: 0.26, color: C.bright, fontSize: 9, bold: true });
    s.addText(phase.period, { x: x + 0.12, y: 1.37, w: 2.76, h: 0.22, color: C.lgray, fontSize: 8 });
    s.addText(phase.title, { x: x + 0.12, y: 1.6, w: 2.76, h: 0.32, color: C.white, fontSize: 11, bold: true });
    s.addText(phase.body, { x: x + 0.12, y: 1.96, w: 2.76, h: 1.6, color: C.lgray, fontSize: 8.5, valign: 'top' });
  });

  s.addShape(prs.ShapeType.roundRect, { x: 0.35, y: 4.05, w: 9.3, h: 0.72, fill: { color: C.hero }, line: { color: C.bright, width: 0.8 }, rectRadius: 0.08 });
  s.addText('What I need to build this:', { x: 0.55, y: 4.1, w: 2.5, h: 0.24, color: C.bright, fontSize: 9, bold: true });
  s.addText('API documentation from Carta, Shareworks, Fidelity  ·  Alignment on canonical equity data model with Data Engineering  ·  1 backend engineer (connector framework)  ·  Compliance team input on required fields per country', {
    x: 0.55, y: 4.34, w: 9.0, h: 0.4, color: C.white, fontSize: 8.5
  });

  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.38, w: W, h: 0.245, fill: { color: C.hero } });
  s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', { x: 0.4, y: 5.38, w: 9.2, h: 0.245, color: C.bright, fontSize: 8.5, align: 'center', valign: 'middle' });
}

prs.writeFile({ fileName: 'assets/deel-equityconnect-deck.pptx' })
  .then(() => console.log('✅ deel-equityconnect-deck.pptx saved'))
  .catch(e => console.error(e));
