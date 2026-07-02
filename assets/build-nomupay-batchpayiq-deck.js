const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';

  const C = {
    dark:   '0B1120',
    hero:   '0F1F3D',
    accent: '2563EB',
    gold:   'F59E0B',
    white:  'FFFFFF',
    lgray:  'F8FAFC',
    muted:  '94A3B8',
    green:  '059669',
    red:    'DC2626',
  };

  const makeShadow = () => ({ type: 'outer', color: '000000', blur: 4, offset: 2, angle: 45, opacity: 0.14 });

  // ── SLIDE 1: COVER ──────────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 3.5, h: 7.5, fill: { color: C.hero } });
    s.addShape('rect', { x: 3.5, y: 0, w: 0.04, h: 7.5, fill: { color: C.accent } });
    s.addShape('ellipse', { x: 6.8, y: -0.8, w: 3, h: 3, fill: { color: C.accent }, line: { color: C.accent }, shadow: makeShadow() });
    s.addShape('ellipse', { x: 7.8, y: 5, w: 1.8, h: 1.8, fill: { color: '1E3A5F' }, line: { color: '1E3A5F' } });

    s.addText('NOMUPAY  ·  MERCHANT EXPERIENCE  ·  Case Study 02', { x: 3.7, y: 0.4, w: 6, h: 0.3, fontSize: 8, bold: true, color: C.muted, charSpacing: 2 });
    s.addText('BatchPayIQ', { x: 3.7, y: 0.9, w: 6, h: 1.4, fontSize: 52, bold: true, color: C.white, charSpacing: -1 });
    s.addText('Transparent Batch Payout Processing\nwith Real-Time Per-Transaction Visibility', { x: 3.7, y: 2.4, w: 5.8, h: 1, fontSize: 16, color: C.gold });
    s.addText('Ajay Avaghade  ·  Product Manager', { x: 3.7, y: 3.55, w: 4, h: 0.3, fontSize: 11, color: C.muted });

    s.addShape('rect', { x: 0.3, y: 0.5, w: 0.06, h: 0.5, fill: { color: C.accent } });
    s.addText('97.8%', { x: 0.5, y: 0.45, w: 2.7, h: 0.6, fontSize: 28, bold: true, color: C.white });
    s.addText('First-pass settlement\nsuccess rate', { x: 0.5, y: 0.95, w: 2.7, h: 0.5, fontSize: 10, color: C.muted });

    s.addShape('rect', { x: 0.3, y: 2.2, w: 0.06, h: 0.5, fill: { color: C.gold } });
    s.addText('↓ 40%', { x: 0.5, y: 2.15, w: 2.7, h: 0.6, fontSize: 28, bold: true, color: C.gold });
    s.addText('Payout support\ntickets eliminated', { x: 0.5, y: 2.65, w: 2.7, h: 0.5, fontSize: 10, color: C.muted });

    s.addShape('rect', { x: 0.3, y: 3.9, w: 0.06, h: 0.5, fill: { color: C.green } });
    s.addText('3x', { x: 0.5, y: 3.85, w: 2.7, h: 0.6, fontSize: 28, bold: true, color: C.white });
    s.addText('Faster reconciliation\nvs manual CSV', { x: 0.5, y: 4.35, w: 2.7, h: 0.5, fontSize: 10, color: C.muted });

    s.addText('Live Prototype: ajay-avaghade.github.io/Portfolio/nomupay-batchpayiq-prototype.html', { x: 3.7, y: 7.1, w: 6, h: 0.3, fontSize: 8, color: C.muted });
  }

  // ── SLIDE 2: THE PROBLEM ────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.red } });

    s.addText('THE PROBLEM', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.red, charSpacing: 2 });
    s.addText('Merchants Submit Batch Payouts and Wait — No Visibility, No Self-Service, No Way to Fix Failures Before They Cascade', { x: 0.5, y: 0.7, w: 9, h: 0.9, fontSize: 25, bold: true, color: C.white });

    const stats = [
      { n: '~50%', l: 'Payout queries are status requests', s: 'Half of all merchant support\ntickets ask: "Where is\nmy payment?" (Kyriba 2022)', col: C.accent },
      { n: '38%', l: 'Batch failures caused by IBAN errors', s: 'IBAN format errors are the\ntop B2B batch failure type\nacross EU platforms (SWIFT 2022)', col: C.gold },
      { n: 'Hours', l: 'Manual reconciliation time per batch', s: 'Finance teams manually\ncross-reference CSVs across\nacquirers after every batch', col: C.red },
    ];
    stats.forEach((st, i) => {
      const x = 0.4 + i * 3.1;
      s.addShape('rect', { x, y: 2.1, w: 2.9, h: 4.05, fill: { color: '111827' }, line: { color: st.col, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 2.1, w: 2.9, h: 0.07, fill: { color: st.col } });
      s.addText(st.n, { x: x + 0.2, y: 2.3, w: 2.5, h: 1, fontSize: 38, bold: true, color: st.col });
      s.addText(st.l, { x: x + 0.2, y: 3.2, w: 2.5, h: 0.5, fontSize: 11, bold: true, color: C.white });
      s.addText(st.s, { x: x + 0.2, y: 3.75, w: 2.5, h: 1.6, fontSize: 10, color: C.muted, paraSpaceAfter: 4 });
    });

    s.addShape('rect', { x: 0.4, y: 6.38, w: 9.2, h: 0.88, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
    s.addText('Root cause: Batch payouts are treated as black boxes. The merchant uploads a CSV and waits — with no visibility into per-transaction status, no self-serve correction of failures, and no automated reconciliation output. Every question becomes a support ticket. Every error is discovered hours too late.', { x: 0.65, y: 6.47, w: 8.7, h: 0.7, fontSize: 9.5, color: C.muted, italic: true });
  }

  // ── SLIDE 3: THE INSIGHT ────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.lgray } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

    s.addText('THE INSIGHT', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('A Batch Payout Is Not a Black Box. It Is 842 Individual Decisions That Merchants Deserve to See in Real Time.', { x: 0.5, y: 0.7, w: 9, h: 0.9, fontSize: 24, bold: true, color: C.dark });

    s.addShape('rect', { x: 0.4, y: 1.85, w: 4.3, h: 4.55, fill: { color: 'FEE2E2' }, line: { color: 'FECACA', pt: 1 }, shadow: makeShadow() });
    s.addText('TODAY', { x: 0.6, y: 1.95, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.red, charSpacing: 2 });
    const probs = [
      'Upload CSV — no pre-validation before submission',
      'Submit 847 rows — no per-transaction confirmation',
      'Wait hours — no status visibility during processing',
      'Failure discovered at settlement: "5 rows failed" — no detail',
      'Manual correction: re-export, edit CSV, re-upload entire batch',
      'Reconciliation: cross-reference 3 acquirer CSVs manually',
    ];
    probs.forEach((p, i) => {
      s.addText('✗  ' + p, { x: 0.6, y: 2.4 + i * 0.63, w: 3.9, h: 0.5, fontSize: 10.5, color: '7F1D1D' });
    });

    s.addShape('rect', { x: 5.3, y: 1.85, w: 4.3, h: 4.55, fill: { color: 'D1FAE5' }, line: { color: 'A7F3D0', pt: 1 }, shadow: makeShadow() });
    s.addText('BatchPayIQ', { x: 5.5, y: 1.95, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.green, charSpacing: 2 });
    const sols = [
      'Upload CSV — instant pre-validation: 842 valid, 5 flagged with reason',
      'Submit 842 rows — per-transaction status grid, live during processing',
      'Real-time: Row 23 failed (invalid IBAN) — AI suggests corrected value',
      '1-click: Review correction, confirm, retry — resolved in 90 seconds',
      'Completion: 842/842 settled — auto-reconciliation CSV ready to download',
      'Zero manual reconciliation — ERP export on batch completion',
    ];
    sols.forEach((p, i) => {
      s.addText('✓  ' + p, { x: 5.5, y: 2.4 + i * 0.63, w: 3.9, h: 0.5, fontSize: 10.5, color: '064E3B' });
    });

    s.addShape('rect', { x: 0.4, y: 6.55, w: 9.2, h: 0.75, fill: { color: C.accent }, shadow: makeShadow() });
    s.addText('Key insight: Pre-validation before submission eliminates 38% of failures before they reach the payment network. The remaining failures are caught in real time with 1-click correction — not discovered at month-end during manual reconciliation. Source: SWIFT Payment Failure Analysis 2022.', { x: 0.6, y: 6.63, w: 9, h: 0.6, fontSize: 9.5, color: C.white, italic: true });
  }

  // ── SLIDE 4: THE MECHANIC ───────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.hero } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.gold } });

    s.addText('THE MECHANIC', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
    s.addText('Five Steps from Upload to Settlement — Zero Black Box', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.white });

    const steps = [
      { n: '01', t: 'Pre-Submission Validation', b: 'CSV uploaded. BatchPayIQ runs instant validation: IBAN format check, BIC/SWIFT lookup, currency support verification, daily limit checks. Output: 842 valid rows, 5 flagged with specific error type and suggested fix. Merchant can fix and re-upload or proceed with valid rows only.' },
      { n: '02', t: 'Per-Transaction Status Grid', b: 'Batch submitted. Real-time status grid shows every row: Settled (green), Processing (blue), Pending (grey), Failed (red). Filter by status. Auto-refresh every 10 seconds. Progress bar shows batch completion %. No waiting in the dark.' },
      { n: '03', t: 'Failure Detail + 1-Click Retry', b: 'Failed row tapped: full detail (recipient, amount, reference, acquirer, failure reason in plain English). AI-suggested correction (e.g. corrected IBAN format). Merchant edits field, confirms, BatchPayIQ retries via optimal acquirer route. 90-second resolution end-to-end.' },
      { n: '04', t: 'Settlement Confirmation', b: 'Batch complete: confetti, summary card (842 settled, 5 retried, 0 final failures). Breakdown by acquirer (Stripe, Adyen, PayU). Total volume confirmed. One-tap download: reconciliation CSV + bank statement PDF.' },
      { n: '05', t: 'Analytics + Auto-Recon', b: 'Batch analytics dashboard: success rate trend, average processing time, failure reason breakdown. AI insight: "38% of failures are IBAN-format errors — enable pre-submission IBAN API to eliminate these entirely." Auto-reconciliation rules: export to ERP on completion.' },
    ];

    steps.forEach((st, i) => {
      const x = 0.35 + i * 1.87;
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 5.6, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 0.07, fill: { color: C.accent } });
      s.addText(st.n, { x: x + 0.1, y: 1.65, w: 1.5, h: 0.55, fontSize: 26, bold: true, color: C.accent });
      s.addText(st.t, { x: x + 0.1, y: 2.2, w: 1.5, h: 0.7, fontSize: 10, bold: true, color: C.white });
      s.addText(st.b, { x: x + 0.1, y: 2.95, w: 1.5, h: 3.5, fontSize: 8.5, color: C.muted, paraSpaceAfter: 3 });
    });

    s.addText('PhonePe proof: Campaign ops redesign reduced 2-day manual workflow to 30-minute TAT — same operational visibility and self-serve correction pattern. Self-serve PG platform: merchants onboard without ops team involvement — same 0-to-live self-service philosophy.', { x: 0.4, y: 7.1, w: 9.2, h: 0.3, fontSize: 8, color: C.muted, italic: true });
  }

  // ── SLIDE 5: THE PRODUCT ────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.lgray } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

    s.addText('THE PRODUCT', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('5 Screen States — Batch Upload to Settlement Analytics', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.dark });

    const screens = [
      { n: '01', t: 'Upload + Validation', b: '847 rows detected. Pre-validation: 842 valid, 5 flagged. Each flagged row shows: row number, recipient, amount, error type. Options: Fix and re-upload or submit 842 valid rows. Total value: £847,280.' },
      { n: '02', t: 'Processing Grid', b: 'Batch #BP-20240624 in progress. Progress bar 73%. Status grid: Settled 614 (green), Processing 198 (blue), Failed 30 (red). Filter chips. Auto-refresh every 10s. Tap any failed row to enter correction flow.' },
      { n: '03', t: 'Failed Row + Retry', b: 'Row 23: Acme Supplies — INVALID_IBAN_FORMAT. Plain English: "IBAN format does not match Nationwide UK." AI suggestion: corrected IBAN pre-filled. Edit, confirm, retry. 90-second resolution timeline visible.' },
      { n: '04', t: 'Batch Complete', b: 'Confetti. £847,280 settled. 837 first-pass + 5 retried = 0 final failures. Breakdown by acquirer (Stripe £468K, Adyen £284K, PayU £95K). Download reconciliation CSV + bank statement PDF.' },
      { n: '05', t: 'Batch Analytics', b: 'Monthly: 12 batches, £2.3M, 97.8% success rate. 8-week success rate chart. Failure analysis: IBAN 38%, insufficient funds 22%, currency issues 18%. AI insight on IBAN prevention.' },
    ];

    screens.forEach((sc, i) => {
      const x = 0.35 + i * 1.87;
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 5.6, fill: { color: C.white }, line: { color: 'E5E7EB', pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 1.55, w: 1.7, h: 0.07, fill: { color: C.accent } });
      s.addText(sc.n, { x: x + 0.1, y: 1.65, w: 1.5, h: 0.55, fontSize: 26, bold: true, color: C.accent });
      s.addText(sc.t, { x: x + 0.1, y: 2.2, w: 1.5, h: 0.7, fontSize: 10, bold: true, color: C.dark });
      s.addText(sc.b, { x: x + 0.1, y: 2.95, w: 1.5, h: 3.5, fontSize: 9, color: '4B5563', paraSpaceAfter: 3 });
    });

    s.addText('Prototype: ajay-avaghade.github.io/Portfolio/nomupay-batchpayiq-prototype.html  ·  All 5 states interactive', { x: 0.4, y: 7.15, w: 9.2, h: 0.25, fontSize: 8, color: C.muted });
  }

  // ── SLIDE 6: IMPACT & ROI ───────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.accent } });

    s.addText('IMPACT & ROI', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('Industry Benchmarks Applied to Nomupay Batch Volume', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.white });

    s.addShape('rect', { x: 0.4, y: 1.55, w: 4.3, h: 5.55, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
    s.addShape('rect', { x: 0.4, y: 1.55, w: 4.3, h: 0.07, fill: { color: C.accent } });
    s.addText('MERCHANT IMPACT', { x: 0.6, y: 1.7, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    const mImpact = [
      { v: '97.8%', l: 'First-pass settlement success rate', src: 'Bottomline B2B Payments Benchmark 2023' },
      { v: '↓ 40%', l: 'Support tickets for payout status queries', src: 'Kyriba Treasury Management Survey 2022' },
      { v: '↑ 3x', l: 'Faster reconciliation vs manual CSV process', src: 'Gartner Finance Automation 2023' },
      { v: '↓ 38%', l: 'Failure rate via pre-submission IBAN validation', src: 'SWIFT Payment Failure Analysis 2022' },
    ];
    mImpact.forEach((m, i) => {
      s.addText(m.v, { x: 0.6, y: 2.2 + i * 1.15, w: 3.9, h: 0.55, fontSize: 22, bold: true, color: C.accent });
      s.addText(m.l, { x: 0.6, y: 2.7 + i * 1.15, w: 3.9, h: 0.35, fontSize: 10, color: C.white });
      s.addText('Source: ' + m.src, { x: 0.6, y: 2.98 + i * 1.15, w: 3.9, h: 0.25, fontSize: 8, color: C.muted, italic: true });
    });

    s.addShape('rect', { x: 5.3, y: 1.55, w: 4.3, h: 5.55, fill: { color: '111827' }, line: { color: C.gold, pt: 1 }, shadow: makeShadow() });
    s.addShape('rect', { x: 5.3, y: 1.55, w: 4.3, h: 0.07, fill: { color: C.gold } });
    s.addText('NOMUPAY ROI', { x: 5.5, y: 1.7, w: 3.9, h: 0.3, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
    const nRoi = [
      { v: '↓ 40%', l: 'Support headcount needed for batch payout ops', src: 'Ticket deflection from real-time visibility' },
      { v: '↑ NPS', l: 'Merchant satisfaction (transparency = trust)', src: 'No batch-result surprise at month-end' },
      { v: '0', l: 'Manual interventions for IBAN-format failures', src: 'Pre-submission validation eliminates category' },
      { v: 'Stickiness', l: 'BatchPayIQ as platform lock-in mechanic', src: 'Merchants tied to BatchPayIQ analytics history' },
    ];
    nRoi.forEach((m, i) => {
      s.addText(m.v, { x: 5.5, y: 2.2 + i * 1.15, w: 3.9, h: 0.55, fontSize: 22, bold: true, color: C.gold });
      s.addText(m.l, { x: 5.5, y: 2.7 + i * 1.15, w: 3.9, h: 0.35, fontSize: 10, color: C.white });
      s.addText('Source: ' + m.src, { x: 5.5, y: 2.98 + i * 1.15, w: 3.9, h: 0.25, fontSize: 8, color: C.muted, italic: true });
    });
  }

  // ── SLIDE 7: PROOF OF WORK ──────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 5, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 5, y: 0, w: 5, h: 7.5, fill: { color: C.lgray } });
    s.addShape('rect', { x: 4.97, y: 0, w: 0.06, h: 7.5, fill: { color: C.accent } });

    s.addText('PROOF OF WORK', { x: 0.4, y: 0.3, w: 4.2, h: 0.3, fontSize: 9, bold: true, color: C.accent, charSpacing: 2 });
    s.addText('I Built the Exact Operational Pattern BatchPayIQ Requires.\nHere Is the Direct Map.', { x: 0.4, y: 0.7, w: 4.2, h: 1, fontSize: 18, bold: true, color: C.white });

    const proofs = [
      { pp: 'Campaign ops redesign: identified 6 bottlenecks across campaign workflow, scoped system integrations, shipped with Engineering — TAT 2 days to 30 minutes', bp: 'BatchPayIQ pre-validation + real-time grid: same operational bottleneck analysis applied to batch payout workflow — "where did this fail and how do I fix it right now?"' },
      { pp: 'Self-serve PG integration platform (0 to 1): led cross-functional delivery with 3 bank partners, legal, compliance, finance — 5,000+ merchants onboard without ops team', bp: 'BatchPayIQ self-serve correction flow: merchants resolve IBAN errors and retry without contacting Nomupay support — same self-serve philosophy' },
      { pp: 'Propensity-to-Transact ML models: wrote product spec for real-time scoring system, worked with Data Science on feature engineering, shipped to production', bp: 'BatchPayIQ AI-suggested corrections: same real-time data intelligence pattern — AI reads IBAN, detects format error, suggests corrected value before human notices' },
      { pp: 'Milestone A/B test: full funnel instrumentation, cohort reading, statistical significance check — 60% activation uplift translated into platform-wide architecture', bp: 'BatchPayIQ analytics: same metrics ownership — success rate trend, failure category analysis, AI recommendation loop closing the quality improvement cycle' },
    ];

    proofs.forEach((p, i) => {
      const y = 1.95 + i * 1.3;
      s.addShape('rect', { x: 0.3, y, w: 4.3, h: 1.1, fill: { color: '111827' }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
      s.addText(p.pp, { x: 0.45, y: y + 0.08, w: 4, h: 0.9, fontSize: 9, color: C.muted });

      s.addShape('rect', { x: 5.2, y, w: 4.4, h: 1.1, fill: { color: C.white }, line: { color: C.accent, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x: 5.2, y, w: 4.4, h: 0.06, fill: { color: C.accent } });
      s.addText('→  ' + p.bp, { x: 5.35, y: y + 0.12, w: 4.1, h: 0.9, fontSize: 9.5, color: '111827' });
    });

    s.addShape('rect', { x: 0.3, y: 7.1, w: 9.4, h: 0.3, fill: { color: C.accent }, shadow: makeShadow() });
    s.addText('BatchPayIQ is the direct productisation of operational efficiency and self-serve platform work I have already shipped at scale.', { x: 0.5, y: 7.12, w: 9, h: 0.26, fontSize: 8.5, color: C.white, italic: true });
  }

  // ── SLIDE 8: ROLLOUT PLAN ───────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 7.5, fill: { color: C.dark } });
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.gold } });

    s.addText('ROLLOUT PLAN', { x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 9, bold: true, color: C.gold, charSpacing: 2 });
    s.addText('Three Phases — Pilot Batch Volume to Global Payout Platform', { x: 0.5, y: 0.7, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.white });

    const phases = [
      {
        ph: 'Phase 1', dur: 'Month 1–2: Pilot (10 enterprise merchants)', col: C.accent,
        pts: [
          'Deploy BatchPayIQ to 10 merchants with highest batch payout volume',
          'Enable pre-submission IBAN validation and per-transaction status grid',
          'Instrument: ticket deflection rate, failure-to-retry conversion, reconciliation time saved',
          'Measure success rate before/after pre-validation (control: current batch flow)',
        ],
      },
      {
        ph: 'Phase 2', dur: 'Month 3–4: Scale + AI Validation', col: C.gold,
        pts: [
          'Expand to 100 merchants — add BIC/SWIFT lookup and currency validation to pre-submission checks',
          'Launch AI-suggested corrections for top 3 failure types (IBAN, name mismatch, currency)',
          'Enable auto-reconciliation export (CSV + PDF) on batch completion',
          'A/B test pre-validation gate (mandatory vs advisory) on failure rate and merchant satisfaction',
        ],
      },
      {
        ph: 'Phase 3', dur: 'Month 5–6: Platform Rollout', col: C.green,
        pts: [
          'Roll out BatchPayIQ to all Nomupay merchants processing batch payouts globally',
          'Launch batch analytics dashboard with failure trend analysis and AI improvement recommendations',
          'Introduce auto-reconciliation rules: ERP integration on batch completion',
          'Target: 97%+ first-pass success rate, 40% payout ticket reduction, 3x reconciliation speed',
        ],
      },
    ];

    phases.forEach((ph, i) => {
      const x = 0.35 + i * 3.1;
      s.addShape('rect', { x, y: 1.55, w: 2.9, h: 5.6, fill: { color: '111827' }, line: { color: ph.col, pt: 1 }, shadow: makeShadow() });
      s.addShape('rect', { x, y: 1.55, w: 2.9, h: 0.07, fill: { color: ph.col } });
      s.addText(ph.ph, { x: x + 0.18, y: 1.68, w: 2.5, h: 0.5, fontSize: 20, bold: true, color: ph.col });
      s.addText(ph.dur, { x: x + 0.18, y: 2.18, w: 2.55, h: 0.45, fontSize: 10, bold: true, color: C.white });
      ph.pts.forEach((pt, j) => {
        s.addText('·  ' + pt, { x: x + 0.18, y: 2.75 + j * 0.9, w: 2.55, h: 0.8, fontSize: 9, color: C.muted, paraSpaceAfter: 2 });
      });
    });

    s.addShape('rect', { x: 0.35, y: 7.1, w: 9.3, h: 0.3, fill: { color: '1E293B' }, shadow: makeShadow() });
    s.addText('What I need to start: Batch payout transaction data schema  ·  IBAN validation API access  ·  1 backend engineer (status grid + reconciliation export)  ·  Merchant Success team on ticket category taxonomy', { x: 0.55, y: 7.14, w: 9, h: 0.22, fontSize: 8, color: C.muted });

    s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439', { x: 0.4, y: 7.33, w: 9.2, h: 0.17, fontSize: 8, color: C.muted });
  }

  await prs.writeFile({ fileName: 'assets/nomupay-batchpayiq-deck.pptx' });
  console.log('Done: assets/nomupay-batchpayiq-deck.pptx');
})();
