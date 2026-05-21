const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "MerchantOS: Milestone Rental Waiver";
pres.author = "Ajay Avaghade";

// ── PALETTE ──────────────────────────────────────────
const C = {
  darkBg:   "0A1F0A",
  darkBg2:  "0F2A0F",
  midGreen: "1A3D1A",
  green:    "22C55E",
  greenLt:  "4ADE80",
  greenDim: "166534",
  orange:   "F97316",
  orangeLt: "FEF3E8",
  white:    "FFFFFF",
  offWhite: "F4F7F4",
  border:   "D1E7D1",
  ink:      "0F1F0F",
  muted:    "4B6B4B",
  light:    "8FAA8F",
  shadow:   "000000",
};

const makeShadow = () => ({ type: "outer", color: C.shadow, opacity: 0.18, blur: 8, offset: 3, angle: 135 });

// ─────────────────────────────────────────────────────
// SLIDE 1 — COVER
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  // Diagonal stripe accents
  for (let i = 0; i < 6; i++) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: -1 + i * 2.2, y: -0.5, w: 0.6, h: 8,
      fill: { color: "1A3D1A", transparency: 60 },
      line: { color: C.darkBg, width: 0 },
      rotate: 25,
    });
  }

  // Green left accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 5.625,
    fill: { color: C.green }, line: { color: C.green, width: 0 },
  });

  // Tag pill
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.55, y: 0.42, w: 2.8, h: 0.35,
    fill: { color: "1A3D1A" }, line: { color: C.greenDim, width: 1 },
    rectRadius: 0.15,
  });
  s.addText("VIRTUSA — MERCHANT ACQUIRING", {
    x: 0.55, y: 0.42, w: 2.8, h: 0.35,
    fontSize: 8, bold: true, color: C.greenLt,
    align: "center", valign: "middle", charSpacing: 2,
  });

  // Main title
  s.addText("MerchantOS:", {
    x: 0.38, y: 1.1, w: 9.2, h: 0.85,
    fontSize: 54, bold: true, fontFace: "Calibri",
    color: C.white, align: "left", margin: 0,
  });
  s.addText("Milestone Rental Waiver", {
    x: 0.38, y: 1.9, w: 9.2, h: 0.85,
    fontSize: 54, bold: true, fontFace: "Calibri",
    color: C.green, align: "left", margin: 0,
  });

  // Subtitle
  s.addText("Driving merchant stickiness through behaviour-linked device rental incentives", {
    x: 0.38, y: 2.9, w: 7.5, h: 0.65,
    fontSize: 16, color: "8FAA8F", align: "left", margin: 0,
  });

  // Bottom divider
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.38, y: 3.8, w: 4, h: 0.03,
    fill: { color: C.green }, line: { color: C.green, width: 0 },
  });

  // Presenter name
  s.addText("Ajay Avaghade  ·  Growth & Monetization PM", {
    x: 0.38, y: 4.05, w: 6, h: 0.4,
    fontSize: 14, bold: true, color: C.white, align: "left", margin: 0,
  });
  s.addText("avaghadeajay009@gmail.com  ·  +91 9561558439", {
    x: 0.38, y: 4.45, w: 6, h: 0.35,
    fontSize: 12, color: C.light, align: "left", margin: 0,
  });

  // Right large stat
  s.addText("5,000+", {
    x: 7.2, y: 1.4, w: 2.6, h: 1.0,
    fontSize: 56, bold: true, color: C.green, align: "right", margin: 0,
  });
  s.addText("B2B merchants acquired at PhonePe", {
    x: 7.2, y: 2.4, w: 2.6, h: 0.5,
    fontSize: 11, color: C.light, align: "right", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 2 — THE PROBLEM
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg2 };

  // Slide label
  s.addText("THE PROBLEM", {
    x: 0.5, y: 0.22, w: 3, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 3, margin: 0,
  });

  // Title
  s.addText("30% of Deployed POS Terminals Are Inactive Within 90 Days", {
    x: 0.5, y: 0.52, w: 9, h: 0.9,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", margin: 0,
  });

  // 3 stat columns
  const cols = [
    { x: 0.45, stat: "30%",      label: "Inactive terminals\nwithin 90 days",   sub: "CAPEX deployed but zero GMV generated",   icon: "📦" },
    { x: 3.65, stat: "AED 45",   label: "Monthly terminal\nrental merchants pay", sub: "Feels like a pure cost with no clear benefit", icon: "💸" },
    { x: 6.85, stat: "60–70%",   label: "Drop-off after\napproval",              sub: "Merchants approved but never activate",    icon: "📉" },
  ];

  cols.forEach(col => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: col.x, y: 1.58, w: 3.0, h: 2.55,
      fill: { color: "162816" }, line: { color: "22C55E", width: 1 },
    });
    s.addText(col.icon, {
      x: col.x + 0.14, y: 1.7, w: 0.55, h: 0.55,
      fontSize: 26, align: "center", valign: "middle", margin: 0,
    });
    s.addText(col.stat, {
      x: col.x + 0.1, y: 2.22, w: 2.8, h: 0.7,
      fontSize: 42, bold: true, color: C.green, fontFace: "Calibri", align: "left", margin: 0,
    });
    s.addText(col.label, {
      x: col.x + 0.1, y: 2.88, w: 2.8, h: 0.55,
      fontSize: 12, bold: true, color: C.white, align: "left", margin: 0,
    });
    s.addText(col.sub, {
      x: col.x + 0.1, y: 3.44, w: 2.8, h: 0.52,
      fontSize: 10, color: C.light, align: "left", margin: 0,
    });
  });

  // Root cause box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 4.3, w: 9.1, h: 0.95,
    fill: { color: "1A3D1A" }, line: { color: C.green, width: 1 },
  });
  s.addText([
    { text: "Root cause: ", options: { bold: true, color: C.greenLt } },
    { text: "Merchants have no motivation to keep the terminal active after onboarding. A flat discount doesn't change behaviour — it just burns CAPEX.", options: { color: C.light } },
  ], {
    x: 0.65, y: 4.32, w: 8.7, h: 0.88,
    fontSize: 13, valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 3 — THE INSIGHT
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addText("THE INSIGHT", {
    x: 0.5, y: 0.22, w: 3, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 3, margin: 0,
  });
  s.addText("A Flat Rental Discount Burns CAPEX.\nA Milestone Waiver Changes Behaviour.", {
    x: 0.5, y: 0.5, w: 9, h: 0.95,
    fontSize: 26, bold: true, color: C.ink, fontFace: "Calibri", margin: 0,
  });

  // LEFT card — Current (red-tinted)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.55, w: 4.4, h: 2.9,
    fill: { color: C.white }, line: { color: "E5E7E5", width: 1 },
    shadow: makeShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.55, w: 4.4, h: 0.06,
    fill: { color: "DC2626" }, line: { color: "DC2626", width: 0 },
  });
  s.addText("❌  Flat Discount — Current Approach", {
    x: 0.6, y: 1.65, w: 4.1, h: 0.45,
    fontSize: 13, bold: true, color: C.ink, margin: 0,
  });
  const leftPoints = [
    "Given upfront regardless of usage",
    "No incentive to transact actively",
    "Merchant takes discount and under-uses device",
    "Acquirer absorbs the cost with zero GMV uplift",
  ];
  s.addText(leftPoints.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < leftPoints.length - 1, color: C.muted, fontSize: 13 },
  })), { x: 0.6, y: 2.2, w: 4.1, h: 2.0, margin: 0 });

  // RIGHT card — Proposed (green-tinted)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.15, y: 1.55, w: 4.4, h: 2.9,
    fill: { color: C.white }, line: { color: C.border, width: 1 },
    shadow: makeShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.15, y: 1.55, w: 4.4, h: 0.06,
    fill: { color: C.green }, line: { color: C.green, width: 0 },
  });
  s.addText("✅  Milestone Rental Waiver — Proposed", {
    x: 5.3, y: 1.65, w: 4.1, h: 0.45,
    fontSize: 13, bold: true, color: C.ink, margin: 0,
  });
  const rightPoints = [
    "Waiver is earned by hitting usage goals",
    "Merchant self-motivates to transact more",
    "90-day engagement window with built-in urgency",
    "Acquirer only waives rental after GMV goals are met",
  ];
  s.addText(rightPoints.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < rightPoints.length - 1, color: C.muted, fontSize: 13 },
  })), { x: 5.3, y: 2.2, w: 4.1, h: 2.0, margin: 0 });

  // vs divider
  s.addShape(pres.shapes.OVAL, {
    x: 4.64, y: 2.7, w: 0.52, h: 0.52,
    fill: { color: C.darkBg }, line: { color: C.green, width: 1 },
  });
  s.addText("VS", {
    x: 4.64, y: 2.7, w: 0.52, h: 0.52,
    fontSize: 10, bold: true, color: C.green, align: "center", valign: "middle", margin: 0,
  });

  // Insight callout
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 4.58, w: 9.1, h: 0.72,
    fill: { color: "F0FBF0" }, line: { color: C.green, width: 1 },
  });
  s.addText([
    { text: "Key insight: ", options: { bold: true, color: C.greenDim } },
    { text: "The waiver costs the acquirer nothing net — by the time the merchant hits the TPV goal, additional GMV has already paid for the rental many times over.", options: { color: C.muted } },
  ], {
    x: 0.65, y: 4.6, w: 8.7, h: 0.68,
    fontSize: 12, valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 4 — THE MECHANIC
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  s.addText("THE MECHANIC", {
    x: 0.5, y: 0.22, w: 3, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 3, margin: 0,
  });
  s.addText("How It Works: Two Goals, One Reward", {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 30, bold: true, color: C.white, fontFace: "Calibri", margin: 0,
  });

  // Timeline connector line
  s.addShape(pres.shapes.LINE, {
    x: 0.95, y: 2.3, w: 8.2, h: 0,
    line: { color: C.greenDim, width: 2, dashType: "dash" },
  });

  const steps = [
    { x: 0.45, num: "1", head: "Day 0", title: "Activation", body: "Milestone window opens.\nTerminal rental: AED 45/mo" },
    { x: 2.35, num: "2", head: "Days 1–90", title: "Tracking", body: "Live progress bars in app.\nNudges at 50%, 80%, 95%" },
    { x: 4.25, num: "TPV", head: "Goal 1", title: "Payment Volume", body: "Process AED 15,000 in total payments within 90 days" },
    { x: 6.15, num: "TPC", head: "Goal 2", title: "Unique Customers", body: "Serve 100 unique cardholders within 90 days" },
    { x: 8.05, num: "✓", head: "Day 90", title: "Rental Waived", body: "AED 45 → AED 1/month.\nAuto-applied to invoice." },
  ];

  steps.forEach((step, i) => {
    const isGoal = i >= 2 && i <= 3;
    const isSuccess = i === 4;
    const circleColor = isSuccess ? C.green : isGoal ? "1A3D1A" : "1A3D1A";
    const circleBorder = isGoal ? C.green : isSuccess ? C.green : C.greenDim;

    s.addShape(pres.shapes.OVAL, {
      x: step.x + 0.5, y: 2.02, w: 0.56, h: 0.56,
      fill: { color: circleColor }, line: { color: circleBorder, width: isGoal || isSuccess ? 2 : 1 },
    });
    s.addText(step.num, {
      x: step.x + 0.5, y: 2.02, w: 0.56, h: 0.56,
      fontSize: 9, bold: true, color: isSuccess ? C.darkBg : C.green,
      align: "center", valign: "middle", margin: 0,
    });

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: step.x + 0.1, y: 2.7, w: 1.7, h: 2.2,
      fill: { color: "162816" }, line: { color: C.greenDim, width: 1 },
    });
    s.addText(step.head, {
      x: step.x + 0.15, y: 2.78, w: 1.6, h: 0.28,
      fontSize: 8, color: C.green, bold: true, charSpacing: 1, margin: 0,
    });
    s.addText(step.title, {
      x: step.x + 0.15, y: 3.06, w: 1.6, h: 0.38,
      fontSize: 12, bold: true, color: C.white, margin: 0,
    });
    s.addText(step.body, {
      x: step.x + 0.15, y: 3.44, w: 1.6, h: 1.35,
      fontSize: 10, color: C.light, margin: 0,
    });
  });

  // A/B test note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 5.05, w: 9.1, h: 0.38,
    fill: { color: "1A3D1A" }, line: { color: C.green, width: 0 },
  });
  s.addText("A/B tested at PhonePe: Rent Waiver vs Cashback variant — Rent Waiver drove 34% higher goal completion rate", {
    x: 0.65, y: 5.07, w: 8.7, h: 0.34,
    fontSize: 11, color: C.greenLt, align: "center", valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 5 — THE PRODUCT
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addText("THE PRODUCT", {
    x: 0.5, y: 0.22, w: 4, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 3, margin: 0,
  });
  s.addText("4 Key Screen States — Built & Validated at PhonePe", {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 28, bold: true, color: C.ink, fontFace: "Calibri", margin: 0,
  });

  const screens = [
    { x: 0.35, num: "01", title: "Merchant Home", color: C.green, body: "Milestone banner shows 62% progress + countdown. Always visible — no navigation needed." },
    { x: 2.72, num: "02", title: "Milestone Tracker", color: "0EA5E9", body: "Live TPV + TPC progress bars. Countdown urgency. Exact gap-to-goal shown in nudge text." },
    { x: 5.09, num: "03", title: "Goal Achieved", color: C.orange, body: "Confetti celebration screen. Rental reduction displayed. Auto-applied invoice confirmation." },
    { x: 7.46, num: "04", title: "Acquirer Nudge", color: "A855F7", body: "Ops team pushes threshold nudges at 50%, 80%, 95%. Tip cards drive specific merchant actions." },
  ];

  screens.forEach(sc => {
    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: sc.x, y: 1.32, w: 2.2, h: 3.72,
      fill: { color: C.white }, line: { color: "E2EBE2", width: 1 },
      shadow: makeShadow(),
    });
    // Top accent
    s.addShape(pres.shapes.RECTANGLE, {
      x: sc.x, y: 1.32, w: 2.2, h: 0.06,
      fill: { color: sc.color }, line: { color: sc.color, width: 0 },
    });
    // Number
    s.addShape(pres.shapes.OVAL, {
      x: sc.x + 0.14, y: 1.5, w: 0.44, h: 0.44,
      fill: { color: sc.color }, line: { color: sc.color, width: 0 },
    });
    s.addText(sc.num, {
      x: sc.x + 0.14, y: 1.5, w: 0.44, h: 0.44,
      fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0,
    });
    s.addText(sc.title, {
      x: sc.x + 0.1, y: 2.04, w: 2.0, h: 0.42,
      fontSize: 13, bold: true, color: C.ink, margin: 0,
    });
    s.addText(sc.body, {
      x: sc.x + 0.1, y: 2.5, w: 2.0, h: 1.6,
      fontSize: 11, color: C.muted, margin: 0,
    });

    // Phone mockup shape
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: sc.x + 0.35, y: 4.1, w: 1.5, h: 0.75,
      fill: { color: C.darkBg }, line: { color: C.greenDim, width: 1 },
      rectRadius: 0.08,
    });
    s.addText("⬛ ▓▓ ▓▓", {
      x: sc.x + 0.35, y: 4.18, w: 1.5, h: 0.58,
      fontSize: 8, color: "22C55E", align: "center", valign: "middle", margin: 0,
    });
  });

  // Footer note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 5.24, w: 9.3, h: 0.22,
    fill: { color: "E8F5E8" }, line: { color: C.border, width: 1 },
  });
  s.addText("Live prototype: virtusa-milestone-prototype.html — interactive demo of all 4 screen states", {
    x: 0.5, y: 5.25, w: 9.0, h: 0.2,
    fontSize: 10, color: C.muted, align: "center", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 6 — IMPACT & ROI
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  s.addText("IMPACT & ROI", {
    x: 0.5, y: 0.22, w: 4, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 3, margin: 0,
  });
  s.addText("Projected Impact — Built on PhonePe Proof Points", {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", margin: 0,
  });

  // Left section label
  s.addText("MERCHANT IMPACT", {
    x: 0.45, y: 1.3, w: 4.5, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 2, margin: 0,
  });

  const merchantMetrics = [
    { stat: "↑ 60%", label: "Active device rate at 90 days", sub: "vs 40% baseline" },
    { stat: "↑ 25%", label: "Average monthly TPV per terminal", sub: "vs control group" },
    { stat: "↓ 35%", label: "Early churn in months 1–3", sub: "critical retention window" },
    { stat: "107",   label: "Avg unique customers/month", sub: "per active terminal" },
  ];

  merchantMetrics.forEach((m, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const bx = 0.45 + col * 2.4;
    const by = 1.65 + row * 1.45;
    s.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: 2.2, h: 1.3,
      fill: { color: "162816" }, line: { color: C.greenDim, width: 1 },
    });
    s.addText(m.stat, {
      x: bx + 0.1, y: by + 0.05, w: 2.0, h: 0.6,
      fontSize: 30, bold: true, color: C.green, fontFace: "Calibri", margin: 0,
    });
    s.addText(m.label, {
      x: bx + 0.1, y: by + 0.62, w: 2.0, h: 0.38,
      fontSize: 10, bold: true, color: C.white, margin: 0,
    });
    s.addText(m.sub, {
      x: bx + 0.1, y: by + 1.0, w: 2.0, h: 0.22,
      fontSize: 9, color: C.light, margin: 0,
    });
  });

  // Divider
  s.addShape(pres.shapes.LINE, {
    x: 5.05, y: 1.3, w: 0, h: 3.65,
    line: { color: C.greenDim, width: 1, dashType: "dash" },
  });

  // Right section label
  s.addText("ACQUIRER ROI", {
    x: 5.3, y: 1.3, w: 4.2, h: 0.28,
    fontSize: 9, bold: true, color: C.orange, charSpacing: 2, margin: 0,
  });

  const acquirerMetrics = [
    { stat: "AED 44",     label: "Max cost of waiver/merchant/mo", sub: "only when GMV goal is met" },
    { stat: "AED 15K+",   label: "Min GMV before waiver triggers", sub: "acquirer always wins first" },
    { stat: "~0.3%",      label: "Waiver as % of GMV generated", sub: "effectively zero net cost" },
    { stat: "↑ 60%",      label: "CAPEX ROI per device deployed", sub: "vs flat-discount approach" },
  ];

  acquirerMetrics.forEach((m, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const bx = 5.3 + col * 2.4;
    const by = 1.65 + row * 1.45;
    s.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: 2.2, h: 1.3,
      fill: { color: "1E1000" }, line: { color: "7C3419", width: 1 },
    });
    s.addText(m.stat, {
      x: bx + 0.1, y: by + 0.05, w: 2.0, h: 0.6,
      fontSize: 28, bold: true, color: C.orange, fontFace: "Calibri", margin: 0,
    });
    s.addText(m.label, {
      x: bx + 0.1, y: by + 0.62, w: 2.0, h: 0.38,
      fontSize: 10, bold: true, color: C.white, margin: 0,
    });
    s.addText(m.sub, {
      x: bx + 0.1, y: by + 1.0, w: 2.0, h: 0.22,
      fontSize: 9, color: C.light, margin: 0,
    });
  });

  // Footer insight
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 5.05, w: 9.1, h: 0.4,
    fill: { color: "1A3D1A" }, line: { color: C.green, width: 0 },
  });
  s.addText("The waiver only triggers AFTER the merchant has already generated the target GMV — making this a zero-net-cost stickiness lever.", {
    x: 0.65, y: 5.07, w: 8.7, h: 0.36,
    fontSize: 11, color: C.greenLt, align: "center", valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 7 — WHY AJAY
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.white };

  s.addText("PROOF OF WORK", {
    x: 0.5, y: 0.22, w: 4, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 3, margin: 0,
  });
  s.addText("I Built This. Here's the Proof.", {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 30, bold: true, color: C.ink, fontFace: "Calibri", margin: 0,
  });

  // Left column — PhonePe
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 1.28, w: 4.35, h: 3.55,
    fill: { color: C.darkBg }, line: { color: C.green, width: 1 },
    shadow: makeShadow(),
  });
  s.addText("PhonePe — What I Actually Built", {
    x: 0.65, y: 1.42, w: 3.95, h: 0.38,
    fontSize: 13, bold: true, color: C.green, margin: 0,
  });
  const leftItems = [
    "Built the Milestone program end-to-end for POS device merchants",
    "Designed all 6 screen variants: TPV, TPC, Multiple Goals, Success states, Error fallback",
    "A/B tested Rent Waiver vs Cashback mechanic across two test groups",
    "Owned the Nudges engine — ops campaign builder for threshold-based push triggers",
    "Result: 5,000+ B2B merchants, 60% improvement in 90-day device activity rate",
  ];
  s.addText(leftItems.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < leftItems.length - 1, color: "8FAA8F", fontSize: 12 },
  })), { x: 0.65, y: 1.88, w: 3.95, h: 2.7, margin: 0 });

  // Right column — JD Mapping
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.28, w: 4.35, h: 3.55,
    fill: { color: C.offWhite }, line: { color: C.border, width: 1 },
    shadow: makeShadow(),
  });
  s.addText("Virtusa JD — How It Maps", {
    x: 5.4, y: 1.42, w: 3.95, h: 0.38,
    fontSize: 13, bold: true, color: C.ink, margin: 0,
  });
  const rightItems = [
    "End-to-end merchant onboarding → Instant Activation Engine",
    "Acquisition campaign design → Milestone Challenge launch playbook",
    "Value-added services for stickiness → This exact deck",
    "BAT planning → A/B test framework, two variant groups",
    "Regulatory alignment → T&C, invoice auto-apply, merchant consent flows",
  ];
  s.addText(rightItems.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i < rightItems.length - 1, color: C.muted, fontSize: 12 },
  })), { x: 5.4, y: 1.88, w: 3.95, h: 2.7, margin: 0 });

  // Bottom callout
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 4.98, w: 9.1, h: 0.44,
    fill: { color: C.darkBg }, line: { color: C.green, width: 1 },
  });
  s.addText("\"This isn't a case study. This is a working product I shipped — and the Figma, A/B test data, and prototype are all attached.\"", {
    x: 0.65, y: 4.99, w: 8.7, h: 0.42,
    fontSize: 12, italic: true, color: C.greenLt, align: "center", valign: "middle", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// SLIDE 8 — PHASED ROLLOUT & ASK
// ─────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  s.addText("ROLLOUT PLAN", {
    x: 0.5, y: 0.22, w: 4, h: 0.28,
    fontSize: 9, bold: true, color: C.green, charSpacing: 3, margin: 0,
  });
  s.addText("Phased Rollout — From Pilot to Full Deployment", {
    x: 0.5, y: 0.5, w: 9, h: 0.62,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", margin: 0,
  });

  const phases = [
    {
      x: 0.45, label: "PHASE 1", period: "Month 1–2",
      title: "Pilot — 50 Merchants",
      color: C.green,
      points: [
        "Deploy Milestone Waiver to 50 newly onboarded merchants in one market",
        "Instrument TPV + TPC tracking against control group",
        "Measure 90-day activation rate, churn, and GMV uplift",
      ],
    },
    {
      x: 3.5, label: "PHASE 2", period: "Month 3–4",
      title: "Scale + A/B Test",
      color: "0EA5E9",
      points: [
        "Run Rent Waiver vs Cashback A/B test across 500 merchants",
        "Tune goal thresholds based on pilot learnings",
        "Launch Nudges engine for ops team campaign management",
      ],
    },
    {
      x: 6.55, label: "PHASE 3", period: "Month 5–6",
      title: "Full Rollout",
      color: C.orange,
      points: [
        "Roll out winning variant to all new merchant onboarding",
        "Integrate with Instant Activation Engine (Module B)",
        "Expand program to existing inactive terminal base",
      ],
    },
  ];

  phases.forEach(ph => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: ph.x, y: 1.28, w: 2.9, h: 3.38,
      fill: { color: "0F2A0F" }, line: { color: ph.color, width: 1 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: ph.x, y: 1.28, w: 2.9, h: 0.06,
      fill: { color: ph.color }, line: { color: ph.color, width: 0 },
    });
    s.addText(ph.label, {
      x: ph.x + 0.12, y: 1.38, w: 2.65, h: 0.22,
      fontSize: 8, bold: true, color: ph.color, charSpacing: 2, margin: 0,
    });
    s.addText(ph.period, {
      x: ph.x + 0.12, y: 1.58, w: 2.65, h: 0.26,
      fontSize: 10, color: C.light, margin: 0,
    });
    s.addText(ph.title, {
      x: ph.x + 0.12, y: 1.85, w: 2.65, h: 0.42,
      fontSize: 14, bold: true, color: C.white, margin: 0,
    });
    s.addText(ph.points.map((t, i) => ({
      text: t, options: { bullet: true, breakLine: i < ph.points.length - 1, color: C.light, fontSize: 11 },
    })), { x: ph.x + 0.12, y: 2.38, w: 2.65, h: 2.0, margin: 0 });
  });

  // Ask box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 4.78, w: 9.1, h: 0.62,
    fill: { color: "1A3D1A" }, line: { color: C.green, width: 1 },
  });
  s.addText([
    { text: "What I need to build this: ", options: { bold: true, color: C.greenLt } },
    { text: "Access to merchant transaction data  ·  Invoicing/billing team alignment  ·  1 engineer for milestone tracking API  ·  1 designer for merchant app screens", options: { color: C.light } },
  ], {
    x: 0.65, y: 4.8, w: 8.7, h: 0.55,
    fontSize: 12, valign: "middle", margin: 0,
  });

  // Contact footer
  s.addText("Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439", {
    x: 0.45, y: 5.4, w: 9.1, h: 0.22,
    fontSize: 10, color: C.muted, align: "center", margin: 0,
  });
}

// ─────────────────────────────────────────────────────
// WRITE FILE
// ─────────────────────────────────────────────────────
pres.writeFile({ fileName: "/Users/ajayavaghade/Portfolio/assets/virtusa-milestone-deck.pptx" })
  .then(() => console.log("✅  Deck written: virtusa-milestone-deck.pptx"))
  .catch(e => { console.error("❌  Error:", e); process.exit(1); });
