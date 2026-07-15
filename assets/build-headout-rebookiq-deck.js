const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';
  prs.title = 'Headout RebookIQ — Post-Experience Rebooking Engine';

  const C = {
    dark:  '0D0D0D',
    hero:  '1A1A1A',
    coral: 'FF385C',
    gold:  'F59E0B',
    green: '22C55E',
    white: 'FFFFFF',
    lgray: 'F7F7F7',
    muted: '888888',
    bordl: 'E8E8E8',
  };
  const makeShadow = () => ({ type:'outer', color:'000000', blur:6, offset:2, angle:45, opacity:0.18 });

  // ── SLIDE 1 — COVER ──────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    // Diagonal coral accent lines
    for (let i = 0; i < 6; i++) {
      s.addShape('line', { x: -0.5 + i*2, y: 0, w: 0, h: 7.5, line: { color: C.coral, width: 0.5, transparency: 92 } });
    }
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.coral } });

    // Tag
    s.addText('HEADOUT · CASE STUDY 01 · REBOOKIQ', {
      x: 0.4, y: 0.35, w: 9.2, h: 0.25,
      fontSize: 9, bold: true, color: C.coral, charSpacing: 3,
      fontFace: 'Inter',
    });

    // Product name
    s.addText('RebookIQ', {
      x: 0.4, y: 1.0, w: 6, h: 1.4,
      fontSize: 72, bold: true, color: C.white, fontFace: 'Outfit',
    });

    // Subtitle
    s.addText('The Post-Experience\nRebooking Engine', {
      x: 0.4, y: 2.5, w: 6, h: 1.0,
      fontSize: 22, color: C.gold, fontFace: 'Inter', bold: false,
    });

    // Presenter
    s.addText('Ajay Avaghade · Product Manager', {
      x: 0.4, y: 6.8, w: 5, h: 0.3,
      fontSize: 11, color: C.muted, fontFace: 'Inter',
    });

    // Big stat
    s.addShape('rect', { x: 6.5, y: 2.2, w: 3.1, h: 2.0, fill: { color: C.coral }, shadow: makeShadow(), rounding: 0.12 });
    s.addText('+28%', { x: 6.6, y: 2.35, w: 2.9, h: 0.8, fontSize: 48, bold: true, color: C.white, fontFace: 'Outfit', align: 'center' });
    s.addText('D30 Rebooking Rate\nProjected from PhonePe +15% baseline', { x: 6.6, y: 3.1, w: 2.9, h: 0.8, fontSize: 11, color: C.white, fontFace: 'Inter', align: 'center' });

    s.addShape('rect', { x: 6.5, y: 4.4, w: 1.45, h: 1.6, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addText('2.4×', { x: 6.6, y: 4.55, w: 1.25, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: 'Outfit', align: 'center' });
    s.addText('CLTV\nUplift', { x: 6.6, y: 5.1, w: 1.25, h: 0.6, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center' });

    s.addShape('rect', { x: 8.1, y: 4.4, w: 1.45, h: 1.6, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addText('−35%', { x: 8.2, y: 4.55, w: 1.25, h: 0.6, fontSize: 24, bold: true, color: C.green, fontFace: 'Outfit', align: 'center' });
    s.addText('Time to\n2nd Booking', { x: 8.2, y: 5.1, w: 1.25, h: 0.6, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center' });
  }

  // ── SLIDE 2 — THE PROBLEM ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.coral } });

    s.addText('THE PROBLEM', { x: 0.4, y: 0.35, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.coral, charSpacing: 3, fontFace: 'Inter' });
    s.addText('60%+ of First-Time Headout Users Never Book Again — The Post-Experience Moment Is Wasted', {
      x: 0.4, y: 0.75, w: 9.2, h: 1.1,
      fontSize: 26, bold: true, color: C.white, fontFace: 'Outfit', wrap: true,
    });

    // 3 stat columns
    const stats = [
      { val: '60%+', label: 'Single-experience rate', sub: 'Most users book once and never return — app closed at ticket delivery' },
      { val: '0', label: 'Post-experience prompts', sub: 'Nothing fires when you exit the venue — highest-intent moment missed entirely' },
      { val: '47 wks', label: 'Avg gap between bookings', sub: 'Travel apps open 0.8× per month vs 8.4× for rideshare — zero engagement hook' },
    ];
    stats.forEach((st, i) => {
      const x = 0.4 + i * 3.2;
      s.addShape('rect', { x, y: 2.1, w: 3.0, h: 2.8, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
      s.addText(st.val, { x: x+0.15, y: 2.25, w: 2.7, h: 0.9, fontSize: 44, bold: true, color: C.coral, fontFace: 'Outfit', align: 'center' });
      s.addText(st.label, { x: x+0.15, y: 3.1, w: 2.7, h: 0.4, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter', align: 'center' });
      s.addText(st.sub, { x: x+0.15, y: 3.55, w: 2.7, h: 0.9, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center', wrap: true });
    });

    // Insight box
    s.addShape('rect', { x: 0.4, y: 5.2, w: 9.2, h: 1.8, fill: { color: '1A0808' }, rounding: 0.1, shadow: makeShadow() });
    s.addShape('rect', { x: 0.4, y: 5.2, w: 0.06, h: 1.8, fill: { color: C.coral } });
    s.addText('Root cause: The highest-intent moment — the 30-minute window after your experience ends, when the Colosseum is still vivid and you\'re already thinking "what next" — fires with zero product response. RebookIQ closes this gap.', {
      x: 0.65, y: 5.35, w: 8.8, h: 1.4,
      fontSize: 12, color: C.white, fontFace: 'Inter', wrap: true, bold: false,
    });
  }

  // ── SLIDE 3 — THE INSIGHT ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('THE INSIGHT', { x: 0.4, y: 0.35, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.coral, charSpacing: 3, fontFace: 'Inter' });
    s.addText('The Highest-Converting Moment Fires When Your Experience Ends — Not a Day Later', {
      x: 0.4, y: 0.75, w: 9.2, h: 1.0,
      fontSize: 26, bold: true, color: C.dark, fontFace: 'Outfit', wrap: true,
    });

    // Two columns
    const lx = 0.4, rx = 5.3, colW = 4.5, colY = 2.0, colH = 4.5;

    s.addShape('rect', { x: lx, y: colY, w: colW, h: colH, fill: { color: 'FEE2E2' }, rounding: 0.12, shadow: makeShadow() });
    s.addText('❌  Without RebookIQ', { x: lx+0.2, y: colY+0.2, w: colW-0.4, h: 0.4, fontSize: 13, bold: true, color: 'B91C1C', fontFace: 'Inter' });
    const noItems = [
      'Experience ends → app closes → user gone',
      'Next push notification: generic promo, 3 days later',
      'No city badge, no milestone, no identity formed',
      'No recommendations tied to what you just did',
      'One-time visitor → zero rebooking loop',
    ];
    noItems.forEach((t, i) => {
      s.addText('· ' + t, { x: lx+0.2, y: colY+0.75+i*0.65, w: colW-0.4, h: 0.55, fontSize: 11, color: '7F1D1D', fontFace: 'Inter', wrap: true });
    });

    s.addShape('rect', { x: rx, y: colY, w: colW, h: colH, fill: { color: C.dark }, rounding: 0.12, shadow: makeShadow() });
    s.addText('✅  With RebookIQ', { x: rx+0.2, y: colY+0.2, w: colW-0.4, h: 0.4, fontSize: 13, bold: true, color: C.gold, fontFace: 'Inter' });
    const yesItems = [
      'Experience ends → RebookIQ fires within 30 minutes',
      'City badge unlocked → Explorer Score activated',
      'Top 3 personalized next picks surfaced immediately',
      '"94% of Colosseum visitors book this next" social proof',
      'Milestone creates forward pull → multi-city explorer identity',
    ];
    yesItems.forEach((t, i) => {
      s.addText('· ' + t, { x: rx+0.2, y: colY+0.75+i*0.65, w: colW-0.4, h: 0.55, fontSize: 11, color: 'E5E7EB', fontFace: 'Inter', wrap: true });
    });

    s.addText('Key insight: The Explorer Score milestone creates pull — not push. Users return not because of a notification but because they\'re 60 points away from a Paris Explorer badge. The mechanic builds identity, not dependency. (Source: Phocuswright Tours & Attractions 2023)', {
      x: 0.4, y: 6.8, w: 9.2, h: 0.5,
      fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 4 — THE MECHANIC ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.hero };
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.coral } });

    s.addText('THE MECHANIC', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.gold, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Five Steps. One Continuous Rebooking Loop.', { x: 0.4, y: 0.65, w: 9.2, h: 0.7, fontSize: 28, bold: true, color: C.white, fontFace: 'Outfit' });

    const steps = [
      { num: '01', title: 'Post-Experience Trigger', body: 'Experience ticket validity expires → RebookIQ fires within 30 min. Push notification + in-app card: "Colosseum done ✓ — Rome Explorer awaits."' },
      { num: '02', title: 'City Badge Unlock', body: 'Celebration screen: "Rome Unlocked 🏛️". Explorer Score activated. Social share prompt. Creates identity anchor — user is now a "Rome Explorer", not a ticket buyer.' },
      { num: '03', title: 'Personalized Next Picks', body: '3 RebookIQ recommendations ranked by: past experience type × next-experience affinity × availability × time of day × peer behavior signal. Not generic. Tied to what you just did.' },
      { num: '04', title: 'Explorer Score + Milestone', body: 'Persistent score profile: cities unlocked, points, next milestone badge. "20 points from Paris Explorer → book 1 experience in Paris." Pull mechanic builds cross-city expansion.' },
      { num: '05', title: 'Book → Loop Continues', body: 'Second booking triggers new city score if cross-city, or deepens Rome Explorer rank if same-city. Peer feed shows activity. Milestone progress bar accelerates to third booking.' },
    ];

    steps.forEach((step, i) => {
      const x = 0.4 + (i % 3) * 3.2;
      const y = i < 3 ? 1.6 : 4.1;
      const w = i >= 3 ? 4.5 : 3.0;
      if (i === 3) { /* adjust x */ }

      const sx = i < 3 ? (0.4 + i*3.2) : (0.4 + (i-3)*4.65);
      const sy = i < 3 ? 1.6 : 4.1;
      const sw = i < 3 ? 3.0 : 4.5;
      const sh = 2.2;

      s.addShape('rect', { x: sx, y: sy, w: sw, h: sh, fill: { color: '222222' }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: sy, w: sw, h: 0.05, fill: { color: C.coral } });
      s.addText(step.num, { x: sx+0.15, y: sy+0.15, w: 0.5, h: 0.4, fontSize: 11, bold: true, color: C.coral, fontFace: 'Outfit' });
      s.addText(step.title, { x: sx+0.15, y: sy+0.45, w: sw-0.3, h: 0.4, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(step.body, { x: sx+0.15, y: sy+0.9, w: sw-0.3, h: 1.15, fontSize: 10, color: C.muted, fontFace: 'Inter', wrap: true });
    });

    s.addText('A/B tested parallel at PhonePe: 3-category breadth threshold + milestone mechanics → D30 retention +15% across new-user cohort', {
      x: 0.4, y: 6.95, w: 9.2, h: 0.35, fontSize: 10, color: C.gold, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 5 — THE PRODUCT ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('THE PRODUCT', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.coral, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Five Screen States — Post-Experience to Second Booking', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.dark, fontFace: 'Outfit' });

    const screens = [
      { num: '01', title: 'Post-Experience Card', body: 'Fires 30 min after ticket validity end. Rating prompt (⭐⭐⭐⭐⭐). "Your experience just ended ✓ — RebookIQ is finding your next pick." Loading animation.' },
      { num: '02', title: 'Rome Explorer Badge', body: 'City badge unlock celebration. Explorer Score activated: 240 pts. "Rome Explorer · Level 2 Unlocked". 3 pick chips: Vatican · Food Tour · Trevi Night Walk.' },
      { num: '03', title: 'RebookIQ Picks', body: '"What\'s Next in Rome" · 3 ranked cards. "94% of Colosseum visitors book Vatican next." Top pick badge, urgency signals, 1-tap booking CTA.' },
      { num: '04', title: 'Explorer Score Profile', body: 'Cities unlocked (Rome ✓, Paris 🔒, Dubai 🔒). Score ring: 240/300 pts. Milestone progress: "20 pts from Paris Explorer badge." Activity log.' },
      { num: '05', title: 'Booking Confirmed', body: 'Vatican Museums confirmed. Confetti. Score upgrade: 240 → 280 pts. "20 pts from Paris Explorer 🗼" — forward pull into next city.' },
    ];

    screens.forEach((sc, i) => {
      const x = 0.4 + (i % 3) * 3.2;
      const y = i < 3 ? 1.5 : 4.15;
      const sx = i < 3 ? (0.4 + i*3.2) : (0.65 + (i-3)*4.65);
      const sw = i < 3 ? 3.0 : 4.5;
      const sh = 2.4;
      const sy = i < 3 ? 1.5 : 4.1;

      s.addShape('rect', { x: sx, y: sy, w: sw, h: sh, fill: { color: C.dark }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: sy, w: sw, h: 0.05, fill: { color: C.coral } });
      s.addText(sc.num, { x: sx+0.15, y: sy+0.15, w: 0.5, h: 0.3, fontSize: 10, bold: true, color: C.coral, fontFace: 'Outfit' });
      s.addText(sc.title, { x: sx+0.15, y: sy+0.45, w: sw-0.3, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(sc.body, { x: sx+0.15, y: sy+0.85, w: sw-0.3, h: 1.35, fontSize: 10, color: '9CA3AF', fontFace: 'Inter', wrap: true });
    });

    s.addText('Prototype: headout-rebookiq-prototype.html · 5 interactive screens · keyboard navigation', {
      x: 0.4, y: 6.95, w: 9.2, h: 0.35, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 6 — IMPACT & ROI ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.coral } });

    s.addText('IMPACT & ROI', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.coral, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Built on PhonePe Proof Points — Adjusted for Headout Scale', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.white, fontFace: 'Outfit' });

    // User Impact
    s.addText('TRAVELER IMPACT', { x: 0.4, y: 1.55, w: 4.5, h: 0.25, fontSize: 9, bold: true, color: C.coral, charSpacing: 2, fontFace: 'Inter' });
    const userMetrics = [
      { val: '+28%', lbl: 'D30 Rebooking Rate', src: 'PhonePe D30 +15% baseline · conservative Headout projection' },
      { val: '2.4×', lbl: 'CLTV (1→3 experiences)', src: 'Bain & Company Customer Loyalty in Travel 2023' },
      { val: '−35%', lbl: 'Time to Second Booking', src: 'In-moment trigger vs next-day push · Skift Research 2024' },
      { val: '+18%', lbl: 'Cross-City Expansion Rate', src: 'Explorer Score milestone creates Paris/Dubai forward pull' },
    ];
    userMetrics.forEach((m, i) => {
      s.addShape('rect', { x: 0.4, y: 1.9+i*1.15, w: 4.5, h: 1.0, fill: { color: C.hero }, rounding: 0.08, shadow: makeShadow() });
      s.addText(m.val, { x: 0.55, y: 2.0+i*1.15, w: 1.2, h: 0.6, fontSize: 30, bold: true, color: C.coral, fontFace: 'Outfit' });
      s.addText(m.lbl, { x: 1.85, y: 2.05+i*1.15, w: 2.9, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText('(Source: ' + m.src + ')', { x: 1.85, y: 2.42+i*1.15, w: 2.9, h: 0.3, fontSize: 9, color: C.muted, fontFace: 'Inter', italic: true });
    });

    // Headout ROI
    s.addText('HEADOUT BUSINESS ROI', { x: 5.2, y: 1.55, w: 4.5, h: 0.25, fontSize: 9, bold: true, color: C.gold, charSpacing: 2, fontFace: 'Inter' });
    const bizMetrics = [
      { val: '0', lbl: 'Net incremental acquisition cost', sub: 'Rebooking works on existing users — no new CAC' },
      { val: '3×', lbl: 'Revenue per acquired user', sub: 'From 1 experience average → 3 per Explorer Score user' },
      { val: '+15%', lbl: 'DAU/MAU ratio uplift', sub: 'Between-trip engagement via Explorer Score home tab' },
      { val: '50K+', lbl: 'New second bookings/month', sub: 'At Headout MAU scale; conservative 5% trigger-to-book rate' },
    ];
    bizMetrics.forEach((m, i) => {
      s.addShape('rect', { x: 5.2, y: 1.9+i*1.15, w: 4.5, h: 1.0, fill: { color: '1A1000' }, rounding: 0.08, shadow: makeShadow() });
      s.addText(m.val, { x: 5.35, y: 2.0+i*1.15, w: 1.2, h: 0.6, fontSize: 30, bold: true, color: C.gold, fontFace: 'Outfit' });
      s.addText(m.lbl, { x: 6.65, y: 2.05+i*1.15, w: 2.9, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(m.sub, { x: 6.65, y: 2.42+i*1.15, w: 2.9, h: 0.3, fontSize: 9, color: C.muted, fontFace: 'Inter' });
    });

    s.addShape('rect', { x: 0.4, y: 6.55, w: 9.2, h: 0.7, fill: { color: '1A0A0A' }, rounding: 0.08 });
    s.addText('RebookIQ has zero net incremental cost — it fires on the user already acquired. Every second booking is pure margin expansion on the existing CAC spent.', {
      x: 0.6, y: 6.65, w: 8.8, h: 0.5, fontSize: 11, color: C.white, fontFace: 'Inter', bold: true,
    });
  }

  // ── SLIDE 7 — PROOF OF WORK ───────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('PROOF OF WORK', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.coral, charSpacing: 3, fontFace: 'Inter' });
    s.addText('I Built the Underlying Mechanics at PhonePe. Here Is the Direct Map.', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.dark, fontFace: 'Outfit' });

    // Left dark panel
    s.addShape('rect', { x: 0.4, y: 1.5, w: 4.5, h: 5.4, fill: { color: C.dark }, rounding: 0.12, shadow: makeShadow() });
    s.addText('What I Built at PhonePe', { x: 0.6, y: 1.7, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.gold, fontFace: 'Inter' });

    const proof = [
      { icon: '🎯', text: 'Built 3-category breadth threshold + directed incentivized onboarding → D30 retention +15% across new-user cohort (350M MAU scale)' },
      { icon: '📊', text: 'Built milestone-based retention construct: behavior → CRM message vs offer → category breadth milestone → D30 attainment' },
      { icon: '🤖', text: 'Defined ML target as category breadth (not raw transaction count) and designed segment-to-channel intervention logic' },
      { icon: '🔄', text: 'Shipped 7-day activation cliff fix: masthead ad moratorium for new users, replaced by journey-progress bar → week-one retention uplift' },
    ];
    proof.forEach((p, i) => {
      s.addText(p.icon, { x: 0.6, y: 2.25+i*1.15, w: 0.4, h: 0.35, fontSize: 14, fontFace: 'Inter' });
      s.addText(p.text, { x: 1.1, y: 2.2+i*1.15, w: 3.6, h: 0.9, fontSize: 10, color: 'D1D5DB', fontFace: 'Inter', wrap: true });
    });

    // Right light panel
    s.addShape('rect', { x: 5.1, y: 1.5, w: 4.5, h: 5.4, fill: { color: C.white }, rounding: 0.12, shadow: makeShadow() });
    s.addText('How It Maps to RebookIQ', { x: 5.3, y: 1.7, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.coral, fontFace: 'Inter' });

    const maps = [
      '→ Explorer Score milestone (Rome → Paris → Dubai breadth) and milestone-to-next-badge mechanic',
      '→ Post-experience trigger → badge → pick → book loop. Same segment-to-channel logic, travel context',
      '→ RebookIQ picks ranked by experience affinity × next-intent signals, not generic category',
      '→ Post-experience 30-min trigger replaces week-one cliff fix; "journey progress" = Explorer Score ring',
    ];
    maps.forEach((m, i) => {
      s.addText(m, { x: 5.3, y: 2.2+i*1.15, w: 4.1, h: 0.9, fontSize: 10, color: '374151', fontFace: 'Inter', wrap: true });
    });

    s.addText('This is not a concept. The milestone-reward retention architecture shipped at 350M MAU scale. RebookIQ applies it to travel experience breadth.', {
      x: 0.4, y: 7.05, w: 9.2, h: 0.35, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 8 — ROLLOUT PLAN ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.coral } });

    s.addText('ROLLOUT PLAN', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.gold, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Three Phases — Post-Experience Trigger to Full Explorer Ecosystem', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.white, fontFace: 'Outfit' });

    const phases = [
      {
        ph: 'Phase 1', time: 'Month 1–2 · MVP Trigger',
        body: 'Ship post-experience trigger card to 10% of bookings. Instrument: trigger-to-open rate, open-to-pick-click rate, pick-click-to-book rate. A/B test: immediate trigger (30 min) vs next-morning. Measure D7 and D30 rebooking delta vs control cohort.',
        color: C.coral,
      },
      {
        ph: 'Phase 2', time: 'Month 3–4 · Explorer Score + Badge',
        body: 'Launch Explorer Score profile and city badge system. Run A/B: badge-only vs badge + milestone vs control. Add Explorer Score to home tab as persistent engagement surface. Instrument milestone progress vs rebooking rate correlation.',
        color: C.gold,
      },
      {
        ph: 'Phase 3', time: 'Month 5–6 · Full Loop + Cross-City',
        body: 'Full rollout: trigger + badge + picks + Explorer Score + peer feed. Cross-city milestone (Rome Explorer → Paris Explorer → Dubai Explorer) drives geographic expansion. Track CLTV curve vs non-RebookIQ cohort at D30/D60/D90.',
        color: C.green,
      },
    ];

    phases.forEach((ph, i) => {
      s.addShape('rect', { x: 0.4 + i*3.2, y: 1.5, w: 3.0, h: 5.0, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: 0.4 + i*3.2, y: 1.5, w: 3.0, h: 0.05, fill: { color: ph.color } });
      s.addText(ph.ph, { x: 0.55+i*3.2, y: 1.65, w: 2.7, h: 0.35, fontSize: 12, bold: true, color: ph.color, fontFace: 'Outfit' });
      s.addText(ph.time, { x: 0.55+i*3.2, y: 2.0, w: 2.7, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Inter' });
      s.addText(ph.body, { x: 0.55+i*3.2, y: 2.45, w: 2.7, h: 3.8, fontSize: 10.5, color: 'D1D5DB', fontFace: 'Inter', wrap: true });
    });

    s.addShape('rect', { x: 0.4, y: 6.75, w: 9.2, h: 0.55, fill: { color: '1A0808' }, rounding: 0.08 });
    s.addText('What I need to start: ticket validity event from booking system · 1 FE + 1 BE engineer · Design partner for Explorer Score profile · Data Science for affinity model (I can spec it)', {
      x: 0.6, y: 6.85, w: 8.8, h: 0.35, fontSize: 10, color: C.white, fontFace: 'Inter',
    });
    s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio', {
      x: 0.4, y: 7.25, w: 9.2, h: 0.2, fontSize: 9, color: C.muted, fontFace: 'Inter', align: 'right',
    });
  }

  await prs.writeFile({ fileName: 'assets/headout-rebookiq-deck.pptx' });
  console.log('Done: assets/headout-rebookiq-deck.pptx');
})();
