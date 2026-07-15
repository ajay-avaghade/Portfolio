const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';
  prs.title = 'Headout BetweenTrips — Between-Trip Engagement Engine';

  const C = {
    dark:  '0D0D0D',
    hero:  '1A1200',
    card:  '231900',
    amber: 'F59E0B',
    coral: 'FF385C',
    green: '22C55E',
    purple:'8B5CF6',
    white: 'FFFFFF',
    lgray: 'FFFBF0',
    muted: '888880',
    bordl: 'FDE68A',
  };
  const makeShadow = () => ({ type:'outer', color:'000000', blur:6, offset:2, angle:45, opacity:0.18 });

  // ── SLIDE 1 — COVER ──────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    s.addShape('ellipse', { x: 5.5, y: -0.5, w: 5, h: 5, fill: { color: C.amber, transparency: 94 }, line: { color: C.amber, transparency: 100 } });
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.amber } });

    s.addText('HEADOUT · CASE STUDY 03 · BETWEENTRIPS', {
      x: 0.4, y: 0.35, w: 9.2, h: 0.25, fontSize: 9, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter',
    });
    s.addText('BetweenTrips', {
      x: 0.4, y: 1.0, w: 6, h: 1.4, fontSize: 60, bold: true, color: C.white, fontFace: 'Outfit',
    });
    s.addText('The Between-Trip\nEngagement Engine', {
      x: 0.4, y: 2.5, w: 6, h: 1.0, fontSize: 22, color: C.amber, fontFace: 'Inter',
    });
    s.addText('Ajay Avaghade · Product Manager', {
      x: 0.4, y: 6.8, w: 5, h: 0.3, fontSize: 11, color: C.muted, fontFace: 'Inter',
    });

    s.addShape('rect', { x: 6.5, y: 2.2, w: 3.1, h: 2.0, fill: { color: C.amber }, shadow: makeShadow(), rounding: 0.12 });
    s.addText('+3.2×', { x: 6.6, y: 2.35, w: 2.9, h: 0.8, fontSize: 44, bold: true, color: C.dark, fontFace: 'Outfit', align: 'center' });
    s.addText('App Opens/Month\n(0.8× → 2.6× projected)', { x: 6.6, y: 3.1, w: 2.9, h: 0.8, fontSize: 11, color: C.dark, fontFace: 'Inter', align: 'center', bold: true });

    s.addShape('rect', { x: 6.5, y: 4.4, w: 1.45, h: 1.6, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addText('+19%', { x: 6.55, y: 4.55, w: 1.35, h: 0.6, fontSize: 24, bold: true, color: C.green, fontFace: 'Outfit', align: 'center' });
    s.addText('Cross-City\nBooking', { x: 6.55, y: 5.1, w: 1.35, h: 0.6, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center' });

    s.addShape('rect', { x: 8.1, y: 4.4, w: 1.45, h: 1.6, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addText('+44%', { x: 8.15, y: 4.55, w: 1.35, h: 0.6, fontSize: 24, bold: true, color: C.amber, fontFace: 'Outfit', align: 'center' });
    s.addText('Wishlist→\nBook CVR', { x: 8.15, y: 5.1, w: 1.35, h: 0.6, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center' });
  }

  // ── SLIDE 2 — THE PROBLEM ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.amber } });

    s.addText('THE PROBLEM', { x: 0.4, y: 0.35, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Headout\'s App Has No Reason to Exist Between Trips — 0.8 Opens/Month vs 8.4 for Rideshare', {
      x: 0.4, y: 0.75, w: 9.2, h: 1.0, fontSize: 25, bold: true, color: C.white, fontFace: 'Outfit', wrap: true,
    });

    const stats = [
      { val: '0.8×', label: 'Monthly app opens between trips', sub: 'Vs 8.4× for rideshare, 4.2× for food delivery. App opened once per trip, closed at ticket delivery.' },
      { val: '47 wks', label: 'Average gap between bookings', sub: 'The product exists for 6 hours of booking flow and then disappears from users\' lives for nearly a year.' },
      { val: '0%', label: 'Engagement surface between trips', sub: 'No memory layer, no peer activity, no intent capture, no proactive card. Zero reasons to open the app.' },
    ];
    stats.forEach((st, i) => {
      const sx = 0.4 + i*3.2;
      s.addShape('rect', { x: sx, y: 2.1, w: 3.0, h: 2.8, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: 2.1, w: 3.0, h: 0.05, fill: { color: C.amber } });
      s.addText(st.val, { x: sx+0.15, y: 2.2, w: 2.7, h: 0.9, fontSize: 44, bold: true, color: C.amber, fontFace: 'Outfit', align: 'center' });
      s.addText(st.label, { x: sx+0.15, y: 3.1, w: 2.7, h: 0.4, fontSize: 11, bold: true, color: C.white, fontFace: 'Inter', align: 'center' });
      s.addText(st.sub, { x: sx+0.15, y: 3.55, w: 2.7, h: 1.0, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center', wrap: true });
    });

    s.addShape('rect', { x: 0.4, y: 5.2, w: 9.2, h: 1.8, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addShape('rect', { x: 0.4, y: 5.2, w: 0.06, h: 1.8, fill: { color: C.amber } });
    s.addText('Root cause: The app is transactional by design — it serves the booking and ticket delivery, then exits. There is no memory layer (your Rome story), no social layer (what your friends are exploring), and no intent capture layer (you searched Dubai 3 times this week). BetweenTrips builds all three.', {
      x: 0.65, y: 5.35, w: 8.8, h: 1.4, fontSize: 12, color: C.white, fontFace: 'Inter', wrap: true,
    });
  }

  // ── SLIDE 3 — THE INSIGHT ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('THE INSIGHT', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter' });
    s.addText('The 47-Week Gap Is Not Dead Time — It\'s the Highest-Intent Planning Window Headout Is Missing', {
      x: 0.4, y: 0.65, w: 9.2, h: 1.0, fontSize: 24, bold: true, color: C.dark, fontFace: 'Outfit', wrap: true,
    });

    s.addShape('rect', { x: 0.4, y: 1.9, w: 4.5, h: 4.8, fill: { color: 'FEE2E2' }, rounding: 0.12, shadow: makeShadow() });
    s.addText('❌  Today: Transactional App', { x: 0.6, y: 2.1, w: 4.1, h: 0.4, fontSize: 13, bold: true, color: 'B91C1C', fontFace: 'Inter' });
    const noItems = [
      'Post-trip: home screen shows "No upcoming trips"',
      'Zero memory of Rome, Vatican, Colosseum experiences',
      'Generic feed (London Eye, Paris Eiffel) — no personalization',
      'Friends\' travel activity invisible — no social signal',
      'User searches "Dubai activities" → no response',
      'Result: app deleted from homescreen by month 2',
    ];
    noItems.forEach((t, i) => {
      s.addText('· ' + t, { x: 0.6, y: 2.65+i*0.6, w: 4.1, h: 0.5, fontSize: 11, color: '7F1D1D', fontFace: 'Inter', wrap: true });
    });

    s.addShape('rect', { x: 5.1, y: 1.9, w: 4.5, h: 4.8, fill: { color: C.dark }, rounding: 0.12, shadow: makeShadow() });
    s.addText('✅  With BetweenTrips', { x: 5.3, y: 2.1, w: 4.1, h: 0.4, fontSize: 13, bold: true, color: C.amber, fontFace: 'Inter' });
    const yesItems = [
      'Post-trip: "Your Rome 2024 Journey" memory card on home',
      'Journal with photo moments, experiences done, badge earned',
      'Peer feed: "Maya booked Bali Monkey Forest" · "Raj: Dubai Frame"',
      '"You searched Dubai 3 times — Planning Dubai?" proactive card',
      'Dubai wishlist synced — 1-tap to book when ready',
      'Between-trip → next trip intent captured and converted',
    ];
    yesItems.forEach((t, i) => {
      s.addText('· ' + t, { x: 5.3, y: 2.65+i*0.6, w: 4.1, h: 0.5, fontSize: 11, color: 'E5E7EB', fontFace: 'Inter', wrap: true });
    });

    s.addText('Key insight: The user who searched Dubai 3 times this week is Headout\'s highest-intent lead — and the app has no mechanism to capture or respond to it. BetweenTrips closes this gap. (Source: Skift Experiences Platform Engagement 2024)', {
      x: 0.4, y: 6.85, w: 9.2, h: 0.4, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 4 — THE MECHANIC ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.hero };
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.amber } });

    s.addText('THE MECHANIC', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Three Layers: Memory · Social · Intent. One Home Screen Between Trips.', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.white, fontFace: 'Outfit' });

    const steps = [
      { num: '01', title: 'Memory Layer — Travel Journal', body: 'Post-trip: city journey card auto-built from booking data (experiences done, rating given, dates). Photo moments from timestamps. Badge earned. "Share your Rome story" CTA. User returns to relive → repeat engagement.' },
      { num: '02', title: 'Social Layer — Peer Discovery', body: 'Friends\' booking activity feed: "Maya booked Bali Monkey Forest" · "Raj completed Dubai Frame." Peer activity is the highest-trust discovery signal. Opted-in friends. No social graph required — just booking activity.' },
      { num: '03', title: 'Intent Layer — Proactive Planning Card', body: '"Planning Dubai?" fires when in-app search history detects 3+ Dubai searches. Card: top 3 Dubai picks, 2 friends who visited, wishlist add. Same context-aware trigger mechanic as PhonePe cart incentivisation — intent × time → right card.' },
      { num: '04', title: 'Wishlist → Booking Loop', body: 'Saved experiences get real-time availability nudges: "3 slots left tomorrow." Wishlist → Book in 1 tap. No re-discovery friction. Trip planning session stays open until booked.' },
      { num: '05', title: 'Engagement Loop Closes', body: 'New trip booked → new journal entry pre-built → new peer activity → new intent signals. BetweenTrips becomes the between-trip home screen — not a trip-booking app but a travel lifestyle platform.' },
    ];

    steps.forEach((step, i) => {
      const sx = i < 3 ? (0.4 + i*3.2) : (0.4 + (i-3)*4.65);
      const sy = i < 3 ? 1.6 : 4.1;
      const sw = i < 3 ? 3.0 : 4.5;

      s.addShape('rect', { x: sx, y: sy, w: sw, h: 2.2, fill: { color: '1C1400' }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: sy, w: sw, h: 0.05, fill: { color: C.amber } });
      s.addText(step.num, { x: sx+0.15, y: sy+0.15, w: 0.5, h: 0.35, fontSize: 11, bold: true, color: C.amber, fontFace: 'Outfit' });
      s.addText(step.title, { x: sx+0.15, y: sy+0.5, w: sw-0.3, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(step.body, { x: sx+0.15, y: sy+0.9, w: sw-0.3, h: 1.15, fontSize: 10, color: C.muted, fontFace: 'Inter', wrap: true });
    });

    s.addText('Direct proof: PhonePe cart incentivisation — context-aware trigger (cart × intent × time) → 60% abandonment reduction, 20% D30 retention uplift. BetweenTrips applies same intent-trigger architecture to travel planning.', {
      x: 0.4, y: 6.95, w: 9.2, h: 0.35, fontSize: 10, color: C.amber, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 5 — THE PRODUCT ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('THE PRODUCT', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Five Screen States — From Empty App to Between-Trip Engagement to Next Booking', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 25, bold: true, color: C.dark, fontFace: 'Outfit' });

    const screens = [
      { num: '01', title: 'Problem State (Today)', body: '"Good morning, Ajay 👋 · No upcoming trips." Generic London Eye / Eiffel Tower feed. Zero connection to Rome trip 3 weeks ago. Zero intent capture. App about to be deleted.' },
      { num: '02', title: 'BetweenTrips Home (New)', body: 'Memory card: "Rome 2024 Journey · 3 weeks ago." Peer feed: Maya (Bali), Raj (Dubai Frame), Priya (Desert Safari wishlist). Proactive card: "Planning Dubai? → See picks." ' },
      { num: '03', title: 'Rome Memory Journal', body: 'City hero card. 6-photo grid from trip. Experiences done: Colosseum ✓, Vatican ✓, Trevi (saved, not booked → Book → CTA). "Share Rome Story" + "Plan Next Trip" actions.' },
      { num: '04', title: 'Planning Dubai Intent Card', body: 'Intent signals: "Searched Dubai 3 times · 2 friends visited." Top 3 experiences: Desert Safari (Priya saved) · Burj Khalifa (Raj recommends) · Dubai Frame. "Plan my Dubai Trip →" CTA.' },
      { num: '05', title: 'Wishlist → Booked', body: 'Desert Safari confirmed (AED 149). Burj Khalifa: "3 slots left tomorrow" urgency nudge. Dubai Frame flexible. Trip total: AED 417 if all booked. Wishlist-to-book in 1 tap.' },
    ];

    screens.forEach((sc, i) => {
      const sx = i < 3 ? (0.4 + i*3.2) : (0.65 + (i-3)*4.65);
      const sy = i < 3 ? 1.5 : 4.1;
      const sw = i < 3 ? 3.0 : 4.5;

      s.addShape('rect', { x: sx, y: sy, w: sw, h: 2.4, fill: { color: C.dark }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: sy, w: sw, h: 0.05, fill: { color: C.amber } });
      s.addText(sc.num, { x: sx+0.15, y: sy+0.15, w: 0.5, h: 0.3, fontSize: 10, bold: true, color: C.amber, fontFace: 'Outfit' });
      s.addText(sc.title, { x: sx+0.15, y: sy+0.45, w: sw-0.3, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(sc.body, { x: sx+0.15, y: sy+0.85, w: sw-0.3, h: 1.35, fontSize: 10, color: '9CA3AF', fontFace: 'Inter', wrap: true });
    });

    s.addText('Prototype: headout-betweentrips-prototype.html · 5 interactive screens · keyboard + click navigation', {
      x: 0.4, y: 6.95, w: 9.2, h: 0.35, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 6 — IMPACT & ROI ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.amber } });

    s.addText('IMPACT & ROI', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter' });
    s.addText('BetweenTrips Converts the 47-Week Gap into a Revenue-Generating Engagement Window', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 23, bold: true, color: C.white, fontFace: 'Outfit' });

    s.addText('ENGAGEMENT IMPACT', { x: 0.4, y: 1.5, w: 4.5, h: 0.25, fontSize: 9, bold: true, color: C.amber, charSpacing: 2, fontFace: 'Inter' });
    const userMetrics = [
      { val: '+3.2×', lbl: 'Monthly App Opens (0.8 → 2.6×)', src: 'App Annie/data.ai State of Mobile 2024; travel app benchmark' },
      { val: '+19%', lbl: 'Cross-City Booking Rate', src: 'Skift Research Experiences Platform Engagement 2024' },
      { val: '+44%', lbl: 'Wishlist-to-Book Conversion', src: 'Contentsquare Digital Experience Benchmark 2024; availability nudge effect' },
      { val: '60%', lbl: 'Abandonment Reduction (PhonePe proof)', src: 'Cart incentivisation shipped at PhonePe; same intent-trigger mechanic' },
    ];
    userMetrics.forEach((m, i) => {
      s.addShape('rect', { x: 0.4, y: 1.9+i*1.15, w: 4.5, h: 1.0, fill: { color: C.hero }, rounding: 0.08, shadow: makeShadow() });
      s.addText(m.val, { x: 0.55, y: 2.0+i*1.15, w: 1.3, h: 0.6, fontSize: 26, bold: true, color: C.amber, fontFace: 'Outfit' });
      s.addText(m.lbl, { x: 1.95, y: 2.05+i*1.15, w: 2.8, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText('(Source: ' + m.src + ')', { x: 1.95, y: 2.42+i*1.15, w: 2.8, h: 0.3, fontSize: 9, color: C.muted, fontFace: 'Inter', italic: true });
    });

    s.addText('HEADOUT BUSINESS ROI', { x: 5.2, y: 1.5, w: 4.5, h: 0.25, fontSize: 9, bold: true, color: C.green, charSpacing: 2, fontFace: 'Inter' });
    const bizMetrics = [
      { val: '0', lbl: 'Additional CAC required', sub: 'All intent captured from existing users who are already thinking about their next trip' },
      { val: '2.6×', lbl: 'Sessions/month per active user', sub: 'From 0.8× → 2.6× via journal + peer feed + intent cards. More sessions = more booking opportunities' },
      { val: '+19%', lbl: 'Revenue/user (cross-city expansion)', sub: 'Explorer mechanic + peer social proof drives Rome users to book Dubai/Paris/Barcelona' },
      { val: 'Q1', lbl: 'Measurable DAU/MAU shift', sub: 'A/B: BetweenTrips home vs standard home. DAU/MAU and session frequency delta measurable in 4 weeks' },
    ];
    bizMetrics.forEach((m, i) => {
      s.addShape('rect', { x: 5.2, y: 1.9+i*1.15, w: 4.5, h: 1.0, fill: { color: '1A1400' }, rounding: 0.08, shadow: makeShadow() });
      s.addText(m.val, { x: 5.35, y: 2.0+i*1.15, w: 1.3, h: 0.6, fontSize: 26, bold: true, color: C.green, fontFace: 'Outfit' });
      s.addText(m.lbl, { x: 6.75, y: 2.05+i*1.15, w: 2.8, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(m.sub, { x: 6.75, y: 2.42+i*1.15, w: 2.8, h: 0.3, fontSize: 9, color: C.muted, fontFace: 'Inter' });
    });

    s.addShape('rect', { x: 0.4, y: 6.55, w: 9.2, h: 0.7, fill: { color: C.hero }, rounding: 0.08 });
    s.addText('BetweenTrips turns the 47-week gap from dead time into a planning funnel. The user who searched Dubai 3 times is Headout\'s highest-intent lead — and this is the product that captures and converts them.', {
      x: 0.6, y: 6.65, w: 8.8, h: 0.5, fontSize: 11, color: C.white, fontFace: 'Inter', bold: true,
    });
  }

  // ── SLIDE 7 — PROOF OF WORK ───────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('PROOF OF WORK', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter' });
    s.addText('I Shipped the Trigger-Based Engagement System at PhonePe. Here Is the Direct Map.', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 24, bold: true, color: C.dark, fontFace: 'Outfit' });

    s.addShape('rect', { x: 0.4, y: 1.5, w: 4.5, h: 5.4, fill: { color: C.dark }, rounding: 0.12, shadow: makeShadow() });
    s.addText('What I Built at PhonePe', { x: 0.6, y: 1.7, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.amber, fontFace: 'Inter' });

    const proof = [
      { icon: '🎯', text: 'Built dynamic cart incentivisation engine for Pincode (quick commerce) — context-aware triggers (cart value × user intent × time of day) → 35% AOV uplift, 60% cart abandonment reduction, 20% D30 retention uplift' },
      { icon: '🔄', text: 'Identified cart abandonment as the core acquisition-to-retention problem (not PMF) — ran 3 successive A/B experiments, killed 2 underperforming variants, scaled the winner based on D7/D30 cohort retention curves' },
      { icon: '📊', text: 'Built D30 retention lifecycle: 7-day activation cliff identification → incentivized onboarding → journey-progress bar → CRM intervention at right segment state → +15% D30 retention across new-user cohort' },
      { icon: '⚡', text: 'Negotiated a 7-day masthead ad moratorium with the ads team (backed by CLTV model) — replaced ads with journey-progress bar for new users. Same "right content at right moment" logic as BetweenTrips proactive card' },
    ];
    proof.forEach((p, i) => {
      s.addText(p.icon, { x: 0.6, y: 2.25+i*1.15, w: 0.4, h: 0.35, fontSize: 14, fontFace: 'Inter' });
      s.addText(p.text, { x: 1.1, y: 2.2+i*1.15, w: 3.6, h: 0.9, fontSize: 10, color: 'D1D5DB', fontFace: 'Inter', wrap: true });
    });

    s.addShape('rect', { x: 5.1, y: 1.5, w: 4.5, h: 5.4, fill: { color: C.white }, rounding: 0.12, shadow: makeShadow() });
    s.addText('How It Maps to BetweenTrips', { x: 5.3, y: 1.7, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.amber, fontFace: 'Inter' });

    const maps = [
      '→ "Planning Dubai?" proactive card: search intent (3 Dubai searches) × time signal → right card at right moment. Same trigger architecture as cart incentivisation',
      '→ Same A/B methodology: Memory feed vs control, peer feed vs control, intent card vs no card. Kill/scale decisions based on session frequency and booking rate delta',
      '→ BetweenTrips engagement loop = travel version of D30 retention lifecycle. Memory journal replaces journey-progress bar. Peer feed replaces ads. Intent card = CRM intervention',
      '→ Wishlist availability nudge ("3 slots left tomorrow") = same urgency mechanic as cart composition × time signal. Converts saved intent to booked trip',
    ];
    maps.forEach((m, i) => {
      s.addText(m, { x: 5.3, y: 2.2+i*1.15, w: 4.1, h: 0.9, fontSize: 10, color: '374151', fontFace: 'Inter', wrap: true });
    });

    s.addText('The cart incentivisation system is in production at PhonePe — 350M MAU, shipped and measured. BetweenTrips applies the same context-aware trigger architecture to travel planning intent.', {
      x: 0.4, y: 7.05, w: 9.2, h: 0.35, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 8 — ROLLOUT PLAN ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.amber } });

    s.addText('ROLLOUT PLAN', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.amber, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Three Phases — Memory Layer to Full Between-Trip Engagement Platform', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 25, bold: true, color: C.white, fontFace: 'Outfit' });

    const phases = [
      {
        ph: 'Phase 1', time: 'Month 1–2 · Memory Layer (Journal)',
        body: 'Auto-build trip journal from booking data for all past trips. A/B test: BetweenTrips home (with journey card) vs standard home. Measure: 30-day app open rate, session length, D30 retention delta. Low engineering cost — purely reads existing booking data.',
        color: C.amber,
      },
      {
        ph: 'Phase 2', time: 'Month 3–4 · Social Layer + Intent Card',
        body: 'Launch opt-in peer activity feed (friends\' bookings). Ship "Planning [City]?" intent trigger based on search history (3+ searches = fire card). A/B: intent card vs no card → measure search-to-book conversion rate. Instrument peer-activity click-to-wishlist and wishlist-to-book.',
        color: C.green,
      },
      {
        ph: 'Phase 3', time: 'Month 5–6 · Full Loop + Wishlist Intelligence',
        body: 'Wishlist real-time availability nudges ("3 slots left tomorrow"). BetweenTrips home tab becomes default between-trip surface. Cross-city intent detection (Rome user searching Paris → trigger Paris card). Track CLTV curve: BetweenTrips vs control at D30/D90/D180.',
        color: C.purple,
      },
    ];

    phases.forEach((ph, i) => {
      s.addShape('rect', { x: 0.4 + i*3.2, y: 1.5, w: 3.0, h: 5.0, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: 0.4 + i*3.2, y: 1.5, w: 3.0, h: 0.05, fill: { color: ph.color } });
      s.addText(ph.ph, { x: 0.55+i*3.2, y: 1.65, w: 2.7, h: 0.35, fontSize: 12, bold: true, color: ph.color, fontFace: 'Outfit' });
      s.addText(ph.time, { x: 0.55+i*3.2, y: 2.0, w: 2.7, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Inter' });
      s.addText(ph.body, { x: 0.55+i*3.2, y: 2.45, w: 2.7, h: 3.8, fontSize: 10.5, color: 'D1D5DB', fontFace: 'Inter', wrap: true });
    });

    s.addShape('rect', { x: 0.4, y: 6.75, w: 9.2, h: 0.55, fill: { color: C.hero }, rounding: 0.08 });
    s.addText('What I need to start: past booking data access (Phase 1 reads existing data only) · push notification infrastructure · 1 FE engineer · Design partner for journal card', {
      x: 0.6, y: 6.85, w: 8.8, h: 0.35, fontSize: 10, color: C.white, fontFace: 'Inter',
    });
    s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio', {
      x: 0.4, y: 7.25, w: 9.2, h: 0.2, fontSize: 9, color: C.muted, fontFace: 'Inter', align: 'right',
    });
  }

  await prs.writeFile({ fileName: 'assets/headout-betweentrips-deck.pptx' });
  console.log('Done: assets/headout-betweentrips-deck.pptx');
})();
