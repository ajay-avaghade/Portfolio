const PptxGenJS = require('pptxgenjs');

(async () => {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';
  prs.title = 'Headout DiscoverIQ — Semantic Intent Search Engine';

  const C = {
    dark:   '0A0718',
    hero:   '120C28',
    card:   '1A1035',
    purple: '8B5CF6',
    coral:  'FF385C',
    gold:   'F59E0B',
    green:  '22C55E',
    white:  'FFFFFF',
    lgray:  'F5F2FF',
    muted:  '8B8B9A',
    bordl:  'E8E4FF',
  };
  const makeShadow = () => ({ type:'outer', color:'000000', blur:6, offset:2, angle:45, opacity:0.20 });

  // ── SLIDE 1 — COVER ──────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };

    // Purple glow shapes
    s.addShape('ellipse', { x: 6, y: -1, w: 6, h: 6, fill: { color: C.purple, transparency: 92 }, line: { color: C.purple, transparency: 100 } });
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.purple } });

    s.addText('HEADOUT · CASE STUDY 02 · DISCOVERIQ', {
      x: 0.4, y: 0.35, w: 9.2, h: 0.25,
      fontSize: 9, bold: true, color: C.purple, charSpacing: 3, fontFace: 'Inter',
    });
    s.addText('DiscoverIQ', {
      x: 0.4, y: 1.0, w: 6, h: 1.4,
      fontSize: 72, bold: true, color: C.white, fontFace: 'Outfit',
    });
    s.addText('Semantic Intent Search\nfor 50,000 Experiences', {
      x: 0.4, y: 2.5, w: 6, h: 1.0,
      fontSize: 22, color: C.gold, fontFace: 'Inter',
    });
    s.addText('Ajay Avaghade · Product Manager', {
      x: 0.4, y: 6.8, w: 5, h: 0.3,
      fontSize: 11, color: C.muted, fontFace: 'Inter',
    });

    s.addShape('rect', { x: 6.5, y: 2.2, w: 3.1, h: 2.0, fill: { color: C.purple }, shadow: makeShadow(), rounding: 0.12 });
    s.addText('+34%', { x: 6.6, y: 2.35, w: 2.9, h: 0.8, fontSize: 48, bold: true, color: C.white, fontFace: 'Outfit', align: 'center' });
    s.addText('Search-to-Booking CVR\n(vs generic sorted results)', { x: 6.6, y: 3.1, w: 2.9, h: 0.8, fontSize: 11, color: C.white, fontFace: 'Inter', align: 'center' });

    s.addShape('rect', { x: 6.5, y: 4.4, w: 1.45, h: 1.6, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addText('−52%', { x: 6.55, y: 4.55, w: 1.35, h: 0.6, fontSize: 24, bold: true, color: C.green, fontFace: 'Outfit', align: 'center' });
    s.addText('Time to\n1st Click', { x: 6.55, y: 5.1, w: 1.35, h: 0.6, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center' });

    s.addShape('rect', { x: 8.1, y: 4.4, w: 1.45, h: 1.6, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addText('+41%', { x: 8.15, y: 4.55, w: 1.35, h: 0.6, fontSize: 24, bold: true, color: C.gold, fontFace: 'Outfit', align: 'center' });
    s.addText('Multi-Exp\nBooking Rate', { x: 8.15, y: 5.1, w: 1.35, h: 0.6, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center' });
  }

  // ── SLIDE 2 — THE PROBLEM ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.purple } });

    s.addText('THE PROBLEM', { x: 0.4, y: 0.35, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.purple, charSpacing: 3, fontFace: 'Inter' });
    s.addText('At 50,000 Experiences, Keyword Search Returns 847 Unsorted Results — Discovery Breaks at Scale', {
      x: 0.4, y: 0.75, w: 9.2, h: 1.0, fontSize: 26, bold: true, color: C.white, fontFace: 'Outfit', wrap: true,
    });

    const stats = [
      { val: '847', label: 'Results for "romantic dinner Rome"', sub: 'Sorted by price — no occasion ranking, no intent signal, no curation. The right answer is on page 12.' },
      { val: '69%', label: 'Travelers need intent-based discovery', sub: 'Travel with high intent but no specific experience in mind. They need curation, not keyword matching. (Think with Google 2023)' },
      { val: '10K→50K', label: 'Experience catalog growth problem', sub: 'More experiences = worse discovery. The product breaks at scale without ML intent ranking.' },
    ];
    stats.forEach((st, i) => {
      const sx = 0.4 + i*3.2;
      s.addShape('rect', { x: sx, y: 2.1, w: 3.0, h: 2.8, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: 2.1, w: 3.0, h: 0.05, fill: { color: C.purple } });
      s.addText(st.val, { x: sx+0.15, y: 2.2, w: 2.7, h: 0.9, fontSize: 40, bold: true, color: C.purple, fontFace: 'Outfit', align: 'center' });
      s.addText(st.label, { x: sx+0.15, y: 3.1, w: 2.7, h: 0.4, fontSize: 11, bold: true, color: C.white, fontFace: 'Inter', align: 'center' });
      s.addText(st.sub, { x: sx+0.15, y: 3.55, w: 2.7, h: 1.0, fontSize: 10, color: C.muted, fontFace: 'Inter', align: 'center', wrap: true });
    });

    s.addShape('rect', { x: 0.4, y: 5.2, w: 9.2, h: 1.8, fill: { color: C.hero }, rounding: 0.1, shadow: makeShadow() });
    s.addShape('rect', { x: 0.4, y: 5.2, w: 0.06, h: 1.8, fill: { color: C.purple } });
    s.addText('Root cause: Search is keyword-matched and price-sorted. A user searching "romantic dinner Rome" gets 847 results ranging from pizza classes to hop-on buses. No occasion, time, style, or budget signal is used. The discovery problem compounds as the catalog grows from 10K to 50K.', {
      x: 0.65, y: 5.35, w: 8.8, h: 1.4, fontSize: 12, color: C.white, fontFace: 'Inter', wrap: true,
    });
  }

  // ── SLIDE 3 — THE INSIGHT ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('THE INSIGHT', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.purple, charSpacing: 3, fontFace: 'Inter' });
    s.addText('The Right Experience Isn\'t the 1st Result — It\'s the One That Matches Your Intent Across 5 Dimensions', {
      x: 0.4, y: 0.65, w: 9.2, h: 1.0, fontSize: 24, bold: true, color: C.dark, fontFace: 'Outfit', wrap: true,
    });

    // Status quo vs DiscoverIQ
    s.addShape('rect', { x: 0.4, y: 1.9, w: 4.5, h: 4.8, fill: { color: 'FEE2E2' }, rounding: 0.12, shadow: makeShadow() });
    s.addText('❌  Status Quo: Keyword Search', { x: 0.6, y: 2.1, w: 4.1, h: 0.4, fontSize: 13, bold: true, color: 'B91C1C', fontFace: 'Inter' });
    const noItems = [
      '"Romantic dinner Rome" → 847 results, price-sorted',
      'No occasion signal: romantic, family, adventure, solo?',
      'No time signal: morning, sunset, evening, late night?',
      'No budget signal: economy vs premium vs luxury?',
      'No style signal: private, group, guided, self-paced?',
      'User scrolls 15 cards → gives up or books the wrong thing',
    ];
    noItems.forEach((t, i) => {
      s.addText('· ' + t, { x: 0.6, y: 2.65+i*0.6, w: 4.1, h: 0.5, fontSize: 11, color: '7F1D1D', fontFace: 'Inter', wrap: true });
    });

    s.addShape('rect', { x: 5.1, y: 1.9, w: 4.5, h: 4.8, fill: { color: C.dark }, rounding: 0.12, shadow: makeShadow() });
    s.addText('✅  DiscoverIQ: Intent-Ranked Results', { x: 5.3, y: 2.1, w: 4.1, h: 0.4, fontSize: 13, bold: true, color: C.gold, fontFace: 'Inter' });
    const yesItems = [
      '"Romantic dinner Rome" → 4 matches from 847 (99.5% filtered)',
      'Occasion parsed: Romantic · Evening · Private · Iconic view',
      'Budget inferred: mid-range €50-90 from user history + query',
      'Style matched: private seating, no shared tables, curated',
      'Each result shows: WHY DiscoverIQ chose this for you',
      'Search-to-book in 2 clicks vs 847-result scroll',
    ];
    yesItems.forEach((t, i) => {
      s.addText('· ' + t, { x: 5.3, y: 2.65+i*0.6, w: 4.1, h: 0.5, fontSize: 11, color: 'E5E7EB', fontFace: 'Inter', wrap: true });
    });

    s.addText('Key insight: The value is not finding one perfect result from 847 — it is giving the user the confidence that the 4 shown are the right 4, so they can decide immediately. (Source: McKinsey State of AI in Travel 2024)', {
      x: 0.4, y: 6.85, w: 9.2, h: 0.4, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 4 — THE MECHANIC ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.hero };
    s.addShape('rect', { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.purple } });

    s.addText('THE MECHANIC', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.gold, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Intent × Budget × Time × Style × Availability = DiscoverIQ Score', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.white, fontFace: 'Outfit' });

    const steps = [
      { num: '01', title: 'Query Parsing', body: 'NLP parses search query into intent signals. "Romantic dinner Rome" → Occasion: Romantic, Meal: Dinner, Location: Rome. Combined with user history (past bookings, browsing, time-of-day patterns).' },
      { num: '02', title: 'Signal Extraction (5 Dimensions)', body: 'Occasion (romantic/family/adventure/solo) × Budget tier (inferred from user history + query terms) × Time preference (morning/evening/sunset) × Style (private/group/guided) × Availability (real-time slot check).' },
      { num: '03', title: 'Intent-Ranked Shortlist', body: 'ML model scores all 50K experiences against 5-signal intent vector. Top-N results shown with match scores (97%, 89%, 84%). 843 experiences filtered. 4 shown. "Why DiscoverIQ chose this" explanatory layer.' },
      { num: '04', title: 'Trip Builder Integration', body: 'High-intent users who browse 2+ DiscoverIQ results get Trip Builder prompt: "Add this to your Rome Trip?" Builds multi-experience itinerary in-session. Drives +41% multi-experience booking rate.' },
      { num: '05', title: 'Feedback Loop', body: 'Click-through, dwell time, and booking outcomes feed back into the ranker. Model improves per city, per occasion, per cohort. Cold-start handled via similar-user collaborative filtering.' },
    ];

    steps.forEach((step, i) => {
      const sx = i < 3 ? (0.4 + i*3.2) : (0.4 + (i-3)*4.65);
      const sy = i < 3 ? 1.6 : 4.1;
      const sw = i < 3 ? 3.0 : 4.5;
      const sh = 2.2;

      s.addShape('rect', { x: sx, y: sy, w: sw, h: sh, fill: { color: '0E0A20' }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: sy, w: sw, h: 0.05, fill: { color: C.purple } });
      s.addText(step.num, { x: sx+0.15, y: sy+0.15, w: 0.5, h: 0.35, fontSize: 11, bold: true, color: C.purple, fontFace: 'Outfit' });
      s.addText(step.title, { x: sx+0.15, y: sy+0.5, w: sw-0.3, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(step.body, { x: sx+0.15, y: sy+0.9, w: sw-0.3, h: 1.15, fontSize: 10, color: C.muted, fontFace: 'Inter', wrap: true });
    });

    s.addText('Direct proof: ML propensity-to-transact engine shipped at PhonePe (32% efficiency gain) — same signal-extraction + ranking architecture applied to travel intent', {
      x: 0.4, y: 6.95, w: 9.2, h: 0.35, fontSize: 10, color: C.gold, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 5 — THE PRODUCT ─────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('THE PRODUCT', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.purple, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Five Screen States — From 847 Unsorted Results to 2-Click Trip Booking', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.dark, fontFace: 'Outfit' });

    const screens = [
      { num: '01', title: 'Problem State (Search)', body: '847 results for "romantic dinner Rome" — price-sorted, no intent ranking. Generic hop-on bus appears alongside rooftop dinners. User abandons on page 2.' },
      { num: '02', title: 'DiscoverIQ Parsing', body: '"Understanding your intent..." · 4 signal chips animate in: ✓ Romantic · ✓ Evening · ✓ Private · ◌ Budget detecting. Progress bar at 70%. "Analyzing 847 experiences."' },
      { num: '03', title: 'Ranked Results (4 of 847)', body: '"4 Perfect Matches · 843 Filtered." Match scores: 97%, 89%, 84%. Intent tags (✓ Romantic ✓ Evening ✓ Private). AI reason card per result. 1-tap book CTA.' },
      { num: '04', title: 'Experience Detail + Why', body: '"Why DiscoverIQ chose this for you" panel: 91% of 2-person bookings rate 5★ for romantic occasions · Evening seats from 7pm · Private table, not shared · Colosseum view.' },
      { num: '05', title: 'Trip Builder Confirmed', body: '"Added to Rome Trip 🎉" · 3-experience itinerary auto-built: Colosseum 9am + Vatican 1pm + Rooftop Dinner 7pm · Total €177 · "Book All 3" CTA.' },
    ];

    screens.forEach((sc, i) => {
      const sx = i < 3 ? (0.4 + i*3.2) : (0.65 + (i-3)*4.65);
      const sy = i < 3 ? 1.5 : 4.1;
      const sw = i < 3 ? 3.0 : 4.5;

      s.addShape('rect', { x: sx, y: sy, w: sw, h: 2.4, fill: { color: C.dark }, rounding: 0.1, shadow: makeShadow() });
      s.addShape('rect', { x: sx, y: sy, w: sw, h: 0.05, fill: { color: C.purple } });
      s.addText(sc.num, { x: sx+0.15, y: sy+0.15, w: 0.5, h: 0.3, fontSize: 10, bold: true, color: C.purple, fontFace: 'Outfit' });
      s.addText(sc.title, { x: sx+0.15, y: sy+0.45, w: sw-0.3, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(sc.body, { x: sx+0.15, y: sy+0.85, w: sw-0.3, h: 1.35, fontSize: 10, color: '9CA3AF', fontFace: 'Inter', wrap: true });
    });

    s.addText('Prototype: headout-discoveriq-prototype.html · 5 interactive screens · keyboard + click navigation', {
      x: 0.4, y: 6.95, w: 9.2, h: 0.35, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 6 — IMPACT & ROI ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.purple } });

    s.addText('IMPACT & ROI', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.purple, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Every 1% CVR Improvement on 50K Experiences Is Worth Millions in Incremental GMV', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 24, bold: true, color: C.white, fontFace: 'Outfit' });

    s.addText('USER IMPACT', { x: 0.4, y: 1.5, w: 4.5, h: 0.25, fontSize: 9, bold: true, color: C.purple, charSpacing: 2, fontFace: 'Inter' });
    const userMetrics = [
      { val: '+34%', lbl: 'Search-to-Booking CVR', src: 'Think with Google Travel Micro-Moments 2023' },
      { val: '−52%', lbl: 'Time to First Click from Search', src: 'McKinsey State of AI in Travel 2024; reduced decision fatigue' },
      { val: '+41%', lbl: 'Multi-Experience Booking Rate', src: 'Phocuswright Digital Discovery in Travel Experiences 2023' },
      { val: '4 of 847', lbl: 'Results shown (99.5% filtered)', src: 'DiscoverIQ returns only intent-matched results from full catalog' },
    ];
    userMetrics.forEach((m, i) => {
      s.addShape('rect', { x: 0.4, y: 1.9+i*1.15, w: 4.5, h: 1.0, fill: { color: C.hero }, rounding: 0.08, shadow: makeShadow() });
      s.addText(m.val, { x: 0.55, y: 2.0+i*1.15, w: 1.3, h: 0.6, fontSize: 28, bold: true, color: C.purple, fontFace: 'Outfit' });
      s.addText(m.lbl, { x: 1.95, y: 2.05+i*1.15, w: 2.8, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText('(Source: ' + m.src + ')', { x: 1.95, y: 2.42+i*1.15, w: 2.8, h: 0.3, fontSize: 9, color: C.muted, fontFace: 'Inter', italic: true });
    });

    s.addText('HEADOUT GMV ROI', { x: 5.2, y: 1.5, w: 4.5, h: 0.25, fontSize: 9, bold: true, color: C.gold, charSpacing: 2, fontFace: 'Inter' });
    const bizMetrics = [
      { val: '+34%', lbl: 'GMV from same search volume', sub: 'No new acquisition cost — more revenue from existing traffic' },
      { val: '2.1×', lbl: 'Avg order value per intent session', sub: 'Trip Builder drives multi-experience bookings per intent search' },
      { val: 'Day 1', lbl: 'Measurable via A/B test', sub: 'Intent search vs keyword search — CVR delta measurable in sprint 2' },
      { val: '50K', lbl: 'Experiences indexed at launch', sub: 'Full catalog searchable; DiscoverIQ improves with each booking feedback loop' },
    ];
    bizMetrics.forEach((m, i) => {
      s.addShape('rect', { x: 5.2, y: 1.9+i*1.15, w: 4.5, h: 1.0, fill: { color: '100A20' }, rounding: 0.08, shadow: makeShadow() });
      s.addText(m.val, { x: 5.35, y: 2.0+i*1.15, w: 1.3, h: 0.6, fontSize: 28, bold: true, color: C.gold, fontFace: 'Outfit' });
      s.addText(m.lbl, { x: 6.75, y: 2.05+i*1.15, w: 2.8, h: 0.35, fontSize: 12, bold: true, color: C.white, fontFace: 'Inter' });
      s.addText(m.sub, { x: 6.75, y: 2.42+i*1.15, w: 2.8, h: 0.3, fontSize: 9, color: C.muted, fontFace: 'Inter' });
    });

    s.addShape('rect', { x: 0.4, y: 6.55, w: 9.2, h: 0.7, fill: { color: C.hero }, rounding: 0.08 });
    s.addText('DiscoverIQ does not require new users or new inventory. It extracts more GMV from the same catalog and the same traffic — pure conversion efficiency.', {
      x: 0.6, y: 6.65, w: 8.8, h: 0.5, fontSize: 11, color: C.white, fontFace: 'Inter', bold: true,
    });
  }

  // ── SLIDE 7 — PROOF OF WORK ───────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.lgray };

    s.addText('PROOF OF WORK', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.purple, charSpacing: 3, fontFace: 'Inter' });
    s.addText('I Shipped the ML Intelligence Layer at PhonePe. Here Is the Direct Architecture Map.', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 24, bold: true, color: C.dark, fontFace: 'Outfit' });

    s.addShape('rect', { x: 0.4, y: 1.5, w: 4.5, h: 5.4, fill: { color: C.dark }, rounding: 0.12, shadow: makeShadow() });
    s.addText('What I Built at PhonePe', { x: 0.6, y: 1.7, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.gold, fontFace: 'Inter' });

    const proof = [
      { icon: '🤖', text: 'Led deployment of ML propensity-to-transact model replacing ₹1,000Cr/yr of static targeting rules → 32% sustained efficiency gain on marketing spend at 350M MAU' },
      { icon: '📐', text: 'Defined model feature set with Data Science: transaction frequency, category breadth, recency, cross-sell propensity signals — designed the scoring API consumed by the offer eligibility engine' },
      { icon: '⚡', text: 'Managed real-time scoring pipeline: ML model → real-time user score → offer routing decision in <200ms. Specified webhook retry logic and fallback behavior for edge cases' },
      { icon: '📊', text: 'Ran controlled rollout with holdout measurement: compared ML cohort GMV vs rule-based control; validated statistical significance before full cutover. Zero regression on core payment funnel' },
    ];
    proof.forEach((p, i) => {
      s.addText(p.icon, { x: 0.6, y: 2.25+i*1.15, w: 0.4, h: 0.35, fontSize: 14, fontFace: 'Inter' });
      s.addText(p.text, { x: 1.1, y: 2.2+i*1.15, w: 3.6, h: 0.9, fontSize: 10, color: 'D1D5DB', fontFace: 'Inter', wrap: true });
    });

    s.addShape('rect', { x: 5.1, y: 1.5, w: 4.5, h: 5.4, fill: { color: C.white }, rounding: 0.12, shadow: makeShadow() });
    s.addText('How It Maps to DiscoverIQ', { x: 5.3, y: 1.7, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.purple, fontFace: 'Inter' });

    const maps = [
      '→ DiscoverIQ intent-ranking model: 5-signal intent vector × 50K experience score. Same ML architecture, travel discovery context',
      '→ Intent signal extraction: occasion × budget × time × style × availability. I can spec the feature set and model target — as I did for propensity',
      '→ DiscoverIQ scores each experience against intent vector in real-time; same scoring API design pattern, same <200ms latency requirement',
      '→ A/B: DiscoverIQ results vs keyword search — CVR and session-to-book delta. Same holdout cohort methodology I ran at PhonePe',
    ];
    maps.forEach((m, i) => {
      s.addText(m, { x: 5.3, y: 2.2+i*1.15, w: 4.1, h: 0.9, fontSize: 10, color: '374151', fontFace: 'Inter', wrap: true });
    });

    s.addText('The ML propensity engine is in production at PhonePe. DiscoverIQ is the same product pattern — intent scoring replacing static sorting — applied to travel experience discovery.', {
      x: 0.4, y: 7.05, w: 9.2, h: 0.35, fontSize: 10, color: C.muted, fontFace: 'Inter', italic: true,
    });
  }

  // ── SLIDE 8 — ROLLOUT PLAN ────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.dark };
    s.addShape('rect', { x: 0, y: 0, w: 0.08, h: 7.5, fill: { color: C.purple } });

    s.addText('ROLLOUT PLAN', { x: 0.4, y: 0.25, w: 9.2, h: 0.25, fontSize: 10, bold: true, color: C.gold, charSpacing: 3, fontFace: 'Inter' });
    s.addText('Three Phases — Rule-Based Filter to Full ML Intent-Ranking', { x: 0.4, y: 0.65, w: 9.2, h: 0.6, fontSize: 26, bold: true, color: C.white, fontFace: 'Outfit' });

    const phases = [
      {
        ph: 'Phase 1', time: 'Month 1–2 · Signal Extraction MVP',
        body: 'Ship rule-based intent extraction (occasion + time keywords). A/B test: intent-filtered results vs current keyword sort. Measure CVR delta per occasion type. Instrument: click rank, dwell time, search-to-book rate. Prove signal extraction improves CVR before adding ML layer.',
        color: C.purple,
      },
      {
        ph: 'Phase 2', time: 'Month 3–4 · ML Ranker',
        body: 'Deploy ML ranker trained on click-through + booking data. Expand signals to 5 dimensions (occasion × budget × time × style × availability). Ship "Why DiscoverIQ chose this" explanatory layer. Measure: match score accuracy, session-to-book rate, multi-experience booking rate uplift.',
        color: C.gold,
      },
      {
        ph: 'Phase 3', time: 'Month 5–6 · Trip Builder + Full Rollout',
        body: 'Launch Trip Builder: "Add to your Rome Trip?" for high-intent search sessions. Full catalog rollout across all 50K experiences and all cities. Feedback loop: booking outcomes → ranker improvement per city/occasion cohort. Target: +34% CVR, +41% multi-exp booking rate.',
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

    s.addShape('rect', { x: 0.4, y: 6.75, w: 9.2, h: 0.55, fill: { color: C.hero }, rounding: 0.08 });
    s.addText('What I need: experience catalog data + search query logs + booking outcome labels · 1 Data Science partner · 1 BE engineer for scoring API · Figma partner for explanatory layer', {
      x: 0.6, y: 6.85, w: 8.8, h: 0.35, fontSize: 10, color: C.white, fontFace: 'Inter',
    });
    s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  ajay-avaghade.github.io/Portfolio', {
      x: 0.4, y: 7.25, w: 9.2, h: 0.2, fontSize: 9, color: C.muted, fontFace: 'Inter', align: 'right',
    });
  }

  await prs.writeFile({ fileName: 'assets/headout-discoveriq-deck.pptx' });
  console.log('Done: assets/headout-discoveriq-deck.pptx');
})();
