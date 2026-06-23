const PptxGenJS = require('pptxgenjs');

// ── SHARED CONSTANTS ──
const C = {
  dark:'080C16', hero:'111827', orange:'FF5C00', orangeL:'F97316',
  green:'22C55E', purple:'8B5CF6', white:'FFFFFF', lgray:'F1F5F9',
  muted:'94A3B8', dark2:'0F172A'
};
const sh = () => ({type:'outer',color:'000000',blur:4,offset:2,angle:45,opacity:0.14});

function buildDeck(title, subtitle, coverStat, coverStatLbl, accentColor, slides, outFile) {
  const prs = new PptxGenJS();
  prs.layout = 'LAYOUT_16x9';
  const A = accentColor;

  // COVER
  {
    const s = prs.addSlide();
    s.background={color:C.dark};
    for(let i=0;i<12;i++) s.addShape(prs.ShapeType.rect,{x:i*0.9-0.2,y:7.5-i*0.65,w:0.06,h:8,fill:{color:A,transparency:92},line:{color:A,transparency:92}});
    s.addText('Snabbit',{x:0.5,y:0.5,w:5,h:0.4,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:3});
    s.addText('PRODUCT CASE STUDY',{x:0.5,y:0.85,w:5,h:0.3,fontSize:8,color:C.muted,fontFace:'Calibri',charSpacing:2});
    s.addText(title,{x:0.5,y:1.4,w:8.5,h:1.1,fontSize:40,bold:true,color:C.white,fontFace:'Calibri'});
    s.addText(subtitle,{x:0.5,y:2.55,w:7.5,h:0.5,fontSize:17,color:C.muted,fontFace:'Calibri'});
    s.addShape(prs.ShapeType.rect,{x:0.5,y:3.15,w:1.8,h:0.05,fill:{color:A},line:{color:A}});
    s.addText('Presented by Ajay Avaghade · PM Candidate',{x:0.5,y:3.35,w:6,h:0.3,fontSize:10,color:C.muted,fontFace:'Calibri'});
    s.addShape(prs.ShapeType.rect,{x:7.2,y:5.2,w:2.5,h:1.8,fill:{color:A,transparency:90},line:{color:A,transparency:70},rounding:true});
    s.addText(coverStat,{x:7.2,y:5.4,w:2.5,h:0.7,fontSize:34,bold:true,color:A,align:'center',fontFace:'Calibri'});
    s.addText(coverStatLbl,{x:7.1,y:6.05,w:2.7,h:0.6,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
  }

  // REMAINING SLIDES
  slides.forEach(sl => {
    const s = prs.addSlide();
    s.background = {color: sl.dark ? C.dark : C.lgray};
    const tc = sl.dark ? C.white : C.dark;
    s.addText(sl.tag,{x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:3});
    s.addText(sl.title,{x:0.5,y:0.7,w:9,h:0.75,fontSize:sl.titleSize||24,bold:true,color:tc,fontFace:'Calibri'});
    sl.render(s, prs, A, tc);
  });

  prs.writeFile({fileName:outFile}).then(()=>console.log('✅  '+outFile));
}

// ════════════════════════════════════════════════
// DECK A — SnabChef
// ════════════════════════════════════════════════
buildDeck(
  'SnabChef',
  'Cooking Vertical: From Trust-Building to Recurring Cook Relationship',
  '₹290Cr+',
  'estimated cooking ARR\npotential in current 5 cities',
  'FF5C00',
  [
    // SLIDE 2 — PROBLEM
    {tag:'THE PROBLEM',title:'India\'s urban household has no reliable cook.\nEvery solution fails the same way.',dark:true,titleSize:22,
     render(s,prs,A){
      const stats=[
        {v:'72%',l:'Cook Churn\nRate Month 1',s:'Cuisine mismatch + timing failure = most households replace their cook within 30 days'},
        {v:'6:30AM',l:'The Timing\nMismatch',s:'Dual-income households need breakfast by 7:30 AM. Cooks live in distant suburbs. They cannot arrive.'},
        {v:'40%',l:'Households\nWith No Cook',s:'Simply skip breakfast, eat outside, or wake earlier — because a reliable cook doesn\'t exist'}
      ];
      stats.forEach((st,i)=>{
        const x=0.4+i*3.1;
        s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.8,fill:{color:'111827'},line:{color:A,transparency:75},rounding:true,shadow:sh()});
        s.addText(st.v,{x,y:1.95,w:2.85,h:0.7,fontSize:30,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(st.l,{x,y:2.65,w:2.85,h:0.55,fontSize:10,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(st.s,{x:x+0.15,y:3.25,w:2.55,h:0.95,fontSize:8.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:4.75,w:9.2,h:0.9,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
      s.addText('Key memo insight (Bijal Dave, Apr 2026): "Cooking is a relationship product, not a logistics product. Treating it as a cleaning extension is the predictable failure mode." Snabbit\'s moat: 40K+ daily bookings, existing trust from cleaning, and a path to wage upgrades for workers. No competitor has all three.',{x:0.6,y:4.8,w:8.8,h:0.8,fontSize:9,color:'FFFFFF',fontFace:'Calibri'});
    }},
    // SLIDE 3 — INSIGHT
    {tag:'THE INSIGHT',title:'Cooking ≠ Cleaning extension.\nCooking = Relationship product.',dark:false,titleSize:22,
     render(s,prs,A){
      s.addShape(prs.ShapeType.rect,{x:0.4,y:1.65,w:4.15,h:3.8,fill:{color:'FEF2F2'},line:{color:'FCA5A5'},rounding:true});
      s.addText('❌  URBAN COMPANY / INSTA HELP',{x:0.6,y:1.85,w:3.7,h:0.35,fontSize:9,bold:true,color:'991B1B',fontFace:'Calibri'});
      ['Cook treated as fungible worker','Random assignment every session','No cuisine matching','Labor union backlash (IFAT, TGPWU)','Reputational damage, stuck between tiers'].forEach((t,i)=>{
        s.addText('• '+t,{x:0.7,y:2.3+i*0.5,w:3.5,h:0.42,fontSize:10,color:'7F1D1D',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:4.8,y:1.65,w:4.75,h:3.8,fill:{color:'F0FDF4'},line:{color:'86EFAC'},rounding:true});
      s.addText('✅  SNABCHEF APPROACH',{x:5.0,y:1.85,w:4.3,h:0.35,fontSize:9,bold:true,color:'166534',fontFace:'Calibri'});
      ['Home Food Passport — cuisine match first','Same cook first 3 sessions, no random swap','Lock-in mechanic after trust is built','Per-meal pricing (not per-hour) matches mental model','Backup cook guarantee if primary unavailable'].forEach((t,i)=>{
        s.addText('✓ '+t,{x:5.0,y:2.3+i*0.5,w:4.3,h:0.42,fontSize:10,color:'166534',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:4.5,y:1.65,w:0.5,h:3.8,fill:{color:A,transparency:85},line:{color:A,transparency:75}});
      s.addText('VS',{x:4.45,y:3.2,w:0.6,h:0.5,fontSize:11,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.6,w:9.2,h:0.7,fill:{color:A,transparency:90},line:{color:A,transparency:70},rounding:true});
      s.addText('Three cooking tiers from the supply memo: Tier 1 (daily recurring cook, nuclear family), Tier 2 (shared cook for flatmates, app-native cost split), Tier 3 (on-demand flexible). Cleaning is the cold-start distribution vehicle — existing trust dramatically reduces CAC.',{x:0.6,y:5.65,w:8.8,h:0.6,fontSize:9.5,color:'1e293b',fontFace:'Calibri'});
    }},
    // SLIDE 4 — MECHANIC
    {tag:'THE MECHANIC',title:'Food Passport → Match → Trial → Lock-In → Subscription',dark:true,titleSize:20,
     render(s,prs,A){
      const steps=[
        {n:'01',t:'Food Passport',d:'Cuisine, allergies, meal timings, household size collected. This is the match signal, not just a preference list.'},
        {n:'02',t:'Cuisine Match',d:'Cook profile shown with 98% match score, sample menu tailored to passport, certification badges, past household type.'},
        {n:'03',t:'Trial (Sessions 1–3)',d:'Same cook every session. Parent present in cooking equivalent. Progress bar shows trust journey. 3-question post-session feedback.'},
        {n:'04',t:'Lock-In Moment',d:'After Session 3: "Lock Meera in as My Cook." Priority slots, backup guarantee, direct in-app messaging channel.'},
        {n:'05',t:'Subscription',d:'Per-month retainer. Direct cook-customer channel for menu requests. Monthly occasion cooking add-on for parties.'},
      ];
      steps.forEach((st,i)=>{
        const x=0.4+i*1.88;
        s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.4,fill:{color:'111827'},line:{color:A,transparency:70},rounding:true,shadow:sh()});
        s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.7,fontSize:8.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
        if(i<4) s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:A,align:'center',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.1,w:9.2,h:0.65,fill:{color:'111827'},line:{color:A,transparency:80},rounding:true});
      s.addText('Ajay\'s parallel: PhonePe milestone gamification — "complete X steps to unlock benefit" drove 63% task completion rate and 85% escalation reduction. The lock-in mechanic after 3 sessions is exactly this pattern applied to a cook relationship.',{x:0.6,y:5.15,w:8.8,h:0.55,fontSize:9,color:'94A3B8',fontFace:'Calibri'});
    }},
    // SLIDE 5 — PRODUCT SCREENS
    {tag:'THE PRODUCT',title:'4 screens. Passport to subscription.',dark:false,titleSize:22,
     render(s,prs,A){
      const screens=[
        {n:'01',t:'Home Food Passport',d:'Cuisine, allergies, household size, meal timings collected. The matching signal. Never shared without consent.'},
        {n:'02',t:'Cook Match Profile',d:'98% cuisine match score, sample 3-day menu, food safety & police verification badges, past household type, 30-sec video intro.'},
        {n:'03',t:'My Cook Dashboard',d:'Trust journey progress bar, next session with menu, session history, direct message button. Cook feels like a relationship, not a booking.'},
        {n:'04',t:'Lock-In + Subscription',d:'3-question post-session rating → "Lock in Meera" CTA with perks: priority slots, backup cook, direct chat. ₹4,999/mo subscription.'},
      ];
      screens.forEach((sc,i)=>{
        const x=0.4+(i%2)*4.75,y=1.45+Math.floor(i/2)*2.35;
        s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:2.15,fill:{color:'FFFFFF'},line:{color:'E2E8F0'},rounding:true,shadow:sh()});
        s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:0.06,fill:{color:A},line:{color:A}});
        s.addText(sc.n,{x:x+0.15,y:y+0.15,w:0.5,h:0.35,fontSize:10,bold:true,color:A,fontFace:'Calibri'});
        s.addText(sc.t,{x:x+0.65,y:y+0.15,w:3.6,h:0.35,fontSize:11,bold:true,color:'1E293B',fontFace:'Calibri'});
        s.addText(sc.d,{x:x+0.15,y:y+0.58,w:4.1,h:1.2,fontSize:9.5,color:'475569',fontFace:'Calibri'});
      });
    }},
    // SLIDE 6 — IMPACT
    {tag:'IMPACT & ROI',title:'What SnabChef moves for Snabbit',dark:true,titleSize:22,
     render(s,prs,A){
      const metrics=[
        {v:'₹290Cr',l:'Cooking ARR Potential',s:'In current 5 cities at steady state (Bijal Dave memo, Apr 2026)'},
        {v:'2×',l:'Implied ARR Uplift',s:'Cooking alone roughly doubles current estimated cleaning ARR'},
        {v:'70%+',l:'Target Week-8 Retention',s:'Tier 1 recurring cook after lock-in — decision gate from memo'},
        {v:'₹1,500+',l:'Contribution Margin',s:'Per Tier 1 customer per month at target scale'},
        {v:'<8%',l:'Cook Attrition Target',s:'Per month — the supply-side health metric that drives everything'},
        {v:'0',l:'Incremental CAC',s:'Existing cleaning customers as primary acquisition channel — trust already exists'},
        {v:'3 tiers',l:'Day-1 Architecture',s:'Recurring Daily, Flatmate Shared, Flexible On-demand — different pricing and supply for each'},
        {v:'18 mo',l:'Competitive Window',s:'Urban Company Insta Help + Pronto already piloting. Speed matters.'},
      ];
      metrics.forEach((m,i)=>{
        const x=0.4+(i%4)*2.35,y=1.5+Math.floor(i/4)*2.0;
        s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:'111827'},line:{color:A,transparency:80},rounding:true});
        s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.65,fontSize:24,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(m.l,{x,y:y+0.78,w:2.15,h:0.4,fontSize:8.5,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(m.s,{x:x+0.1,y:y+1.18,w:1.95,h:0.45,fontSize:7.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
      });
    }},
    // SLIDE 7 — PROOF OF WORK
    {tag:'PROOF OF WORK',title:'What I built at PhonePe maps directly to this.',dark:false,titleSize:22,
     render(s,prs,A){
      s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:4.55,h:4.7,fill:{color:'080C16'},line:{color:A,transparency:80},rounding:true});
      s.addText('AT PHONEPE',{x:0.6,y:1.65,w:4.1,h:0.35,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:2});
      [
        'Milestone gamification capability: 63% task completion, 85% merchant escalation reduction — the "trust → lock-in" mechanic pattern',
        'Rebuilt ₹100 Cr/yr rewards into ML marketplace: 500+ brand partners, 11% YoY revenue growth — marketplace at scale',
        'Pincode cart interventions: 35% AOV uplift, 60% cart abandonment reduction — consumer behaviour change at hyperlocal level',
        'Propensity-to-Transact ML model: 32% marketing burn reduction — demand prediction and user scoring',
      ].forEach((t,i)=>{
        s.addText('→ '+t,{x:0.65,y:2.15+i*0.88,w:4.1,h:0.78,fontSize:9,color:'94A3B8',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:5.2,y:1.45,w:4.35,h:4.7,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
      s.addText('MAPS TO SNABCHEF',{x:5.4,y:1.65,w:3.9,h:0.35,fontSize:9,bold:true,color:'1E293B',fontFace:'Calibri',charSpacing:2});
      [
        'Milestone mechanic → 3-session trust journey + lock-in: same "complete steps to unlock a deeper commitment" architecture',
        'ML marketplace → cuisine match scoring: same ranking and personalisation logic applied to cook-household fit',
        'Cart abandonment reduction → cook churn reduction: same North Star (completing the committed action), same real-time intervention design',
        'Propensity model → demand prediction for cooking slots: predict which households need a cook before they search',
      ].forEach((t,i)=>{
        s.addText('→ '+t,{x:5.4,y:2.15+i*0.88,w:3.9,h:0.78,fontSize:9,color:'475569',fontFace:'Calibri'});
      });
    }},
    // SLIDE 8 — ROLLOUT
    {tag:'ROLLOUT PLAN',title:'Phase 0 → PMF in 90 days (per memo framework)',dark:true,titleSize:22,
     render(s,prs,A){
      const phases=[
        {p:'Phase 0',d:'Weeks 1–4',t:'Diagnostic',b:['30 customer conversations across 6 archetypes','10 cook conversations (offline + Snabbit cleaners)','Fake-door test: bundle offer to 500 cleaning customers','Gate: >8% click-through + 3/6 archetypes at ₹5K WTP']},
        {p:'Phase 1',d:'Weeks 5–12',t:'Launch via Cleaning Bundle',b:['Bundle: 2 cleaning + 2 cooking trials as entry offer','1 micro-market only (HSR/Koramangala)','Same cook guaranteed for all trial sessions','Gate: 60% trial→subscription conversion at week 8']},
        {p:'Phase 2',d:'Weeks 13–20',t:'Standalone Subscriptions',b:['Launch Tier 1 + Tier 2 standalone pricing','Build cook-customer pairing + backup cook system','Recruit 50 experienced offline cooks','Gate: Tier 1 retention >70%, CM >₹1,500/month']},
      ];
      phases.forEach((ph,i)=>{
        const x=0.4+i*3.1;
        s.addShape(prs.ShapeType.rect,{x,y:1.5,w:2.9,h:3.6,fill:{color:'111827'},line:{color:A,transparency:70},rounding:true,shadow:sh()});
        s.addText(ph.p,{x,y:1.65,w:2.9,h:0.35,fontSize:9,bold:true,color:A,align:'center',fontFace:'Calibri',charSpacing:2});
        s.addText(ph.d,{x,y:1.95,w:2.9,h:0.3,fontSize:8,color:'94A3B8',align:'center',fontFace:'Calibri'});
        s.addText(ph.t,{x,y:2.25,w:2.9,h:0.4,fontSize:10.5,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.15,y:2.75+j*0.62,w:2.6,h:0.55,fontSize:8.5,color:'94A3B8',fontFace:'Calibri'});});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.25,w:9.2,h:0.9,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
      s.addText('What I need from Snabbit: 1 micro-market as sandbox · 50 cooking-willing professional conversations · FSSAI registration research · Fake-door test infrastructure · Access to cleaning customer base for bundle offer',{x:0.6,y:5.3,w:8.8,h:0.8,fontSize:9,color:'FFFFFF',fontFace:'Calibri'});
    }},
  ],
  'assets/snabbit-snabchef-deck.pptx'
);

// ════════════════════════════════════════════════
// DECK B — SnabNanny
// ════════════════════════════════════════════════
buildDeck(
  'SnabNanny',
  'Childcare: Trust Infrastructure → Recurring Nanny Relationship',
  '0',
  'platform-grade childcare\nsolutions in India today',
  '6366F1',
  [
    {tag:'THE PROBLEM',title:'The nanny problem in urban India is not a product problem.\nIt is a trust infrastructure problem.',dark:true,titleSize:20,
     render(s,prs,A){
      const stats=[
        {v:'₹0',l:'Platform-Grade\nChildcare in India',s:'No trusted, on-demand, verified nanny solution exists for temporary or emergency childcare for urban families'},
        {v:'60%+',l:'Crèche Waitlist\nRate in Bengaluru',s:'Premium crèches in HSR, Koramangala, Indiranagar have 6–18 month waitlists — nuclear families have no fallback'},
        {v:'0',l:'Informal Referral\nNetworks Left',s:'Nuclear apartment living has destroyed the "neighbour\'s nanny" trust network that urban India relied on for decades'},
      ];
      stats.forEach((st,i)=>{
        const x=0.4+i*3.1;
        s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.8,fill:{color:'111827'},line:{color:A,transparency:75},rounding:true,shadow:sh()});
        s.addText(st.v,{x,y:1.95,w:2.85,h:0.7,fontSize:30,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(st.l,{x,y:2.65,w:2.85,h:0.55,fontSize:10,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(st.s,{x:x+0.15,y:3.25,w:2.55,h:0.95,fontSize:8.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:4.75,w:9.2,h:0.9,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
      s.addText('The three use cases with no solution: (1) Sick day cover — regular nanny didn\'t show, WFH meeting in 1 hour. (2) Event/party — need a responsible person for 3–4 hours while you host. (3) Trial-before-full-time — want to assess a nanny before committing to a full-time arrangement. All three require trust. None have a platform.',{x:0.6,y:4.8,w:8.8,h:0.8,fontSize:9,color:'FFFFFF',fontFace:'Calibri'});
    }},
    {tag:'THE INSIGHT',title:'Trust is not a feature for childcare.\nTrust IS the product.',dark:false,titleSize:22,
     render(s,prs,A){
      s.addShape(prs.ShapeType.rect,{x:0.4,y:1.65,w:4.15,h:3.8,fill:{color:'FEF2F2'},line:{color:'FCA5A5'},rounding:true});
      s.addText('❌  CURRENT STATE (URBAN INDIA)',{x:0.6,y:1.85,w:3.7,h:0.35,fontSize:9,bold:true,color:'991B1B',fontFace:'Calibri'});
      ['WhatsApp groups — unverified strangers','Maid agencies — no accountability, cash only','Nextdoor/Facebook — no screening, no recourse','No platform trust layer exists','Result: parents don\'t leave. They cancel their plans.'].forEach((t,i)=>{
        s.addText('• '+t,{x:0.7,y:2.3+i*0.5,w:3.5,h:0.42,fontSize:10,color:'7F1D1D',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:4.8,y:1.65,w:4.75,h:3.8,fill:{color:'F0FDF4'},line:{color:'86EFAC'},rounding:true});
      s.addText('✅  SNABNANNY TRUST STACK',{x:5.0,y:1.85,w:4.3,h:0.35,fontSize:9,bold:true,color:'166534',fontFace:'Calibri'});
      ['Police verification with badge # visible','First Aid + child development certification','30-sec video introduction (nanny, not algorithm)','3 verified family references in-app','Nanny Score (composite) — not just star rating','Progressive trust: trial → regular → My Nanny'].forEach((t,i)=>{
        s.addText('✓ '+t,{x:5.0,y:2.3+i*0.45,w:4.3,h:0.38,fontSize:10,color:'166534',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:4.5,y:1.65,w:0.5,h:3.8,fill:{color:A,transparency:85},line:{color:A,transparency:75}});
      s.addText('VS',{x:4.45,y:3.2,w:0.6,h:0.5,fontSize:11,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.6,w:9.2,h:0.7,fill:{color:A,transparency:90},line:{color:A,transparency:70},rounding:true});
      s.addText('The gradual trust model: Session 1 (parent present, observer mode) → Session 2 (parent at home, different room) → Session 3+ (parent can leave, live photo check-ins every 45 min) → 5 sessions: My Nanny locked in. Switching cost becomes emotional, not rational.',{x:0.6,y:5.65,w:8.8,h:0.6,fontSize:9.5,color:'1e293b',fontFace:'Calibri'});
    }},
    {tag:'THE MECHANIC',title:'Verified Profile → Book → Live Check-ins → My Nanny',dark:true,titleSize:22,
     render(s,prs,A){
      const steps=[
        {n:'01',t:'Trust Profile',d:'Police verification #, First Aid cert, Child Dev training, 3 family references, 30-sec video intro, Nanny Score visible upfront.'},
        {n:'02',t:'Session Type',d:'Trial (parent present), Regular (can leave, live check-ins), Emergency Cover (substitute within 30 min — SLA guaranteed).'},
        {n:'03',t:'Live Session',d:'Photo check-in every 45 min from nanny. Parent sees activity, meals, mood. Emergency button always visible. Session timing monitored.'},
        {n:'04',t:'Post-Session Debrief',d:'What child did, ate, mood + energy report. Nanny logs, parent reads. Daycare-style report builds familiarity and trust.'},
        {n:'05',t:'My Nanny Lock-in',d:'After 5 sessions: priority booking, rate freeze, emergency cover guarantee. Child knows her. Switching cost is emotional.'},
      ];
      steps.forEach((st,i)=>{
        const x=0.4+i*1.88;
        s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.4,fill:{color:'111827'},line:{color:A,transparency:70},rounding:true,shadow:sh()});
        s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.7,fontSize:8.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
        if(i<4) s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:A,align:'center',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.1,w:9.2,h:0.65,fill:{color:'111827'},line:{color:A,transparency:80},rounding:true});
      s.addText('Emergency Cover SLA: if My Nanny is unavailable, a verified substitute is dispatched within 30 minutes — guaranteed. Higher bar than any other category. This SLA is the competitive moat for the product.',{x:0.6,y:5.15,w:8.8,h:0.55,fontSize:9,color:'94A3B8',fontFace:'Calibri'});
    }},
    {tag:'THE PRODUCT',title:'4 screens. Trust profile to My Nanny.',dark:false,titleSize:22,
     render(s,prs,A){
      const screens=[
        {n:'01',t:'Nanny Trust Profile',d:'Police verification badge # shown. First Aid, Child Dev certs. 3 verified family references. 30-sec video intro. Nanny Score (composite, not just stars).'},
        {n:'02',t:'Booking + Session Type',d:'3 session types: Trial (parent present), Regular (live check-ins), Emergency Cover (30-min SLA). Child age + special needs captured for matching.'},
        {n:'03',t:'Live Session Check-in',d:'Nanny sends photo update every 45 min with activity, meals, mood notes. Emergency button always visible. Session timing monitored by Snabbit.'},
        {n:'04',t:'Debrief + My Nanny',d:'What child did, ate, mood report (daycare-style). Post-Session 5: lock-in CTA with priority booking, rate freeze, emergency cover guarantee.'},
      ];
      screens.forEach((sc,i)=>{
        const x=0.4+(i%2)*4.75,y=1.45+Math.floor(i/2)*2.35;
        s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:2.15,fill:{color:'FFFFFF'},line:{color:'E2E8F0'},rounding:true,shadow:sh()});
        s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:0.06,fill:{color:A},line:{color:A}});
        s.addText(sc.n,{x:x+0.15,y:y+0.15,w:0.5,h:0.35,fontSize:10,bold:true,color:A,fontFace:'Calibri'});
        s.addText(sc.t,{x:x+0.65,y:y+0.15,w:3.6,h:0.35,fontSize:11,bold:true,color:'1E293B',fontFace:'Calibri'});
        s.addText(sc.d,{x:x+0.15,y:y+0.58,w:4.1,h:1.2,fontSize:9.5,color:'475569',fontFace:'Calibri'});
      });
    }},
    {tag:'IMPACT & ROI',title:'What SnabNanny moves for Snabbit',dark:true,titleSize:22,
     render(s,prs,A){
      const metrics=[
        {v:'0 → 1',l:'New Category Entirely',s:'No platform-grade childcare solution exists in India — first-mover advantage is structural'},
        {v:'₹499',l:'Trial Session Price',s:'vs. ₹0 for a cancelled plan or ₹15K/mo for a full-time arrangement — perfect price-value entry'},
        {v:'4×',l:'LTV vs Episodic',s:'A My Nanny subscriber has 4× LTV vs. one-time bookers (recurring + emergency coverage usage)'},
        {v:'30 min',l:'Emergency SLA',s:'Guaranteed substitute dispatch — the trust anchor that makes the premium subscription worth it'},
        {v:'5 sessions',l:'Lock-in Threshold',s:'After 5 sessions, child knows the nanny — emotional switching cost converts to subscription naturally'},
        {v:'Society',l:'B2B Distribution',s:'RWA/gated community partnerships: nanny availability as an amenity for 500+ resident households'},
        {v:'0',l:'Incremental CAC',s:'Existing cleaning customers in dual-income households are the primary acquisition channel'},
        {v:'3× Repeat',l:'Emergency → Regular',s:'Families who use emergency cover 3× convert to My Nanny subscription within 60 days'},
      ];
      metrics.forEach((m,i)=>{
        const x=0.4+(i%4)*2.35,y=1.5+Math.floor(i/4)*2.0;
        s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:'111827'},line:{color:A,transparency:80},rounding:true});
        s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.65,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(m.l,{x,y:y+0.78,w:2.15,h:0.4,fontSize:8.5,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(m.s,{x:x+0.1,y:y+1.18,w:1.95,h:0.45,fontSize:7.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
      });
    }},
    {tag:'PROOF OF WORK',title:'What I built at PhonePe maps directly to this.',dark:false,titleSize:22,
     render(s,prs,A){
      s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:4.55,h:4.7,fill:{color:'080C16'},line:{color:A,transparency:80},rounding:true});
      s.addText('AT PHONEPE',{x:0.6,y:1.65,w:4.1,h:0.35,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:2});
      [
        'Built B2B self-serve platform from zero: 5,000+ merchants onboarded with zero ops touch — trust-building at scale via structured verification flow',
        'Milestone gamification: 63% task completion, 85% escalation reduction — the "complete N sessions to unlock deeper commitment" pattern',
        'Pincode real-time interventions: consumer behaviour change via contextual nudges at exactly the right moment in the user journey',
        'Offer checkout redesign (350M+ MAU): structured UX to build confidence before a financial commitment — same trust-before-conversion principle',
      ].forEach((t,i)=>{
        s.addText('→ '+t,{x:0.65,y:2.15+i*0.88,w:4.1,h:0.78,fontSize:9,color:'94A3B8',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:5.2,y:1.45,w:4.35,h:4.7,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
      s.addText('MAPS TO SNABNANNY',{x:5.4,y:1.65,w:3.9,h:0.35,fontSize:9,bold:true,color:'1E293B',fontFace:'Calibri',charSpacing:2});
      [
        'Zero-ops-touch merchant onboarding → Nanny verification pipeline: same problem (build trust at scale without human review for every profile)',
        'Milestone mechanic → 5-session trust journey: same "progressive commitment unlock" architecture applied to a care relationship',
        'Contextual nudges → live check-in design: surface the right trust signal at exactly the moment a parent needs reassurance most',
        'Confidence-before-conversion UX → trial session design: "session 1, you stay home" is the same principle as showing full offer terms before payment',
      ].forEach((t,i)=>{
        s.addText('→ '+t,{x:5.4,y:2.15+i*0.88,w:3.9,h:0.78,fontSize:9,color:'475569',fontFace:'Calibri'});
      });
    }},
    {tag:'ROLLOUT PLAN',title:'0 → Category PMF in 90 days',dark:true,titleSize:22,
     render(s,prs,A){
      const phases=[
        {p:'Phase 1',d:'Days 1–30',t:'Trust Stack + Trial',b:['Launch police verification + First Aid cert pipeline','Build nanny video intro feature + Nanny Score','Beta: 50 nannies, 200 trial bookings in 2 societies','Gate: NPS > 60, trial→repeat rate > 40%']},
        {p:'Phase 2',d:'Days 31–60',t:'Regular + Emergency SLA',b:['Launch regular session (live check-in flow)','Build emergency 30-min dispatch capability','Nanny backup matching system for My Nanny tier','Gate: emergency SLA met >90%, regular session NPS > 65'],},
        {p:'Phase 3',d:'Days 61–90',t:'My Nanny Subscription',b:['Launch My Nanny lock-in + subscription pricing','Society partnership pilot (2 gated communities)','Track: subscription conversion at session 5','Gate: My Nanny LTV > 3× trial-only users at 90 days']},
      ];
      phases.forEach((ph,i)=>{
        const x=0.4+i*3.1;
        s.addShape(prs.ShapeType.rect,{x,y:1.5,w:2.9,h:3.6,fill:{color:'111827'},line:{color:A,transparency:70},rounding:true,shadow:sh()});
        s.addText(ph.p,{x,y:1.65,w:2.9,h:0.35,fontSize:9,bold:true,color:A,align:'center',fontFace:'Calibri',charSpacing:2});
        s.addText(ph.d,{x,y:1.95,w:2.9,h:0.3,fontSize:8,color:'94A3B8',align:'center',fontFace:'Calibri'});
        s.addText(ph.t,{x,y:2.25,w:2.9,h:0.4,fontSize:10.5,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.15,y:2.75+j*0.62,w:2.6,h:0.55,fontSize:8.5,color:'94A3B8',fontFace:'Calibri'});});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.25,w:9.2,h:0.9,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
      s.addText('What I need from Snabbit: Background verification API partner (AuthBridge/IDfy) · First Aid training vendor · 50 nanny recruits for beta · 2 society partnerships for launch · ₹20L subsidy budget for trial session pricing',{x:0.6,y:5.3,w:8.8,h:0.8,fontSize:9,color:'FFFFFF',fontFace:'Calibri'});
    }},
  ],
  'assets/snabbit-snabnanny-deck.pptx'
);

// ════════════════════════════════════════════════
// DECK C — SnabHome
// ════════════════════════════════════════════════
buildDeck(
  'SnabHome',
  'Home Care Subscription — Converting Transactions into Relationships',
  '4×',
  'higher LTV for\nsubscription vs. episodic users',
  'FF5C00',
  [
    {tag:'THE PROBLEM',title:'Every Snabbit booking is a transaction.\nNo booking is a relationship.',dark:true,titleSize:22,
     render(s,prs,A){
      const stats=[
        {v:'1.8×',l:'Average Bookings\nBefore Churn',s:'Most urban home service users book 1–2 times then return to informal arrangements — no habit loop forms'},
        {v:'₹800+',l:'CAC Per New User',s:'Performance ads for each episodic booking means Snabbit re-acquires the same user multiple times at full CAC'},
        {v:'0',l:'Predictable Demand\nfor Professionals',s:'No recurring schedule = professional income is unpredictable = professional churn = supply quality failure'},
      ];
      stats.forEach((st,i)=>{
        const x=0.4+i*3.1;
        s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.8,fill:{color:'111827'},line:{color:A,transparency:75},rounding:true,shadow:sh()});
        s.addText(st.v,{x,y:1.95,w:2.85,h:0.7,fontSize:30,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(st.l,{x,y:2.65,w:2.85,h:0.55,fontSize:10,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(st.s,{x:x+0.15,y:3.25,w:2.55,h:0.95,fontSize:8.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:4.75,w:9.2,h:0.9,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
      s.addText('The compound problem: episodic users have low LTV, high re-acquisition cost, and unpredictable demand for professionals. SnabHome solves all three simultaneously — a subscription converts a transaction into a relationship on both sides of the marketplace.',{x:0.6,y:4.8,w:8.8,h:0.8,fontSize:9,color:'FFFFFF',fontFace:'Calibri'});
    }},
    {tag:'THE INSIGHT',title:'A monthly plan converts CAC into a compounding asset.\nSame professional = trust that prevents churn.',dark:false,titleSize:20,
     render(s,prs,A){
      const cols=[
        {h:'❌  EPISODIC BOOKING',items:['Book → forget → re-book later','Re-acquire at full CAC every time','Random professional each session','No trust accumulation','Professional has no income stability','Snabbit is "just an app" to the user'],c1:'FEF2F2',c2:'FCA5A5',tc:'7F1D1D',hc:'991B1B'},
        {h:'✅  SNABHOME SUBSCRIPTION',items:['Commit once → auto-booked forever','CAC paid once, 12+ months of revenue','Same professional every session','Trust deepens with every visit','Professional has guaranteed income','Snabbit becomes a home utility'],c1:'F0FDF4',c2:'86EFAC',tc:'166534',hc:'166534'},
      ];
      cols.forEach((col,i)=>{
        const x=i===0?0.4:4.8;const w=i===0?4.15:4.75;
        s.addShape(prs.ShapeType.rect,{x,y:1.65,w,h:3.8,fill:{color:col.c1},line:{color:col.c2},rounding:true});
        s.addText((i===0?'❌  ':'✅  ')+col.h.replace('❌  ','').replace('✅  ',''),{x:x+0.2,y:1.85,w:w-0.4,h:0.35,fontSize:9,bold:true,color:col.hc,fontFace:'Calibri'});
        col.items.forEach((t,j)=>{s.addText((i===0?'• ':'✓ ')+t,{x:x+0.2,y:2.3+j*0.45,w:w-0.35,h:0.38,fontSize:10,color:col.tc,fontFace:'Calibri'});});
      });
      s.addShape(prs.ShapeType.rect,{x:4.5,y:1.65,w:0.5,h:3.8,fill:{color:A,transparency:85},line:{color:A,transparency:75}});
      s.addText('VS',{x:4.45,y:3.2,w:0.6,h:0.5,fontSize:11,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.6,w:9.2,h:0.7,fill:{color:A,transparency:90},line:{color:A,transparency:70},rounding:true});
      s.addText('The supply-side benefit is equal to the consumer benefit: a professional with a recurring SnabHome subscriber has predictable weekly income → professional churn drops → supply quality improves → consumer experience improves → subscription NPS rises. It is a self-reinforcing loop.',{x:0.6,y:5.65,w:8.8,h:0.6,fontSize:9.5,color:'1e293b',fontFace:'Calibri'});
    }},
    {tag:'THE MECHANIC',title:'Plan Selection → Schedule Builder → My Home → My Professional',dark:true,titleSize:22,
     render(s,prs,A){
      const steps=[
        {n:'01',t:'Plan Selection',d:'3 tiers: Basic (4 cleans), Pro (8 cleans + deep clean), Premium (unlimited + maintenance). Clear feature diff, price anchoring.'},
        {n:'02',t:'Schedule Builder',d:'Pick recurring days + time slot + services. Summary confirms plan. Snabbit assigns same professional for all booked slots.'},
        {n:'03',t:'My Home Dashboard',d:'"This week" calendar, upcoming sessions, session history with ratings. My Professional profile with lock status always visible.'},
        {n:'04',t:'My Professional',d:'Dedicated professional profile: sessions together count, your ratings, their full verification history, past review excerpts.'},
        {n:'05',t:'Retention Loop',d:'Weekly sessions build professional-household relationship. Reschedule or skip — never lose your professional. Monthly plan renewal is effortless.'},
      ];
      steps.forEach((st,i)=>{
        const x=0.4+i*1.88;
        s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.4,fill:{color:'111827'},line:{color:A,transparency:70},rounding:true,shadow:sh()});
        s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.7,fontSize:8.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
        if(i<4) s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:A,align:'center',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.1,w:9.2,h:0.65,fill:{color:'111827'},line:{color:A,transparency:80},rounding:true});
      s.addText('Add-on upsell within the subscription: SnabChef cooking sessions, SnabNanny childcare sessions, quarterly home maintenance audits — the subscription becomes the entry point to the full Snabbit household platform.',{x:0.6,y:5.15,w:8.8,h:0.55,fontSize:9,color:'94A3B8',fontFace:'Calibri'});
    }},
    {tag:'THE PRODUCT',title:'4 screens. Plan to relationship.',dark:false,titleSize:22,
     render(s,prs,A){
      const screens=[
        {n:'01',t:'Plan Selection',d:'3 tiers with clear feature diff. Pro highlighted as most popular. Price anchoring via Premium shown first. Framing: "subscribe once, home is covered."'},
        {n:'02',t:'Schedule Builder',d:'Day picker (M/T/W/T/F/S/S), time slot selector, service checklist. Live summary confirms plan. One CTA to confirm.'},
        {n:'03',t:'My Home Dashboard',d:'"This week" with upcoming sessions. My Professional card with lock status. Session history with ratings. Zero friction to see full history.'},
        {n:'04',t:'My Professional Profile',d:'Sessions together count, your avg rating, their full verification. Past customer reviews (anonymised). Reschedule button always available.'},
      ];
      screens.forEach((sc,i)=>{
        const x=0.4+(i%2)*4.75,y=1.45+Math.floor(i/2)*2.35;
        s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:2.15,fill:{color:'FFFFFF'},line:{color:'E2E8F0'},rounding:true,shadow:sh()});
        s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:0.06,fill:{color:A},line:{color:A}});
        s.addText(sc.n,{x:x+0.15,y:y+0.15,w:0.5,h:0.35,fontSize:10,bold:true,color:A,fontFace:'Calibri'});
        s.addText(sc.t,{x:x+0.65,y:y+0.15,w:3.6,h:0.35,fontSize:11,bold:true,color:'1E293B',fontFace:'Calibri'});
        s.addText(sc.d,{x:x+0.15,y:y+0.58,w:4.1,h:1.2,fontSize:9.5,color:'475569',fontFace:'Calibri'});
      });
    }},
    {tag:'IMPACT & ROI',title:'What SnabHome moves for Snabbit',dark:true,titleSize:22,
     render(s,prs,A){
      const metrics=[
        {v:'4×',l:'LTV vs Episodic',s:'Subscription users have 4× lifetime value vs. one-time bookers at equivalent session count'},
        {v:'₹0',l:'Re-acquisition Cost',s:'Monthly renewal requires zero marketing spend — the subscription IS the retention mechanism'},
        {v:'8 sessions',l:'Pro Plan Monthly',s:'8 guaranteed sessions per subscriber creates predictable weekly demand for professionals'},
        {v:'60%',l:'Professional Churn Reduction',s:'Recurring income stability reduces professional churn — supply quality directly improves'},
        {v:'3 tiers',l:'Revenue Segmentation',s:'₹1,299 Basic / ₹2,499 Pro / ₹4,299 Premium — captures full willingness-to-pay curve'},
        {v:'12+ mo',l:'Average Subscriber Tenure',s:'Once a household has a trusted professional, switching cost is emotional — churn near zero'},
        {v:'Gateway',l:'To Full Platform',s:'SnabHome subscriber is 3× more likely to book SnabChef or SnabNanny — subscription as top of funnel'},
        {v:'NPS',l:'Structural Driver',s:'Same professional every session = trust compounds = NPS rises = organic referral flywheel'},
      ];
      metrics.forEach((m,i)=>{
        const x=0.4+(i%4)*2.35,y=1.5+Math.floor(i/4)*2.0;
        s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:'111827'},line:{color:A,transparency:80},rounding:true});
        s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.65,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
        s.addText(m.l,{x,y:y+0.78,w:2.15,h:0.4,fontSize:8.5,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        s.addText(m.s,{x:x+0.1,y:y+1.18,w:1.95,h:0.45,fontSize:7.5,color:'94A3B8',align:'center',fontFace:'Calibri'});
      });
    }},
    {tag:'PROOF OF WORK',title:'What I built at PhonePe maps directly to this.',dark:false,titleSize:22,
     render(s,prs,A){
      s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:4.55,h:4.7,fill:{color:'080C16'},line:{color:A,transparency:80},rounding:true});
      s.addText('AT PHONEPE',{x:0.6,y:1.65,w:4.1,h:0.35,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:2});
      [
        'Rebuilt ₹100 Cr/yr rewards portfolio into ML marketplace: 500+ brand partners, 11% YoY revenue growth, 26% engagement increase — subscription-style recurring engagement mechanics',
        'Milestone gamification: 63% task completion, 85% escalation reduction — the "progressive commitment unlock" architecture that drives subscription conversion',
        'Pincode cart interventions: 35% AOV uplift, 60% cart abandonment reduction — changing episodic behaviour into recurring commitment',
        'Self-serve merchant platform (5,000+ merchants, zero ops touch): same "onboard once, operate continuously" design philosophy as SnabHome',
      ].forEach((t,i)=>{
        s.addText('→ '+t,{x:0.65,y:2.15+i*0.88,w:4.1,h:0.78,fontSize:9,color:'94A3B8',fontFace:'Calibri'});
      });
      s.addShape(prs.ShapeType.rect,{x:5.2,y:1.45,w:4.35,h:4.7,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
      s.addText('MAPS TO SNABHOME',{x:5.4,y:1.65,w:3.9,h:0.35,fontSize:9,bold:true,color:'1E293B',fontFace:'Calibri',charSpacing:2});
      [
        'Rewards marketplace recurring engagement → SnabHome recurring schedule: same North Star (DAU/WAU of returning users, not new user count)',
        'Milestone mechanic → subscription commitment: "complete N actions to unlock better tier" drives plan upgrades from Basic to Pro',
        'Cart abandonment reduction → subscription churn reduction: same framework (identify drop-off point, design intervention for that exact moment)',
        'Zero-ops merchant platform → zero-friction subscription management: reschedule, pause, change professional — all self-serve, no support ticket',
      ].forEach((t,i)=>{
        s.addText('→ '+t,{x:5.4,y:2.15+i*0.88,w:3.9,h:0.78,fontSize:9,color:'475569',fontFace:'Calibri'});
      });
    }},
    {tag:'ROLLOUT PLAN',title:'0 → Subscription PMF in 90 days',dark:true,titleSize:22,
     render(s,prs,A){
      const phases=[
        {p:'Phase 1',d:'Days 1–30',t:'Plan + Schedule MVP',b:['Launch 3-tier plan selection for cleaning only','Professional assignment for recurring slots','My Home dashboard v1','Gate: plan selection rate > 15% of active users']},
        {p:'Phase 2',d:'Days 31–60',t:'My Professional + Retention',b:['Launch My Professional profile + lock-in feature','A/B test: plan upgrade nudges at session 4','Measure: subscription renewal rate at day 30','Gate: 60-day subscriber retention > 70%']},
        {p:'Phase 3',d:'Days 61–90',t:'Expansion + Platform',b:['Bundle: SnabHome subscribers get priority SnabChef access','Society partnership: SnabHome as resident amenity','Track: cross-vertical booking rate for subscribers','Gate: subscriber LTV > 4× episodic users at 90 days']},
      ];
      phases.forEach((ph,i)=>{
        const x=0.4+i*3.1;
        s.addShape(prs.ShapeType.rect,{x,y:1.5,w:2.9,h:3.6,fill:{color:'111827'},line:{color:A,transparency:70},rounding:true,shadow:sh()});
        s.addText(ph.p,{x,y:1.65,w:2.9,h:0.35,fontSize:9,bold:true,color:A,align:'center',fontFace:'Calibri',charSpacing:2});
        s.addText(ph.d,{x,y:1.95,w:2.9,h:0.3,fontSize:8,color:'94A3B8',align:'center',fontFace:'Calibri'});
        s.addText(ph.t,{x,y:2.25,w:2.9,h:0.4,fontSize:10.5,bold:true,color:'FFFFFF',align:'center',fontFace:'Calibri'});
        ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.15,y:2.75+j*0.62,w:2.6,h:0.55,fontSize:8.5,color:'94A3B8',fontFace:'Calibri'});});
      });
      s.addShape(prs.ShapeType.rect,{x:0.4,y:5.25,w:9.2,h:0.9,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
      s.addText('What I need from Snabbit: Professional assignment algorithm for recurring slots · Plan management backend · Subscription billing infrastructure · 2 society partnerships for SnabHome amenity pilot · ₹15L subsidy for first-month free offer',{x:0.6,y:5.3,w:8.8,h:0.8,fontSize:9,color:'FFFFFF',fontFace:'Calibri'});
    }},
  ],
  'assets/snabbit-snabhome-deck.pptx'
);
