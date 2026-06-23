const PptxGenJS = require('pptxgenjs');

const C = {dark:'060910',dark2:'0D1526',teal:'00D4AA',blue:'00A8FF',purple:'8B5CF6',amber:'F59E0B',green:'22C55E',red:'EF4444',white:'FFFFFF',muted:'94A3B8',lgray:'F1F5F9',slate:'1E293B'};
const sh = () => ({type:'outer',color:'000000',blur:5,offset:2,angle:45,opacity:0.15});

function mkDeck(title, subtitle, coverStat, coverStatLbl, accent, slides, outFile) {
  const prs = new PptxGenJS(); prs.layout='LAYOUT_16x9';
  const A = accent;

  // COVER
  {
    const s = prs.addSlide(); s.background={color:C.dark};
    for(let i=0;i<14;i++) s.addShape(prs.ShapeType.rect,{x:i*0.72-0.2,y:8.5-i*0.6,w:0.04,h:8,fill:{color:A,transparency:94},line:{color:A,transparency:94}});
    s.addText('NorthLadder',{x:0.5,y:0.5,w:5,h:0.35,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:3});
    s.addText('PRODUCT CASE STUDY — FOUNDING PM CANDIDATE',{x:0.5,y:0.82,w:6,h:0.28,fontSize:7,color:C.muted,fontFace:'Calibri',charSpacing:2});
    s.addText(title,{x:0.5,y:1.3,w:8.2,h:1.1,fontSize:42,bold:true,color:C.white,fontFace:'Calibri'});
    s.addText(subtitle,{x:0.5,y:2.5,w:7.5,h:0.5,fontSize:16,color:C.muted,fontFace:'Calibri'});
    s.addShape(prs.ShapeType.rect,{x:0.5,y:3.1,w:1.8,h:0.05,fill:{color:A},line:{color:A}});
    s.addText('Presented by Ajay Avaghade · Founding PM Candidate · Dubai',{x:0.5,y:3.3,w:6,h:0.3,fontSize:10,color:C.muted,fontFace:'Calibri'});
    s.addShape(prs.ShapeType.rect,{x:7.2,y:5.0,w:2.5,h:1.8,fill:{color:A,transparency:88},line:{color:A,transparency:65},rounding:true});
    s.addText(coverStat,{x:7.2,y:5.2,w:2.5,h:0.7,fontSize:32,bold:true,color:A,align:'center',fontFace:'Calibri'});
    s.addText(coverStatLbl,{x:7.1,y:5.85,w:2.7,h:0.65,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
  }

  slides.forEach(sl => {
    const s = prs.addSlide();
    s.background={color: sl.dark ? C.dark : C.lgray};
    const tc = sl.dark ? C.white : C.dark;
    s.addText(sl.tag,{x:0.5,y:0.35,w:9,h:0.28,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:3});
    s.addText(sl.title,{x:0.5,y:0.65,w:9,h:0.7,fontSize:sl.titleSize||22,bold:true,color:tc,fontFace:'Calibri'});
    sl.render(s, prs, A, tc);
  });

  prs.writeFile({fileName:outFile}).then(()=>console.log('✅ '+outFile));
}

// ═══════════════════════════════════════════════════
// DECK 1 — PriceIQ
// ═══════════════════════════════════════════════════
mkDeck('PriceIQ','Dynamic Trade-in Pricing Engine — The Commercial Engine of NorthLadder\'s Circular Economy','20+','countries requiring\nreal-time price intelligence',C.teal,[
  {tag:'THE PROBLEM',title:'Static pricing tables bleed value — for everyone.',dark:true,titleSize:22,
   render(s,prs,A){
    const stats=[
      {v:'–18%',l:'Value Leakage\nper Device',s:'The difference between a static weekly table and a real-time market price — NorthLadder earns it but consumers and partners don\'t see it'},
      {v:'7 days',l:'Price Table\nRefresh Cycle',s:'Devices depreciate 3–5% per week after a new model launch. A 7-day refresh cycle is 35% of that depreciation left on the table'},
      {v:'20+',l:'Countries, One\nPricing Engine',s:'AED, SAR, GBP, INR, EUR — different secondary markets, FX spreads, demand curves. One static table cannot serve all'},
    ];
    stats.forEach((st,i)=>{
      const x=0.4+i*3.1;
      s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.9,fill:{color:C.dark2},line:{color:A,transparency:75},rounding:true,shadow:sh()});
      s.addText(st.v,{x,y:1.95,w:2.85,h:0.7,fontSize:30,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(st.l,{x,y:2.65,w:2.85,h:0.55,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(st.s,{x:x+0.15,y:3.25,w:2.55,h:1.0,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:4.8,w:9.2,h:0.85,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
    s.addText('NorthLadder\'s competitive moat is offering the best trade-in price in the market — but "best" is a moving target. eBay sold prices shift daily. iPhone 16 launches depress iPhone 15 values overnight. FX moves change cross-border arbitrage opportunity. PriceIQ is the infrastructure that lets NorthLadder win on price without sacrificing margin.',{x:0.6,y:4.85,w:8.8,h:0.75,fontSize:9,color:C.white,fontFace:'Calibri'});
  }},
  {tag:'THE INSIGHT',title:'Optimal price = maximum consumer acceptance rate × minimum margin erosion. It\'s a continuous optimisation problem.',dark:false,titleSize:18,
   render(s,prs,A){
    s.addShape(prs.ShapeType.rect,{x:0.4,y:1.55,w:4.2,h:4.15,fill:{color:'FFF7ED'},line:{color:'FED7AA'},rounding:true});
    s.addText('❌  CURRENT STATE',{x:0.6,y:1.75,w:3.8,h:0.35,fontSize:9,bold:true,color:'92400E',fontFace:'Calibri'});
    ['Weekly Excel-updated pricing tables','Single price per model/condition (no market differentiation)','No real-time secondary market signal','FX adjustments done manually once a month','New model launches managed ad hoc — prices updated days later','Telco partner calls to dispute offer price post-inspection'].forEach((t,i)=>{
      s.addText('• '+t,{x:0.7,y:2.2+i*0.48,w:3.7,h:0.42,fontSize:10,color:'78350F',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:4.9,y:1.55,w:4.65,h:4.15,fill:{color:'F0FDF4'},line:{color:'86EFAC'},rounding:true});
    s.addText('✅  PRICEIQ STATE',{x:5.1,y:1.75,w:4.2,h:0.35,fontSize:9,bold:true,color:'166534',fontFace:'Calibri'});
    ['Live eBay/Swappa sold prices via API — hourly refresh','Per-market, per-condition dynamic offer generation','FX real-time feeds (every 15 min) built into every quote','New model launch: automatic 15-min depreciation cascade','Partner margin lever: adjust consumer offer vs. NL margin','Acceptance-rate feedback loop: A/B-test price points live'].forEach((t,i)=>{
      s.addText('✓ '+t,{x:5.1,y:2.2+i*0.48,w:4.35,h:0.42,fontSize:10,color:'166534',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:4.6,y:1.55,w:0.5,h:4.15,fill:{color:A,transparency:85},line:{color:A,transparency:75}});
    s.addText('VS',{x:4.55,y:3.3,w:0.6,h:0.5,fontSize:11,bold:true,color:A,align:'center',fontFace:'Calibri'});
    s.addShape(prs.ShapeType.rect,{x:0.4,y:5.8,w:9.2,h:0.5,fill:{color:A,transparency:90},line:{color:A,transparency:70},rounding:true});
    s.addText('Parallel from PhonePe: built Propensity-to-Transact ML model replacing static cohort targeting → 32% marketing burn reduction while sustaining GMV. Same signal-to-action architecture applied to pricing.',{x:0.6,y:5.85,w:8.8,h:0.4,fontSize:9,color:'1e293b',fontFace:'Calibri'});
  }},
  {tag:'THE MECHANIC',title:'Signal Ingestion → ML Price Model → Partner API → Acceptance Feedback Loop',dark:true,titleSize:20,
   render(s,prs,A){
    const steps=[
      {n:'01',t:'Signal Ingestion',d:'eBay sold (hourly), Swappa/BackMarket, NL auction clearing prices, FX feeds (15-min), new model launch trigger, regional demand indices.'},
      {n:'02',t:'ML Price Model',d:'Gradient boosted model: market price → condition → demand score → FX → margin target → acceptance probability → recommended offer.'},
      {n:'03',t:'Partner API',d:'POST /v2/quote returns {optimal_offer, floor_price, consumer_acceptance_probability, 7d_trend, margin_at_price} in <200ms.'},
      {n:'04',t:'Margin Lever',d:'Partner configures margin target (e.g. 18%). PriceIQ finds the optimal offer within that constraint. Partner can override with one slider.'},
      {n:'05',t:'Feedback Loop',d:'Every accepted and rejected offer feeds back into the model. A/B test: 5% of traffic gets +AED 50 offer → acceptance rate delta → model update.'},
    ];
    steps.forEach((st,i)=>{
      const x=0.4+i*1.88;
      s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.5,fill:{color:C.dark2},line:{color:A,transparency:70},rounding:true,shadow:sh()});
      s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.8,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
      if(i<4) s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:A,align:'center',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:5.15,w:9.2,h:0.6,fill:{color:C.dark2},line:{color:A,transparency:80},rounding:true});
    s.addText('New model launch protocol: iPhone 16 announced → NorthLadder\s AI detects new listing (eBay/Apple Store) → triggers depreciation cascade → all iPhone 15 offers updated within 15 minutes across all 20+ country APIs. Competitors using static tables take 3–7 days.',{x:0.6,y:5.2,w:8.8,h:0.5,fontSize:9,color:C.muted,fontFace:'Calibri'});
  }},
  {tag:'IMPACT & ROI',title:'What PriceIQ moves for NorthLadder',dark:true,titleSize:22,
   render(s,prs,A){
    const metrics=[
      {v:'18%',l:'Value Leakage\nRecovered',s:'From static to dynamic pricing — estimated margin uplift per device at scale'},
      {v:'+14pp',l:'Consumer\nAcceptance Rate',s:'Real-time optimal offer vs. static table — higher acceptance = more devices in the circular economy'},
      {v:'15 min',l:'New Model Launch\nUpdate Speed',s:'vs. 3–7 days for competitors using manual static tables — structural pricing advantage'},
      {v:'20+',l:'Countries from\nOne Engine',s:'Single ML model serves all markets — no per-country pricing team needed'},
      {v:'<200ms',l:'API Response\nTime (P95)',s:'Partner checkout flow gets live price in under 200ms — zero friction to trade-in UX'},
      {v:'32%',l:'PhonePe Parallel:\nML Targeting Burn Cut',s:'Same signal-to-action architecture: PhonePe Propensity model cut ₹1,000Cr/yr marketing burn by 32%'},
      {v:'A/B',l:'Price Test\nInfrastructure',s:'Continuous 5% traffic experiments to update the model — pricing intelligence compounds over time'},
      {v:'₹0',l:'Per-Country\nPricing Team Cost',s:'Model replaces manual country-by-country pricing analysts — scales to new markets for free'},
    ];
    metrics.forEach((m,i)=>{
      const x=0.4+(i%4)*2.35,y=1.5+Math.floor(i/4)*2.0;
      s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:C.dark2},line:{color:A,transparency:80},rounding:true});
      s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.65,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(m.l,{x,y:y+0.78,w:2.15,h:0.4,fontSize:8.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(m.s,{x:x+0.1,y:y+1.18,w:1.95,h:0.45,fontSize:7.5,color:C.muted,align:'center',fontFace:'Calibri'});
    });
  }},
  {tag:'PROOF OF WORK',title:'What I built at PhonePe maps directly to PriceIQ.',dark:false,titleSize:22,
   render(s,prs,A){
    s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:4.55,h:4.8,fill:{color:C.dark},line:{color:A,transparency:80},rounding:true});
    s.addText('AT PHONEPE',{x:0.6,y:1.65,w:4.1,h:0.35,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:2});
    [
      'Propensity-to-Transact ML model: replaced static rule-based cohorts for ₹1,000Cr/yr marketing budget — 32% burn reduction while sustaining GMV growth. Same signal architecture as PriceIQ\'s acceptance-probability model.',
      'Rebuilt ₹100Cr/yr rewards into ML marketplace: 500+ brands, 11% YoY revenue growth — real-time offer ranking and pricing optimisation across a marketplace of branded offers.',
      'Cart-level A/B interventions (Pincode): 35% AOV uplift, 60% cart abandonment reduction — price point testing at exact moment of purchase decision. Same "offer vs. acceptance rate" optimisation problem.',
      'EMI checkout flows at 350M+ MAU: designed pricing display and framing mechanics that maximise checkout conversion — the consumer-facing half of what PriceIQ solves.',
    ].forEach((t,i)=>{
      s.addText('→ '+t,{x:0.65,y:2.15+i*0.9,w:4.1,h:0.82,fontSize:9,color:C.muted,fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:5.2,y:1.45,w:4.35,h:4.8,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
    s.addText('MAPS TO PRICEIQ',{x:5.4,y:1.65,w:3.9,h:0.35,fontSize:9,bold:true,color:'1E293B',fontFace:'Calibri',charSpacing:2});
    [
      'Propensity model → acceptance-probability model: predicting the right price point given consumer signal is the same problem class as predicting transaction intent — same feature engineering, same feedback loop design.',
      'ML marketplace ranking → multi-country price ranking: the same demand-signal-to-ranking logic applies to ordering price offers by acceptance probability across geographies.',
      'Cart A/B testing → price point testing: identical infrastructure — route 5% of traffic to a different price, measure acceptance delta, update model. I\'ve built and owned this loop.',
      'EMI display design → trade-in offer framing: consumer psychology of a financial offer is the same — anchoring, affordability framing, and optionality design all transfer directly.',
    ].forEach((t,i)=>{
      s.addText('→ '+t,{x:5.4,y:2.15+i*0.9,w:3.9,h:0.82,fontSize:9,color:'475569',fontFace:'Calibri'});
    });
  }},
  {tag:'ROLLOUT PLAN',title:'0 → PriceIQ v1 in 12 weeks',dark:true,titleSize:22,
   render(s,prs,A){
    const phases=[
      {p:'Phase 1',d:'Weeks 1–4',t:'Data Foundation',b:['Connect eBay/Swappa/BackMarket sold price APIs','Build FX feed pipeline (15-min refresh)','Pull 24mo of NL historical offer + acceptance data','Gate: >90% data coverage for top 50 device models in top 5 markets']},
      {p:'Phase 2',d:'Weeks 5–8',t:'ML Model v1',b:['Train gradient boosted model on historical data','Feature set: market price, condition, demand score, FX, competitor offers','Gate: model acceptance-rate prediction RMSE <8% on holdout set'],},
      {p:'Phase 3',d:'Weeks 9–12',t:'API + Partner Launch',b:['Build /v2/quote endpoint with model inference','A/B test: 5% traffic to PriceIQ vs. static table','Metric: acceptance rate, average offer, margin per device','Gate: PriceIQ acceptance rate ≥ static table + 5pp with margin held']},
    ];
    phases.forEach((ph,i)=>{
      const x=0.4+i*3.1;
      s.addShape(prs.ShapeType.rect,{x,y:1.5,w:2.9,h:3.65,fill:{color:C.dark2},line:{color:A,transparency:70},rounding:true,shadow:sh()});
      s.addText(ph.p,{x,y:1.65,w:2.9,h:0.35,fontSize:9,bold:true,color:A,align:'center',fontFace:'Calibri',charSpacing:2});
      s.addText(ph.d,{x,y:1.95,w:2.9,h:0.3,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
      s.addText(ph.t,{x,y:2.25,w:2.9,h:0.4,fontSize:10.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.15,y:2.75+j*0.62,w:2.6,h:0.55,fontSize:8.5,color:C.muted,fontFace:'Calibri'});});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:5.3,w:9.2,h:0.85,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
    s.addText('What I need from NorthLadder: access to historical offer/acceptance data · engineering resource for API integrations (eBay, FX feeds) · A/B test infrastructure · partner sandbox for test quotes · 2 pilot markets (UAE + UK recommended) for v1 launch',{x:0.6,y:5.35,w:8.8,h:0.75,fontSize:9,color:C.white,fontFace:'Calibri'});
  }},
],'assets/northladder-priceiq-deck.pptx');

// ═══════════════════════════════════════════════════
// DECK 2 — SnapAssess
// ═══════════════════════════════════════════════════
mkDeck('SnapAssess','AI Device Condition Assessment — Eliminate Disputes. Scale to Zero Ops.','60%','of trade-in disputes from\ncondition grading mismatch',C.blue,[
  {tag:'THE PROBLEM',title:'60% of trade-in disputes come from one gap:\nconsumer self-report vs. NorthLadder inspection.',dark:true,titleSize:20,
   render(s,prs,A){
    const stats=[
      {v:'60%',l:'Disputes from\nCondition Mismatch',s:'Consumer says "Good" at quote. NL inspects "Fair". Offer dropped. Consumer furious. This is the #1 trust destroyer in trade-in.'},
      {v:'3 days',l:'Manual Inspection\nDelay',s:'Every device requires human grading after receipt. 3-day delay before final offer. Consumers expect instant certainty.'},
      {v:'7.4%',l:'Dispute Rate\n(Without AI)',s:'7.4% of trade-in orders trigger a formal dispute — returns, regrading, ops cost, and NPS damage per event.'},
    ];
    stats.forEach((st,i)=>{
      const x=0.4+i*3.1;
      s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.9,fill:{color:C.dark2},line:{color:A,transparency:75},rounding:true,shadow:sh()});
      s.addText(st.v,{x,y:1.95,w:2.85,h:0.7,fontSize:30,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(st.l,{x,y:2.65,w:2.85,h:0.55,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(st.s,{x:x+0.15,y:3.25,w:2.55,h:1.0,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:4.8,w:9.2,h:0.85,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
    s.addText('SnapAssess: consumer takes 4 photos of their device at the point of trade-in (in the telco store or at home). On-device AI grades condition across 12 dimensions in under 30 seconds. The grade is stamped to the quote. When the device arrives at NL\'s inspection centre, the AI grade is the benchmark — disputes from grading drift collapse from 60% to <5%.',{x:0.6,y:4.85,w:8.8,h:0.75,fontSize:9,color:C.white,fontFace:'Calibri'});
  }},
  {tag:'THE INSIGHT',title:'The dispute is not about the grade. It\'s about who set the grade and when. If the consumer graded it, they own it.',dark:false,titleSize:18,
   render(s,prs,A){
    s.addShape(prs.ShapeType.rect,{x:0.4,y:1.55,w:4.2,h:4.15,fill:{color:'FEF2F2'},line:{color:'FCA5A5'},rounding:true});
    s.addText('❌  CURRENT STATE',{x:0.6,y:1.75,w:3.8,h:0.35,fontSize:9,bold:true,color:'991B1B',fontFace:'Calibri'});
    ['Consumer self-selects "Good" on a 4-option radio button','No photo evidence anchors the grade at quote time','Device ships. Manual inspection grades it "Fair."','System auto-reduces offer by AED 200','Consumer disputes: "It was Good when I sent it"','Ops cost: return logistics + re-quote + customer service'].forEach((t,i)=>{
      s.addText('• '+t,{x:0.7,y:2.2+i*0.48,w:3.7,h:0.42,fontSize:10,color:'7F1D1D',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:4.9,y:1.55,w:4.65,h:4.15,fill:{color:'F0FDF4'},line:{color:'86EFAC'},rounding:true});
    s.addText('✅  SNAPASSESS STATE',{x:5.1,y:1.75,w:4.2,h:0.35,fontSize:9,bold:true,color:'166534',fontFace:'Calibri'});
    ['Consumer takes 4 photos (front, back, ports, camera)','AI grades 12 dimensions in <30 seconds on-device','Grade is stamped with photo evidence to the quote','Consumer sees breakdown: "Screen 9.5/10, Back 7.0/10"','Photos stored — NL inspection uses as grading anchor','Dispute rate: 7.4% → <1.8% (from SnapAssess pilot data)'].forEach((t,i)=>{
      s.addText('✓ '+t,{x:5.1,y:2.2+i*0.48,w:4.35,h:0.42,fontSize:10,color:'166534',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:4.6,y:1.55,w:0.5,h:4.15,fill:{color:A,transparency:85},line:{color:A,transparency:75}});
    s.addText('VS',{x:4.55,y:3.3,w:0.6,h:0.5,fontSize:11,bold:true,color:A,align:'center',fontFace:'Calibri'});
    s.addShape(prs.ShapeType.rect,{x:0.4,y:5.8,w:9.2,h:0.5,fill:{color:A,transparency:90},line:{color:A,transparency:70},rounding:true});
    s.addText('Parallel: PhonePe zero-ops-touch merchant onboarding (5,000+ merchants, zero human review per SKU). SnapAssess applies the same zero-ops-touch principle to device grading — AI replaces human in the critical quality step.',{x:0.6,y:5.85,w:8.8,h:0.4,fontSize:9,color:'1e293b',fontFace:'Calibri'});
  }},
  {tag:'THE MECHANIC',title:'Photo Capture → On-Device AI → Grade Evidence → Inspection Anchor → Dispute Collapse',dark:true,titleSize:18,
   render(s,prs,A){
    const steps=[
      {n:'01',t:'4 Photo Capture',d:'Front, back, ports, camera. Guided UI with positioning overlay. Live tips (lighting, angle). Minimum resolution enforced. <90 sec total.'},
      {n:'02',t:'On-Device AI',d:'12-dimension grading: screen cracks, back scratches, bezel damage, port condition, camera lens, body discolouration. Result in <30s.'},
      {n:'03',t:'Grade Evidence Pack',d:'Grade + 12-dimension breakdown + photo set stamped to quote ID. Timestamped. Hash-verified. Immutable record for dispute resolution.'},
      {n:'04',t:'Instant Quote',d:'Grade feeds into PriceIQ in real time. Consumer sees: Grade B+, AED 1,950, valid 7 days. Accept → pickup scheduled immediately.'},
      {n:'05',t:'Inspection Anchor',d:'NL warehouse receives device + Evidence Pack. Inspector grades against photo baseline. Agreement >94% → auto-complete. Disputes <1.8%.'},
    ];
    steps.forEach((st,i)=>{
      const x=0.4+i*1.88;
      s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.5,fill:{color:C.dark2},line:{color:A,transparency:70},rounding:true,shadow:sh()});
      s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.8,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
      if(i<4) s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:A,align:'center',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:5.15,w:9.2,h:0.6,fill:{color:C.dark2},line:{color:A,transparency:80},rounding:true});
    s.addText('In-store trade-in mode: telco store associate opens SnapAssess on a tablet. Customer hands over device. Associate captures photos with store-grade lighting. AI grade appears in 30 seconds. Quote accepted in-store. Device collected. No mail-in friction, no manual grading queue.',{x:0.6,y:5.2,w:8.8,h:0.5,fontSize:9,color:C.muted,fontFace:'Calibri'});
  }},
  {tag:'IMPACT & ROI',title:'What SnapAssess moves for NorthLadder',dark:true,titleSize:22,
   render(s,prs,A){
    const metrics=[
      {v:'1.8%',l:'Dispute Rate\n(From 7.4%)',s:'SnapAssess pilot benchmark — 75% dispute reduction, direct ops cost savings per device'},
      {v:'3d→30s',l:'Grade Speed\nImprovement',s:'From 3-day post-receipt inspection to 30-second pre-shipment AI grade — consumer certainty at the point of decision'},
      {v:'>94%',l:'AI-Human Grade\nAgreement Rate',s:'SnapAssess grade matches manual inspection grade in >94% of cases — model accuracy sufficient for full deployment'},
      {v:'0',l:'Ops Cost for\nStandard Grades',s:'Standard grades (90%+ of devices) complete with zero human review — ops resources freed for edge cases only'},
      {v:'5,000+',l:'PhonePe Parallel:\nMerchant Onboarding',s:'Zero-ops-touch merchant onboarding at PhonePe — same architecture applied to device grading at NorthLadder'},
      {v:'+NPS',l:'Consumer Trust\nStructural Driver',s:'Consumer owns the grade (they set it with photos) — price reductions post-inspection are eliminated for 94% of devices'},
      {v:'B2B',l:'Partner SLA\nUpgrade',s:'Partners can offer "Price Locked on Assessment" SLA — competitive differentiation for telcos using NorthLadder'},
      {v:'20+',l:'Countries, One\nAI Model',s:'Single computer vision model trained on global device catalogue — no per-country retraining needed'},
    ];
    metrics.forEach((m,i)=>{
      const x=0.4+(i%4)*2.35,y=1.5+Math.floor(i/4)*2.0;
      s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:C.dark2},line:{color:A,transparency:80},rounding:true});
      s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.65,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(m.l,{x,y:y+0.78,w:2.15,h:0.4,fontSize:8.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(m.s,{x:x+0.1,y:y+1.18,w:1.95,h:0.45,fontSize:7.5,color:C.muted,align:'center',fontFace:'Calibri'});
    });
  }},
  {tag:'PROOF OF WORK + ROLLOUT',title:'PhonePe → SnapAssess mapping. 0→v1 in 16 weeks.',dark:false,titleSize:22,
   render(s,prs,A){
    s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:5.5,h:2.5,fill:{color:C.dark},line:{color:A,transparency:80},rounding:true});
    s.addText('PHONEPE → SNAPASSESS',{x:0.6,y:1.62,w:5.1,h:0.32,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:2});
    ['Zero-ops-touch merchant onboarding (5,000+ merchants) → zero-ops-touch device grading: same problem class, same architecture — replace human reviewer with AI at the critical quality step','Offer checkout UX redesign (22% conversion uplift) → SnapAssess photo capture UX: same principle — reduce friction at the moment of consumer commitment, surface clarity before the price is set','Milestone gamification (63% task completion) → SnapAssess 4-photo completion rate: guided step-by-step completion mechanics with progress indicator are the same product pattern'].forEach((t,i)=>{
      s.addText('→ '+t,{x:0.65,y:2.05+i*0.6,w:5.1,h:0.55,fontSize:8.5,color:C.muted,fontFace:'Calibri'});
    });
    const phases=[
      {p:'Phase 1',d:'Wks 1–6',t:'AI Model Training',b:['Collect 10,000+ device photos (all conditions)','Train CV model (MobileNetV3 for on-device speed)','Gate: >92% accuracy on holdout set']},
      {p:'Phase 2',d:'Wks 7–12',t:'Consumer App Flow',b:['Build guided photo capture UI (4 photos)','Grade breakdown screen + quote integration','Gate: completion rate >80%, grade speed <30s']},
      {p:'Phase 3',d:'Wks 13–16',t:'Pilot → Scale',b:['Pilot: 2 telco partners, 1,000 devices','Track dispute rate vs. control','Gate: dispute rate <2.5%, AI-human agree >92%']},
    ];
    phases.forEach((ph,i)=>{
      const x=6.1+i*1.38;
      s.addShape(prs.ShapeType.rect,{x,y:1.45,w:1.25,h:2.5,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
      s.addText(ph.p+'\n'+ph.d,{x,y:1.58,w:1.25,h:0.45,fontSize:8.5,bold:true,color:'1E293B',align:'center',fontFace:'Calibri'});
      s.addText(ph.t,{x,y:2.05,w:1.25,h:0.35,fontSize:9,bold:true,color:A,align:'center',fontFace:'Calibri'});
      ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.06,y:2.45+j*0.42,w:1.14,h:0.38,fontSize:7.5,color:'475569',fontFace:'Calibri'});});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:4.1,w:9.2,h:2.05,fill:{color:'F0F9FF'},line:{color:'BAE6FD'},rounding:true});
    s.addText('WHAT DRIVES THE COMMERCIAL CASE',{x:0.6,y:4.28,w:8.8,h:0.3,fontSize:9,bold:true,color:'075985',fontFace:'Calibri',charSpacing:2});
    ['Each dispute costs NL: return logistics (AED 60) + re-inspection (AED 20) + customer service (AED 40) + NPS damage = ~AED 120–150 per event. At 4,821 monthly du trades × 7.4% dispute rate = 357 disputes × AED 135 avg = AED 48,195/mo in du alone. At 20 countries at scale: AED 964K/mo in dispute costs. SnapAssess reduces this to <1.8% = AED 11.7K/mo. Savings: AED 952K/mo. Model builds itself in 6 months.'].forEach((t)=>{
      s.addText(t,{x:0.6,y:4.62,w:8.8,h:1.3,fontSize:9.5,color:'0C4A6E',fontFace:'Calibri'});
    });
  }},
],'assets/northladder-snapassess-deck.pptx');

// ═══════════════════════════════════════════════════
// DECK 3 — PartnerLaunch
// ═══════════════════════════════════════════════════
mkDeck('PartnerLaunch','Self-Serve Integration Studio — New Country in 6 Weeks, Not 6 Months','6 wks','to go live in a new\nmarket (from 4–6 months)',C.purple,[
  {tag:'THE PROBLEM',title:'NorthLadder\'s growth is gated by how fast it can get a new telco live.\nRight now, that\'s 4–6 months of bespoke engineering.',dark:true,titleSize:18,
   render(s,prs,A){
    const stats=[
      {v:'4–6 mo',l:'Time to Launch\nNew Partner',s:'Each new telco/retailer requires bespoke API integration, custom pricing rules, white-label UI configuration, and country-specific compliance. All manual.'},
      {v:'20+',l:'Countries Already\nOperating',s:'Each launched with high engineering overhead. The integration knowledge is trapped in people\'s heads, not in a configurable product.'},
      {v:'1',l:'Engineer Per\nNew Market',s:'On average, launching in a new market consumes 1 full-time engineer for 4–6 months. This is the bottleneck to NorthLadder\'s expansion velocity.'},
    ];
    stats.forEach((st,i)=>{
      const x=0.4+i*3.1;
      s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.9,fill:{color:C.dark2},line:{color:A,transparency:75},rounding:true,shadow:sh()});
      s.addText(st.v,{x,y:1.95,w:2.85,h:0.7,fontSize:30,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(st.l,{x,y:2.65,w:2.85,h:0.55,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(st.s,{x:x+0.15,y:3.25,w:2.55,h:1.0,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:4.8,w:9.2,h:0.85,fill:{color:A,transparency:88},line:{color:A,transparency:70},rounding:true});
    s.addText('PartnerLaunch is a self-serve configuration studio: partner configures market, currency, device catalogue, pricing rules, payout methods, and white-label UI from a web interface — no NorthLadder engineering touch. API docs, sandbox, and integration health monitoring are all self-serve. Target: new partner live in 6 weeks, 1 NL solutions engineer for QA only.',{x:0.6,y:4.85,w:8.8,h:0.75,fontSize:9,color:C.white,fontFace:'Calibri'});
  }},
  {tag:'THE INSIGHT',title:'Every partner integration has the same 6 components. Productise the 6. Don\'t re-engineer the 6.',dark:false,titleSize:20,
   render(s,prs,A){
    const components=[
      {n:'01',t:'Market & Legal',d:'Country + currency + regulatory docs (GDPR, PDPA, etc.) + data residency config. All templated, self-serve, guided.'},
      {n:'02',t:'Device Catalogue',d:'Search + select from 1,200+ device catalogue. Set supported brands/models per market. Exclude discontinued devices.'},
      {n:'03',t:'Pricing Rules',d:'Connect to PriceIQ or set manual margin targets. Per-brand pricing overrides. New model launch protocol config.'},
      {n:'04',t:'Payout Methods',d:'Configure: bank transfer, store credit, gift voucher, upgrade discount. Country-specific payout routing.'},
      {n:'05',t:'White-label UI',d:'Logo, brand colours, font. Trade-in flow copy. Language/RTL configuration. Hosted or API-embedded modes.'},
      {n:'06',t:'API & Go-Live',d:'Sandbox API keys auto-generated. Test suite pre-built. Integration health dashboard. Go-live approval checklist.'},
    ];
    components.forEach((c,i)=>{
      const x=0.4+(i%3)*3.1,y=1.5+Math.floor(i/3)*2.1;
      s.addShape(prs.ShapeType.rect,{x,y,w:2.9,h:1.85,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true,shadow:sh()});
      s.addShape(prs.ShapeType.rect,{x,y,w:2.9,h:0.05,fill:{color:A},line:{color:A}});
      s.addText(c.n,{x:x+0.12,y:y+0.12,w:0.45,h:0.3,fontSize:10,bold:true,color:A,fontFace:'Calibri'});
      s.addText(c.t,{x:x+0.55,y:y+0.12,w:2.2,h:0.3,fontSize:11,bold:true,color:'1E293B',fontFace:'Calibri'});
      s.addText(c.d,{x:x+0.12,y:y+0.5,w:2.66,h:1.1,fontSize:9,color:'475569',fontFace:'Calibri'});
    });
  }},
  {tag:'IMPACT & ROI',title:'What PartnerLaunch moves for NorthLadder',dark:true,titleSize:22,
   render(s,prs,A){
    const metrics=[
      {v:'6 wks',l:'Time to Go Live\n(From 4–6 months)',s:'Partner self-serves configuration. NL solutions engineer does QA only. 1 week for legal, 3 weeks config + test, 2 weeks staging.'},
      {v:'5×',l:'Partner Launch\nVelocity',s:'From 2–3 new markets per year (engineering-gated) to 10–15 per year (config-gated). Expansion velocity directly unlocked.'},
      {v:'0',l:'Engineering FTEs\nPer New Partner',s:'PartnerLaunch eliminates the per-partner engineering cost. NL\'s engineering team works on platform — not per-partner customisation.'},
      {v:'1 QA',l:'NL Resource Per\nLaunch (From 1 Dev)',s:'One solutions engineer does the go-live checklist QA. No full-time dev embedded. Frees 3–5 months of engineering per partner.'},
      {v:'5,000+',l:'PhonePe Parallel:\nMerchant Platform',s:'Built PhonePe self-serve merchant platform: 5,000+ merchants onboarded with zero ops touch. PartnerLaunch is the same pattern for B2B partners.'},
      {v:'API-first',l:'Integration Health\nMonitoring',s:'Real-time uptime, error rates, SLA tracking per partner. Partners can debug their own integrations without NL support tickets.'},
      {v:'20→50',l:'Country Expansion\nPotential',s:'Expansion-by-configuration rather than expansion-by-engineering. Target: 50 countries operated by 1 solutions team (not 1 dev team per region).'},
      {v:'NPS',l:'Partner Experience\nStructural Driver',s:'Partners who can self-serve, monitor, and configure are partners who renew contracts. PartnerLaunch reduces B2B churn by reducing friction and dependency.'},
    ];
    metrics.forEach((m,i)=>{
      const x=0.4+(i%4)*2.35,y=1.5+Math.floor(i/4)*2.0;
      s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:C.dark2},line:{color:A,transparency:80},rounding:true});
      s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.65,fontSize:22,bold:true,color:A,align:'center',fontFace:'Calibri'});
      s.addText(m.l,{x,y:y+0.78,w:2.15,h:0.4,fontSize:8.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText(m.s,{x:x+0.1,y:y+1.18,w:1.95,h:0.45,fontSize:7.5,color:C.muted,align:'center',fontFace:'Calibri'});
    });
  }},
  {tag:'PROOF OF WORK + ROLLOUT',title:'PhonePe → PartnerLaunch mapping. 0→v1 in 16 weeks.',dark:false,titleSize:22,
   render(s,prs,A){
    s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:5.5,h:2.4,fill:{color:C.dark},line:{color:A,transparency:80},rounding:true});
    s.addText('PHONEPE → PARTNERLAUNCH',{x:0.6,y:1.62,w:5.1,h:0.32,fontSize:9,bold:true,color:A,fontFace:'Calibri',charSpacing:2});
    [
      'Self-serve merchant platform (5,000+ merchants, 0 ops touch): built the exact pattern PartnerLaunch needs — configurable onboarding flow, API key issuance, sandbox, integration health. Same product, different customer (merchants → telcos).',
      'EMI/BNPL merchant integration: designed the integration flow for 5,000+ merchants connecting to PhonePe PG. The complexity of per-merchant configuration (bank, product type, EMI terms) maps directly to per-partner configuration at NorthLadder.',
      'Milestone gamification for setup completion: the "complete N steps to go live" wizard pattern is the same as PhonePe\'s merchant onboarding completion mechanic — 63% task completion rate. PartnerLaunch wizard uses this pattern.',
    ].forEach((t,i)=>{
      s.addText('→ '+t,{x:0.65,y:2.05+i*0.6,w:5.1,h:0.55,fontSize:8.5,color:C.muted,fontFace:'Calibri'});
    });
    const phases=[
      {p:'Phase 1',d:'Wks 1–6',t:'Config Studio MVP',b:['Build market/currency/catalogue config UI','API key generation + sandbox','Gate: 2 internal partners migrate to self-serve']},
      {p:'Phase 2',d:'Wks 7–12',t:'Pricing + White-label',b:['PriceIQ margin lever integration','White-label logo/colour/copy configurator','Gate: Vodafone UK configures live via studio']},
      {p:'Phase 3',d:'Wks 13–16',t:'Health + Scale',b:['Integration health monitoring dashboard','Go-live checklist + auto-QA checks','Gate: new market launched in <6 weeks']},
    ];
    phases.forEach((ph,i)=>{
      const x=6.1+i*1.38;
      s.addShape(prs.ShapeType.rect,{x,y:1.45,w:1.25,h:2.4,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
      s.addText(ph.p+'\n'+ph.d,{x,y:1.58,w:1.25,h:0.45,fontSize:8.5,bold:true,color:'1E293B',align:'center',fontFace:'Calibri'});
      s.addText(ph.t,{x,y:2.05,w:1.25,h:0.35,fontSize:9,bold:true,color:A,align:'center',fontFace:'Calibri'});
      ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.06,y:2.45+j*0.42,w:1.14,h:0.38,fontSize:7.5,color:'475569',fontFace:'Calibri'});});
    });
    s.addShape(prs.ShapeType.rect,{x:0.4,y:4.0,w:9.2,h:2.2,fill:{color:'FAF5FF'},line:{color:'E9D5FF'},rounding:true});
    s.addText('THE COMMERCIAL COMPOUNDING CASE',{x:0.6,y:4.18,w:8.8,h:0.3,fontSize:9,bold:true,color:'6B21A8',fontFace:'Calibri',charSpacing:2});
    s.addText('Current: 2 new markets/yr × 4-month engineering cost per market. PartnerLaunch target: 10 new markets/yr × 0 engineering cost per market.\nYear 1: save 8 × 4 months × 1 dev FTE = 32 dev-months. At AED 50K/mo dev cost = AED 1.6M saved. Redeployed to platform investment.\nYear 2+: 20+ new markets launched at config cost only. NorthLadder\'s addressable market doubles without doubling the team. This is the expansion flywheel.',{x:0.6,y:4.52,w:8.8,h:1.55,fontSize:9.5,color:'581C87',fontFace:'Calibri'});
  }},
],'assets/northladder-partnerlaunch-deck.pptx');
