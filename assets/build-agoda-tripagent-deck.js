const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
const C={dark:'0B1F3A',hero:'1A3458',bright:'FF6B35',gold:'FFB800',green:'22C55E',purple:'8B5CF6',white:'FFFFFF',lgray:'F0F4F8',muted:'64748B',ink:'0F172A'};
const makeShadow=()=>({type:'outer',color:'000000',blur:3,offset:1,angle:45,opacity:0.12});

// SLIDE 1 — COVER
{const s=pptx.addSlide();s.background={color:C.dark};
for(let i=0;i<8;i++)s.addShape(pptx.ShapeType.line,{x:7.4+i*.22,y:0,w:.01,h:7.5,line:{color:C.bright,width:.8,transparency:82},rotate:35});
s.addShape(pptx.ShapeType.rect,{x:.45,y:.35,w:1.6,h:.3,fill:{color:C.hero},line:{color:C.hero}});
s.addText('AGODA',{x:.45,y:.35,w:1.6,h:.3,fontSize:8,bold:true,color:C.bright,align:'center',valign:'middle',charSpacing:5});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.45,w:.05,h:1.6,fill:{color:C.bright},line:{color:C.bright}});
s.addText('Agoda TripAgent',{x:.65,y:1.4,w:6.5,h:.85,fontSize:42,bold:true,color:C.white,fontFace:'Georgia'});
s.addText('AI agent that plans flights, hotels, cabs & activities end-to-end in one conversation',{x:.65,y:2.28,w:6.8,h:.52,fontSize:15,color:C.gold,fontFace:'Georgia',italic:true});
s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM',{x:.65,y:2.94,w:6,h:.3,fontSize:11,color:'AAAAAA'});
s.addShape(pptx.ShapeType.rect,{x:7.2,y:5.6,w:2.5,h:1.6,fill:{color:C.bright},line:{color:C.bright}});
s.addText('3.4×',{x:7.2,y:5.75,w:2.5,h:.65,fontSize:44,bold:true,color:C.white,align:'center',fontFace:'Georgia'});
s.addText('higher booking value\nvs hotel-only',{x:7.2,y:6.38,w:2.5,h:.65,fontSize:10,color:'DDDDDD',align:'center'});
s.addText('Agoda  ·  AI Travel Planning & New Revenue Expansion',{x:.45,y:7.05,w:6.5,h:.25,fontSize:9,color:'555555',italic:true});}

// SLIDE 2 — PROBLEM
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('THE PROBLEM',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Planning a 3-Day Trip Takes 6 Hours Across 7 Different Platforms',{x:.45,y:.55,w:9.1,h:.7,fontSize:24,bold:true,color:C.white,fontFace:'Georgia'});
const cols=[{icon:'⏱',stat:'6h',lbl:'Average time spent planning a 3-night trip',sub:'Skyscanner, Google Maps, Agoda, Uber, Viator, TripAdvisor, WhatsApp group chat'},{icon:'💸',stat:'$0',lbl:'Agoda revenue from flights, cabs & activities today',sub:'Agoda owns only the hotel layer — missing 70%+ of the total trip economy'},{icon:'🚖',stat:'0',lbl:'Ground transport options inside Agoda',sub:'Airport transfers and inter-city cabs are a completely unclaimed revenue vertical'}];
cols.forEach((col,i)=>{const x=.45+i*3.1;s.addShape(pptx.ShapeType.rect,{x,y:1.55,w:2.85,h:3.2,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(col.icon,{x,y:1.7,w:2.85,h:.5,fontSize:26,align:'center'});s.addText(col.stat,{x,y:2.22,w:2.85,h:.65,fontSize:40,bold:true,color:C.bright,align:'center',fontFace:'Georgia'});s.addText(col.lbl,{x:x+.12,y:2.88,w:2.6,h:.6,fontSize:11,bold:true,color:C.white,align:'center',wrap:true});s.addText(col.sub,{x:x+.12,y:3.5,w:2.6,h:.8,fontSize:9.5,color:'AAAAAA',align:'center',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.0,w:9.1,h:.95,fill:{color:'1A1A00'},line:{color:C.gold,width:.5}});
s.addText('Root cause: Agoda is a hotel booking app in a world where travelers want a trip planner. TripAgent closes the gap — and in doing so, unlocks ground transport as a net-new revenue stream.',{x:.65,y:5.08,w:8.7,h:.8,fontSize:10.5,color:C.white,italic:true,wrap:true,valign:'middle'});}

// SLIDE 3 — INSIGHT
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('THE INSIGHT',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('Travelers Want a Planner, Not a Search Bar. The Agent Is the Product.',{x:.45,y:.55,w:9.1,h:.65,fontSize:22,bold:true,color:C.ink,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.42,w:4.0,h:3.75,fill:{color:'EAE8E0'},line:{color:'D8D0C4'}});
s.addText('Without TripAgent (Today)',{x:.6,y:1.58,w:3.7,h:.32,fontSize:13,bold:true,color:C.bright});
['Open Skyscanner → find flights (30 min)', 'Open Agoda → search hotels (45 min)','Open Uber/Grab for airport transfer → may not pre-book','Open Viator for activities → separate checkout (60 min)','WhatsApp group to coordinate → more friction','Total: 4–6 hours, 5–7 apps, 3–5 separate payment flows'].forEach((r,i)=>s.addText((i===5?'→  ':'✕  ')+r,{x:.62,y:2.05+i*.46,w:3.6,h:.4,fontSize:10,color:i===5?'1A56DB':'8B0000',wrap:true,bold:i===5}));
s.addShape(pptx.ShapeType.ellipse,{x:4.52,y:3.0,w:.6,h:.6,fill:{color:C.dark},line:{color:C.dark}});
s.addText('VS',{x:4.52,y:3.0,w:.6,h:.6,fontSize:9,bold:true,color:C.white,align:'center',valign:'middle'});
s.addShape(pptx.ShapeType.rect,{x:5.2,y:1.42,w:4.35,h:3.75,fill:{color:C.dark},line:{color:C.dark}});
s.addText('With TripAgent (Proposed)',{x:5.35,y:1.58,w:4.0,h:.32,fontSize:13,bold:true,color:C.gold});
['"Plan Chiang Mai, 3 days, $500, I like hiking"','Agent returns: cheapest flights + hotel + cabs + activities','Full day-by-day itinerary with pre-scheduled cab transfers','One checkout — entire bundle paid in a single flow','Cab pick-up times auto-synced with flight arrival','Total: under 2 hours, 1 app, 1 payment — $77 saved vs ad-hoc booking'].forEach((r,i)=>s.addText((i===5?'→  ':'✓  ')+r,{x:5.35,y:2.05+i*.46,w:4.0,h:.4,fontSize:10,color:i===5?C.gold:'A8D5A2',wrap:true,bold:i===5}));
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.38,w:9.1,h:.62,fill:{color:C.dark},line:{color:C.dark}});
s.addText('Key insight: Ground transport is not a feature — it is a new revenue vertical. Every cab Agoda pre-books through TripAgent is GMV that currently goes to Grab, Bolt, or local operators.',{x:.65,y:5.43,w:8.7,h:.52,fontSize:10,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 4 — MECHANIC
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('THE MECHANIC',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('"Plan Chiang Mai, 3 Days, $500, I Like Hiking" — Done in 47 Seconds',{x:.45,y:.55,w:9.1,h:.55,fontSize:22,bold:true,color:C.white,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.line,{x:.88,y:1.6,w:0,h:4.5,line:{color:C.bright,width:1,dashType:'dash',transparency:50}});
[{num:'01',title:'Natural Language Intent Capture',body:'User types (or speaks) destination, dates, budget, and preferences. Agent asks 2 clarifying questions (solo vs group, travel style). Takes under 30 seconds.'},{num:'02',title:'AI Builds Full Itinerary',body:'Agent queries Agoda flight API (cheapest options), hotel API (persona-matched), ground transport API (cab pre-bookings), and activities database. Returns complete day-by-day plan with costs.'},{num:'03',title:'User Reviews and Adjusts',body:'Full itinerary displayed: flight card, hotel card, cab cards (airport + inter-day), activity cards. User can swap any component — "different hotel" or "skip activity 2" — agent re-prices instantly.'},{num:'04',title:'One-Tap Bundle Checkout',body:'Single payment flow covers all components. AgodaCoins awarded at 2x rate on full-bundle bookings. Cab transfer times auto-synced with flight arrival — no manual coordination.'},{num:'05',title:'Post-Trip Intelligence Loop',body:'Agent learns from the trip: what was actually visited vs planned, satisfaction scores, budget vs spend. Next itinerary is more personalised. High-frequency travelers develop a "travel DNA" profile.'}].forEach((step,i)=>{const y=1.38+i*.92;s.addShape(pptx.ShapeType.ellipse,{x:.6,y:y+.05,w:.55,h:.55,fill:{color:C.bright},line:{color:C.bright}});s.addText(step.num,{x:.6,y:y+.05,w:.55,h:.55,fontSize:9,bold:true,color:C.dark,align:'center',valign:'middle'});s.addText(step.title,{x:1.3,y,w:3.6,h:.3,fontSize:12,bold:true,color:C.white});s.addText(step.body,{x:1.3,y:y+.3,w:8.1,h:.52,fontSize:9.5,color:'BBBBBB',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:6.18,w:9.1,h:.55,fill:{color:C.hero},line:{color:C.hero}});
s.addText('PhonePe proof: Self-serve merchant onboarding platform cut TAT from 2 days to 30 minutes — same principle: eliminate friction at every step of a multi-step flow. TripAgent does this for trip planning.',{x:.65,y:6.24,w:8.7,h:.44,fontSize:9.5,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 5 — PRODUCT
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('THE PRODUCT',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('4 Screens — Chat to Bundle Confirmed in Under 2 Minutes',{x:.45,y:.55,w:9.1,h:.5,fontSize:24,bold:true,color:C.ink,fontFace:'Georgia'});
[{num:'01',title:'Agent Chat',desc:'Natural language input. Intent clarification chips (solo/group, budget/luxury). Typing indicator while agent builds plan. Full context preserved across sessions.'},{num:'02',title:'Full Itinerary',desc:'Day-by-day plan: flight card + hotel card + cab cards + activity cards. Total cost shown vs budget. Swap any component. All pre-scheduled — no manual coordination.'},{num:'03',title:'Bundle Confirmed',desc:'Single checkout for all components. Breakdown by category. 2x AgodaCoins on bundles. Cab pick-ups pre-scheduled. All confirmation emails in one.'},{num:'04',title:'Ops Dashboard',desc:'Avg bundle value (3.4x vs hotel-only). Planning time reduction (-45%). Cab GMV ($52 avg/booking, 12% of total). Agent re-use rate (68%). Top intent destinations.'}].forEach((card,i)=>{const x=.4+i*2.38;s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.18,h:4.55,fill:{color:C.dark},line:{color:C.dark},shadow:makeShadow()});s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.18,h:.05,fill:{color:C.bright},line:{color:C.bright}});s.addText(card.num,{x:x+.12,y:1.4,w:.5,h:.3,fontSize:9,bold:true,color:C.bright,fontFace:'Courier New'});s.addText(card.title,{x:x+.12,y:1.7,w:1.92,h:.5,fontSize:12,bold:true,color:C.white,wrap:true});s.addShape(pptx.ShapeType.rect,{x:x+.12,y:2.28,w:1.92,h:2.05,fill:{color:C.hero},line:{color:'2A3A5A'}});const mockups=['┌─────────────┐\n│ 🤖 TripAgent│\n│ Chiang Mai │\n│ 3d · $500  │\n│ 🎒 Solo    │\n│ Building.. │\n└─────────────┘','┌─────────────┐\n│ ✈ $164 flt │\n│ 🏨 $87 htl │\n│ 🚖 $52 cab │\n│ 🗺 $120 act│\n│ Total: $423 │\n└─────────────┘','┌─────────────┐\n│ ✓ BOOKED!  │\n│ $164 + $87 │\n│ + $52 + $120│\n│ = $423 total│\n│ 2x Coins!  │\n└─────────────┘','┌─────────────┐\n│Bundle: 3.4× │\n│Cab GMV:$52 │\n│Plan time:   │\n│  -45%      │\n│Reuse: 68%  │\n└─────────────┘'];s.addText(mockups[i],{x:x+.15,y:2.35,w:1.86,h:1.9,fontSize:7,color:'88AACC',fontFace:'Courier New',wrap:true,valign:'top'});s.addText(card.desc,{x:x+.1,y:4.42,w:1.98,h:1.3,fontSize:9,color:'BBBBBB',wrap:true});});
s.addText('Interactive prototype: agoda-tripagent-prototype.html  ·  All 4 screen states live',{x:.45,y:6.02,w:9.1,h:.22,fontSize:9,color:C.muted,italic:true});}

// SLIDE 6 — IMPACT
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('IMPACT & ROI',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('3.4x Booking Value + A Net-New Revenue Vertical',{x:.45,y:.55,w:9.1,h:.55,fontSize:24,bold:true,color:C.white,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.3,w:4.4,h:.3,fill:{color:C.bright},line:{color:C.bright}});
s.addText('TRAVELER IMPACT',{x:.45,y:1.3,w:4.4,h:.3,fontSize:8,bold:true,color:C.white,align:'center',valign:'middle',charSpacing:3});
[{val:'−45%',lbl:'Trip planning time',sub:'6 hours reduced to under 2 hours'},{val:'3.4×',lbl:'Avg bundle value vs hotel-only',sub:'$423 bundle vs $112 solo hotel'},{val:'68%',lbl:'Agent re-use rate',sub:'Users return to plan their next trip'},{val:'$77',lbl:'Saved vs ad-hoc booking',sub:'Pre-scheduled cabs avoid surge pricing'}].forEach((m,i)=>{const y=1.75+i*1.0;s.addShape(pptx.ShapeType.rect,{x:.45,y,w:4.4,h:.88,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(m.val,{x:.6,y:y+.06,w:1.1,h:.52,fontSize:30,bold:true,color:C.bright,fontFace:'Georgia',align:'center',valign:'middle'});s.addText(m.lbl,{x:1.75,y:y+.06,w:2.95,h:.32,fontSize:11,bold:true,color:C.white});s.addText(m.sub,{x:1.75,y:y+.42,w:2.95,h:.32,fontSize:9,color:'888888'});});
s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.3,w:4.4,h:.3,fill:{color:C.hero},line:{color:'2A3A5A'}});
s.addText('AGODA BUSINESS ROI',{x:5.15,y:1.3,w:4.4,h:.3,fontSize:8,bold:true,color:C.gold,align:'center',valign:'middle',charSpacing:3});
[{val:'$52',lbl:'Avg cab GMV per bundle booking',sub:'New vertical — 18% commission = $9.36/booking'},{val:'+12%',lbl:'Total GMV uplift from cab commissions',sub:'Scales with bundle volume, zero incremental cost'},{val:'4.3×',lbl:'GMV per TripAgent session vs search',sub:'Flights + hotel + cabs vs hotel alone'},{val:'2× coins',lbl:'Bundle loyalty driver',sub:'Users return to spend coins — higher repeat rate'}].forEach((m,i)=>{const y=1.75+i*1.0;s.addShape(pptx.ShapeType.rect,{x:5.15,y,w:4.4,h:.88,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(m.val,{x:5.3,y:y+.06,w:1.2,h:.52,fontSize:26,bold:true,color:'C4B5FD',fontFace:'Georgia',align:'center',valign:'middle'});s.addText(m.lbl,{x:6.55,y:y+.06,w:2.85,h:.32,fontSize:11,bold:true,color:C.white});s.addText(m.sub,{x:6.55,y:y+.42,w:2.85,h:.32,fontSize:9,color:'888888'});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.95,w:9.1,h:.55,fill:{color:C.bright},line:{color:C.bright}});
s.addText('TripAgent makes cab integration the killer feature — once users trust Agoda to pre-book their airport transfer, they stop using Grab for every travel need.',{x:.65,y:6.0,w:8.7,h:.45,fontSize:10,color:C.white,italic:true,wrap:true,valign:'middle'});}

// SLIDE 7 — PROOF
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('PROOF OF WORK',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('I Built the Infrastructure. Here Is the Parallel.',{x:.45,y:.55,w:9.1,h:.5,fontSize:26,bold:true,color:C.ink,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.22,w:4.4,h:4.7,fill:{color:C.dark},line:{color:C.dark}});
s.addText('What I Built at PhonePe',{x:.65,y:1.38,w:4.0,h:.32,fontSize:13,bold:true,color:C.gold});
['Zero-to-one self-serve merchant onboarding platform — cut onboarding TAT from 2 days to 30 minutes; eliminated every friction step that caused merchants to abandon mid-flow','Multi-tenant referral engine — 5,000+ B2B merchants, 5Mn+ users/month, 23% lower CAC vs prior manual process; built entire platform 0→1 with cross-functional team','Dynamic incentivisation engine at Pincode — real-time context-aware triggers across cart value, user intent, time signals; 60% cart abandonment reduction','ML propensity models at 350M MAU scale — every architectural decision on latency, real-time data, and fallback had to work at national UPI volume'].forEach((p,i)=>s.addText('→  '+p,{x:.65,y:1.82+i*.78,w:4.05,h:.68,fontSize:9.5,color:'CCCCCC',wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.22,w:4.4,h:4.7,fill:{color:'EDE8E0'},line:{color:'D8D0C4'}});
s.addText('How It Maps to TripAgent',{x:5.35,y:1.38,w:4.0,h:.32,fontSize:13,bold:true,color:C.bright});
['Merchant onboarding TAT reduction  →  Trip planning time reduction (same "remove friction from multi-step flow" design principle)','Multi-tenant referral engine (0→1 platform)  →  TripAgent is also 0→1 infrastructure for Agoda','Context-aware incentivisation  →  Agent personalizes bundle based on user intent signals in real time','350M MAU API architecture  →  TripAgent queries multiple APIs (flights, hotels, cabs) with sub-second response requirements','23% CAC reduction  →  TripAgent drives direct bookings — less reliance on paid acquisition for high-value users'].forEach((m,i)=>s.addText('✓  '+m,{x:5.35,y:1.82+i*.78,w:4.05,h:.68,fontSize:9.5,color:C.ink,wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:.45,y:6.1,w:9.1,h:.6,fill:{color:C.dark},line:{color:C.dark}});
s.addText('"I turned a 2-day manual merchant onboarding into a 30-minute self-serve flow. TripAgent is the same problem: a 6-hour planning ordeal turned into a 2-minute conversation."',{x:.65,y:6.15,w:8.7,h:.5,fontSize:10,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 8 — ROLLOUT
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('ROLLOUT PLAN',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Phased Launch — 6-Month Plan',{x:.45,y:.55,w:9.1,h:.5,fontSize:26,bold:true,color:C.white,fontFace:'Georgia'});
[{phase:'Phase 1',period:'Month 1–2: Pilot',body:'Launch TripAgent for top 3 origin-destination pairs (BKK, BLR, SIN outbound). Enable flights + hotels only. A/B test vs search control. Measure bundle value, planning time, and rebooking rate.'},{phase:'Phase 2',period:'Month 3–4: Cab Integration',body:'Partner with top cab operators in Bangkok, Bali, Chiang Mai (Grab, Bolt, local). Add cab layer to TripAgent bundle. Instrument cab GMV separately. Target $52 avg cab value per booking.'},{phase:'Phase 3',period:'Month 5–6: Full Expansion',body:'Expand to all Agoda markets. Add activities layer (Viator API or Agoda Experiences). Launch "Travel DNA" profile — agent learns user preferences across trips. Target 3.4x avg bundle value baseline.'}].forEach((p,i)=>{const x=.45+i*3.12;s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.88,h:3.55,fill:{color:C.hero},line:{color:'2A3A5A'},shadow:makeShadow()});s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.88,h:.06,fill:{color:C.bright},line:{color:C.bright}});s.addText(p.phase,{x:x+.14,y:1.42,w:2.6,h:.28,fontSize:9,bold:true,color:C.bright,charSpacing:3});s.addText(p.period,{x:x+.14,y:1.72,w:2.6,h:.38,fontSize:12,bold:true,color:C.white});s.addText(p.body,{x:x+.14,y:2.18,w:2.6,h:2.5,fontSize:9.5,color:'CCCCCC',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.02,w:9.1,h:.92,fill:{color:C.bright},line:{color:C.bright}});
s.addText('What I Need to Build This',{x:.65,y:5.1,w:9,h:.28,fontSize:11,bold:true,color:C.white});
s.addText('Flights API integration (Agoda or Skyscanner partner)  ·  Cab operator API partnerships (Grab, Bolt)  ·  AI agent LLM layer (GPT-4o or equivalent)  ·  2 backend engineers  ·  A/B test vs search control group',{x:.65,y:5.4,w:8.7,h:.46,fontSize:9.5,color:'FFEEEE',wrap:true});
s.addShape(pptx.ShapeType.rect,{x:0,y:6.9,w:10,h:.6,fill:{color:C.hero},line:{color:C.hero}});
s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade',{x:.45,y:6.95,w:9.1,h:.45,fontSize:10,color:'AAAAAA',align:'center'});}

pptx.writeFile({fileName:'assets/agoda-tripagent-deck.pptx'}).then(()=>console.log('agoda-tripagent-deck.pptx done')).catch(e=>console.error(e));
