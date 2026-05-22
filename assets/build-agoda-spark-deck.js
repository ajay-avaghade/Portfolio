const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
const C={dark:'0B1F3A',hero:'1A3458',bright:'FF6B35',gold:'FFB800',green:'22C55E',white:'FFFFFF',lgray:'F0F4F8',muted:'64748B',ink:'0F172A'};
const makeShadow=()=>({type:'outer',color:'000000',blur:3,offset:1,angle:45,opacity:0.12});

// SLIDE 1 — COVER
{const s=pptx.addSlide();s.background={color:C.dark};
for(let i=0;i<8;i++)s.addShape(pptx.ShapeType.line,{x:7.4+i*.22,y:0,w:.01,h:7.5,line:{color:C.bright,width:.8,transparency:82},rotate:35});
s.addShape(pptx.ShapeType.rect,{x:.45,y:.35,w:1.6,h:.3,fill:{color:C.hero},line:{color:C.hero}});
s.addText('AGODA',{x:.45,y:.35,w:1.6,h:.3,fontSize:8,bold:true,color:C.bright,align:'center',valign:'middle',charSpacing:5});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.45,w:.05,h:1.6,fill:{color:C.bright},line:{color:C.bright}});
s.addText('Agoda Spark',{x:.65,y:1.4,w:6.5,h:.85,fontSize:46,bold:true,color:C.white,fontFace:'Georgia'});
s.addText('Short-form travel content that collapses the path from inspiration to booking',{x:.65,y:2.25,w:6.8,h:.5,fontSize:16,color:C.gold,fontFace:'Georgia',italic:true});
s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM',{x:.65,y:2.9,w:6,h:.3,fontSize:11,color:'AAAAAA'});
s.addShape(pptx.ShapeType.rect,{x:7.2,y:5.6,w:2.5,h:1.6,fill:{color:C.bright},line:{color:C.bright}});
s.addText('35%',{x:7.2,y:5.75,w:2.5,h:.65,fontSize:44,bold:true,color:C.white,align:'center',fontFace:'Georgia'});
s.addText('of bookings from\nSpark (target)',{x:7.2,y:6.38,w:2.5,h:.65,fontSize:10,color:'DDDDDD',align:'center'});
s.addText('Agoda  ·  Travel Discovery & Conversion',{x:.45,y:7.05,w:6.5,h:.25,fontSize:9,color:'555555',italic:true});}

// SLIDE 2 — PROBLEM
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('THE PROBLEM',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Agoda Only Captures Travelers Who Already Know Where They Want to Go',{x:.45,y:.55,w:9.1,h:.7,fontSize:24,bold:true,color:C.white,fontFace:'Georgia'});
const cols=[{icon:'📱',stat:'68%',lbl:'Travelers discover destinations on short-form social content',sub:'Instagram, TikTok, YouTube — before they ever open Agoda or a search bar'},{icon:'🔍',stat:'0%',lbl:'Of this inspiration-stage traffic captured by Agoda today',sub:'Agoda is a search-intent product. It misses the entire top of the travel funnel'},{icon:'💸',stat:'8.3×',lbl:'Higher LTV of travelers who book their first trip via discovery content',sub:'Inspiration-led travelers explore more destinations and book more frequently'}];
cols.forEach((col,i)=>{const x=.45+i*3.1;s.addShape(pptx.ShapeType.rect,{x,y:1.55,w:2.85,h:3.2,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(col.icon,{x,y:1.7,w:2.85,h:.5,fontSize:26,align:'center'});s.addText(col.stat,{x,y:2.22,w:2.85,h:.65,fontSize:40,bold:true,color:C.bright,align:'center',fontFace:'Georgia'});s.addText(col.lbl,{x:x+.12,y:2.88,w:2.6,h:.6,fontSize:11,bold:true,color:C.white,align:'center',wrap:true});s.addText(col.sub,{x:x+.12,y:3.5,w:2.6,h:.8,fontSize:9.5,color:'AAAAAA',align:'center',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.0,w:9.1,h:.95,fill:{color:'1A1A00'},line:{color:C.gold,width:.5}});
s.addText('Root cause: Travel is a dream before it is a decision. Agoda sits at step 4 of a 4-step journey (dream → research → compare → book) and captures only users who have already completed steps 1-3 elsewhere.',{x:.65,y:5.08,w:8.7,h:.8,fontSize:10.5,color:C.white,italic:true,wrap:true,valign:'middle'});}

// SLIDE 3 — INSIGHT
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('THE INSIGHT',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('Shoppable Content Beats a Search Bar at the Top of the Travel Funnel',{x:.45,y:.55,w:9.1,h:.65,fontSize:22,bold:true,color:C.ink,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.42,w:4.0,h:3.75,fill:{color:'EAE8E0'},line:{color:'D8D0C4'}});
s.addText('Generic Discovery (Today)',{x:.6,y:1.58,w:3.7,h:.32,fontSize:13,bold:true,color:C.bright});
['User has travel intent but no destination in mind','Opens TikTok or Instagram to browse for ideas','Finds destination inspiration in content','Leaves social platform and opens Agoda separately','Searches from scratch — loses momentum, may not book'].forEach((r,i)=>s.addText('✕  '+r,{x:.62,y:2.05+i*.48,w:3.6,h:.42,fontSize:10,color:'8B0000',wrap:true}));
s.addShape(pptx.ShapeType.ellipse,{x:4.52,y:3.0,w:.6,h:.6,fill:{color:C.dark},line:{color:C.dark}});
s.addText('VS',{x:4.52,y:3.0,w:.6,h:.6,fontSize:9,bold:true,color:C.white,align:'center',valign:'middle'});
s.addShape(pptx.ShapeType.rect,{x:5.2,y:1.42,w:4.35,h:3.75,fill:{color:C.dark},line:{color:C.dark}});
s.addText('Agoda Spark (Proposed)',{x:5.35,y:1.58,w:4.0,h:.32,fontSize:13,bold:true,color:C.gold});
['Spark feed shows shoppable destination clips in the Agoda home tab','User sees Kyoto clip — flights + hotel cost overlaid on the content','Taps "Book This Trip" — auto-populated bundle appears in-app','One more tap confirms flights + hotel in a single session','Zero platform switching — dream to booked in under 3 minutes'].forEach((r,i)=>s.addText('✓  '+r,{x:5.35,y:2.05+i*.48,w:4.0,h:.42,fontSize:10,color:'A8D5A2',wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.38,w:9.1,h:.62,fill:{color:C.dark},line:{color:C.dark}});
s.addText('Key insight: Agoda already has the inventory (flights + hotels). Spark is the discovery layer that brings users to it before a competitor does.',{x:.65,y:5.43,w:8.7,h:.52,fontSize:10,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 4 — MECHANIC
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('THE MECHANIC',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('From Scroll to Booked — Five Steps Inside One App',{x:.45,y:.55,w:9.1,h:.55,fontSize:24,bold:true,color:C.white,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.line,{x:.88,y:1.6,w:0,h:4.5,line:{color:C.bright,width:1,dashType:'dash',transparency:50}});
[{num:'01',title:'Spark Feed Loads on Home Tab',body:'Personalised short-form clips ranked by: past booking history match, seasonal best-time-to-visit relevance, and price range alignment to the user\'s historical booking behavior.'},{num:'02',title:'User Engages with a Clip',body:'Tap or swipe on any clip. Destination details appear as an overlay: destination name, season tag ("Cherry blossom season"), and estimated bundle cost — flights + hotel combined.'},{num:'03',title:'Trip Card Auto-Populates',body:'One tap on "Book This Trip" surfaces a pre-filled bundle: cheapest flights for the user\'s home city, top-matched hotel, total cost — no search required. User can adjust dates or swap components.'},{num:'04',title:'Bundle Confirmed at Checkout',body:'User pays for flights + hotel in a single checkout flow. AgodaCoins awarded on the full bundle value. Booking sourced and attributed to Spark content — creator gets a performance signal.'},{num:'05',title:'Demand Signal Feeds Supply Team',body:'Spark engagement (saves, views, tap-through rate) by destination becomes a real-time demand signal. Supply team uses it to prioritise hotel inventory acquisition and flight partnerships in high-interest markets.'}].forEach((step,i)=>{const y=1.38+i*.92;s.addShape(pptx.ShapeType.ellipse,{x:.6,y:y+.05,w:.55,h:.55,fill:{color:C.bright},line:{color:C.bright}});s.addText(step.num,{x:.6,y:y+.05,w:.55,h:.55,fontSize:9,bold:true,color:C.dark,align:'center',valign:'middle'});s.addText(step.title,{x:1.3,y,w:3.6,h:.3,fontSize:12,bold:true,color:C.white});s.addText(step.body,{x:1.3,y:y+.3,w:8.1,h:.52,fontSize:9.5,color:'BBBBBB',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:6.18,w:9.1,h:.55,fill:{color:C.hero},line:{color:C.hero}});
s.addText('PhonePe proof: Cart incentivisation engine with context-aware triggers (cart value x user intent x time) drove 60% cart abandonment reduction and 20% 30-day retention uplift. Spark applies the same context-awareness to travel discovery.',{x:.65,y:6.24,w:8.7,h:.44,fontSize:9.5,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 5 — PRODUCT
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('THE PRODUCT',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('4 Screens — Discovery to Booking in One Session',{x:.45,y:.55,w:9.1,h:.5,fontSize:26,bold:true,color:C.ink,fontFace:'Georgia'});
[{num:'01',title:'Spark Feed',desc:'Shoppable clips in Agoda home tab. Destination name, season tag, and total bundle cost overlaid on content. "Book This Trip" CTA embedded in-frame. Ranked by persona match + price alignment.'},{num:'02',title:'Trip Card',desc:'Auto-populated bundle: cheapest flights + top hotel + estimated total. Match score shown ("94% for solo travelers like you"). Swap any component. No manual search.'},{num:'03',title:'Booking Confirmed',desc:'Single-checkout for flights + hotel. AgodaCoins 2x on bundle bookings. "Sparked by @creator" attribution. Share trip to earn bonus coins if friends book.'},{num:'04',title:'Partner Ops',desc:'Content-to-booking conversion rate by destination. Avg bundle value ($487 vs $112 hotel-only). Top converting destinations. Real-time demand signals for supply team.'}].forEach((card,i)=>{const x=.4+i*2.38;s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.18,h:4.55,fill:{color:C.dark},line:{color:C.dark},shadow:makeShadow()});s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.18,h:.05,fill:{color:C.bright},line:{color:C.bright}});s.addText(card.num,{x:x+.12,y:1.4,w:.5,h:.3,fontSize:9,bold:true,color:C.bright,fontFace:'Courier New'});s.addText(card.title,{x:x+.12,y:1.7,w:1.92,h:.5,fontSize:12,bold:true,color:C.white,wrap:true});s.addShape(pptx.ShapeType.rect,{x:x+.12,y:2.28,w:1.92,h:2.05,fill:{color:C.hero},line:{color:'2A3A5A'}});const mockups=['┌─────────────┐\n│⚡ SPARK     │\n│ 🏯 Kyoto   │\n│  Apr best  │\n│  $487 total │\n│[Book This] │\n└─────────────┘','┌─────────────┐\n│Kyoto Bundle │\n│ Match: 94%  │\n│ ✈ $189 flt │\n│ 🏨 $297 htl │\n│ Total $487  │\n└─────────────┘','┌─────────────┐\n│ ✓ BOOKED!  │\n│ Kyoto 🏯    │\n│ Apr 25–28  │\n│ Total $487  │\n│⚡ Sparked by│\n└─────────────┘','┌─────────────┐\n│ Spark Ops  │\n│Conv: 35%   │\n│Avg: $487   │\n│New users:  │\n│ +28%       │\n└─────────────┘'];s.addText(mockups[i],{x:x+.15,y:2.35,w:1.86,h:1.9,fontSize:7,color:'88AACC',fontFace:'Courier New',wrap:true,valign:'top'});s.addText(card.desc,{x:x+.1,y:4.42,w:1.98,h:1.3,fontSize:9,color:'BBBBBB',wrap:true});});
s.addText('Interactive prototype: agoda-spark-prototype.html  ·  All 4 screen states live',{x:.45,y:6.02,w:9.1,h:.22,fontSize:9,color:C.muted,italic:true});}

// SLIDE 6 — IMPACT
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('IMPACT & ROI',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Projected Impact — Built on PhonePe Proof Points',{x:.45,y:.55,w:9.1,h:.55,fontSize:24,bold:true,color:C.white,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.3,w:4.4,h:.3,fill:{color:C.bright},line:{color:C.bright}});
s.addText('TRAVELER IMPACT',{x:.45,y:1.3,w:4.4,h:.3,fontSize:8,bold:true,color:C.white,align:'center',valign:'middle',charSpacing:3});
[{val:'35%',lbl:'Bookings initiated from Spark',sub:'New discovery funnel — currently 0%'},{val:'28%',lbl:'New user activation rate',sub:'First booking via Spark content'},{val:'2.1×',lbl:'Session depth increase',sub:'Content → property → booking in one session'},{val:'8.3×',lbl:'Higher LTV vs search-intent users',sub:'Inspiration-led travelers book more trips'}].forEach((m,i)=>{const y=1.75+i*1.0;s.addShape(pptx.ShapeType.rect,{x:.45,y,w:4.4,h:.88,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(m.val,{x:.6,y:y+.06,w:1.1,h:.52,fontSize:30,bold:true,color:C.bright,fontFace:'Georgia',align:'center',valign:'middle'});s.addText(m.lbl,{x:1.75,y:y+.06,w:2.95,h:.32,fontSize:11,bold:true,color:C.white});s.addText(m.sub,{x:1.75,y:y+.42,w:2.95,h:.32,fontSize:9,color:'888888'});});
s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.3,w:4.4,h:.3,fill:{color:C.hero},line:{color:'2A3A5A'}});
s.addText('AGODA BUSINESS ROI',{x:5.15,y:1.3,w:4.4,h:.3,fontSize:8,bold:true,color:C.gold,align:'center',valign:'middle',charSpacing:3});
[{val:'$487',lbl:'Avg Spark bundle value',sub:'vs $112 hotel-only booking baseline'},{val:'4.3×',lbl:'GMV per Spark booking vs solo hotel',sub:'Flights + hotel booked together'},{val:'+12%',lbl:'New supply demand signal accuracy',sub:'Spark saves predict future high-demand routes'},{val:'2× coins',lbl:'AgodaCoins on bundle — retention driver',sub:'Users return to spend coins on next Spark trip'}].forEach((m,i)=>{const y=1.75+i*1.0;s.addShape(pptx.ShapeType.rect,{x:5.15,y,w:4.4,h:.88,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(m.val,{x:5.3,y:y+.06,w:1.2,h:.52,fontSize:28,bold:true,color:'60A5FA',fontFace:'Georgia',align:'center',valign:'middle'});s.addText(m.lbl,{x:6.55,y:y+.06,w:2.85,h:.32,fontSize:11,bold:true,color:C.white});s.addText(m.sub,{x:6.55,y:y+.42,w:2.85,h:.32,fontSize:9,color:'888888'});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.95,w:9.1,h:.55,fill:{color:C.bright},line:{color:C.bright}});
s.addText('Spark shifts Agoda from a last-click conversion tool to a full-funnel travel brand — owning the moment travelers decide where to dream next.',{x:.65,y:6.0,w:8.7,h:.45,fontSize:10,color:C.white,italic:true,wrap:true,valign:'middle'});}

// SLIDE 7 — PROOF
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('PROOF OF WORK',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('I Built the Mechanic. Here Is the Parallel.',{x:.45,y:.55,w:9.1,h:.5,fontSize:26,bold:true,color:C.ink,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.22,w:4.4,h:4.7,fill:{color:C.dark},line:{color:C.dark}});
s.addText('What I Built at PhonePe',{x:.65,y:1.38,w:4.0,h:.32,fontSize:13,bold:true,color:C.gold});
['Dynamic cart incentivisation engine for Pincode (Q-commerce) — context-aware triggers across cart value, user intent, and real-time time signals; 60% cart abandonment reduction','Propensity-to-Transact ML replacing manual cohort targeting with real-time user-level scoring — cut 32% marketing burn while sustaining GMV growth across all channels','Rebuilt static rewards portfolio into ML-driven marketplace with audience segmentation and brand self-serve tooling — drove 26% user engagement uplift','Result: 35% AOV uplift, 60% cart abandonment drop, 22% checkout conversion lift, 26% engagement increase across 350M+ MAU'].forEach((p,i)=>s.addText('→  '+p,{x:.65,y:1.82+i*.78,w:4.05,h:.68,fontSize:9.5,color:'CCCCCC',wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.22,w:4.4,h:4.7,fill:{color:'EDE8E0'},line:{color:'D8D0C4'}});
s.addText('How It Maps to Spark',{x:5.35,y:1.38,w:4.0,h:.32,fontSize:13,bold:true,color:C.bright});
['Context-aware cart triggers  →  Context-aware Spark clip ranking (season + persona + price)','Real-time intent scoring  →  Spark "Book This Trip" CTA surfaces only when match score is high enough','Dynamic incentivisation at scale  →  Bundle pricing auto-populated per user location + dates','Audience segmentation + ML personalization  →  Spark ranked by past trip history, not generic popularity','22% checkout conversion lift at PhonePe  →  Spark targets 28% new user activation from discovery content'].forEach((m,i)=>s.addText('✓  '+m,{x:5.35,y:1.82+i*.78,w:4.05,h:.68,fontSize:9.5,color:C.ink,wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:.45,y:6.1,w:9.1,h:.6,fill:{color:C.dark},line:{color:C.dark}});
s.addText('"I built the context-aware trigger engine that made PhonePe\'s cart mechanic work at 350M MAU. Spark is the same architecture — applied to the moment a traveler sees a beach and thinks \'I should go there.\'"',{x:.65,y:6.15,w:8.7,h:.5,fontSize:10,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 8 — ROLLOUT
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('ROLLOUT PLAN',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Phased Launch — 6-Month Plan',{x:.45,y:.55,w:9.1,h:.5,fontSize:26,bold:true,color:C.white,fontFace:'Georgia'});
[{phase:'Phase 1',period:'Month 1–2: Pilot',body:'Launch Spark feed for 5% of Agoda app users (Bangkok, Bali, Tokyo markets). Seed with 500 curated clips from existing travel content partners. Instrument content-to-booking funnel, A/B test vs search-only control group.'},{phase:'Phase 2',period:'Month 3–4: Personalise + Scale',body:'Train clip ranking model on pilot engagement data. Add persona-based ranking (solo vs family vs couple clips). Open creator program — travel influencers can submit clips, earn commission on bookings. Scale to 20% of users.'},{phase:'Phase 3',period:'Month 5–6: Full Launch + Supply Signal',body:'Roll out to all Agoda markets. Pipe Spark save and engagement data to Supply team as real-time demand signal. Launch "Sparked by" attribution for creators. Target 35% of new bookings originating from Spark.'}].forEach((p,i)=>{const x=.45+i*3.12;s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.88,h:3.55,fill:{color:C.hero},line:{color:'2A3A5A'},shadow:makeShadow()});s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.88,h:.06,fill:{color:C.bright},line:{color:C.bright}});s.addText(p.phase,{x:x+.14,y:1.42,w:2.6,h:.28,fontSize:9,bold:true,color:C.bright,charSpacing:3});s.addText(p.period,{x:x+.14,y:1.72,w:2.6,h:.38,fontSize:12,bold:true,color:C.white});s.addText(p.body,{x:x+.14,y:2.18,w:2.6,h:2.5,fontSize:9.5,color:'CCCCCC',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.02,w:9.1,h:.92,fill:{color:C.bright},line:{color:C.bright}});
s.addText('What I Need to Build This',{x:.65,y:5.1,w:9,h:.28,fontSize:11,bold:true,color:C.white});
s.addText('500 seed travel clips from content partnerships  ·  Clip ranking ML model (2 engineers)  ·  Bundle auto-population API (flights + hotels)  ·  Creator attribution tracking  ·  A/B test framework vs search control group',{x:.65,y:5.4,w:8.7,h:.46,fontSize:9.5,color:'FFEEEE',wrap:true});
s.addShape(pptx.ShapeType.rect,{x:0,y:6.9,w:10,h:.6,fill:{color:C.hero},line:{color:C.hero}});
s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade',{x:.45,y:6.95,w:9.1,h:.45,fontSize:10,color:'AAAAAA',align:'center'});}

pptx.writeFile({fileName:'assets/agoda-spark-deck.pptx'}).then(()=>console.log('agoda-spark-deck.pptx done')).catch(e=>console.error(e));
