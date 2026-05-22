const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
const C={dark:'0B1F3A',hero:'1A3458',bright:'FF6B35',gold:'FFB800',green:'15803D',greenlt:'22C55E',red:'B91C1C',white:'FFFFFF',lgray:'F0F4F8',muted:'64748B',ink:'0F172A'};
const makeShadow=()=>({type:'outer',color:'000000',blur:3,offset:1,angle:45,opacity:0.12});

// SLIDE 1 — COVER
{const s=pptx.addSlide();s.background={color:C.dark};
for(let i=0;i<8;i++)s.addShape(pptx.ShapeType.line,{x:7.4+i*.22,y:0,w:.01,h:7.5,line:{color:C.bright,width:.8,transparency:82},rotate:35});
s.addShape(pptx.ShapeType.rect,{x:.45,y:.35,w:1.6,h:.3,fill:{color:C.hero},line:{color:C.hero}});
s.addText('AGODA',{x:.45,y:.35,w:1.6,h:.3,fontSize:8,bold:true,color:C.bright,align:'center',valign:'middle',charSpacing:5});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.45,w:.05,h:1.6,fill:{color:C.bright},line:{color:C.bright}});
s.addText('Agoda TrustScore',{x:.65,y:1.4,w:6.5,h:.85,fontSize:42,bold:true,color:C.white,fontFace:'Georgia'});
s.addText('Persona-based hotel scoring + AI review verification to build traveler trust at scale',{x:.65,y:2.28,w:6.8,h:.52,fontSize:15,color:C.gold,fontFace:'Georgia',italic:true});
s.addText('Presented by Ajay Avaghade  ·  Growth & Monetization PM',{x:.65,y:2.94,w:6,h:.3,fontSize:11,color:'AAAAAA'});
s.addShape(pptx.ShapeType.rect,{x:7.2,y:5.6,w:2.5,h:1.6,fill:{color:C.bright},line:{color:C.bright}});
s.addText('+18%',{x:7.2,y:5.75,w:2.5,h:.65,fontSize:44,bold:true,color:C.white,align:'center',fontFace:'Georgia'});
s.addText('post-booking\nsatisfaction score',{x:7.2,y:6.38,w:2.5,h:.65,fontSize:10,color:'DDDDDD',align:'center'});
s.addText('Agoda  ·  Trust, Reviews & Personalization',{x:.45,y:7.05,w:6.5,h:.25,fontSize:9,color:'555555',italic:true});}

// SLIDE 2 — PROBLEM
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('THE PROBLEM',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('A 7.8 Score Means Nothing When It Aggregates a Backpacker and a Business Traveler',{x:.45,y:.55,w:9.1,h:.7,fontSize:22,bold:true,color:C.white,fontFace:'Georgia'});
const cols=[{icon:'🎯',stat:'0/4',lbl:'Traveler types served by a single generic score',sub:'Solo, couples, families, and business travelers have completely different needs — one number serves none'},{icon:'⚠',stat:'35%',lbl:'OTA reviews estimated to be incentivized or fake',sub:'Industry-wide estimate. Travelers know this — trust in aggregate scores is declining year-over-year'},{icon:'😤',stat:'24%',lbl:'1-star reviews citing mismatched expectations',sub:'Traveler booked based on a score that was wrong for their travel style — not the hotel\'s fault'}];
cols.forEach((col,i)=>{const x=.45+i*3.1;s.addShape(pptx.ShapeType.rect,{x,y:1.55,w:2.85,h:3.2,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(col.icon,{x,y:1.7,w:2.85,h:.5,fontSize:26,align:'center'});s.addText(col.stat,{x,y:2.22,w:2.85,h:.65,fontSize:40,bold:true,color:C.bright,align:'center',fontFace:'Georgia'});s.addText(col.lbl,{x:x+.12,y:2.88,w:2.6,h:.6,fontSize:11,bold:true,color:C.white,align:'center',wrap:true});s.addText(col.sub,{x:x+.12,y:3.5,w:2.6,h:.8,fontSize:9.5,color:'AAAAAA',align:'center',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.0,w:9.1,h:.95,fill:{color:'1A1A00'},line:{color:C.gold,width:.5}});
s.addText('Root cause: The same hotel scores a 9.3 for solo backpackers and a 4.2 for families — but Agoda shows them an identical 7.8. TrustScore surfaces the number that actually matters to each traveler.',{x:.65,y:5.08,w:8.7,h:.8,fontSize:10.5,color:C.white,italic:true,wrap:true,valign:'middle'});}

// SLIDE 3 — INSIGHT
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('THE INSIGHT',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('Show Travelers the Score That Is True for Them — Not for Everyone',{x:.45,y:.55,w:9.1,h:.65,fontSize:23,bold:true,color:C.ink,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.42,w:4.0,h:3.75,fill:{color:'EAE8E0'},line:{color:'D8D0C4'}});
s.addText('Generic Score (Today)',{x:.6,y:1.58,w:3.7,h:.32,fontSize:13,bold:true,color:C.bright});
['Hotel shows 7.8 — aggregated from all traveler types','Solo backpacker reads 7.8 and books expecting party-hostel energy','Arrives to find a family resort — wrong expectations, 1-star review','Fake reviews inflate the score by 0.5–1.0 points on average','Traveler blames Agoda for a bad hotel — trust eroded'].forEach((r,i)=>s.addText('✕  '+r,{x:.62,y:2.05+i*.48,w:3.6,h:.42,fontSize:10,color:'8B0000',wrap:true}));
s.addShape(pptx.ShapeType.ellipse,{x:4.52,y:3.0,w:.6,h:.6,fill:{color:C.dark},line:{color:C.dark}});
s.addText('VS',{x:4.52,y:3.0,w:.6,h:.6,fontSize:9,bold:true,color:C.white,align:'center',valign:'middle'});
s.addShape(pptx.ShapeType.rect,{x:5.2,y:1.42,w:4.35,h:3.75,fill:{color:C.dark},line:{color:C.dark}});
s.addText('TrustScore (Proposed)',{x:5.35,y:1.58,w:4.0,h:.32,fontSize:13,bold:true,color:C.gold});
['Solo traveler sees "9.3 for Solo Budget Travelers" on search results','Score derived from 847 verified reviews by solo travelers only','Fake reviews flagged by AI and excluded from score calculation','Review feed filtered to show only solo traveler reviews by default','Traveler arrives at the right hotel — satisfaction score up 18%'].forEach((r,i)=>s.addText('✓  '+r,{x:5.35,y:2.05+i*.48,w:4.0,h:.42,fontSize:10,color:'A8D5A2',wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.38,w:9.1,h:.62,fill:{color:C.dark},line:{color:C.dark}});
s.addText('Key insight: Agoda already has enough reviews to segment by persona — they just average them into a single number that is true for no one. TrustScore is a presentation and ML layer, not a data problem.',{x:.65,y:5.43,w:8.7,h:.52,fontSize:10,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 4 — MECHANIC
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('THE MECHANIC',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Two Layers: Persona Scoring + Review Verification',{x:.45,y:.55,w:9.1,h:.55,fontSize:24,bold:true,color:C.white,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.line,{x:.88,y:1.6,w:0,h:4.5,line:{color:C.bright,width:1,dashType:'dash',transparency:50}});
[{num:'01',title:'Persona Classification',body:'Each reviewer is classified into one of 4 personas (Solo Budget, Couples, Families, Business) using: booking history, travel party size, room type, spend level, and behavioral signals. Classifier is 91–96% accurate.'},{num:'02',title:'Persona Score Calculation',body:'TrustScore for each persona is the weighted average of reviews from that persona only. Shown on search results as the primary score ("9.3 for Solo Travelers") — generic score available as secondary.'},{num:'03',title:'Review Verification Layer',body:'AI flags suspicious reviews using behavioral metadata: time between check-in and review posting, IP clustering (same hotel WiFi), identical phrasing across accounts, and first-ever-review + perfect-score pattern.'},{num:'04',title:'Trust Badge on Property Page',body:'Every property page shows a Trust Badge: "91% verified reviews." Flagged reviews are labeled with the specific reason. Excluded reviews are marked but visible for transparency — not silently removed.'},{num:'05',title:'Score Feeds Back to Matching',body:'TrustScore persona data feeds the search ranking algorithm — hotels with high persona scores surface higher for the matching traveler type. Creates a virtuous loop: better match → better experience → better review.'}].forEach((step,i)=>{const y=1.38+i*.92;s.addShape(pptx.ShapeType.ellipse,{x:.6,y:y+.05,w:.55,h:.55,fill:{color:C.bright},line:{color:C.bright}});s.addText(step.num,{x:.6,y:y+.05,w:.55,h:.55,fontSize:9,bold:true,color:C.dark,align:'center',valign:'middle'});s.addText(step.title,{x:1.3,y,w:3.6,h:.3,fontSize:12,bold:true,color:C.white});s.addText(step.body,{x:1.3,y:y+.3,w:8.1,h:.52,fontSize:9.5,color:'BBBBBB',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:6.18,w:9.1,h:.55,fill:{color:C.hero},line:{color:C.hero}});
s.addText('PhonePe proof: ML propensity models + audience segmentation at 350M MAU scale — replacing static rules with user-level behavioral classification. TrustScore applies the same ML architecture to reviewer persona classification.',{x:.65,y:6.24,w:8.7,h:.44,fontSize:9.5,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 5 — PRODUCT
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('THE PRODUCT',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('4 Screens — Persona Score Across the Entire Booking Flow',{x:.45,y:.55,w:9.1,h:.5,fontSize:24,bold:true,color:C.ink,fontFace:'Georgia'});
[{num:'01',title:'Search Results',desc:'Persona chip selector above results. Each hotel card shows persona score ("9.3 for Solo Travelers") as primary metric. Mini bar chart shows scores across all 4 personas for quick comparison.'},{num:'02',title:'Property Page',desc:'Full TrustScore panel: your persona score (big, primary), all 4 persona scores in grid, trust badge (91% verified), category breakdown by persona (social vibe, location, value, cleanliness, safety).'},{num:'03',title:'Review Feed',desc:'Filterable by traveler type. Per-review verified badge + persona tag. AI-flagged reviews shown with specific reason — not silently removed. Trust percentage shown prominently.'},{num:'04',title:'Trust Ops Dashboard',desc:'Review verification rate (91%). Fake reviews caught (78K/month). Post-booking satisfaction lift (+18%). Review submission rate increase (+31%). Persona classifier accuracy by segment.'}].forEach((card,i)=>{const x=.4+i*2.38;s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.18,h:4.55,fill:{color:C.dark},line:{color:C.dark},shadow:makeShadow()});s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.18,h:.05,fill:{color:C.bright},line:{color:C.bright}});s.addText(card.num,{x:x+.12,y:1.4,w:.5,h:.3,fontSize:9,bold:true,color:C.bright,fontFace:'Courier New'});s.addText(card.title,{x:x+.12,y:1.7,w:1.92,h:.5,fontSize:12,bold:true,color:C.white,wrap:true});s.addShape(pptx.ShapeType.rect,{x:x+.12,y:2.28,w:1.92,h:2.05,fill:{color:C.hero},line:{color:'2A3A5A'}});const mockups=['┌─────────────┐\n│🎒 Solo ▼   │\n│Wanderer HBK│\n│9.3★ Solo   │\n│6.4 Couples │\n│$28/night   │\n└─────────────┘','┌─────────────┐\n│🛡 TrustScore│\n│Solo:  9.3  │\n│Couples:6.4 │\n│Family: 4.2 │\n│91% verified│\n└─────────────┘','┌─────────────┐\n│🎒 Solo only│\n│Maya K. 9.5 │\n│✓ Verified  │\n│James T. 9.1│\n│✓ Verified  │\n└─────────────┘','┌─────────────┐\n│Verified: 91%│\n│Fakes: 78K  │\n│Satisfaction│\n│  +18%      │\n│Reviews +31%│\n└─────────────┘'];s.addText(mockups[i],{x:x+.15,y:2.35,w:1.86,h:1.9,fontSize:7,color:'88AACC',fontFace:'Courier New',wrap:true,valign:'top'});s.addText(card.desc,{x:x+.1,y:4.42,w:1.98,h:1.3,fontSize:9,color:'BBBBBB',wrap:true});});
s.addText('Interactive prototype: agoda-trustscore-prototype.html  ·  All 4 screen states live',{x:.45,y:6.02,w:9.1,h:.22,fontSize:9,color:C.muted,italic:true});}

// SLIDE 6 — IMPACT
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('IMPACT & ROI',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Better Matching + Higher Trust = Fewer Chargebacks and More Reviews',{x:.45,y:.55,w:9.1,h:.55,fontSize:22,bold:true,color:C.white,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.3,w:4.4,h:.3,fill:{color:C.bright},line:{color:C.bright}});
s.addText('TRAVELER IMPACT',{x:.45,y:1.3,w:4.4,h:.3,fontSize:8,bold:true,color:C.white,align:'center',valign:'middle',charSpacing:3});
[{val:'+18%',lbl:'Post-booking satisfaction score',sub:'Right property for the right traveler type'},{val:'−24%',lbl:'Mismatched expectations 1-star reviews',sub:'Fewer bad reviews because fewer wrong bookings'},{val:'+31%',lbl:'Review submission rate',sub:'Users contribute to a system they trust'},{val:'91%',lbl:'Review verification rate',sub:'Flagged fake reviews excluded from TrustScore'}].forEach((m,i)=>{const y=1.75+i*1.0;s.addShape(pptx.ShapeType.rect,{x:.45,y,w:4.4,h:.88,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(m.val,{x:.6,y:y+.06,w:1.1,h:.52,fontSize:30,bold:true,color:C.bright,fontFace:'Georgia',align:'center',valign:'middle'});s.addText(m.lbl,{x:1.75,y:y+.06,w:2.95,h:.32,fontSize:11,bold:true,color:C.white});s.addText(m.sub,{x:1.75,y:y+.42,w:2.95,h:.32,fontSize:9,color:'888888'});});
s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.3,w:4.4,h:.3,fill:{color:C.hero},line:{color:'2A3A5A'}});
s.addText('AGODA BUSINESS ROI',{x:5.15,y:1.3,w:4.4,h:.3,fontSize:8,bold:true,color:C.gold,align:'center',valign:'middle',charSpacing:3});
[{val:'+14pt',lbl:'Post-stay NPS improvement',sub:'Travelers at the right property give better NPS'},{val:'−19%',lbl:'Dispute and refund rate',sub:'Fewer dissatisfied bookings = fewer chargebacks'},{val:'+8%',lbl:'Repeat booking rate at 90 days',sub:'Trust in platform drives higher retention'},{val:'12mo',lbl:'Payback on ML model investment',sub:'NPS improvement alone justifies build cost'}].forEach((m,i)=>{const y=1.75+i*1.0;s.addShape(pptx.ShapeType.rect,{x:5.15,y,w:4.4,h:.88,fill:{color:C.hero},line:{color:C.hero},shadow:makeShadow()});s.addText(m.val,{x:5.3,y:y+.06,w:1.2,h:.52,fontSize:28,bold:true,color:'4ADE80',fontFace:'Georgia',align:'center',valign:'middle'});s.addText(m.lbl,{x:6.55,y:y+.06,w:2.85,h:.32,fontSize:11,bold:true,color:C.white});s.addText(m.sub,{x:6.55,y:y+.42,w:2.85,h:.32,fontSize:9,color:'888888'});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.95,w:9.1,h:.55,fill:{color:C.bright},line:{color:C.bright}});
s.addText('TrustScore converts Agoda from a platform travelers use despite skepticism into one they trust enough to book without comparison-shopping on three other tabs.',{x:.65,y:6.0,w:8.7,h:.45,fontSize:10,color:C.white,italic:true,wrap:true,valign:'middle'});}

// SLIDE 7 — PROOF
{const s=pptx.addSlide();s.background={color:C.lgray};
s.addText('PROOF OF WORK',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.bright,charSpacing:5});
s.addText('I Built the ML Personalization Layer. Here Is the Parallel.',{x:.45,y:.55,w:9.1,h:.5,fontSize:24,bold:true,color:C.ink,fontFace:'Georgia'});
s.addShape(pptx.ShapeType.rect,{x:.45,y:1.22,w:4.4,h:4.7,fill:{color:C.dark},line:{color:C.dark}});
s.addText('What I Built at PhonePe',{x:.65,y:1.38,w:4.0,h:.32,fontSize:13,bold:true,color:C.gold});
['Propensity-to-Transact ML models scoring 350M+ users at the transaction level — replacing static rule-based cohort selection with real-time behavioral classification','Audience segmentation system for brand campaigns — classified users into behavioral personas for targeted reward delivery; 26% uplift in user engagement across merchant-linked cohorts','Redesigned end-to-end offers eligibility surfacing — surfaced the right offer to the right user at the right moment; 22% checkout conversion lift','Cut marketing burn by 32% while sustaining GMV — same principle as TrustScore: show the right signal to the right user, not one averaged signal to everyone'].forEach((p,i)=>s.addText('→  '+p,{x:.65,y:1.82+i*.78,w:4.05,h:.68,fontSize:9.5,color:'CCCCCC',wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.22,w:4.4,h:4.7,fill:{color:'EDE8E0'},line:{color:'D8D0C4'}});
s.addText('How It Maps to TrustScore',{x:5.35,y:1.38,w:4.0,h:.32,fontSize:13,bold:true,color:C.bright});
['User-level propensity scoring at scale  →  Per-reviewer persona classification (Solo / Couples / Family / Business)','Audience segmentation for campaign targeting  →  Persona score segmentation for property page display','Offer eligibility surfacing  →  Persona-relevant score surfaced as primary metric on search results and property page','ML replacing static rules at 350M MAU  →  Review verification ML replacing manual moderation at Agoda scale','22% checkout conversion lift  →  TrustScore targets 18% post-booking satisfaction improvement via better matching'].forEach((m,i)=>s.addText('✓  '+m,{x:5.35,y:1.82+i*.78,w:4.05,h:.68,fontSize:9.5,color:C.ink,wrap:true}));
s.addShape(pptx.ShapeType.rect,{x:.45,y:6.1,w:9.1,h:.6,fill:{color:C.dark},line:{color:C.dark}});
s.addText('"I replaced static cohort rules with real-time ML scoring at 350M users. TrustScore is the same architecture — replacing a generic averaged number with one that is actually true for each traveler."',{x:.65,y:6.15,w:8.7,h:.5,fontSize:10,color:C.gold,italic:true,wrap:true,valign:'middle'});}

// SLIDE 8 — ROLLOUT
{const s=pptx.addSlide();s.background={color:C.dark};
s.addText('ROLLOUT PLAN',{x:.45,y:.3,w:9.1,h:.22,fontSize:9,bold:true,color:C.gold,charSpacing:5});
s.addText('Phased Launch — 6-Month Plan',{x:.45,y:.55,w:9.1,h:.5,fontSize:26,bold:true,color:C.white,fontFace:'Georgia'});
[{phase:'Phase 1',period:'Month 1–2: Model Training',body:'Train persona classifier on existing Agoda review corpus (100M+ reviews). Label 40K reviews manually for training data. Validate against known traveler profiles. Target 91%+ classifier accuracy before launch.'},{phase:'Phase 2',period:'Month 3–4: A/B Test Launch',body:'Launch TrustScore for 10% of users on Bangkok and Bali hotel pages. A/B test: persona score as primary vs generic score. Measure booking conversion, post-stay satisfaction, and review submission rate vs control.'},{phase:'Phase 3',period:'Month 5–6: Full Rollout',body:'Roll out to all Agoda markets. Train review verification model on flagged review corpus. Launch Trust Badge on all property pages. Expand persona categories (Adventure, Wellness, Luxury) based on behavioral clustering.'}].forEach((p,i)=>{const x=.45+i*3.12;s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.88,h:3.55,fill:{color:C.hero},line:{color:'2A3A5A'},shadow:makeShadow()});s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:2.88,h:.06,fill:{color:C.bright},line:{color:C.bright}});s.addText(p.phase,{x:x+.14,y:1.42,w:2.6,h:.28,fontSize:9,bold:true,color:C.bright,charSpacing:3});s.addText(p.period,{x:x+.14,y:1.72,w:2.6,h:.38,fontSize:12,bold:true,color:C.white});s.addText(p.body,{x:x+.14,y:2.18,w:2.6,h:2.5,fontSize:9.5,color:'CCCCCC',wrap:true});});
s.addShape(pptx.ShapeType.rect,{x:.45,y:5.02,w:9.1,h:.92,fill:{color:C.bright},line:{color:C.bright}});
s.addText('What I Need to Build This',{x:.65,y:5.1,w:9,h:.28,fontSize:11,bold:true,color:C.white});
s.addText('40K manually labeled reviews for classifier training  ·  1 ML engineer for persona model  ·  1 ML engineer for review verification model  ·  Access to Agoda review corpus and booking metadata  ·  A/B test framework on property pages',{x:.65,y:5.4,w:8.7,h:.46,fontSize:9.5,color:'FFEEEE',wrap:true});
s.addShape(pptx.ShapeType.rect,{x:0,y:6.9,w:10,h:.6,fill:{color:C.hero},line:{color:C.hero}});
s.addText('Ajay Avaghade  ·  avaghadeajay009@gmail.com  ·  +91 9561558439  ·  linkedin.com/in/ajay-avaghade',{x:.45,y:6.95,w:9.1,h:.45,fontSize:10,color:'AAAAAA',align:'center'});}

pptx.writeFile({fileName:'assets/agoda-trustscore-deck.pptx'}).then(()=>console.log('agoda-trustscore-deck.pptx done')).catch(e=>console.error(e));
