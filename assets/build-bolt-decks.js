const PptxGenJS = require('pptxgenjs');

const C = {
  dark:   '0D1A10',
  hero:   '162815',
  green:  '34D186',
  gold:   'F59E0B',
  purple: '8B5CF6',
  red:    'EF4444',
  white:  'FFFFFF',
  lgray:  'F0FFF4',
  muted:  '687A70',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

// ═══════════════════════════════════════════════════════════════
// DECK 1: RIDEGUIDE
// ═══════════════════════════════════════════════════════════════
function buildRideGuide(){
  const pptx = new PptxGenJS(); pptx.layout='LAYOUT_16x9';

  // SLIDE 1
  {const s=pptx.addSlide();s.background={color:C.dark};
  for(let i=0;i<18;i++) s.addShape(pptx.ShapeType.line,{x:-0.5+i*0.7,y:0,w:4,h:0,line:{color:'162815',width:18},rotate:35});
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
  s.addText('BOLT · MICROMOBILITY COMPLIANCE · Case Study 01',{x:0.35,y:0.28,w:7,h:0.22,fontSize:8,bold:true,color:C.muted,charSpacing:1.5,fontFace:'Calibri'});
  s.addText('RideGuide',{x:0.35,y:1.1,w:7,h:1.3,fontSize:72,bold:true,color:C.white,fontFace:'Calibri'});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:2.3,w:2.8,h:0.06,fill:{color:C.green}});
  s.addText('Compliance UX That Riders Accept, Not Resist',{x:0.35,y:2.5,w:8.5,h:0.45,fontSize:18,color:C.gold,fontFace:'Calibri'});
  s.addText('Ajay Avaghade · Product Manager',{x:0.35,y:3.12,w:5,h:0.3,fontSize:11,color:C.muted,fontFace:'Calibri'});
  s.addShape(pptx.ShapeType.rect,{x:6.8,y:5.8,w:3.0,h:1.45,fill:{color:C.hero},line:{color:C.green,width:1.5},shadow:makeShadow()});
  s.addText('−40%',{x:6.8,y:5.85,w:3.0,h:0.55,fontSize:38,bold:true,color:C.gold,align:'center',fontFace:'Calibri'});
  s.addText('end-ride drop-off rate',{x:6.8,y:6.38,w:3.0,h:0.25,fontSize:11,color:C.white,align:'center',fontFace:'Calibri'});
  s.addText('vs hard enforcement baseline',{x:6.8,y:6.63,w:3.0,h:0.2,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});}

  // SLIDE 2 — PROBLEM
  {const s=pptx.addSlide();s.background={color:C.dark};
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
  s.addText('THE PROBLEM',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
  s.addText('Hard Compliance Rules Create the Worst Rider Moments.\nJarring Cuts. Confusing Rejections. 40% Drop-Off at End-Ride.',{x:0.35,y:0.62,w:9.3,h:1.1,fontSize:24,bold:true,color:C.white,fontFace:'Calibri',breakLine:true});
  const cols=[{icon:'⚡',stat:'3×',label:'Complaint rate at zone boundaries',sub:'Instant throttle cuts = jarring experience = NPS damage that shows up in churn'},{icon:'🅿️',stat:'40%',label:'End-ride drop-off rate',sub:'Riders abandon the parking validator because they cannot find valid zones'},{icon:'🛡️',stat:'38%',label:'Onboarding age verification drop-off',sub:'Document upload abandonment — riders give up before their first ride'}];
  cols.forEach((c,i)=>{const x=0.35+i*3.2;
    s.addShape(pptx.ShapeType.rect,{x,y:2.0,w:3.0,h:2.75,fill:{color:C.hero},line:{color:C.green,width:1},shadow:makeShadow()});
    s.addText(c.icon,{x:x+0.15,y:2.1,w:0.6,h:0.5,fontSize:22});
    s.addText(c.stat,{x:x+0.15,y:2.58,w:2.7,h:0.7,fontSize:44,bold:true,color:C.gold,fontFace:'Calibri'});
    s.addText(c.label,{x:x+0.15,y:3.28,w:2.7,h:0.38,fontSize:11,bold:true,color:C.white,fontFace:'Calibri'});
    s.addText(c.sub,{x:x+0.15,y:3.68,w:2.7,h:0.85,fontSize:10,color:C.muted,fontFace:'Calibri'});});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.0,w:9.3,h:1.15,fill:{color:'1F2E22'},line:{color:C.green,width:1}});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.0,w:0.05,h:1.15,fill:{color:C.green}});
  s.addText('Root cause: Compliance is implemented as hard technical rules, not UX design. The geofence boundary, the parking validator error, and the document upload form are all compliance-correct but UX-terrible. The worst moments in a Bolt ride are all compliance moments — and none of them need to be.',{x:0.55,y:5.07,w:9.0,h:1.0,fontSize:11,color:C.white,fontFace:'Calibri'});}

  // SLIDE 3 — INSIGHT
  {const s=pptx.addSlide();s.background={color:C.lgray};
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
  s.addText('THE INSIGHT',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
  s.addText('Every Compliance Failure is a UX Failure First.\nFix the Experience — the Compliance Follows.',{x:0.35,y:0.62,w:9.3,h:1.0,fontSize:27,bold:true,color:C.dark,fontFace:'Calibri',breakLine:true});
  s.addShape(pptx.ShapeType.ellipse,{x:4.7,y:2.0,w:0.6,h:0.6,fill:{color:C.green}});
  s.addText('VS',{x:4.7,y:2.05,w:0.6,h:0.5,fontSize:11,bold:true,color:'000000',align:'center',fontFace:'Calibri'});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.82,w:4.2,h:3.45,fill:{color:'FDE8E8'},line:{color:'F59090',width:1}});
  s.addText('❌  Today',{x:0.5,y:1.92,w:3.9,h:0.35,fontSize:13,bold:true,color:'B91C1C',fontFace:'Calibri'});
  ['Throttle cuts instantly at zone boundary — no warning','Parking validator shows error — no nearest spot shown','Age verification requires document upload — 38% quit','Compliance events = worst NPS moments in the ride','City authority sees violation counts — no context'].forEach((t,i)=>s.addText('•  '+t,{x:0.5,y:2.36+i*0.52,w:3.9,h:0.45,fontSize:11,color:'374151',fontFace:'Calibri'}));
  s.addShape(pptx.ShapeType.rect,{x:5.45,y:1.82,w:4.2,h:3.45,fill:{color:'ECFDF5'},line:{color:'6EE7B7',width:1}});
  s.addText('✅  RideGuide',{x:5.6,y:1.92,w:3.9,h:0.35,fontSize:13,bold:true,color:'065F46',fontFace:'Calibri'});
  ['200m warning + 12-second soft deceleration curve','Parking validator: valid ✓ or nearest spot in 40m shown','Selfie liveness check — 8 seconds, instant approval','Compliance events become neutral or positive moments','Rider compliance data exported for city tender reports'].forEach((t,i)=>s.addText('•  '+t,{x:5.6,y:2.36+i*0.52,w:3.9,h:0.45,fontSize:11,color:'065F46',fontFace:'Calibri'}));
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.42,w:9.3,h:1.18,fill:{color:'FFF3CD'},line:{color:C.gold,width:1.5}});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.42,w:0.05,h:1.18,fill:{color:C.gold}});
  s.addText('Key insight: PhonePe cart incentivization used context-aware triggers (cart value x user intent x time) to reduce 60% cart abandonment. RideGuide applies the same contextual trigger architecture to compliance — trajectory + speed + distance = proactive intervention. Same architecture, zero new principles.',{x:0.55,y:5.49,w:9.0,h:1.0,fontSize:10.5,color:'1F1109',fontFace:'Calibri'});}

  // SLIDE 4 — MECHANIC
  {const s=pptx.addSlide();s.background={color:C.hero};
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.gold}});
  s.addText('THE MECHANIC',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.gold,charSpacing:2,fontFace:'Calibri'});
  s.addText('Trajectory Prediction x Soft Deceleration x Smart Parking x Selfie Verification',{x:0.35,y:0.62,w:9.3,h:0.65,fontSize:22,bold:true,color:C.white,fontFace:'Calibri'});
  const steps=[
    {n:'1',t:'Trajectory Prediction',b:'GPS + speed + heading sampled every 500ms. ZoneIQ ML model calculates P(zone entry) in next 30 seconds. At confidence >= 75%, trigger soft alert. Works 30 seconds before the boundary — not at it.'},
    {n:'2',t:'Soft Alert + Graduated Decel',b:'Audio cue + haptic at 200m. 12-second deceleration curve starts 20 seconds before zone entry. Jerk limited to 1.2 m/s3 — smoother than a car braking gently. Rider arrives at zone boundary at compliant speed.'},
    {n:'3',t:'Parking Validator',b:'GPS cross-referenced with city GeoJSON (updated nightly from LaunchKit). If valid zone: green confirmation, ride ends. If invalid: nearest 3 valid spots ranked by distance. "Navigate" button turns on turn-by-turn audio cues.'},
    {n:'4',t:'Selfie Liveness Verification',b:'Onfido or Veriff API: 8-second selfie check. ML model detects age band with 94.2% accuracy. No data stored beyond 48h (GDPR). Same legal compliance as document upload — zero friction for the rider.'},
    {n:'5',t:'Compliance Logging',b:'Every zone event (entry speed, decel curve used, parking outcome) logged per session. Aggregated in CityScore (Case Study 3). Exportable for city authority tender submissions. Compliance UX and compliance data are one system.'},
  ];
  steps.forEach((step,i)=>{const y=1.48+i*1.04;
    s.addShape(pptx.ShapeType.ellipse,{x:0.35,y:y+0.02,w:0.45,h:0.45,fill:{color:C.green}});
    s.addText(step.n,{x:0.35,y:y+0.04,w:0.45,h:0.4,fontSize:14,bold:true,color:'000000',align:'center',fontFace:'Calibri'});
    if(i<4) s.addShape(pptx.ShapeType.line,{x:0.575,y:y+0.47,w:0,h:0.57,line:{color:C.muted,width:1,dashType:'dash'}});
    s.addText(step.t,{x:0.95,y:y,w:9.0,h:0.28,fontSize:13,bold:true,color:C.gold,fontFace:'Calibri'});
    s.addText(step.b,{x:0.95,y:y+0.28,w:8.9,h:0.65,fontSize:10.5,color:C.white,fontFace:'Calibri'});});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.85,w:9.3,h:0.42,fill:{color:C.dark},line:{color:C.green,width:0.8}});
  s.addText('PhonePe proof: Dynamic cart incentivization (context-aware triggers: cart value x intent x time) drove 35% AOV uplift, 60% cart abandonment reduction. RideGuide = same trigger architecture applied to physical compliance events.',{x:0.5,y:6.88,w:9.1,h:0.36,fontSize:9.5,color:C.muted,fontFace:'Calibri'});}

  // SLIDE 5 — PRODUCT
  {const s=pptx.addSlide();s.background={color:C.lgray};
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
  s.addText('THE PRODUCT',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
  s.addText('4 Screen States — Ride to Verify',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:26,bold:true,color:C.dark,fontFace:'Calibri'});
  const cards=[
    {n:'01',t:'Active Ride',b:'Map with soft geofence trajectory arc. RideGuide Active indicator. Zone approach warning card (tap for details). Compliance status checklist (age, helmet, upcoming zone). End ride CTA.'},
    {n:'02',t:'Zone Approach Alert',b:'Countdown ring (30 seconds to entry). Speed change display (22 km/h → 25 km/h limit). Decel curve visual. Hard cut vs soft ramp comparison. Continue riding CTA.'},
    {n:'03',t:'Parking Validator',b:'Map with valid zone ✓ confirmation + nearest spots. Trip summary: distance, duration, cost, compliance 100%. "Lock Scooter & End Ride" CTA.'},
    {n:'04',t:'Age Verification',b:'4-step progress (phone, payment, age, start). Selfie liveness frame with scan animation. "Why selfie not doc?" comparison. Document upload fallback available.'},
  ];
  const mocks=[
    ['╔══════════╗','🟢 RideGuide On','─────────────','19 km/h  4:12  €1.40','⚠ Zone in 200m →','✓ Age  ✓ Helmet','[End Ride]'],
    ['╔══════════╗','🧠 Zone in 30s','─────────────','22 → ≤25 km/h','████████░░ 12s ramp','Hard❌ vs Soft✓','[Continue Riding]'],
    ['╔══════════╗','✓ Valid Parking!','─────────────','Zone 4B · 0m away','1.4km · 7m32s','✓ 100% compliant','€2.45 [End Ride]'],
    ['╔══════════╗','Age Verification','─────────────','Step 3/4','[🤳 Selfie frame]','8 sec · GDPR safe','[Verify Age]'],
  ];
  cards.forEach((c,i)=>{const x=0.35+i*2.4;
    s.addShape(pptx.ShapeType.rect,{x,y:1.35,w:2.25,h:5.65,fill:{color:C.white},line:{color:'C8E8D8',width:1},shadow:makeShadow()});
    s.addShape(pptx.ShapeType.rect,{x,y:1.35,w:2.25,h:0.05,fill:{color:C.green}});
    s.addText(c.n,{x:x+0.12,y:1.45,w:1.0,h:0.28,fontSize:9,bold:true,color:C.green,fontFace:'Calibri',charSpacing:1});
    s.addText(c.t,{x:x+0.12,y:1.72,w:2.0,h:0.52,fontSize:12,bold:true,color:C.dark,fontFace:'Calibri'});
    s.addShape(pptx.ShapeType.rect,{x:x+0.12,y:2.32,w:2.0,h:2.2,fill:{color:'0D1A10'},line:{color:'1E3422',width:1}});
    mocks[i].forEach((line,li)=>s.addText(line,{x:x+0.15,y:2.35+li*0.275,w:1.95,h:0.26,fontSize:7.5,color:li===0?C.green:'86EFAC',fontFace:'Courier New',bold:li===0}));
    s.addText(c.b,{x:x+0.12,y:4.62,w:2.0,h:2.25,fontSize:9.5,color:'4B6B5A',fontFace:'Calibri'});});
  s.addText('Prototype: bolt-rideguide-prototype.html · All 4 states interactive · Click-based navigation',{x:0.35,y:7.12,w:9.3,h:0.22,fontSize:8.5,color:C.muted,italic:true,fontFace:'Calibri'});}

  // SLIDE 6 — IMPACT
  {const s=pptx.addSlide();s.background={color:C.dark};
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
  s.addText('IMPACT & ROI',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
  s.addText('PhonePe Context-Aware Trigger Architecture — Applied to Compliance',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:22,bold:true,color:C.white,fontFace:'Calibri'});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.35,w:4.5,h:0.38,fill:{color:C.green}});
  s.addText('RIDER EXPERIENCE IMPACT',{x:0.35,y:1.37,w:4.5,h:0.34,fontSize:10,bold:true,color:'000000',align:'center',charSpacing:1.5,fontFace:'Calibri'});
  s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.35,w:4.5,h:0.38,fill:{color:'2C1A00'},line:{color:C.gold,width:1}});
  s.addText('BOLT / CITY AUTHORITY ROI',{x:5.15,y:1.37,w:4.5,h:0.34,fontSize:10,bold:true,color:C.gold,align:'center',charSpacing:1.5,fontFace:'Calibri'});
  const lm=[{v:'−40%',l:'End-ride drop-off rate',s:'Parking validator + nearest spot map replaces error state with actionable UX'},{v:'+35%',l:'Onboarding completion rate',s:'Selfie liveness (8s) vs document upload (38% abandon) — same legal compliance'},{v:'73%',l:'Fewer sudden brake events',s:'12-second soft ramp vs instant throttle cut — measured against hard geofence baseline'},{v:'+28%',l:'Invalid-zone-to-valid-spot rate',s:'Riders who complete the walk to nearest valid spot when shown navigation'}];
  const rm=[{v:'+22%',l:'NPS score (compliance moments)',s:'Compliance interactions become neutral or positive — no longer the worst moment in the ride'},{v:'91%',l:'Compliance score (Warsaw pilot)',s:'RideGuide-enabled cities achieve higher city authority compliance scores → stronger tenders'},{v:'€0',l:'Fine exposure per compliant ride',s:'Proactive compliance vs reactive enforcement eliminates fine liability at source'},{v:'+35%',l:'Tender win rate (est.)',s:'Cities see live compliance data via CityScore — Bolt demonstrates compliance, not claims it'}];
  lm.forEach((m,i)=>{const y=1.9+i*1.25;
    s.addShape(pptx.ShapeType.rect,{x:0.35,y,w:4.5,h:1.15,fill:{color:C.hero},line:{color:'1E3422',width:1}});
    s.addText(m.v,{x:0.5,y:y+0.08,w:1.6,h:0.55,fontSize:34,bold:true,color:C.gold,fontFace:'Calibri'});
    s.addText(m.l,{x:0.5,y:y+0.6,w:4.1,h:0.25,fontSize:12,bold:true,color:C.white,fontFace:'Calibri'});
    s.addText(m.s,{x:0.5,y:y+0.83,w:4.1,h:0.28,fontSize:9,color:C.muted,fontFace:'Calibri'});});
  rm.forEach((m,i)=>{const y=1.9+i*1.25;
    s.addShape(pptx.ShapeType.rect,{x:5.15,y,w:4.5,h:1.15,fill:{color:'1F1200'},line:{color:'4A3300',width:1}});
    s.addText(m.v,{x:5.3,y:y+0.08,w:1.6,h:0.55,fontSize:34,bold:true,color:C.gold,fontFace:'Calibri'});
    s.addText(m.l,{x:5.3,y:y+0.6,w:4.1,h:0.25,fontSize:12,bold:true,color:C.white,fontFace:'Calibri'});
    s.addText(m.s,{x:5.3,y:y+0.83,w:4.1,h:0.28,fontSize:9,color:C.muted,fontFace:'Calibri'});});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.9,w:9.3,h:0.38,fill:{color:'100F08'},line:{color:C.green,width:1}});
  s.addText('RideGuide makes compliance invisible to compliant riders and unavoidable for non-compliant ones. The same geofence, radically different experience.',{x:0.5,y:6.93,w:9.0,h:0.32,fontSize:9.5,color:C.white,fontFace:'Calibri'});}

  // SLIDE 7 — PROOF
  {const s=pptx.addSlide();s.background={color:C.lgray};
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
  s.addText('PROOF OF WORK',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
  s.addText('I Built the Underlying Architecture at PhonePe.',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:26,bold:true,color:C.dark,fontFace:'Calibri'});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.35,w:4.5,h:4.6,fill:{color:C.dark}});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.35,w:4.5,h:0.32,fill:{color:C.green}});
  s.addText('What I shipped at PhonePe',{x:0.5,y:1.37,w:4.2,h:0.28,fontSize:10,bold:true,color:'000000',fontFace:'Calibri'});
  [{h:'Dynamic Cart Incentivization (Pincode)',b:'Context-aware triggers: cart value x user intent x time-of-day. 35% AOV uplift, 60% cart abandonment reduction, 20% 30-day retention improvement. The trigger architecture is identical to RideGuide.'},{h:'Propensity-to-Transact ML (350M MAU)',b:'Real-time per-user scoring replacing manual cohort targeting. 32% marketing burn reduction. ML scoring applied to physical trajectory is the same model pattern applied to user purchase probability.'},{h:'Offers Placement + Eligibility Redesign',b:'Redesigned compliance-like rules (offer eligibility, terms) into frictionless UX. 22% checkout conversion lift. Same design principle: make required checks invisible.'},{h:'Unified Ops Intelligence Layer',b:'Replaced 6 siloed workflows with single console. TAT 2d to 4h, -68% escalation tickets. CityScore (Case Study 3) is the compliance version of this ops intelligence platform.'}].forEach((item,i)=>{const y=1.82+i*1.02;
    s.addShape(pptx.ShapeType.ellipse,{x:0.48,y:y+0.05,w:0.22,h:0.22,fill:{color:C.green}});
    s.addText(item.h,{x:0.82,y:y,w:3.85,h:0.28,fontSize:11,bold:true,color:C.gold,fontFace:'Calibri'});
    s.addText(item.b,{x:0.82,y:y+0.28,w:3.85,h:0.62,fontSize:9.5,color:C.white,fontFace:'Calibri'});});
  s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.35,w:4.5,h:4.6,fill:{color:'F0FAF4'}});
  s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.35,w:4.5,h:0.32,fill:{color:'A7F3D0'}});
  s.addText('How it maps to RideGuide',{x:5.3,y:1.37,w:4.2,h:0.28,fontSize:10,bold:true,color:C.dark,fontFace:'Calibri'});
  ['-> RideGuide zone alert trigger: trajectory x speed x distance x time. Same context-aware architecture, physical compliance instead of cart conversion.','-> ZoneIQ trajectory ML: same per-user real-time scoring, applied to movement prediction instead of purchase propensity.','-> RideGuide parking validator and selfie verification: complex rules made frictionless. Compliance UX as product craft.','-> CityScore ops console (Case Study 3): same unified intelligence architecture, compliance data instead of ops workflows.'].forEach((t,i)=>{const y=1.82+i*1.02;
    s.addShape(pptx.ShapeType.rect,{x:5.22,y:y+0.04,w:0.05,h:0.88,fill:{color:C.green}});
    s.addText(t,{x:5.4,y:y+0.04,w:4.1,h:0.88,fontSize:10.5,color:'1F3328',fontFace:'Calibri'});});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.1,w:9.3,h:1.1,fill:{color:C.dark},line:{color:C.green,width:1.5}});
  s.addText('"Compliance is not a feature I bolt onto a product. It is a UX problem I design through. RideGuide is what that looks like when you take it seriously."',{x:0.55,y:6.18,w:9.0,h:0.75,fontSize:12,color:C.white,italic:true,fontFace:'Calibri',align:'center'});
  s.addText('- Ajay Avaghade',{x:0.55,y:6.88,w:9.0,h:0.25,fontSize:10,color:C.muted,align:'center',fontFace:'Calibri'});}

  // SLIDE 8 — ROLLOUT
  {const s=pptx.addSlide();s.background={color:C.dark};
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.gold}});
  s.addText('ROLLOUT PLAN',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.gold,charSpacing:2,fontFace:'Calibri'});
  s.addText('Three Phases — Warsaw Pilot to Global Compliance Platform',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:23,bold:true,color:C.white,fontFace:'Calibri'});
  [{ph:'Phase 1',tl:'Month 1–2',sub:'Warsaw Pilot · 5K rides/day',body:'Deploy RideGuide soft deceleration + parking validator. A/B test vs hard geofence control (50/50 split). Measure: complaint rate, end-ride drop-off, compliance score delta. Selfie verification for all new onboarding.'},{ph:'Phase 2',tl:'Month 3–4',sub:'5 Cities · Europe',body:'Roll out RideGuide to top 5 European cities. Tune jerk limit and confidence threshold per city via Zone Editor. Launch CityScore integration for compliance logging. Feed Warsaw A/B results into city configs.'},{ph:'Phase 3',tl:'Month 5–6',sub:'All Active Cities',body:'RideGuide becomes the default compliance layer for all cities. LaunchKit auto-configures RideGuide parameters at city launch. CityScore tender exports enabled for all cities. Compliance UX = platform standard.'}].forEach((p,i)=>{const x=0.35+i*3.2;
    s.addShape(pptx.ShapeType.rect,{x,y:1.4,w:3.0,h:4.65,fill:{color:C.hero},line:{color:i===0?C.green:i===1?C.gold:'34D186',width:2},shadow:makeShadow()});
    s.addShape(pptx.ShapeType.rect,{x,y:1.4,w:3.0,h:0.06,fill:{color:i===0?C.green:i===1?C.gold:'34D186'}});
    s.addText(p.ph,{x:x+0.15,y:1.52,w:2.7,h:0.35,fontSize:18,bold:true,color:C.white,fontFace:'Calibri'});
    s.addText(p.tl,{x:x+0.15,y:1.88,w:2.7,h:0.28,fontSize:11,bold:true,color:i===0?C.green:i===1?C.gold:'34D186',fontFace:'Calibri'});
    s.addText(p.sub,{x:x+0.15,y:2.14,w:2.7,h:0.28,fontSize:11,color:C.muted,fontFace:'Calibri'});
    s.addShape(pptx.ShapeType.line,{x:x+0.15,y:2.47,w:2.6,h:0,line:{color:'1E3422',width:0.8}});
    s.addText(p.body,{x:x+0.15,y:2.55,w:2.7,h:3.4,fontSize:10.5,color:C.white,fontFace:'Calibri'});});
  s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.2,w:9.3,h:0.9,fill:{color:'1F1200'},line:{color:C.gold,width:1.5}});
  s.addText('What I need to ship Phase 1:',{x:0.55,y:6.27,w:3,h:0.28,fontSize:11,bold:true,color:C.gold,fontFace:'Calibri'});
  s.addText('Warsaw ride telemetry access (GPS + speed log)  x  1 mobile engineer (decel curve + validator UI)  x  Alignment with compliance team on thresholds',{x:0.55,y:6.55,w:9.0,h:0.42,fontSize:10,color:C.white,fontFace:'Calibri'});
  s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  linkedin.com/in/ajay-avaghade',{x:0.35,y:7.2,w:9.3,h:0.22,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});}

  return pptx.writeFile({fileName:'/Users/ajayavaghade/Portfolio/assets/bolt-rideguide-deck.pptx'});
}

// ═══════════════════════════════════════════════════════════════
// DECK 2: LAUNCHKIT
// ═══════════════════════════════════════════════════════════════
function buildLaunchKit(){
  const pptx=new PptxGenJS();pptx.layout='LAYOUT_16x9';
  const slides=[
    // S1 cover
    ()=>{const s=pptx.addSlide();s.background={color:C.dark};
      for(let i=0;i<18;i++) s.addShape(pptx.ShapeType.line,{x:-0.5+i*0.7,y:0,w:4,h:0,line:{color:'162815',width:18},rotate:35});
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
      s.addText('BOLT · EXPANSION TOOLKIT · Case Study 02',{x:0.35,y:0.28,w:7,h:0.22,fontSize:8,bold:true,color:C.muted,charSpacing:1.5,fontFace:'Calibri'});
      s.addText('LaunchKit',{x:0.35,y:1.1,w:7,h:1.3,fontSize:72,bold:true,color:C.white,fontFace:'Calibri'});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:2.3,w:2.9,h:0.06,fill:{color:C.green}});
      s.addText('Self-Serve City Configuration — 6 Weeks to 3 Days',{x:0.35,y:2.5,w:8.5,h:0.45,fontSize:18,color:C.gold,fontFace:'Calibri'});
      s.addText('Ajay Avaghade · Product Manager',{x:0.35,y:3.12,w:5,h:0.3,fontSize:11,color:C.muted,fontFace:'Calibri'});
      s.addShape(pptx.ShapeType.rect,{x:6.8,y:5.8,w:3.0,h:1.45,fill:{color:C.hero},line:{color:C.green,width:1.5},shadow:makeShadow()});
      s.addText('3 days',{x:6.8,y:5.85,w:3.0,h:0.55,fontSize:34,bold:true,color:C.gold,align:'center',fontFace:'Calibri'});
      s.addText('avg city launch time',{x:6.8,y:6.38,w:3.0,h:0.25,fontSize:11,color:C.white,align:'center',fontFace:'Calibri'});
      s.addText('down from 6 weeks with engineering',{x:6.8,y:6.63,w:3.0,h:0.2,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});},
    // S2 problem
    ()=>{const s=pptx.addSlide();s.background={color:C.dark};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
      s.addText('THE PROBLEM',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
      s.addText('Every New City Costs 6 Weeks of Engineering Time.\nBolt\'s Expansion Velocity is Bottlenecked by Config, Not Strategy.',{x:0.35,y:0.62,w:9.3,h:1.1,fontSize:24,bold:true,color:C.white,fontFace:'Calibri',breakLine:true});
      [{icon:'⏱️',stat:'6 wks',label:'Per-city engineering config time',sub:'Speed zones, parking zones, age rules, fleet caps — all require code, PRs, and deploys'},{icon:'💰',stat:'€18K',label:'Avg engineering cost per city',sub:'Sprint allocation x eng hourly rate x QA x deploy cycles per new city'},{icon:'📉',stat:'−40%',label:'Cities launched vs BD pipeline',sub:'BD closes cities faster than engineering can configure them — revenue left on the table'}].forEach((c,i)=>{const x=0.35+i*3.2;
        s.addShape(pptx.ShapeType.rect,{x,y:2.0,w:3.0,h:2.75,fill:{color:C.hero},line:{color:C.green,width:1},shadow:makeShadow()});
        s.addText(c.icon,{x:x+0.15,y:2.1,w:0.6,h:0.5,fontSize:22});
        s.addText(c.stat,{x:x+0.15,y:2.58,w:2.7,h:0.7,fontSize:42,bold:true,color:C.gold,fontFace:'Calibri'});
        s.addText(c.label,{x:x+0.15,y:3.28,w:2.7,h:0.38,fontSize:11,bold:true,color:C.white,fontFace:'Calibri'});
        s.addText(c.sub,{x:x+0.15,y:3.68,w:2.7,h:0.85,fontSize:10,color:C.muted,fontFace:'Calibri'});});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.0,w:9.3,h:1.15,fill:{color:'1F2E22'},line:{color:C.green,width:1}});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.0,w:0.05,h:1.15,fill:{color:C.green}});
      s.addText('Root cause: The bottleneck was never the regulation. It was the handoff. PM writes spec, ops writes ticket, engineer codes zones, QA tests, deploy. Four steps with handoff lag at each. LaunchKit collapses this to one step: ops configures, checklist passes, activate.',{x:0.55,y:5.07,w:9.0,h:1.0,fontSize:11,color:C.white,fontFace:'Calibri'});},
    // S3 insight
    ()=>{const s=pptx.addSlide();s.background={color:C.lgray};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
      s.addText('THE INSIGHT',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
      s.addText('City Configuration Is a Data Problem, Not an\nEngineering Problem.',{x:0.35,y:0.62,w:9.3,h:1.0,fontSize:27,bold:true,color:C.dark,fontFace:'Calibri',breakLine:true});
      s.addShape(pptx.ShapeType.ellipse,{x:4.7,y:2.0,w:0.6,h:0.6,fill:{color:C.green}});
      s.addText('VS',{x:4.7,y:2.05,w:0.6,h:0.5,fontSize:11,bold:true,color:'000000',align:'center',fontFace:'Calibri'});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.82,w:4.2,h:3.45,fill:{color:'FDE8E8'},line:{color:'F59090',width:1}});
      s.addText('Today (Engineering-Dependent)',{x:0.5,y:1.92,w:3.9,h:0.35,fontSize:12,bold:true,color:'B91C1C',fontFace:'Calibri'});
      ['PM spec -> Eng ticket -> Sprint planning (1-2 weeks)','Code: define zones in GeoJSON, write rule logic (2 weeks)','QA: test on staging, validate GeoJSON accuracy (1 week)','Deploy: push to prod, monitor for regressions (1 week)','Total: 6 weeks, 1 sprint, 1 engineer, 18K euros'].forEach((t,i)=>s.addText('•  '+t,{x:0.5,y:2.36+i*0.52,w:3.9,h:0.45,fontSize:11,color:'374151',fontFace:'Calibri'}));
      s.addShape(pptx.ShapeType.rect,{x:5.45,y:1.82,w:4.2,h:3.45,fill:{color:'ECFDF5'},line:{color:'6EE7B7',width:1}});
      s.addText('LaunchKit (Ops Self-Serve)',{x:5.6,y:1.92,w:3.9,h:0.35,fontSize:12,bold:true,color:'065F46',fontFace:'Calibri'});
      ['Upload city GeoJSON or draw zones on map (30 min)','Set rules: age, speed, parking, fleet cap (20 min)','Run checklist: all gates pass or flag blockers (10 min)','Click Activate City: live in 1 click, 0 code changes','Total: 3 days, 0 engineers, 0 euros incremental cost'].forEach((t,i)=>s.addText('•  '+t,{x:5.6,y:2.36+i*0.52,w:3.9,h:0.45,fontSize:11,color:'065F46',fontFace:'Calibri'}));
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.42,w:9.3,h:1.18,fill:{color:'FFF3CD'},line:{color:C.gold,width:1.5}});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.42,w:0.05,h:1.18,fill:{color:C.gold}});
      s.addText('Key insight: PhonePe self-serve PG integration platform eliminated engineering dependency for merchant onboarding. 5,000+ merchants acquired, 23% lower CAC. LaunchKit applies the same principle: move city configuration from engineering (code) to ops (UI toggles). The product is the configuration layer itself.',{x:0.55,y:5.49,w:9.0,h:1.0,fontSize:10.5,color:'1F1109',fontFace:'Calibri'});},
    // S4 mechanic
    ()=>{const s=pptx.addSlide();s.background={color:C.hero};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.gold}});
      s.addText('THE MECHANIC',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.gold,charSpacing:2,fontFace:'Calibri'});
      s.addText('Zone Config + Rule Library + Launch Checklist + Compliance Health',{x:0.35,y:0.62,w:9.3,h:0.65,fontSize:22,bold:true,color:C.white,fontFace:'Calibri'});
      [{n:'1',t:'Zone Map Configurator',b:'Ops uploads city GeoJSON or draws freehand on Mapbox. Three zone types: Speed Limit (sets km/h threshold), Parking Zone (designates valid spots), Restricted (no-ride area). Drag to resize, right-click to delete. No code, no PR, no deploy.'},{n:'2',t:'Rule Library',b:'Per-city toggles: age verification method (selfie/document/none), minimum age, speed limit per zone, fleet size cap, unlock fee, per-minute pricing, dynamic pricing toggle. Template inheritance from nearest peer city — 80% of rules pre-filled.'},{n:'3',t:'Launch Checklist',b:'Machine-enforced gate: Required items (zones, rules, insurance) must all pass before Activate button is unlocked. Optional items can be pending with PM override. Automated check runs on every rule change — no manual verification.'},{n:'4',t:'Compliance Health Dashboard',b:'Post-launch: real-time compliance score per city (speed adherence, valid parking rate, age-verified rider %). Anomaly detection: >10% drop in 4h fires Critical alert. Market leaderboard. Feeds directly into CityScore tender exports.'},{n:'5',t:'Rollout Velocity Tracking',b:'LaunchKit dashboard: cities launched per quarter, avg launch time, engineering hours saved, incremental revenue from faster launches. The PM owns city launch velocity as a business metric, not just an ops workflow.'}].forEach((step,i)=>{const y=1.48+i*1.04;
        s.addShape(pptx.ShapeType.ellipse,{x:0.35,y:y+0.02,w:0.45,h:0.45,fill:{color:C.green}});
        s.addText(step.n,{x:0.35,y:y+0.04,w:0.45,h:0.4,fontSize:14,bold:true,color:'000000',align:'center',fontFace:'Calibri'});
        if(i<4) s.addShape(pptx.ShapeType.line,{x:0.575,y:y+0.47,w:0,h:0.57,line:{color:C.muted,width:1,dashType:'dash'}});
        s.addText(step.t,{x:0.95,y:y,w:9.0,h:0.28,fontSize:13,bold:true,color:C.gold,fontFace:'Calibri'});
        s.addText(step.b,{x:0.95,y:y+0.28,w:8.9,h:0.65,fontSize:10.5,color:C.white,fontFace:'Calibri'});});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.85,w:9.3,h:0.42,fill:{color:C.dark},line:{color:C.gold,width:0.8}});
      s.addText('PhonePe proof: Self-serve PG integration platform (zero-to-one) -> 5,000+ B2B merchants onboarded, 23% lower CAC, 5M+ new users/month. LaunchKit = same self-serve principle applied to city compliance configuration.',{x:0.5,y:6.88,w:9.1,h:0.36,fontSize:9.5,color:C.muted,fontFace:'Calibri'});},
    // S5 product
    ()=>{const s=pptx.addSlide();s.background={color:C.lgray};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
      s.addText('THE PRODUCT',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
      s.addText('4 Screen States — Configure to Launch',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:26,bold:true,color:C.dark,fontFace:'Calibri'});
      [{n:'01',t:'Zone Configurator',b:'City map with draw/upload interface. Three zone types selectable. Zone summary table with edit buttons. Runs Checklist CTA. No code required.'},{n:'02',t:'Rule Library',b:'Age verification toggles (selfie/doc). Speed limit sliders per zone. Pricing grid (unlock, per-min, fleet cap). Dynamic pricing toggle. Template from peer city pre-filled.'},{n:'03',t:'Launch Checklist',b:'Machine-enforced gate: 4/7 items in example. Required vs Optional items. Activate City button locks until required items pass. Progress bar shows 57% completion.'},{n:'04',t:'Compliance Health',b:'4 KPI cards: 14 active cities, 84% score, 3-day launch avg, 0 eng cost. City score cards with ring charts. Global breakdown bars.'}].forEach((c,i)=>{const x=0.35+i*2.4;
        s.addShape(pptx.ShapeType.rect,{x,y:1.35,w:2.25,h:5.65,fill:{color:C.white},line:{color:'C8E8D8',width:1},shadow:makeShadow()});
        s.addShape(pptx.ShapeType.rect,{x,y:1.35,w:2.25,h:0.05,fill:{color:C.green}});
        s.addText(c.n,{x:x+0.12,y:1.45,w:1.0,h:0.28,fontSize:9,bold:true,color:C.green,fontFace:'Calibri',charSpacing:1});
        s.addText(c.t,{x:x+0.12,y:1.72,w:2.0,h:0.52,fontSize:12,bold:true,color:C.dark,fontFace:'Calibri'});
        s.addShape(pptx.ShapeType.rect,{x:x+0.12,y:2.32,w:2.0,h:2.2,fill:{color:'0D1A10'},line:{color:'1E3422',width:1}});
        [['╔══════════╗','Zone Map Editor','─────────────','⚡ 25km/h zone','🅿 Park Zone A','🚫 Airport zone','[Run Checklist]'],['╔══════════╗','Rule Library','─────────────','18+ Selfie on','Speed: 25 km/h','Fleet: 150 scoot','[Save Rules]'],['╔══════════╗','Launch Checklist','─────────────','✓ Zones 4/7','! Insurance pend','— Auth sign-off','[Activate City]'],['╔══════════╗','Compliance Health','─────────────','14 cities active','Avg Score: 84','🇵🇱91 🇷🇴88 🇳🇬61','[Export PDF]']][i].forEach((line,li)=>s.addText(line,{x:x+0.15,y:2.35+li*0.275,w:1.95,h:0.26,fontSize:7.5,color:li===0?C.green:'86EFAC',fontFace:'Courier New',bold:li===0}));
        s.addText(c.b,{x:x+0.12,y:4.62,w:2.0,h:2.25,fontSize:9.5,color:'4B6B5A',fontFace:'Calibri'});});
      s.addText('Prototype: bolt-launchkit-prototype.html · All 4 states interactive · Toggles + sliders live',{x:0.35,y:7.12,w:9.3,h:0.22,fontSize:8.5,color:C.muted,italic:true,fontFace:'Calibri'});},
    // S6 impact
    ()=>{const s=pptx.addSlide();s.background={color:C.dark};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
      s.addText('IMPACT & ROI',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
      s.addText('PhonePe Self-Serve Platform Proof — Applied to City Launch',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:22,bold:true,color:C.white,fontFace:'Calibri'});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.35,w:4.5,h:0.38,fill:{color:C.green}});
      s.addText('OPS / LAUNCH VELOCITY IMPACT',{x:0.35,y:1.37,w:4.5,h:0.34,fontSize:10,bold:true,color:'000000',align:'center',charSpacing:1.5,fontFace:'Calibri'});
      s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.35,w:4.5,h:0.38,fill:{color:'2C1A00'},line:{color:C.gold,width:1}});
      s.addText('BUSINESS / EXPANSION ROI',{x:5.15,y:1.37,w:4.5,h:0.34,fontSize:10,bold:true,color:C.gold,align:'center',charSpacing:1.5,fontFace:'Calibri'});
      [{v:'3 days',l:'Avg city launch time',s:'Down from 6 weeks. 3 days includes city authority docs + test rides — not engineering'},{v:'−85%',l:'Engineering hours per city',s:'From ~1 full sprint to near zero. Engineers freed for product features, not config work'},{v:'€0',l:'Incremental eng cost per launch',s:'LaunchKit amortizes over all city launches. Marginal cost of each new city = zero'},{v:'+40%',l:'Cities launched per quarter',s:'BD closes more contracts; ops can now configure at the speed of business development'}].forEach((m,i)=>{const y=1.9+i*1.25;
        s.addShape(pptx.ShapeType.rect,{x:0.35,y,w:4.5,h:1.15,fill:{color:C.hero},line:{color:'1E3422',width:1}});
        s.addText(m.v,{x:0.5,y:y+0.08,w:1.8,h:0.55,fontSize:32,bold:true,color:C.gold,fontFace:'Calibri'});
        s.addText(m.l,{x:0.5,y:y+0.6,w:4.1,h:0.25,fontSize:12,bold:true,color:C.white,fontFace:'Calibri'});
        s.addText(m.s,{x:0.5,y:y+0.83,w:4.1,h:0.28,fontSize:9,color:C.muted,fontFace:'Calibri'});});
      [{v:'+18%',l:'Revenue from faster launches',s:'Each city delayed by 1 week = ~18% of that quarter revenue. 3 days vs 6 weeks = significant rev recovery'},{v:'×3',l:'Expansion team leverage',s:'Same ops team can manage 3x city launches with LaunchKit vs manual config process'},{v:'84%',l:'Avg compliance score (LaunchKit cities)',s:'Pre-configured rule templates from peer cities = higher first-week compliance than hand-coded configs'},{v:'Tenders',l:'Compliance health exportable for tenders',s:'LaunchKit data feeds CityScore. Every city has a compliance PDF ready for city authority review'}].forEach((m,i)=>{const y=1.9+i*1.25;
        s.addShape(pptx.ShapeType.rect,{x:5.15,y,w:4.5,h:1.15,fill:{color:'1F1200'},line:{color:'4A3300',width:1}});
        s.addText(m.v,{x:5.3,y:y+0.08,w:1.8,h:0.55,fontSize:32,bold:true,color:C.gold,fontFace:'Calibri'});
        s.addText(m.l,{x:5.3,y:y+0.6,w:4.1,h:0.25,fontSize:12,bold:true,color:C.white,fontFace:'Calibri'});
        s.addText(m.s,{x:5.3,y:y+0.83,w:4.1,h:0.28,fontSize:9,color:C.muted,fontFace:'Calibri'});});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.9,w:9.3,h:0.38,fill:{color:'100F08'},line:{color:C.green,width:1}});
      s.addText('The configuration was never the product. LaunchKit makes the configuration invisible, so city launches compound without engineering cost.',{x:0.5,y:6.93,w:9.0,h:0.32,fontSize:9.5,color:C.white,fontFace:'Calibri'});},
    // S7 proof
    ()=>{const s=pptx.addSlide();s.background={color:C.lgray};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.green}});
      s.addText('PROOF OF WORK',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
      s.addText('I Built the Self-Serve Platform Architecture at PhonePe.',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:26,bold:true,color:C.dark,fontFace:'Calibri'});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.35,w:4.5,h:4.6,fill:{color:C.dark}});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:1.35,w:4.5,h:0.32,fill:{color:C.green}});
      s.addText('What I shipped at PhonePe',{x:0.5,y:1.37,w:4.2,h:0.28,fontSize:10,bold:true,color:'000000',fontFace:'Calibri'});
      [{h:'Self-Serve PG Integration Platform (0 to 1)',b:'Built zero-to-one self-serve merchant onboarding with Instant Discount and EMI subvention as core value prop. 5,000+ B2B merchants acquired, 23% lower CAC, 5M+ new users/month. The self-serve principle is identical to LaunchKit.'},{h:'Offer Operations Workflow Redesign',b:'Replaced manual offer creation workflow with self-serve ops tooling. TAT 30 min (down from 2 days). Same ops-layer configuration approach as LaunchKit rule library and launch checklist.'},{h:'Unified Ops Intelligence Layer',b:'Replaced 6 siloed ops workflows with single intelligence console. TAT 2d to 4h, -68% escalation tickets. LaunchKit compliance health dashboard = same consolidation principle.'},{h:'A/B Testing and Experimentation at Scale',b:'Ran ML model A/B tests across 350M MAU with shared measurement framework. LaunchKit A/B tests decel thresholds per city — same experimentation architecture, physical compliance instead of digital offers.'}].forEach((item,i)=>{const y=1.82+i*1.02;
        s.addShape(pptx.ShapeType.ellipse,{x:0.48,y:y+0.05,w:0.22,h:0.22,fill:{color:C.green}});
        s.addText(item.h,{x:0.82,y:y,w:3.85,h:0.28,fontSize:11,bold:true,color:C.gold,fontFace:'Calibri'});
        s.addText(item.b,{x:0.82,y:y+0.28,w:3.85,h:0.62,fontSize:9.5,color:C.white,fontFace:'Calibri'});});
      s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.35,w:4.5,h:4.6,fill:{color:'F0FAF4'}});
      s.addShape(pptx.ShapeType.rect,{x:5.15,y:1.35,w:4.5,h:0.32,fill:{color:'A7F3D0'}});
      s.addText('How it maps to LaunchKit',{x:5.3,y:1.37,w:4.2,h:0.28,fontSize:10,bold:true,color:C.dark,fontFace:'Calibri'});
      ['-> LaunchKit zone + rule configuration: same self-serve onboarding principle. Ops team = merchants. City config = PG integration. Zero engineering dependency in both cases.','-> LaunchKit rule library and checklist: same 30-min ops TAT principle. Compliance configuration that was weeks of eng now done in hours by ops analysts.','-> LaunchKit compliance health: same unified intelligence console. Compliance data from 14+ cities in one dashboard instead of 14 Notion docs.','-> LaunchKit A/B city configs: test different decel thresholds, age rules, parking configs across similar cities. Data-driven iteration at the city level.'].forEach((t,i)=>{const y=1.82+i*1.02;
        s.addShape(pptx.ShapeType.rect,{x:5.22,y:y+0.04,w:0.05,h:0.88,fill:{color:C.green}});
        s.addText(t,{x:5.4,y:y+0.04,w:4.1,h:0.88,fontSize:10.5,color:'1F3328',fontFace:'Calibri'});});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.1,w:9.3,h:1.1,fill:{color:C.dark},line:{color:C.green,width:1.5}});
      s.addText('"The bottleneck is always the handoff, not the work. LaunchKit eliminates the handoff between ops and engineering for city launches — the same way the PG platform eliminated it for merchant onboarding."',{x:0.55,y:6.18,w:9.0,h:0.75,fontSize:12,color:C.white,italic:true,fontFace:'Calibri',align:'center'});
      s.addText('- Ajay Avaghade',{x:0.55,y:6.88,w:9.0,h:0.25,fontSize:10,color:C.muted,align:'center',fontFace:'Calibri'});},
    // S8 rollout
    ()=>{const s=pptx.addSlide();s.background={color:C.dark};
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:C.gold}});
      s.addText('ROLLOUT PLAN',{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.gold,charSpacing:2,fontFace:'Calibri'});
      s.addText('Three Phases — Internal Tool to Expansion Platform',{x:0.35,y:0.62,w:9.3,h:0.55,fontSize:24,bold:true,color:C.white,fontFace:'Calibri'});
      [{ph:'Phase 1',tl:'Month 1–2',sub:'Internal pilot: 3 cities',body:'Build LaunchKit for 3 upcoming city launches. Measure: launch time delta vs prior 3 cities, engineering hours saved. Validate checklist completeness. Test GeoJSON import with Warsaw, Riga, Vilnius configs.'},{ph:'Phase 2',tl:'Month 3–4',sub:'All new city launches',body:'LaunchKit becomes the required process for all new city launches. Rule library expanded with all regulatory edge cases (electric bike rules, scooter sub-types, permit caps). Compliance health dashboard live for all 14+ active cities.'},{ph:'Phase 3',tl:'Month 5–6',sub:'Expansion platform + tenders',body:'LaunchKit feeds CityScore for tender data exports. BD team uses LaunchKit to demo Bolt regulatory capability during city pitches. Template inheritance expanded — any market can launch from day 1 with peer city config.'}].forEach((p,i)=>{const x=0.35+i*3.2;
        s.addShape(pptx.ShapeType.rect,{x,y:1.4,w:3.0,h:4.65,fill:{color:C.hero},line:{color:i===0?C.green:i===1?C.gold:'34D186',width:2},shadow:makeShadow()});
        s.addShape(pptx.ShapeType.rect,{x,y:1.4,w:3.0,h:0.06,fill:{color:i===0?C.green:i===1?C.gold:'34D186'}});
        s.addText(p.ph,{x:x+0.15,y:1.52,w:2.7,h:0.35,fontSize:18,bold:true,color:C.white,fontFace:'Calibri'});
        s.addText(p.tl,{x:x+0.15,y:1.88,w:2.7,h:0.28,fontSize:11,bold:true,color:i===0?C.green:i===1?C.gold:'34D186',fontFace:'Calibri'});
        s.addText(p.sub,{x:x+0.15,y:2.14,w:2.7,h:0.28,fontSize:11,color:C.muted,fontFace:'Calibri'});
        s.addShape(pptx.ShapeType.line,{x:x+0.15,y:2.47,w:2.6,h:0,line:{color:'1E3422',width:0.8}});
        s.addText(p.body,{x:x+0.15,y:2.55,w:2.7,h:3.4,fontSize:10.5,color:C.white,fontFace:'Calibri'});});
      s.addShape(pptx.ShapeType.rect,{x:0.35,y:6.2,w:9.3,h:0.9,fill:{color:'1F1200'},line:{color:C.gold,width:1.5}});
      s.addText('What I need to ship Phase 1:',{x:0.55,y:6.27,w:3,h:0.28,fontSize:11,bold:true,color:C.gold,fontFace:'Calibri'});
      s.addText('Access to city GeoJSON files for 3 pilot cities  x  1 frontend engineer (Mapbox + toggle UI)  x  Alignment with ops team and compliance on rule library structure',{x:0.55,y:6.55,w:9.0,h:0.42,fontSize:10,color:C.white,fontFace:'Calibri'});
      s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  linkedin.com/in/ajay-avaghade',{x:0.35,y:7.2,w:9.3,h:0.22,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});}
  ];
  slides.forEach(fn=>fn());
  return pptx.writeFile({fileName:'/Users/ajayavaghade/Portfolio/assets/bolt-launchkit-deck.pptx'});
}

// ═══════════════════════════════════════════════════════════════
// DECK 3: CITYSCORE — condensed
// ═══════════════════════════════════════════════════════════════
function buildCityScore(){
  const pptx=new PptxGenJS();pptx.layout='LAYOUT_16x9';
  function slide(bg,fn){const s=pptx.addSlide();s.background={color:bg};fn(s);}
  const addBar=(s,x,y,w,h,color)=>s.addShape(pptx.ShapeType.rect,{x,y,w:0.06,h,fill:{color}});
  // S1 cover
  slide(C.dark,s=>{for(let i=0;i<18;i++) s.addShape(pptx.ShapeType.line,{x:-0.5+i*0.7,y:0,w:4,h:0,line:{color:'162815',width:18},rotate:35});
    addBar(s,0,0,0,7.5,C.green);
    s.addText('BOLT · COMPLIANCE INTELLIGENCE · Case Study 03',{x:0.35,y:0.28,w:7,h:0.22,fontSize:8,bold:true,color:C.muted,charSpacing:1.5,fontFace:'Calibri'});
    s.addText('CityScore',{x:0.35,y:1.1,w:7,h:1.3,fontSize:72,bold:true,color:C.white,fontFace:'Calibri'});
    s.addShape(pptx.ShapeType.rect,{x:0.35,y:2.3,w:2.8,h:0.06,fill:{color:C.green}});
    s.addText('Real-Time Compliance Intelligence for Ops & Tenders',{x:0.35,y:2.5,w:8.5,h:0.45,fontSize:18,color:C.gold,fontFace:'Calibri'});
    s.addText('Ajay Avaghade · Product Manager',{x:0.35,y:3.12,w:5,h:0.3,fontSize:11,color:C.muted,fontFace:'Calibri'});
    s.addShape(pptx.ShapeType.rect,{x:6.8,y:5.8,w:3.0,h:1.45,fill:{color:C.hero},line:{color:C.green,width:1.5},shadow:makeShadow()});
    s.addText('+35%',{x:6.8,y:5.85,w:3.0,h:0.55,fontSize:38,bold:true,color:C.gold,align:'center',fontFace:'Calibri'});
    s.addText('tender win rate (est.)',{x:6.8,y:6.38,w:3.0,h:0.25,fontSize:11,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText('with live compliance evidence vs verbal claims',{x:6.8,y:6.63,w:3.0,h:0.2,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});});
  // S2–S8 condensed
  [{eyebrow:'THE PROBLEM',bg:C.dark,h1:'Compliance Data Lives in Silos. Bolt Cannot Demonstrate Performance to City Authorities.',cols:[{i:'📊',s:'0',l:'Unified compliance view',b:'14 cities, 14 separate spreadsheets, no single score language'},{i:'⏱️',s:'4h',l:'Anomaly detection lag',b:'Parking violation spike noticed by user complaints — not monitoring'},{i:'🤝',s:'?',l:'Tender evidence provided',b:'Verbal claims only — no data export, no compliance audit trail'}],insight:'Root cause: Compliance data is generated but never aggregated. Each city ops team has their own Notion doc with violation counts, but there is no cross-city view, no anomaly detection, and nothing exportable for city authorities. CityScore makes compliance data a product asset, not a byproduct.'},{eyebrow:'THE INSIGHT',bg:C.lgray,h1:'Compliance Visibility is a Competitive Advantage — Not Just an Ops Nicety.',left:['No cross-city compliance score language','Anomalies discovered by user complaints','Tender submissions: verbal and un-evidenced','Rider data siloed — no cohort compliance view','City authority has no visibility into Bolt performance'],right:['84/100 unified score across all cities + dimensions','3sigma anomaly classifier: 12-min detection vs 4h lag','CityScore PDF export: machine-generated, tamper-proof','New rider cohort view: Day 1-7 has 3x violation rate','City authority gets live compliance dashboard access'],callout:'Key insight: PhonePe ops intelligence layer (unified console, -68% escalation tickets) applies directly here. CityScore is the compliance version: one console, all cities, real-time. The architecture is the same.'},{eyebrow:'THE MECHANIC',bg:C.hero,h1:'Composite Score x Anomaly Detection x Rider Segments x Tender Export',steps:[{n:'1',t:'Composite Score Engine',b:'5 dimensions weighted by city regulation: speed adherence (30%), valid parking (30%), age verification (20%), restricted zone (15%), no-helmet (5%). Score refreshes per hour. Market leaderboard ranks all active cities.'},{n:'2',t:'3sigma Anomaly Classifier',b:'Rolling 7-day baseline per city x dimension. When a metric drops >10% in 4h, CityScore fires a Critical alert with auto-classified root cause: sign removal, GPS accuracy, rule change, weather. Ops acts within 12 minutes vs 4h lag.'},{n:'3',t:'Rider Segment Breakdown',b:'New riders (Day 1-7) have 3x violation rate vs Day 15-30. CityScore surfaces this cohort view so PMs can target interventions at the right moment. Age group breakdown informs onboarding design decisions.'},{n:'4',t:'Tender Report Generator',b:'One-click PDF export: per-city compliance summary with ride count, score, dimension breakdown, improvement trend. Data is machine-generated and cannot be manually edited — city authorities trust the source integrity.'},{n:'5',t:'Recommended Actions Layer',b:'CityScore recommends fixes (notify authority email + rider nudge) with est. impact and 1-click send. Same recommended-fix architecture as AgentEval (DH) and PhonePe ops console.'}],proof:'PhonePe proof: Unified ops intelligence layer replacing siloed workflows. TAT 2d to 4h, -68% escalation tickets. CityScore = same consolidation architecture applied to compliance operations.'},{eyebrow:'THE PRODUCT',bg:C.lgray,h1:'4 Screen States — Overview to Tender',cards:[{n:'01',t:'City Overview',b:'4 KPI cards (score, adherence, parking, anomaly time). City score cards with ring charts. Lagos anomaly alert at bottom. Click-to-drill.'},{n:'02',t:'Violation Heatmap',b:'4x6 heatmap grid (violation type x day of week). Lagos Island cluster highlighted. Top violation patterns ranked. 1-click notify authority + rider nudge.'},{n:'03',t:'Rider Segments',b:'Three segments: compliant 62%, occasional 28%, repeat 10%. Age group bars. New rider cohort Day 1-7 breakdown (42% violation rate highlighted).'},{n:'04',t:'Tender Report',b:'Live document preview. 4 KPI tiles. Machine-generated summary paragraph. Tamper-proof footer. Export PDF CTA.'}],note:'Prototype: bolt-cityscore-prototype.html'},{eyebrow:'IMPACT & ROI',bg:C.dark,h1:'PhonePe Unified Intelligence Architecture — Applied to Compliance',lm:[{v:'+35%',l:'Tender win rate (est.)',s:'Live compliance evidence vs verbal claims in city authority submissions'},{v:'12 min',l:'Anomaly detection time',s:'Down from 4h manual discovery via user complaints or ops escalation'},{v:'−60%',l:'Compliance investigation time',s:'Root cause auto-classified: ops acts on CityScore alert, not raw data dig'},{v:'84',l:'Avg compliance score (LaunchKit cities)',s:'CityScore visibility creates accountability -> ops teams actively improve scores'}],rm:[{v:'€0',l:'Marginal cost per compliance report',s:'Machine-generated PDF from live data -- no analyst time, no manual compilation'},{v:'3×',l:'City ops team leverage',s:'Same ops team manages 3x cities when CityScore replaces per-city Notion docs'},{v:'+22%',l:'Contract renewal rate (est.)',s:'City authorities renew contracts with operators who demonstrate compliance data transparently'},{v:'Moat',l:'Data advantage vs competitors',s:'No other micromobility operator has CityScore-level compliance data visibility'}],insight:'CityScore turns compliance from a cost center into a growth asset. Every ride generates a compliance data point that compounds into tender wins, contract renewals, and market expansion.'},{eyebrow:'PROOF OF WORK',bg:C.lgray,h1:'Same Unified Intelligence Architecture I Built at PhonePe.',phLeft:[{h:'Unified Ops Intelligence Layer',b:'6 siloed ops workflows -> single intelligence console. TAT 2d to 4h, -68% escalation tickets. Same consolidation principle as CityScore.'},{h:'A/B Testing + Data Framework',b:'Shared measurement framework across 350M MAU. CityScore cross-city leaderboard is the same benchmarking architecture applied to compliance dimensions.'},{h:'Offers Eligibility Redesign',b:'Redesigned complex eligibility rules into frictionless PM tools. 22% checkout conversion lift. CityScore rule layer is the compliance equivalent.'},{h:'Propensity ML Deployment',b:'ML models for real-time per-user scoring. CityScore anomaly classifier is the same statistical anomaly detection applied to compliance dimensions.'}],phRight:['-> CityScore unified console: 14+ cities in one view, replacing per-city spreadsheets. Same single-source-of-truth architecture.','-> CityScore city leaderboard + trend comparison: same shared measurement framework, compliance dimensions instead of conversion metrics.','-> CityScore recommended fix layer: same actionable intelligence design. Compliance anomaly -> root cause -> fix recommendation in one screen.','-> CityScore anomaly detection: same real-time per-city scoring, replacing manual compliance review lag.'],quote:'"Compliance data is only valuable if someone acts on it. CityScore makes acting on it the path of least resistance."'},{eyebrow:'ROLLOUT PLAN',bg:C.dark,h1:'Three Phases — Internal Ops Tool to City Authority Platform',phases:[{ph:'Phase 1',tl:'Month 1–2',sub:'Internal ops · 5 cities',body:'Deploy CityScore for ops teams in Warsaw, Bucharest, Nairobi, Lagos, Tallinn. Validate anomaly detector on historical data (tune 3sigma threshold). Build compliance score formula with compliance team. Measure: ops time saved per anomaly.'},{ph:'Phase 2',tl:'Month 3–4',sub:'Tender mode',body:'Build tender report PDF export. Pilot with Warsaw city authority: share CityScore dashboard read-only access. A/B test: cities with CityScore evidence vs without in tender submissions. Measure: tender win rate, contract renewal rate.'},{ph:'Phase 3',tl:'Month 5–6',sub:'All cities + authorities',body:'Roll out CityScore to all 14+ active cities. City authorities get read-only dashboard access (live, not PDF). BD team uses CityScore in city pitches. LaunchKit config data flows into CityScore automatically.'}],ask:'What I need: Read access to compliance event logs across 5 pilot cities  x  1 data engineer (anomaly classifier + score aggregation)  x  City authority relationship for Warsaw pilot access'}];
  // Build S2-S8
  const[prob,ins,mech,prod,imp,proof,roll]=[];
  // simplified -- just use pptx builder inline
  [{e:'THE PROBLEM',b:C.dark,H:'Compliance Data Lives in Silos. Bolt Cannot Demonstrate Performance to City Authorities.',cn:[{i:'📊',s:'0',l:'Unified compliance view',x:'14 cities, 14 Notion docs. No cross-city score, no single source of truth.'},{i:'⏱️',s:'4h',l:'Anomaly detection time',x:'Violations discovered via user complaints -- hours after they start compounding.'},{i:'🤝',s:'None',l:'Tender evidence',x:'Bolt relies on verbal claims in city tenders. Competitors with data will win.'}],ins:'Root cause: Compliance events are logged per city but never aggregated. CityScore creates the intelligence layer on top -- one composite score, real-time anomaly detection, and machine-generated tender exports that cannot be manually edited.'}
  ].forEach(d=>{slide(d.b,s=>{
    addBar(s,0,0,0,7.5,C.green);
    s.addText(d.e,{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:C.green,charSpacing:2,fontFace:'Calibri'});
    s.addText(d.H,{x:0.35,y:0.62,w:9.3,h:1.1,fontSize:24,bold:true,color:C.white,fontFace:'Calibri',breakLine:true});
    d.cn.forEach((c,i)=>{const x=0.35+i*3.2;
      s.addShape(pptx.ShapeType.rect,{x,y:2.0,w:3.0,h:2.75,fill:{color:C.hero},line:{color:C.green,width:1},shadow:makeShadow()});
      s.addText(c.i,{x:x+0.15,y:2.1,w:0.6,h:0.5,fontSize:22});
      s.addText(c.s,{x:x+0.15,y:2.58,w:2.7,h:0.7,fontSize:42,bold:true,color:C.gold,fontFace:'Calibri'});
      s.addText(c.l,{x:x+0.15,y:3.28,w:2.7,h:0.38,fontSize:11,bold:true,color:C.white,fontFace:'Calibri'});
      s.addText(c.x,{x:x+0.15,y:3.68,w:2.7,h:0.85,fontSize:10,color:C.muted,fontFace:'Calibri'});});
    s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.0,w:9.3,h:1.15,fill:{color:'1F2E22'},line:{color:C.green,width:1}});
    s.addShape(pptx.ShapeType.rect,{x:0.35,y:5.0,w:0.05,h:1.15,fill:{color:C.green}});
    s.addText(d.ins,{x:0.55,y:5.07,w:9.0,h:1.0,fontSize:11,color:C.white,fontFace:'Calibri'});});});
  // slides 3-8 for CityScore
  ['THE INSIGHT','THE MECHANIC','THE PRODUCT','IMPACT & ROI','PROOF OF WORK','ROLLOUT PLAN'].forEach((eyebrow,idx)=>{
    const bg=[C.lgray,C.hero,C.lgray,C.dark,C.lgray,C.dark][idx];
    slide(bg,s=>{
      addBar(s,0,0,0,7.5,idx%2===0?C.green:C.gold);
      s.addText(eyebrow,{x:0.35,y:0.28,w:5,h:0.22,fontSize:9,bold:true,color:idx%2===0?C.green:C.gold,charSpacing:2,fontFace:'Calibri'});
      const heads=['Compliance Visibility is a Competitive Advantage — Not Just an Ops Tool.','Score Engine x Anomaly Detector x Rider Segments x Tender Export','4 Screen States — Overview to Tender Report','PhonePe Unified Intelligence — Applied to Compliance Ops','Same Architecture I Built at PhonePe — Applied to Compliance Data.','Three Phases — Internal Ops Tool to City Authority Platform'];
      s.addText(heads[idx],{x:0.35,y:0.62,w:9.3,h:0.7,fontSize:24,bold:true,color:idx%2===0?C.dark:C.white,fontFace:'Calibri'});
      const bodies=['CityScore consolidates 14+ city compliance data streams into one composite score per dimension. The PhonePe ops intelligence layer did the same for 6 siloed workflows, compressing TAT from 2 days to 4 hours and cutting escalation tickets by 68%. Same architecture, compliance domain.\n\nThe tender angle is the key insight: when a city authority asks "show us your compliance record," CityScore generates a machine-produced, tamper-proof PDF with 6 months of data in one click. No competitor can match this today.',
      '1. Composite Score Engine: 5 weighted dimensions per city (speed 30%, parking 30%, age 20%, restricted 15%, helmet 5%). Refreshes hourly. Market leaderboard.\n\n2. 3-sigma Anomaly Classifier: rolling 7-day baseline per city x dimension. >10% drop in 4h fires Critical alert with auto-classified root cause. 12-minute detection vs 4h lag.\n\n3. Rider Segment View: new riders (Day 1-7) have 3x violation rate. Cohort compliance view surfaces the right intervention target.\n\n4. Tender PDF Export: 1-click machine-generated report. Tamper-proof. City authorities trust source integrity.',
      'Screen 1 - City Overview: 4 KPI tiles, city score ring cards (Warsaw 91, Bucharest 88, Nairobi 75, Lagos 61), active anomaly alert.\n\nScreen 2 - Violation Heatmap: 4x6 grid (violation type x day), Lagos Island cluster highlighted, top patterns ranked, recommended actions.\n\nScreen 3 - Rider Segments: compliance by age group and new rider cohort (Day 1-7 at 42% violation rate highlighted).\n\nScreen 4 - Tender Report: live document preview with machine-generated summary, export PDF CTA.\n\nPrototype: bolt-cityscore-prototype.html — all 4 screens interactive.',
      'Rider/Ops Impact:\n- +35% est. tender win rate with live compliance evidence\n- 12 min anomaly detection (down from 4h lag)\n- -60% compliance investigation time per incident\n- 84/100 avg score across LaunchKit-configured cities\n\nBusiness ROI:\n- 0 euro marginal cost per compliance report\n- 3x city ops team leverage vs per-city Notion docs\n- +22% est. contract renewal rate with data transparency\n- Compliance moat: no competitor has equivalent data layer',
      'PhonePe Parallel:\n- Unified Ops Intelligence Layer (6 siloed workflows -> 1 console, -68% tickets, TAT 2d->4h) = CityScore unified city compliance console\n- A/B testing + shared measurement framework (350M MAU) = CityScore city leaderboard + cross-city benchmarking\n- Propensity ML deployment (real-time per-user scoring) = CityScore 3-sigma anomaly classifier\n- Offers eligibility redesign (complex rules -> frictionless UX, +22% conversion) = CityScore recommended fix layer\n\nThis is not a case study I built for the application. It is the architecture I have shipped, applied to compliance data.',
      'Phase 1 (Month 1-2): Internal ops tool for 5 pilot cities. Build score engine, anomaly classifier, violation heatmap. Baseline compliance time savings.\n\nPhase 2 (Month 3-4): Tender mode. Build PDF export. Pilot Warsaw city authority dashboard access. A/B test tender win rate with vs without CityScore evidence.\n\nPhase 3 (Month 5-6): All cities + authority access. BD team uses CityScore in city pitches. LaunchKit config flows into CityScore automatically.\n\nWhat I need: Compliance event logs for 5 pilot cities x 1 data engineer (classifier + aggregation) x Warsaw city authority relationship for pilot.\n\nAjay Avaghade | avaghadeajay009@gmail.com | +91 9561558439'];
      s.addText(bodies[idx],{x:0.35,y:1.5,w:9.3,h:5.6,fontSize:11,color:idx%2===0?C.dark:C.white,fontFace:'Calibri'});});});
  return pptx.writeFile({fileName:'/Users/ajayavaghade/Portfolio/assets/bolt-cityscore-deck.pptx'});
}

// ═══════════════════════════════════════════════════════════════
// DECK 4: ZONEIQ
// ═══════════════════════════════════════════════════════════════
function buildZoneIQ(){
  const pptx=new PptxGenJS();pptx.layout='LAYOUT_16x9';
  function slide(bg,fn){const s=pptx.addSlide();s.background={color:bg};fn(s);}
  const eyebrows=['BOLT · ML COMPLIANCE PLATFORM · Case Study 04','THE PROBLEM','THE INSIGHT','THE MECHANIC','THE PRODUCT','IMPACT & ROI','PROOF OF WORK','ROLLOUT PLAN'];
  const heads=['ZoneIQ','Hard Geofences Fail Twice: Bad UX for Compliant Riders, Bypassable by Non-Compliant Ones.','The Geofence is a Prediction Problem, Not a Boundary Problem.','Trajectory Prediction x Soft Decel x Nearest-Spot x Configurable Platform','4 Screen States — Trajectory to Zone Editor','ML Prediction Architecture — Applied to Physical Compliance','I Built the ML Prediction Architecture at PhonePe.','Three Phases — Warsaw ML Pilot to Global Compliance Platform'];
  const subtitles=['ML-Powered Smart Geofence — Predict, Warn, Decelerate Gently',null,null,null,null,null,null,null];
  const bodies=[null,'Hard geofences have two distinct failure modes. For compliant riders: sudden throttle cuts at zone boundaries create jarring, NPS-destroying moments -- the worst experience in a Bolt ride. For non-compliant riders: knowing exactly where the boundary is makes it trivially easy to end rides just outside the geofence to avoid restrictions.\n\nStats: 45% of speed zone entries involve sudden brake events (hard geofence). 22% of riders who experience a sudden cut do not rebook within 7 days. 18% of invalid parking events occur <20m outside the boundary -- a clear bypass pattern.\n\nBoth problems are solved by the same insight: treat zone entry as a trajectory prediction problem, not a boundary detection problem.',
  'ZoneIQ reframes the geofence question. Instead of "is the rider inside or outside the boundary?" it asks "where will this rider be in 30 seconds, with what confidence?"\n\nThis changes everything:\n- Compliant riders get a 30-second proactive warning and a smooth 12-second deceleration curve. No sudden cut.\n- Non-compliant riders are decelerated before they know the boundary -- the bypass strategy fails.\n- Ops teams can tune prediction sensitivity and deceleration behavior per city via a UI, not code.\n\nPhonePe parallel: Propensity-to-Transact ML model replaced manual cohort rules with real-time per-user scoring. ZoneIQ applies the same predictive scoring architecture to physical movement instead of purchase intent. Trajectory confidence is the compliance propensity score.',
  '1. Trajectory Prediction Engine: GPS + speed + heading sampled every 500ms. LSTM model (or simpler: speed vector + heading + zone GeoJSON distance) outputs P(zone entry in 30s) per rider. At confidence >= 75%, fire soft alert.\n\n2. Soft Alert (30-second countdown): audio + haptic at confidence threshold. Countdown ring in UI. Speed change shown (22 km/h -> 25 km/h limit). Decel curve visual. No sudden surprise.\n\n3. Graduated Deceleration Curve: jerk-limited deceleration (1.2 m/s3 default). 12-second ramp to zone-compliant speed. Rider arrives at boundary already compliant. No throttle cut needed.\n\n4. Nearest-Valid-Spot Engine: if rider tries to end ride in invalid zone, GPS cross-reference with LaunchKit zone GeoJSON. Return 3 nearest valid spots ranked by walking distance. Navigate CTA.\n\n5. Zone Editor (Ops Config): confidence threshold, alert distance, ramp duration, jerk limit all tunable per city via UI. No code. Impact preview shows est. effect before saving.',
  'Screen 1 - Trajectory Map: Active ride with ZoneIQ purple trajectory arc showing predicted path. 88% confidence badge. Zone approach warning card (tap for details). Speed / zone limit / ETA indicators.\n\nScreen 2 - Smart Alert: 30-second countdown ring. Speed change display. Decel curve chart (smooth purple line vs flat gold limit line). Hard cut vs soft ramp comparison. Continue riding CTA.\n\nScreen 3 - Nearest Valid Spot: Map with user dot in red no-park zone. 2 nearest valid spots ranked. Best spot badge on P-7A (38m). Navigate CTA.\n\nScreen 4 - Ops Zone Editor: Sliders for alert distance, confidence threshold, ramp duration, jerk limit. Impact preview showing -45% brake events, -28% invalid parking.\n\nPrototype: bolt-zoneiq-prototype.html -- all 4 screens interactive, countdown live.',
  'Rider Experience:\n- -45% sudden brake events at speed zone entries\n- -28% invalid parking rate (nearest-spot engine)\n- +22% est. NPS improvement at compliance moments\n- 73% less jerk force vs hard geofence (physics: 1.2 m/s3 vs ~4-6 m/s3)\n\nBolt / City ROI:\n- -32% bypass attempts (soft decel starts before riders see the boundary)\n- 88% trajectory prediction accuracy (Warsaw pilot projection)\n- 0 code changes per city threshold tuning (Zone Editor UI)\n- Full integration with CityScore: every ZoneIQ event logged for compliance reporting',
  'PhonePe Propensity-to-Transact ML Model:\nReplaced manual cohort targeting with real-time per-user ML scoring (propensity to transact) across 350M+ MAU. 32% marketing burn reduction, sustained GMV growth.\n\nZoneIQ parallel: Trajectory confidence score is the compliance version of propensity score. Same LSTM/regression model architecture, physical movement instead of purchase behavior. Same result: predictive intervention beats reactive enforcement.\n\nPhonePe Dynamic Cart Incentivization:\nContext-aware triggers: cart value x user intent x time-of-day. 35% AOV uplift, 60% cart abandonment reduction.\n\nZoneIQ parallel: Zone alert trigger: trajectory x speed x heading x zone proximity. Same multi-signal context engine, physical compliance trigger instead of purchase trigger. The architecture is identical -- the domain is different.',
  'Phase 1 (Month 1-2): Warsaw ML Pilot. Deploy ZoneIQ trajectory prediction on 1,000 rides/day sample. A/B test vs hard geofence (50/50). Measure: sudden brake events, invalid parking rate, compliance score delta, 7-day rebooking rate post-zone-event.\n\nPhase 2 (Month 3-4): Tune + Expand. Use Warsaw A/B data to tune confidence threshold and jerk limit via Zone Editor. Deploy to 5 cities. Integrate with LaunchKit -- new city launches automatically configure ZoneIQ parameters from peer city template.\n\nPhase 3 (Month 5-6): Global Standard. ZoneIQ becomes the default compliance engine for all cities. CityScore integration: every trajectory event logged for city authority compliance reports. Hard geofence deprecated.\n\nWhat I need: Warsaw GPS + speed telemetry (last 90 days) x 1 ML engineer (trajectory prediction model) x alignment with compliance team on confidence thresholds.\n\nAjay Avaghade | avaghadeajay009@gmail.com | +91 9561558439'];
  const bgs=[C.dark,C.dark,C.lgray,C.hero,C.lgray,C.dark,C.lgray,C.dark];
  const accentColors=[C.green,C.green,C.green,C.gold,C.green,C.green,C.green,C.gold];
  bgs.forEach((bg,idx)=>{
    slide(bg,s=>{
      s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:7.5,fill:{color:accentColors[idx]}});
      s.addText(eyebrows[idx],{x:0.35,y:0.28,w:7,h:0.22,fontSize:idx===0?8:9,bold:true,color:idx===0?C.muted:accentColors[idx],charSpacing:idx===0?1.5:2,fontFace:'Calibri'});
      if(idx===0){
        for(let i=0;i<18;i++) s.addShape(pptx.ShapeType.line,{x:-0.5+i*0.7,y:0,w:4,h:0,line:{color:'162815',width:18},rotate:35});
        s.addText(heads[0],{x:0.35,y:1.1,w:7,h:1.3,fontSize:72,bold:true,color:C.white,fontFace:'Calibri'});
        s.addShape(pptx.ShapeType.rect,{x:0.35,y:2.3,w:2.4,h:0.06,fill:{color:C.purple}});
        s.addText(subtitles[0],{x:0.35,y:2.5,w:8.5,h:0.45,fontSize:18,color:C.gold,fontFace:'Calibri'});
        s.addText('Ajay Avaghade · Product Manager',{x:0.35,y:3.12,w:5,h:0.3,fontSize:11,color:C.muted,fontFace:'Calibri'});
        s.addShape(pptx.ShapeType.rect,{x:6.8,y:5.8,w:3.0,h:1.45,fill:{color:C.hero},line:{color:C.purple,width:1.5},shadow:makeShadow()});
        s.addText('-45%',{x:6.8,y:5.85,w:3.0,h:0.55,fontSize:38,bold:true,color:C.gold,align:'center',fontFace:'Calibri'});
        s.addText('sudden brake events',{x:6.8,y:6.38,w:3.0,h:0.25,fontSize:11,color:C.white,align:'center',fontFace:'Calibri'});
        s.addText('vs hard geofence baseline',{x:6.8,y:6.63,w:3.0,h:0.2,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
      } else {
        s.addText(heads[idx],{x:0.35,y:0.62,w:9.3,h:1.0,fontSize:24,bold:true,color:idx%2===0?C.dark:C.white,fontFace:'Calibri'});
        s.addText(bodies[idx],{x:0.35,y:1.7,w:9.3,h:5.5,fontSize:10.5,color:idx%2===0?'374151':C.white,fontFace:'Calibri'});
      }
    });
  });
  return pptx.writeFile({fileName:'/Users/ajayavaghade/Portfolio/assets/bolt-zoneiq-deck.pptx'});
}

// Build all
Promise.all([
  buildRideGuide().then(()=>console.log('✅ bolt-rideguide-deck.pptx')),
  buildLaunchKit().then(()=>console.log('✅ bolt-launchkit-deck.pptx')),
  buildCityScore().then(()=>console.log('✅ bolt-cityscore-deck.pptx')),
  buildZoneIQ().then(()=>console.log('✅ bolt-zoneiq-deck.pptx')),
]).then(()=>console.log('🎉 All 4 Bolt decks saved!')).catch(e=>{console.error(e);process.exit(1);});
