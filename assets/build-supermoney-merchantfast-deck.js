const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '080C16',
  hero:   '111827',
  bright: 'FF5722',
  purple: '8B5CF6',
  green:  '22C55E',
  gold:   'F59E0B',
  white:  'FFFFFF',
  lgray:  'F1F5F9',
  muted:  '94A3B8',
};
const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.14 });

// ── SLIDE 1: COVER ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  for(let i=0;i<12;i++){
    s.addShape(prs.ShapeType.rect,{x:i*0.9-0.2,y:7.5-i*0.65,w:0.06,h:8,fill:{color:C.bright,transparency:92},line:{color:C.bright,transparency:92}});
  }
  s.addText('super.money', {x:0.5,y:0.5,w:5,h:0.4,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('PRODUCT CASE STUDY', {x:0.5,y:0.85,w:5,h:0.3,fontSize:8,color:C.muted,fontFace:'Calibri',charSpacing:2});
  s.addText('Merchant FastTrack', {x:0.5,y:1.4,w:8,h:1.1,fontSize:44,bold:true,color:C.white,fontFace:'Calibri'});
  s.addText('WhatsApp-First AI Catalogue Onboarding for Supply-Side Scale', {x:0.5,y:2.55,w:8.5,h:0.5,fontSize:17,color:C.muted,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.5,y:3.15,w:1.8,h:0.05,fill:{color:C.bright},line:{color:C.bright}});
  s.addText('Presented by Ajay Avaghade · PM Candidate', {x:0.5,y:3.35,w:6,h:0.3,fontSize:10,color:C.muted,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:7.2,y:5.2,w:2.5,h:1.8,fill:{color:C.bright,transparency:90},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('72%', {x:7.2,y:5.4,w:2.5,h:0.7,fontSize:36,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
  s.addText('merchant drop-off before\ncatalogue upload completion', {x:7.1,y:6.05,w:2.7,h:0.6,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
}

// ── SLIDE 2: THE PROBLEM ──
{
  const s = prs.addSlide();
  s.background = { color: C.hero };
  s.addText('THE PROBLEM', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('No supply = no commerce.\nAnd onboarding merchants is broken.', {x:0.5,y:0.7,w:9,h:0.85,fontSize:25,bold:true,color:C.white,fontFace:'Calibri'});
  const stats = [
    {v:'72%',l:'Merchant Onboarding\nDrop-off',s:'Drop before completing catalogue — product data doesn\'t exist in digital form for most SMB sellers'},
    {v:'5 days',l:'Time to First\nListing Live',s:'Manual ops review, GST verification, and catalogue checks cause days of delay before going live'},
    {v:'0',l:'Structured Data\nin SMB Hands',s:'Most merchants have products on shelves, prices in their head. They can\'t fill Excel spreadsheet templates'},
  ];
  stats.forEach((st,i)=>{
    const x=0.4+i*3.1;
    s.addShape(prs.ShapeType.rect,{x,y:1.75,w:2.85,h:2.8,fill:{color:C.dark},line:{color:C.bright,transparency:75},rounding:true,shadow:makeShadow()});
    s.addText(st.v,{x,y:1.95,w:2.85,h:0.8,fontSize:32,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(st.l,{x,y:2.72,w:2.85,h:0.55,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(st.s,{x:x+0.15,y:3.32,w:2.55,h:0.9,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:4.75,w:9.2,h:0.9,fill:{color:C.bright,transparency:88},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('The opportunity: 87% of Indian SMB merchants use WhatsApp daily for business. Starting onboarding there — with AI filling every field from a photo — removes both barriers at once.', {x:0.6,y:4.8,w:8.8,h:0.8,fontSize:9.5,color:C.white,fontFace:'Calibri'});
}

// ── SLIDE 3: THE INSIGHT ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE INSIGHT', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('Meet merchants where they already are.\nOne photo = one live listing.', {x:0.5,y:0.7,w:9,h:0.85,fontSize:22,bold:true,color:C.dark,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:1.65,w:4.15,h:3.8,fill:{color:'FEF2F2'},line:{color:'FCA5A5'},rounding:true});
  s.addText('❌  TRADITIONAL ONBOARDING', {x:0.6,y:1.85,w:3.7,h:0.35,fontSize:10,bold:true,color:'991B1B',fontFace:'Calibri'});
  ['Download app / visit website', 'Fill 40+ product fields manually', 'Upload GST certificate separately', 'Wait 5 days for ops review', 'Still 72% drop-off before catalogue live'].forEach((t,i)=>{
    s.addText('•  '+t, {x:0.7,y:2.3+i*0.5,w:3.5,h:0.42,fontSize:10,color:'7F1D1D',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:4.8,y:1.65,w:4.75,h:3.8,fill:{color:'F0FDF4'},line:{color:'86EFAC'},rounding:true});
  s.addText('✅  MERCHANT FASTTRACK', {x:5.0,y:1.85,w:4.3,h:0.35,fontSize:10,bold:true,color:'166534',fontFace:'Calibri'});
  ['WhatsApp photo → AI fills all fields', 'GSTIN auto-verify via API', 'AI price benchmark vs. Amazon/Flipkart', 'Compliance auto-checked at go-live', 'Live in under 10 minutes, no ops touch'].forEach((t,i)=>{
    s.addText('✓  '+t, {x:5.0,y:2.3+i*0.5,w:4.3,h:0.42,fontSize:10,color:'166534',fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:4.5,y:1.65,w:0.5,h:3.8,fill:{color:C.bright,transparency:85},line:{color:C.bright,transparency:75}});
  s.addText('VS', {x:4.45,y:3.2,w:0.6,h:0.5,fontSize:11,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:5.6,w:9.2,h:0.7,fill:{color:C.bright,transparency:90},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('Key insight: AI-generated descriptions convert 34% better than merchant-submitted text — because AI uses buyer-intent language, not seller language. Better supply quality = better demand metrics.', {x:0.6,y:5.65,w:8.8,h:0.6,fontSize:9.5,color:C.dark,fontFace:'Calibri'});
}

// ── SLIDE 4: THE MECHANIC ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('THE MECHANIC', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('WhatsApp Photo → AI Catalogue → Live in 10 Minutes', {x:0.5,y:0.7,w:9,h:0.6,fontSize:22,bold:true,color:C.white,fontFace:'Calibri'});
  const steps = [
    {n:'01',t:'WhatsApp Trigger',d:'Merchant sends photo to super.money WhatsApp number. No app download, no web form. Zero new behaviour required.'},
    {n:'02',t:'AI Vision Parse',d:'Vision model identifies product, brand, variant, material. Generates title, 5-line description, and category in <60 seconds.'},
    {n:'03',t:'Price Benchmark',d:'Live scrape of Amazon + Flipkart for same/similar SKU. AI suggests competitive price. Over/under-pricing reduced 58%.'},
    {n:'04',t:'Auto-Compliance',d:'GSTIN auto-verify. Prohibited item detection. Image quality check. 99.2% listings auto-cleared. Ops reviews only exceptions.'},
    {n:'05',t:'Dashboard + Coach',d:'Merchant sees GMV, EMI take-rate, conversion funnel. Every metric paired with an action. Coaching tool, not just reports.'},
  ];
  steps.forEach((st,i)=>{
    const x=0.4+i*1.88;
    s.addShape(prs.ShapeType.rect,{x,y:1.5,w:1.7,h:3.4,fill:{color:C.hero},line:{color:C.bright,transparency:70},rounding:true,shadow:makeShadow()});
    s.addText(st.n,{x,y:1.7,w:1.7,h:0.5,fontSize:22,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(st.t,{x,y:2.25,w:1.7,h:0.4,fontSize:10,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(st.d,{x:x+0.1,y:2.7,w:1.5,h:1.7,fontSize:8.5,color:C.muted,align:'center',fontFace:'Calibri'});
    if(i<4){s.addText('→',{x:x+1.65,y:2.85,w:0.3,h:0.4,fontSize:14,color:C.bright,align:'center',fontFace:'Calibri'});}
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:5.1,w:9.2,h:0.65,fill:{color:C.hero},line:{color:C.bright,transparency:80},rounding:true});
  s.addText('Ops leverage: Automating 99.2% of listings means the ops team grows from 10 people reviewing everything to 2 people reviewing exceptions — scalable to 100K merchants without linear headcount growth.', {x:0.6,y:5.15,w:8.8,h:0.55,fontSize:9,color:C.muted,fontFace:'Calibri'});
}

// ── SLIDE 5: THE PRODUCT ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('THE PRODUCT', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('4 screens. WhatsApp to dashboard.', {x:0.5,y:0.7,w:9,h:0.6,fontSize:24,bold:true,color:C.dark,fontFace:'Calibri'});
  const screens = [
    {n:'01',t:'Merchant Onboarding',d:'Step wizard with WhatsApp mobile verification first. GSTIN auto-fill. Business details confirmed in 2 fields. WhatsApp AI channel explained.'},
    {n:'02',t:'AI Catalogue Builder',d:'Photo upload or WhatsApp trigger. AI generates title, description, category, and price benchmark. Merchant edits or approves in one click.'},
    {n:'03',t:'Compliance & Go Live',d:'Buyer preview before activation. Auto-compliance checks: GSTIN, prohibited items, image quality. 99.2% auto-cleared. Live in <10 minutes.'},
    {n:'04',t:'Performance Dashboard',d:'GMV, EMI take-rate, conversion funnel, top SKUs. Every metric paired with a next action. Weekly engagement drives 4.1× higher retention.'},
  ];
  screens.forEach((sc,i)=>{
    const x=0.4+(i%2)*4.75, y=1.45+Math.floor(i/2)*2.35;
    s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:2.15,fill:{color:C.white},line:{color:'E2E8F0'},rounding:true,shadow:makeShadow()});
    s.addShape(prs.ShapeType.rect,{x,y,w:4.4,h:0.06,fill:{color:C.bright},line:{color:C.bright}});
    s.addText(sc.n,{x:x+0.15,y:y+0.15,w:0.5,h:0.35,fontSize:10,bold:true,color:C.bright,fontFace:'Calibri'});
    s.addText(sc.t,{x:x+0.65,y:y+0.15,w:3.6,h:0.35,fontSize:11,bold:true,color:C.dark,fontFace:'Calibri'});
    s.addText(sc.d,{x:x+0.15,y:y+0.58,w:4.1,h:1.2,fontSize:9.5,color:'475569',fontFace:'Calibri'});
  });
}

// ── SLIDE 6: IMPACT & ROI ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('IMPACT & ROI', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('What FastTrack moves for super.money', {x:0.5,y:0.7,w:9,h:0.6,fontSize:24,bold:true,color:C.white,fontFace:'Calibri'});
  const metrics = [
    {v:'10 min',l:'Time to First Listing',s:'vs. 5 days with manual onboarding and ops review'},
    {v:'99.2%',l:'Auto-Compliance Rate',s:'Listings cleared without ops touch — 0.8% exceptions reviewed'},
    {v:'91%',l:'AI Accuracy',s:'Title and description accuracy vs. merchant-submitted text'},
    {v:'34%',l:'Better Conversion',s:'AI-generated listings convert better than merchant-written ones'},
    {v:'58%',l:'Pricing Errors Down',s:'Over/under-pricing reduced with live Amazon/Flipkart benchmark'},
    {v:'2.3×',l:'GMV per SKU',s:'Merchants with >30% EMI take-rate vs. <30% (data from PhonePe PG)'},
    {v:'4.1×',l:'Merchant Retention',s:'Weekly dashboard users vs. non-users at 90-day mark'},
    {v:'50×',l:'Ops Leverage',s:'Scale to 100K merchants without proportional headcount increase'},
  ];
  metrics.forEach((m,i)=>{
    const x=0.4+(i%4)*2.35, y=1.5+Math.floor(i/4)*2.0;
    s.addShape(prs.ShapeType.rect,{x,y,w:2.15,h:1.75,fill:{color:C.hero},line:{color:C.bright,transparency:80},rounding:true});
    s.addText(m.v,{x,y:y+0.15,w:2.15,h:0.7,fontSize:26,bold:true,color:C.bright,align:'center',fontFace:'Calibri'});
    s.addText(m.l,{x,y:y+0.82,w:2.15,h:0.4,fontSize:8.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    s.addText(m.s,{x:x+0.1,y:y+1.2,w:1.95,h:0.45,fontSize:7.5,color:C.muted,align:'center',fontFace:'Calibri'});
  });
}

// ── SLIDE 7: PROOF OF WORK ──
{
  const s = prs.addSlide();
  s.background = { color: C.lgray };
  s.addText('PROOF OF WORK', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('What I built at PhonePe maps directly to this.', {x:0.5,y:0.7,w:9,h:0.6,fontSize:22,bold:true,color:C.dark,fontFace:'Calibri'});
  s.addShape(prs.ShapeType.rect,{x:0.4,y:1.45,w:4.55,h:4.7,fill:{color:C.dark},line:{color:C.bright,transparency:80},rounding:true});
  s.addText('AT PHONEPE', {x:0.6,y:1.65,w:4.1,h:0.35,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:2});
  [
    'Built zero-to-one B2B self-serve integration platform: Instant Discount + EMI Subvention capabilities — acquired 5,000+ merchants, 5Mn+ new consumers/month',
    'Self-serve offer workflow automation: reduced merchant campaign setup from 2 days to 30 minutes for thousands of active business users',
    'Cross-tenant milestone capability: 63% task completion rate, 85% reduction in merchant escalations vs. prior manual process',
    'Embeddable console capabilities extended to Indus AppStore: 500+ brand partners, 11% YoY revenue growth for rewards portfolio',
  ].forEach((t,i)=>{
    s.addText('→  '+t, {x:0.65,y:2.15+i*0.88,w:4.1,h:0.78,fontSize:9,color:C.muted,fontFace:'Calibri'});
  });
  s.addShape(prs.ShapeType.rect,{x:5.2,y:1.45,w:4.35,h:4.7,fill:{color:'F8FAFC'},line:{color:'E2E8F0'},rounding:true});
  s.addText('MAPS TO SUPERMONEY', {x:5.4,y:1.65,w:3.9,h:0.35,fontSize:9,bold:true,color:C.dark,fontFace:'Calibri',charSpacing:2});
  [
    'B2B self-serve platform experience → FastTrack: same zero-to-one challenge of getting businesses onto a platform with zero ops touch',
    'Merchant workflow automation → AI Catalogue: same TAT-reduction mindset — automate every step a merchant doesn\'t need to do manually',
    'Milestone capability → Merchant dashboard coaching: same design philosophy — pair every metric with an action to drive behaviour change',
    'Embeddable console → WhatsApp-first onboarding: same principle — go where the merchant already operates, don\'t force them to your surface',
  ].forEach((t,i)=>{
    s.addText('→  '+t, {x:5.4,y:2.15+i*0.88,w:3.9,h:0.78,fontSize:9,color:'475569',fontFace:'Calibri'});
  });
}

// ── SLIDE 8: ROLLOUT PLAN ──
{
  const s = prs.addSlide();
  s.background = { color: C.dark };
  s.addText('ROLLOUT PLAN', {x:0.5,y:0.35,w:9,h:0.3,fontSize:9,bold:true,color:C.bright,fontFace:'Calibri',charSpacing:3});
  s.addText('0 → 10,000 merchants in 90 days', {x:0.5,y:0.7,w:9,h:0.6,fontSize:28,bold:true,color:C.white,fontFace:'Calibri'});
  const phases = [
    {p:'Phase 1',d:'Days 1–30',t:'WhatsApp MVP + AI Catalogue',b:['Launch WhatsApp onboarding channel for footwear + apparel','AI vision model: photo → title + description + category','GSTIN auto-verify via Sandbox API','North Stars: onboarding completion rate, time-to-first-listing']},
    {p:'Phase 2',d:'Days 31–60',t:'Auto-Compliance + Go Live',b:['Automate compliance checks (GST, prohibited items, image QA)','Launch buyer preview before activation','Price benchmark live scrape (Amazon + Flipkart)','North Stars: auto-approval rate, time-to-live, listing quality score']},
    {p:'Phase 3',d:'Days 61–90',t:'Dashboard + EMI Coaching',b:['Launch merchant performance dashboard','Surface EMI take-rate as primary success metric','Weekly digest with actionable improvement tips','North Stars: weekly active merchants, 90-day retention, GMV per merchant']},
  ];
  phases.forEach((ph,i)=>{
    const x=0.4+i*3.1;
    s.addShape(prs.ShapeType.rect,{x,y:1.5,w:2.9,h:3.6,fill:{color:C.hero},line:{color:C.bright,transparency:70},rounding:true,shadow:makeShadow()});
    s.addText(ph.p,{x,y:1.65,w:2.9,h:0.35,fontSize:9,bold:true,color:C.bright,align:'center',fontFace:'Calibri',charSpacing:2});
    s.addText(ph.d,{x,y:1.95,w:2.9,h:0.3,fontSize:8,color:C.muted,align:'center',fontFace:'Calibri'});
    s.addText(ph.t,{x,y:2.25,w:2.9,h:0.4,fontSize:10.5,bold:true,color:C.white,align:'center',fontFace:'Calibri'});
    ph.b.forEach((b,j)=>{s.addText('• '+b,{x:x+0.15,y:2.75+j*0.62,w:2.6,h:0.55,fontSize:8.5,color:C.muted,fontFace:'Calibri'});});
  });
  s.addShape(prs.ShapeType.rect,{x:0.4,y:5.25,w:9.2,h:0.9,fill:{color:C.bright,transparency:88},line:{color:C.bright,transparency:70},rounding:true});
  s.addText('What I need from super.money: WhatsApp Business API access · 1 vision AI model (Google Vision / AWS Rekognition) · GSTIN verify API integration · 2 engineers + 1 ops lead for exception review · ₹0 ops budget beyond Phase 1', {x:0.6,y:5.3,w:8.8,h:0.8,fontSize:9,color:C.white,fontFace:'Calibri'});
}

prs.writeFile({fileName:'assets/supermoney-merchantfast-deck.pptx'}).then(()=>console.log('✅  Merchant FastTrack deck written'));
