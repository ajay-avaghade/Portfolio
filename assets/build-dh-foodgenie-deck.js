const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0D0808',
  hero:   '1A0F0F',
  red:    'E2001A',
  gold:   'FFB800',
  white:  'FFFFFF',
  lgray:  'FFF8F8',
  green:  '22C55E',
  muted:  '8B7070',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.15 });

// ─── SLIDE 1: COVER ───────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  // diagonal texture lines
  for(let i=0;i<18;i++){
    s.addShape(pptx.ShapeType.line,{
      x: -0.5 + i*0.7, y: 0, w: 4, h: 0,
      line:{ color:'1A0F0F', width:18 },
      rotate: 35, transparency: 0
    });
  }

  // red accent bar left
  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  // top-left tag
  s.addText('DELIVERY HERO · FOODTECH PRODUCT · Case Study 01',{
    x:0.35, y:0.28, w:6, h:0.22,
    fontSize:8, bold:true, color:C.muted,
    charSpacing:1.5, fontFace:'Calibri'
  });

  // product name
  s.addText('FoodGenie',{
    x:0.35, y:1.1, w:7, h:1.4,
    fontSize:72, bold:true, color:C.white,
    fontFace:'Calibri'
  });

  // red underline
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:2.35, w:3.2, h:0.06, fill:{ color:C.red } });

  // subtitle
  s.addText('AI Agent That Knows What You Want Before You Search',{
    x:0.35, y:2.55, w:7.5, h:0.45,
    fontSize:18, bold:false, color:C.gold,
    fontFace:'Calibri'
  });

  // presenter
  s.addText('Ajay Avaghade · Product Manager',{
    x:0.35, y:3.2, w:5, h:0.3,
    fontSize:11, color:C.muted, fontFace:'Calibri'
  });

  // bottom-right stat box
  s.addShape(pptx.ShapeType.rect,{
    x:6.8, y:5.8, w:3.0, h:1.45,
    fill:{ color:C.hero }, line:{ color:C.red, width:1.5 },
    shadow: makeShadow()
  });
  s.addText('+22%',{
    x:6.8, y:5.85, w:3.0, h:0.55,
    fontSize:38, bold:true, color:C.gold,
    align:'center', fontFace:'Calibri'
  });
  s.addText('repeat order rate',{
    x:6.8, y:6.38, w:3.0, h:0.25,
    fontSize:11, color:C.white, align:'center', fontFace:'Calibri'
  });
  s.addText('Pilot projection · PhonePe proof-backed',{
    x:6.8, y:6.63, w:3.0, h:0.2,
    fontSize:8, color:C.muted, align:'center', fontFace:'Calibri'
  });

  // bottom-left DH red dot motif
  s.addShape(pptx.ShapeType.ellipse,{ x:0.3, y:6.5, w:0.6, h:0.6, fill:{ color:C.red }, transparency:60 });
  s.addShape(pptx.ShapeType.ellipse,{ x:0.7, y:6.7, w:0.35, h:0.35, fill:{ color:C.gold }, transparency:70 });
}

// ─── SLIDE 2: THE PROBLEM ─────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('THE PROBLEM',{
    x:0.35, y:0.28, w:4, h:0.22,
    fontSize:9, bold:true, color:C.red,
    charSpacing:2, fontFace:'Calibri'
  });

  s.addText('65% of Sessions End Without an Order.\nGeneric Recs, No Taste Memory, Zero Agent Intelligence.',{
    x:0.35, y:0.62, w:9.3, h:1.1,
    fontSize:26, bold:true, color:C.white,
    fontFace:'Calibri', breakLine:true
  });

  // 3 stat columns
  const cols = [
    { icon:'🔍', stat:'65%', label:'Sessions end without ordering', sub:'Browse → scroll → exit. No resolution.' },
    { icon:'🤖', stat:'0', label:'Taste memory across sessions', sub:'Same generic "Top Picks" shown every time, regardless of history.' },
    { icon:'💸', stat:'−18%', label:'Repeat order rate vs potential', sub:'Users who get relevant recs reorder 3× more — MAF missing this lever.' },
  ];

  cols.forEach((c,i)=>{
    const x = 0.35 + i*3.2;
    s.addShape(pptx.ShapeType.rect,{
      x, y:2.0, w:3.0, h:2.8,
      fill:{ color:C.hero }, line:{ color:C.red, width:1 },
      shadow: makeShadow()
    });
    s.addText(c.icon,{ x:x+0.15, y:2.1, w:0.6, h:0.5, fontSize:22 });
    s.addText(c.stat,{
      x:x+0.15, y:2.6, w:2.7, h:0.7,
      fontSize:44, bold:true, color:C.gold, fontFace:'Calibri'
    });
    s.addText(c.label,{
      x:x+0.15, y:3.3, w:2.7, h:0.4,
      fontSize:12, bold:true, color:C.white, fontFace:'Calibri'
    });
    s.addText(c.sub,{
      x:x+0.15, y:3.72, w:2.7, h:0.7,
      fontSize:10, color:C.muted, fontFace:'Calibri'
    });
  });

  // insight box
  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:5.05, w:9.3, h:1.15,
    fill:{ color:'1F0A0A' }, line:{ color:C.red, width:1 }
  });
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:5.05, w:0.05, h:1.15, fill:{ color:C.red } });
  s.addText('Root cause: DH has rich order history per user but uses none of it for proactive surfacing. The discovery surface is pull-based (search) — forcing the user to know what they want. FoodGenie flips to push-based: the agent surfaces the right restaurant before the user types.',{
    x:0.55, y:5.12, w:9.0, h:1.0,
    fontSize:11, color:C.white, fontFace:'Calibri'
  });
}

// ─── SLIDE 3: THE INSIGHT ─────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.lgray };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('THE INSIGHT',{
    x:0.35, y:0.28, w:4, h:0.22,
    fontSize:9, bold:true, color:C.red,
    charSpacing:2, fontFace:'Calibri'
  });

  s.addText('Users Don\'t Want to Search. They Want an Agent That\nAlready Knows Them.',{
    x:0.35, y:0.62, w:9.3, h:1.0,
    fontSize:28, bold:true, color:C.dark,
    fontFace:'Calibri', breakLine:true
  });

  // VS divider
  s.addShape(pptx.ShapeType.ellipse,{ x:4.7, y:2.05, w:0.6, h:0.6, fill:{ color:C.red } });
  s.addText('VS',{ x:4.7, y:2.1, w:0.6, h:0.5, fontSize:11, bold:true, color:C.white, align:'center', fontFace:'Calibri' });

  // LEFT — today
  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:1.85, w:4.2, h:3.4,
    fill:{ color:'FDE8E8' }, line:{ color:'F59090', width:1 }
  });
  s.addText('❌  Today',{
    x:0.5, y:1.95, w:3.9, h:0.35,
    fontSize:13, bold:true, color:'B91C1C', fontFace:'Calibri'
  });
  [
    'Static "Top Picks" from past 7-day order frequency',
    'User types keyword → scrolls → exits (65% of sessions)',
    'No memory of cuisine mood, budget, or novelty preference',
    'Campaign banners drive brand spend, not user relevance',
    'Session ends: zero signals captured for next time',
  ].forEach((t,i)=>{
    s.addText('•  '+t,{
      x:0.5, y:2.4+i*0.52, w:3.9, h:0.45,
      fontSize:11, color:'374151', fontFace:'Calibri'
    });
  });

  // RIGHT — FoodGenie
  s.addShape(pptx.ShapeType.rect,{
    x:5.45, y:1.85, w:4.2, h:3.4,
    fill:{ color:'ECFDF5' }, line:{ color:'6EE7B7', width:1 }
  });
  s.addText('✅  FoodGenie',{
    x:5.6, y:1.95, w:3.9, h:0.35,
    fontSize:13, bold:true, color:'065F46', fontFace:'Calibri'
  });
  [
    'Proactive banner: "Burger Boutique, as usual?" — 1 tap',
    'Conversational: "something light, under €14" → ranked picks',
    'Taste Profile: 6 confirmed tags + 3 AI-learned signals (91% conf.)',
    'Every completed order updates the model — RLHF-lite',
    'Session ends: 3 new signals stored, next session smarter',
  ].forEach((t,i)=>{
    s.addText('•  '+t,{
      x:5.6, y:2.4+i*0.52, w:3.9, h:0.45,
      fontSize:11, color:'065F46', fontFace:'Calibri'
    });
  });

  // callout
  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:5.45, w:9.3, h:1.15,
    fill:{ color:'FFF3CD' }, line:{ color:C.gold, width:1.5 }
  });
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:5.45, w:0.05, h:1.15, fill:{ color:C.gold } });
  s.addText('Key insight: The PhonePe Propensity-to-Transact ML model replaced manual cohort targeting → 32% marketing burn reduction. FoodGenie applies the same architecture to food discovery: user-level real-time scoring replaces broadcast "Top Picks." The ML is not the innovation — applying it to the discovery surface is.',{
    x:0.55, y:5.52, w:9.0, h:1.0,
    fontSize:10.5, color:'1F1109', fontFace:'Calibri'
  });
}

// ─── SLIDE 4: THE MECHANIC ────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.hero };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.gold } });

  s.addText('THE MECHANIC',{
    x:0.35, y:0.28, w:5, h:0.22,
    fontSize:9, bold:true, color:C.gold,
    charSpacing:2, fontFace:'Calibri'
  });
  s.addText('Taste Profile × Proactive Surface × Conversational Agent × RLHF-Lite Loop',{
    x:0.35, y:0.62, w:9.3, h:0.7,
    fontSize:24, bold:true, color:C.white, fontFace:'Calibri'
  });

  const steps = [
    { n:'1', t:'Taste Profile Build', b:'First 5 orders analyzed for cuisine tags, budget band, delivery time preference, novelty index. Score refreshes after every order. Stored per-user, real-time.' },
    { n:'2', t:'Proactive Banner', b:'App open: FoodGenie checks last 3 orders + time-of-day + day-of-week. Surfaces "Burger Boutique, as usual?" if pattern confidence ≥ 85%. Tap → instant checkout flow.' },
    { n:'3', t:'Conversational Discovery', b:'User types "something light under €14". FoodGenie parses: cuisine-type=light, budget≤14, novelty=preferred. Returns ranked picks with 🏆 Best Match badge. All from taste profile + real-time filters.' },
    { n:'4', t:'Order → Learning Loop', b:'Post-order: cuisine tag confidence updated (+8% for confirmed new preference). Budget signal reinforced. Uncertain patterns (e.g., Friday meals) flagged for user to rate via "Rate it" button.' },
    { n:'5', t:'A/B Console', b:'PM + data team sees per-segment FoodGenie vs control session resolution rate, AOV, repeat rate. Winner rules auto-promoted to production. Full experimentation layer built in from Day 1.' },
  ];

  steps.forEach((step,i)=>{
    const y = 1.52 + i*1.05;
    // number circle
    s.addShape(pptx.ShapeType.ellipse,{ x:0.35, y:y+0.03, w:0.45, h:0.45, fill:{ color:C.red } });
    s.addText(step.n,{ x:0.35, y:y+0.05, w:0.45, h:0.4, fontSize:14, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
    // connector line (not last)
    if(i<4){
      s.addShape(pptx.ShapeType.line,{
        x:0.575, y:y+0.48, w:0, h:0.57,
        line:{ color:C.muted, width:1, dashType:'dash' }
      });
    }
    s.addText(step.t,{
      x:0.95, y:y, w:9.0, h:0.3,
      fontSize:13, bold:true, color:C.gold, fontFace:'Calibri'
    });
    s.addText(step.b,{
      x:0.95, y:y+0.3, w:8.9, h:0.65,
      fontSize:10.5, color:C.white, fontFace:'Calibri'
    });
  });

  // proof note
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:6.85, w:9.3, h:0.42, fill:{ color:C.dark }, line:{ color:C.red, width:0.8 } });
  s.addText('PhonePe proof: ML propensity model replacing manual cohorts → 32% marketing burn reduction, sustained GMV growth. FoodGenie = same ML architecture applied to food discovery surface.',{
    x:0.5, y:6.88, w:9.1, h:0.36,
    fontSize:9.5, color:C.muted, fontFace:'Calibri'
  });
}

// ─── SLIDE 5: THE PRODUCT ─────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.lgray };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('THE PRODUCT',{
    x:0.35, y:0.28, w:4, h:0.22,
    fontSize:9, bold:true, color:C.red, charSpacing:2, fontFace:'Calibri'
  });
  s.addText('4 Screen States — Discovery to Learning',{
    x:0.35, y:0.62, w:9.3, h:0.55,
    fontSize:26, bold:true, color:C.dark, fontFace:'Calibri'
  });

  const cards = [
    {
      n:'01', t:'Home + Proactive Banner',
      b:'Taste-matched "as usual?" banner at top. Chat entry "Tell Genie what you\'re craving..." 3 top cuisine suggestion chips. "23 signals learned" counter shows model depth. 1-tap ordering from banner.'
    },
    {
      n:'02', t:'Agent Chat (Discovery)',
      b:'User: "something light, under €14, quick delivery, haven\'t tried before." Typing indicator → ranked picks in ≤1.2s. 🏆 Best match badge on Levant Kitchen. Cuisine tags + delivery time shown per result.'
    },
    {
      n:'03', t:'Taste Profile',
      b:'6 confirmed tags (Thai, Burgers, etc.). 3 AI-learned signals with confidence bars: 91% "Quick Delivery preference", 84% "Budget-conscious", 88% "Cuisine explorer." User confirms or skips uncertain signals.'
    },
    {
      n:'04', t:'Order Success + Learning',
      b:'Confetti. Order card (Levant Kitchen, €22.40). 3 signal-learned cards shown: new tag added, budget confirmed, uncertain Friday pattern → Rate It button. Model updates real-time post-order.'
    },
  ];

  cards.forEach((c,i)=>{
    const x = 0.35 + i*2.4;
    s.addShape(pptx.ShapeType.rect,{
      x, y:1.35, w:2.25, h:5.65,
      fill:{ color:C.white }, line:{ color:'E5D5D5', width:1 },
      shadow: makeShadow()
    });
    s.addShape(pptx.ShapeType.rect,{ x, y:1.35, w:2.25, h:0.05, fill:{ color:C.red } });
    s.addText(c.n,{
      x:x+0.12, y:1.45, w:1.0, h:0.28,
      fontSize:9, bold:true, color:C.red, fontFace:'Calibri', charSpacing:1
    });
    s.addText(c.t,{
      x:x+0.12, y:1.72, w:2.0, h:0.55,
      fontSize:12, bold:true, color:C.dark, fontFace:'Calibri'
    });

    // ASCII mockup
    s.addShape(pptx.ShapeType.rect,{
      x:x+0.12, y:2.35, w:2.0, h:2.2,
      fill:{ color:'1A0F0F' }, line:{ color:'331515', width:1 }
    });
    const mocks = [
      ['╔══════════╗','🍔 Burger Boutique','\"as usual?\" →','─────────────','💬 Craving... ›','🍜 Thai  🌮 Mex','📊 23 signals'],
      ['╔══════════╗','"light, <€14..."','⏳ Genie thinking','─────────────','🏆 Levant Kitchen','⭐ Bao Bistro','🌟 Sakura Box'],
      ['╔══════════╗','Taste Profile','─────────────','🏷 Thai ✓  Burgers','🧠 Quick delivery','91% ████████░','Budget 84% ██████░'],
      ['╔══════════╗','🎉 Order placed!','Levant Kitchen','€22.40 · 28 min','─────────────','✨ New: Levant tag','💡 Friday? Rate it'],
    ];
    mocks[i].forEach((line,li)=>{
      s.addText(line,{
        x:x+0.15, y:2.38+li*0.28, w:1.95, h:0.27,
        fontSize:7.5, color: li===0?C.red:'6EE7B7',
        fontFace:'Courier New', bold: li===0
      });
    });

    s.addText(c.b,{
      x:x+0.12, y:4.65, w:2.0, h:2.2,
      fontSize:9.5, color:'4B3535', fontFace:'Calibri'
    });
  });

  s.addText('Prototype: dh-foodgenie-prototype.html · All 4 states interactive · Click-based navigation',{
    x:0.35, y:7.12, w:9.3, h:0.22,
    fontSize:8.5, color:C.muted, italic:true, fontFace:'Calibri'
  });
}

// ─── SLIDE 6: IMPACT & ROI ────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('IMPACT & ROI',{
    x:0.35, y:0.28, w:5, h:0.22,
    fontSize:9, bold:true, color:C.red, charSpacing:2, fontFace:'Calibri'
  });
  s.addText('Built on PhonePe ML Proof Points — Adjusted for DH Food Discovery Scale',{
    x:0.35, y:0.62, w:9.3, h:0.55,
    fontSize:22, bold:true, color:C.white, fontFace:'Calibri'
  });

  // left block header
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:1.35, w:4.5, h:0.38, fill:{ color:C.red } });
  s.addText('USER / DEMAND IMPACT',{
    x:0.35, y:1.37, w:4.5, h:0.34,
    fontSize:10, bold:true, color:C.white, align:'center', charSpacing:1.5, fontFace:'Calibri'
  });

  // right block header
  s.addShape(pptx.ShapeType.rect,{ x:5.15, y:1.35, w:4.5, h:0.38, fill:{ color:'2C1A00' }, line:{ color:C.gold, width:1 } });
  s.addText('DH / PLATFORM ROI',{
    x:5.15, y:1.37, w:4.5, h:0.34,
    fontSize:10, bold:true, color:C.gold, align:'center', charSpacing:1.5, fontFace:'Calibri'
  });

  const leftMetrics = [
    { v:'+22%', l:'Repeat order rate', s:'Users w/ taste-matched recs reorder 3× more · PhonePe redemption model parallel' },
    { v:'+18%', l:'Session-to-order conversion', s:'Proactive banner removes search friction — 1-tap vs 5-step' },
    { v:'−35%', l:'Browse-without-order sessions', s:'Agent resolution vs scroll-exit · ML propensity model applied to discovery' },
    { v:'+14%', l:'AOV on Genie-initiated orders', s:'Recs skew toward higher-rated, higher-value restaurants user has not tried' },
  ];

  const rightMetrics = [
    { v:'€0', l:'Incremental cost per session', s:'Taste Profile runs on existing order history — no new data collection required' },
    { v:'+31%', l:'Commission revenue per active user', s:'AOV uplift × repeat rate uplift = compound revenue per user cohort' },
    { v:'−40%', l:'Campaign CPO (cost per order)', s:'Targeted Genie push vs broadcast banner → same GMV, lower media spend' },
    { v:'×3', l:'Lifetime GMV per cohort vs control', s:'Every session adds a signal: FoodGenie compounds while static recs plateau' },
  ];

  leftMetrics.forEach((m,i)=>{
    const y = 1.9 + i*1.25;
    s.addShape(pptx.ShapeType.rect,{ x:0.35, y, w:4.5, h:1.15, fill:{ color:C.hero }, line:{ color:'331515', width:1 } });
    s.addText(m.v,{ x:0.5, y:y+0.08, w:1.5, h:0.55, fontSize:36, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(m.l,{ x:0.5, y:y+0.6, w:4.1, h:0.25, fontSize:12, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s,{ x:0.5, y:y+0.83, w:4.1, h:0.28, fontSize:9, color:C.muted, fontFace:'Calibri' });
  });

  rightMetrics.forEach((m,i)=>{
    const y = 1.9 + i*1.25;
    s.addShape(pptx.ShapeType.rect,{ x:5.15, y, w:4.5, h:1.15, fill:{ color:'1F1200' }, line:{ color:'4A3300', width:1 } });
    s.addText(m.v,{ x:5.3, y:y+0.08, w:1.5, h:0.55, fontSize:36, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(m.l,{ x:5.3, y:y+0.6, w:4.1, h:0.25, fontSize:12, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s,{ x:5.3, y:y+0.83, w:4.1, h:0.28, fontSize:9, color:C.muted, fontFace:'Calibri' });
  });

  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:6.9, w:9.3, h:0.38, fill:{ color:'100505' }, line:{ color:C.red, width:1 } });
  s.addText('FoodGenie compounds: every order makes the model smarter → next session has higher conversion → same MAU delivers more GMV. Unlike discounts, this improvement is permanent.',{
    x:0.5, y:6.93, w:9.0, h:0.32,
    fontSize:9.5, color:C.white, fontFace:'Calibri'
  });
}

// ─── SLIDE 7: PROOF OF WORK ───────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.lgray };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('PROOF OF WORK',{
    x:0.35, y:0.28, w:5, h:0.22,
    fontSize:9, bold:true, color:C.red, charSpacing:2, fontFace:'Calibri'
  });
  s.addText('I Built This at PhonePe. Here\'s the Direct Map.',{
    x:0.35, y:0.62, w:9.3, h:0.55,
    fontSize:26, bold:true, color:C.dark, fontFace:'Calibri'
  });

  // left panel (dark)
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:1.35, w:4.5, h:4.6, fill:{ color:C.dark } });
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:1.35, w:4.5, h:0.32, fill:{ color:C.red } });
  s.addText('What I shipped at PhonePe',{
    x:0.5, y:1.37, w:4.2, h:0.28,
    fontSize:10, bold:true, color:C.white, fontFace:'Calibri'
  });

  const phLeft = [
    { h:'Propensity-to-Transact ML', b:'Replaced static cohort targeting with real-time user-level ML scoring across 350M+ MAU → 32% marketing burn reduction, sustained GMV growth' },
    { h:'Dynamic Cart Incentivization', b:'Context-aware triggers: cart value × user intent × time-of-day → 35% AOV uplift, 60% cart abandonment reduction, 20% retention improvement' },
    { h:'Personalized Offers Placement', b:'Redesigned offers placement + eligibility surfacing with ML-ranked relevance → 22% checkout conversion lift across fragmented offers stack' },
    { h:'Third-Party Rewards Marketplace', b:'Built brand self-serve tooling + audience segmentation → 500+ brand partners, 26% user engagement increase, 11% YoY revenue growth' },
  ];

  phLeft.forEach((item,i)=>{
    const y = 1.82 + i*1.02;
    s.addShape(pptx.ShapeType.ellipse,{ x:0.48, y:y+0.06, w:0.22, h:0.22, fill:{ color:C.red } });
    s.addText(item.h,{ x:0.82, y:y, w:3.85, h:0.28, fontSize:11, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(item.b,{ x:0.82, y:y+0.28, w:3.85, h:0.6, fontSize:9.5, color:C.white, fontFace:'Calibri' });
  });

  // right panel (light)
  s.addShape(pptx.ShapeType.rect,{ x:5.15, y:1.35, w:4.5, h:4.6, fill:{ color:'FDE8E8' } });
  s.addShape(pptx.ShapeType.rect,{ x:5.15, y:1.35, w:4.5, h:0.32, fill:{ color:'F59090' } });
  s.addText('How it maps to FoodGenie at DH',{
    x:5.3, y:1.37, w:4.2, h:0.28,
    fontSize:10, bold:true, color:C.dark, fontFace:'Calibri'
  });

  const phRight = [
    '→ FoodGenie Taste Profile: per-user real-time taste scoring replacing "Top Picks" broadcast feed',
    '→ FoodGenie proactive banner + conversational agent: context × history → 1-tap order vs 5-step search',
    '→ FoodGenie ranked picks surfacing in discovery, not just checkout → ML applied to top-of-funnel',
    '→ FoodGenie A/B console + post-order signal capture → brand-safe personalization with measurement layer',
  ];

  phRight.forEach((t,i)=>{
    const y = 1.82 + i*1.02;
    s.addShape(pptx.ShapeType.rect,{ x:5.22, y:y+0.05, w:0.05, h:0.9, fill:{ color:'B91C1C' } });
    s.addText(t,{ x:5.4, y:y+0.05, w:4.1, h:0.9, fontSize:11, color:'1F0A0A', fontFace:'Calibri' });
  });

  // closing quote
  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:6.1, w:9.3, h:1.1,
    fill:{ color:C.dark }, line:{ color:C.red, width:1.5 }
  });
  s.addText('"This isn\'t a case study built in a weekend. This is architecture I\'ve shipped at scale — applied to DH\'s food discovery problem."',{
    x:0.55, y:6.18, w:9.0, h:0.75,
    fontSize:12, color:C.white, italic:true, fontFace:'Calibri', align:'center'
  });
  s.addText('— Ajay Avaghade',{
    x:0.55, y:6.88, w:9.0, h:0.25,
    fontSize:10, color:C.muted, align:'center', fontFace:'Calibri'
  });
}

// ─── SLIDE 8: ROLLOUT PLAN ────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.gold } });

  s.addText('ROLLOUT PLAN',{
    x:0.35, y:0.28, w:5, h:0.22,
    fontSize:9, bold:true, color:C.gold, charSpacing:2, fontFace:'Calibri'
  });
  s.addText('Three Phases — Zero-Risk Pilot to Full App Rollout',{
    x:0.35, y:0.62, w:9.3, h:0.55,
    fontSize:24, bold:true, color:C.white, fontFace:'Calibri'
  });

  const phases = [
    {
      ph:'Phase 1',
      timeline:'Month 1–2',
      sub:'Berlin Pilot · 50K MAU',
      body:'Deploy FoodGenie to 50K power users in Berlin (highest order frequency cohort). Activate Taste Profile on existing order history (no new data collection). Measure session resolution rate vs control. Ship proactive banner + taste profile screen.'
    },
    {
      ph:'Phase 2',
      timeline:'Month 3–4',
      sub:'A/B vs Keyword Search',
      body:'Run FoodGenie (full agent) vs keyword search control A/B across 500K MAU. Ship conversational discovery + ranked picks. Tune taste profile confidence thresholds. Measure: repeat rate, AOV, session-to-order conversion per variant.'
    },
    {
      ph:'Phase 3',
      timeline:'Month 5–6',
      sub:'Full App Rollout',
      body:'Roll out FoodGenie to all DH markets. Integrate with campaign/voucher stack — FoodGenie-initiated orders get brand-side attribution. Scale RLHF-lite post-order learning loop. Launch A/B console for market-level experimentation.'
    },
  ];

  phases.forEach((p,i)=>{
    const x = 0.35 + i*3.2;
    s.addShape(pptx.ShapeType.rect,{
      x, y:1.4, w:3.0, h:4.65,
      fill:{ color:C.hero }, line:{ color: i===0?C.red:i===1?C.gold:C.green, width:2 },
      shadow: makeShadow()
    });
    s.addShape(pptx.ShapeType.rect,{ x, y:1.4, w:3.0, h:0.06, fill:{ color: i===0?C.red:i===1?C.gold:C.green } });

    s.addText(p.ph,{
      x:x+0.15, y:1.52, w:2.7, h:0.35,
      fontSize:18, bold:true, color:C.white, fontFace:'Calibri'
    });
    s.addText(p.timeline,{
      x:x+0.15, y:1.88, w:2.7, h:0.28,
      fontSize:11, bold:true, color: i===0?C.red:i===1?C.gold:C.green, fontFace:'Calibri'
    });
    s.addText(p.sub,{
      x:x+0.15, y:2.14, w:2.7, h:0.28,
      fontSize:11, color:C.muted, fontFace:'Calibri'
    });
    s.addShape(pptx.ShapeType.line,{
      x:x+0.15, y:2.47, w:2.6, h:0,
      line:{ color:'331515', width:0.8 }
    });
    s.addText(p.body,{
      x:x+0.15, y:2.55, w:2.7, h:3.4,
      fontSize:10.5, color:C.white, fontFace:'Calibri'
    });
  });

  // ask box
  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:6.2, w:9.3, h:0.9,
    fill:{ color:'1F1200' }, line:{ color:C.gold, width:1.5 }
  });
  s.addText('What I need to ship this:',{
    x:0.55, y:6.27, w:3, h:0.28,
    fontSize:11, bold:true, color:C.gold, fontFace:'Calibri'
  });
  s.addText('Anonymized user order history (last 90 days) for taste profile seeding  ·  1 ML engineer (propensity scoring API)  ·  Alignment with Discovery PM team',{
    x:0.55, y:6.55, w:9.0, h:0.42,
    fontSize:10, color:C.white, fontFace:'Calibri'
  });

  s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  linkedin.com/in/ajay-avaghade',{
    x:0.35, y:7.2, w:9.3, h:0.22,
    fontSize:8.5, color:C.muted, align:'center', fontFace:'Calibri'
  });
}

// ─── SAVE ─────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/dh-foodgenie-deck.pptx' })
  .then(()=>console.log('✅  dh-foodgenie-deck.pptx saved'))
  .catch(e=>{ console.error(e); process.exit(1); });
