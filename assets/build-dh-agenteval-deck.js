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
  blue:   '3B82F6',
  muted:  '8B7070',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:4, offset:2, angle:45, opacity:0.15 });

// ─── SLIDE 1: COVER ───────────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  // texture lines
  for(let i=0;i<18;i++){
    s.addShape(pptx.ShapeType.line,{
      x: -0.5+i*0.7, y:0, w:4, h:0, line:{ color:'1A0F0F', width:18 }, rotate:35
    });
  }

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('DELIVERY HERO · AI QUALITY PLATFORM · Case Study 02',{
    x:0.35, y:0.28, w:7, h:0.22,
    fontSize:8, bold:true, color:C.muted, charSpacing:1.5, fontFace:'Calibri'
  });

  s.addText('AgentEval',{
    x:0.35, y:1.1, w:7, h:1.3,
    fontSize:72, bold:true, color:C.white, fontFace:'Calibri'
  });

  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:2.3, w:3.4, h:0.06, fill:{ color:C.red } });

  s.addText('AI Quality Intelligence for Every DH Agent — Detect, Diagnose, Fix.',{
    x:0.35, y:2.5, w:8.5, h:0.45,
    fontSize:18, color:C.gold, fontFace:'Calibri'
  });

  s.addText('Ajay Avaghade · Product Manager',{
    x:0.35, y:3.12, w:5, h:0.3,
    fontSize:11, color:C.muted, fontFace:'Calibri'
  });

  // bottom-right stat
  s.addShape(pptx.ShapeType.rect,{
    x:6.8, y:5.8, w:3.0, h:1.45,
    fill:{ color:C.hero }, line:{ color:C.red, width:1.5 },
    shadow: makeShadow()
  });
  s.addText('−64%',{
    x:6.8, y:5.85, w:3.0, h:0.55,
    fontSize:38, bold:true, color:C.gold, align:'center', fontFace:'Calibri'
  });
  s.addText('hallucination rate',{
    x:6.8, y:6.38, w:3.0, h:0.25,
    fontSize:11, color:C.white, align:'center', fontFace:'Calibri'
  });
  s.addText('FoodGenie v2.1 → v2.2 · Berlin pilot',{
    x:6.8, y:6.63, w:3.0, h:0.2,
    fontSize:8, color:C.muted, align:'center', fontFace:'Calibri'
  });

  s.addShape(pptx.ShapeType.ellipse,{ x:0.3, y:6.5, w:0.6, h:0.6, fill:{ color:C.red }, transparency:60 });
  s.addShape(pptx.ShapeType.ellipse,{ x:0.7, y:6.7, w:0.35, h:0.35, fill:{ color:C.gold }, transparency:70 });
}

// ─── SLIDE 2: THE PROBLEM ─────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('THE PROBLEM',{
    x:0.35, y:0.28, w:5, h:0.22,
    fontSize:9, bold:true, color:C.red, charSpacing:2, fontFace:'Calibri'
  });

  s.addText('DH Has 12+ AI Agents Live. No Unified Quality Layer.\nHallucinations Are Invisible Until Users Complain.',{
    x:0.35, y:0.62, w:9.3, h:1.1,
    fontSize:25, bold:true, color:C.white, fontFace:'Calibri', breakLine:true
  });

  const cols = [
    { icon:'🤖', stat:'12+', label:'AI agents in production', sub:'FoodGenie, OrderAgent, SupportBot, PricingAI — no shared quality framework.' },
    { icon:'🔍', stat:'0', label:'Unified quality monitoring', sub:'Each team tracks its own metrics in silos. Cross-agent patterns invisible.' },
    { icon:'💸', stat:'18%', label:'Avg hallucination rate', sub:'1-in-5 agent responses contain a factual error. Users lose trust silently.' },
  ];

  cols.forEach((c,i)=>{
    const x = 0.35 + i*3.2;
    s.addShape(pptx.ShapeType.rect,{
      x, y:2.0, w:3.0, h:2.75,
      fill:{ color:C.hero }, line:{ color:C.red, width:1 }, shadow: makeShadow()
    });
    s.addText(c.icon,{ x:x+0.15, y:2.1, w:0.6, h:0.5, fontSize:22 });
    s.addText(c.stat,{ x:x+0.15, y:2.58, w:2.7, h:0.7, fontSize:44, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(c.label,{ x:x+0.15, y:3.28, w:2.7, h:0.38, fontSize:12, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(c.sub,{ x:x+0.15, y:3.68, w:2.7, h:0.85, fontSize:10, color:C.muted, fontFace:'Calibri' });
  });

  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:5.0, w:9.3, h:1.15,
    fill:{ color:'1F0A0A' }, line:{ color:C.red, width:1 }
  });
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:5.0, w:0.05, h:1.15, fill:{ color:C.red } });
  s.addText('Root cause: AI teams ship fast but monitor slow. When FoodGenie hallucinates AED in a Turkish session, no system catches it — the user does. By then, 847 sessions have the same error and trust is eroding. The cost of not having a quality layer compounds every sprint.',{
    x:0.55, y:5.07, w:9.0, h:1.0,
    fontSize:11, color:C.white, fontFace:'Calibri'
  });
}

// ─── SLIDE 3: THE INSIGHT ─────────────────────────────────────────────────────
{
  const s = pptx.addSlide();
  s.background = { color: C.lgray };

  s.addShape(pptx.ShapeType.rect,{ x:0, y:0, w:0.06, h:7.5, fill:{ color:C.red } });

  s.addText('THE INSIGHT',{
    x:0.35, y:0.28, w:5, h:0.22,
    fontSize:9, bold:true, color:C.red, charSpacing:2, fontFace:'Calibri'
  });

  s.addText('AI Quality Needs Its Own Product — Not a Dashboard Bolted\nonto an Analytics Tool.',{
    x:0.35, y:0.62, w:9.3, h:1.0,
    fontSize:27, bold:true, color:C.dark, fontFace:'Calibri', breakLine:true
  });

  s.addShape(pptx.ShapeType.ellipse,{ x:4.7, y:2.0, w:0.6, h:0.6, fill:{ color:C.red } });
  s.addText('VS',{ x:4.7, y:2.05, w:0.6, h:0.5, fontSize:11, bold:true, color:C.white, align:'center', fontFace:'Calibri' });

  // left
  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:1.82, w:4.2, h:3.45,
    fill:{ color:'FDE8E8' }, line:{ color:'F59090', width:1 }
  });
  s.addText('❌  Status Quo',{
    x:0.5, y:1.92, w:3.9, h:0.35,
    fontSize:13, bold:true, color:'B91C1C', fontFace:'Calibri'
  });
  [
    'Each agent team monitors quality separately, in silos',
    'Hallucinations surface via user complaints (too late)',
    'No shared quality score language across product + eng',
    'A/B tests measure conversion — not agent quality',
    'Root cause analysis: manual, hours per incident',
  ].forEach((t,i)=>{
    s.addText('•  '+t,{
      x:0.5, y:2.36+i*0.52, w:3.9, h:0.45,
      fontSize:11, color:'374151', fontFace:'Calibri'
    });
  });

  // right
  s.addShape(pptx.ShapeType.rect,{
    x:5.45, y:1.82, w:4.2, h:3.45,
    fill:{ color:'ECFDF5' }, line:{ color:'6EE7B7', width:1 }
  });
  s.addText('✅  AgentEval',{
    x:5.6, y:1.92, w:3.9, h:0.35,
    fontSize:13, bold:true, color:'065F46', fontFace:'Calibri'
  });
  [
    'One quality score language across all 12+ agents',
    '3σ anomaly detection catches spikes before users do',
    'Session replay: see the exact turn where quality broke',
    'A/B test runner: quality score + conversion per variant',
    'Root cause classifier: 30-second triage, not 3-hour dig',
  ].forEach((t,i)=>{
    s.addText('•  '+t,{
      x:5.6, y:2.36+i*0.52, w:3.9, h:0.45,
      fontSize:11, color:'065F46', fontFace:'Calibri'
    });
  });

  // callout
  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:5.42, w:9.3, h:1.18,
    fill:{ color:'FFF3CD' }, line:{ color:C.gold, width:1.5 }
  });
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:5.42, w:0.05, h:1.18, fill:{ color:C.gold } });
  s.addText('Key insight: At PhonePe, replacing siloed ops workflows with a unified intelligence layer reduced TAT from 2 days → 4 hours and cut escalation tickets by 68%. AgentEval applies the same unification principle to AI quality: one layer that every team reads from and every PM acts on.',{
    x:0.55, y:5.49, w:9.0, h:1.0,
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
    fontSize:9, bold:true, color:C.gold, charSpacing:2, fontFace:'Calibri'
  });
  s.addText('Quality Score × 3σ Detection × Session Replay × Fix Loop',{
    x:0.35, y:0.62, w:9.3, h:0.65,
    fontSize:24, bold:true, color:C.white, fontFace:'Calibri'
  });

  const steps = [
    { n:'1', t:'Quality Score Engine', b:'Every agent response scored on 5 dimensions: Factual Accuracy, Intent Resolution, Relevance, Hallucination-Free rate, Tone + Safety. Composite score refreshes per session. Weighted by intent criticality (food safety > recommendation).' },
    { n:'2', t:'3σ Anomaly Classifier', b:'Rolling 7-day baseline per agent × market × intent. When a metric deviates >3 standard deviations, AgentEval fires a Critical alert within 4 minutes. Root cause label auto-applied: currency bleed, geo-mismatch, taxonomy error, etc.' },
    { n:'3', t:'Session Replay + Inspector', b:'Every flagged session is surfaced with: conversation turn-by-turn, quality scores per exchange, AgentEval failure annotation (shown inline in the message), and 1-click Jira + retrain actions.' },
    { n:'4', t:'A/B Test Runner', b:'Quality dimensions measured per variant — not just conversion. PM sees: "Variant B resolves 13.6% more intents AND hallucinates 11.4% less." Winner badge fires at p < 0.05. 1-click promote to production.' },
    { n:'5', t:'Fix → Verify Loop', b:'Recommended fixes (e.g., context isolation, geo-coord filter) are surfaced with est. impact and sprint cost. Post-deploy: AgentEval auto-runs regression check and marks fix as verified if metric returns to baseline.' },
  ];

  steps.forEach((step,i)=>{
    const y = 1.48 + i*1.04;
    s.addShape(pptx.ShapeType.ellipse,{ x:0.35, y:y+0.02, w:0.45, h:0.45, fill:{ color:C.red } });
    s.addText(step.n,{ x:0.35, y:y+0.04, w:0.45, h:0.4, fontSize:14, bold:true, color:C.white, align:'center', fontFace:'Calibri' });
    if(i<4){
      s.addShape(pptx.ShapeType.line,{
        x:0.575, y:y+0.47, w:0, h:0.57,
        line:{ color:C.muted, width:1, dashType:'dash' }
      });
    }
    s.addText(step.t,{ x:0.95, y:y, w:9.0, h:0.28, fontSize:13, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(step.b,{ x:0.95, y:y+0.28, w:8.9, h:0.65, fontSize:10.5, color:C.white, fontFace:'Calibri' });
  });

  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:6.85, w:9.3, h:0.42, fill:{ color:C.dark }, line:{ color:C.gold, width:0.8 } });
  s.addText('PhonePe proof: Rebuilt 6 siloed ops workflows into unified intelligence layer → TAT 2d→4h, −68% escalation tickets. AgentEval = same architectural principle applied to AI quality operations.',{
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
    x:0.35, y:0.28, w:5, h:0.22,
    fontSize:9, bold:true, color:C.red, charSpacing:2, fontFace:'Calibri'
  });
  s.addText('4 Screen States — Overview to Action',{
    x:0.35, y:0.62, w:9.3, h:0.55,
    fontSize:26, bold:true, color:C.dark, fontFace:'Calibri'
  });

  const cards = [
    {
      n:'01', t:'Quality Overview',
      b:'Live KPIs: avg quality score, hallucination rate, intent resolution, response time. Active anomalies ranked by severity. Quality score breakdown bar chart per agent. Market leaderboard — accountability surface for regional PMs.'
    },
    {
      n:'02', t:'Session Inspector',
      b:'Turn-by-turn replay of flagged session. Inline AgentEval annotations on failure turns. Quality scores per dimension. Root cause label with classifier confidence (91%). 1-click Jira bug creation + retrain flag.'
    },
    {
      n:'03', t:'Intent Heatmap',
      b:'Failure rate by intent × market. Color-coded cells (green <5%, amber 5–15%, red >15%). Click any cell → drill into session batch. Top failure patterns auto-ranked. Recommended fixes with estimated impact and sprint cost.'
    },
    {
      n:'04', t:'A/B Test Runner',
      b:'Side-by-side quality + conversion metrics per variant. Statistical significance badge (p-value shown). Winner declaration with GMV impact estimate. 1-click promote to production. Queued experiment pipeline.'
    },
  ];

  cards.forEach((c,i)=>{
    const x = 0.35 + i*2.4;
    s.addShape(pptx.ShapeType.rect,{
      x, y:1.35, w:2.25, h:5.65,
      fill:{ color:C.white }, line:{ color:'E5D5D5', width:1 }, shadow: makeShadow()
    });
    s.addShape(pptx.ShapeType.rect,{ x, y:1.35, w:2.25, h:0.05, fill:{ color:C.red } });
    s.addText(c.n,{ x:x+0.12, y:1.45, w:1.0, h:0.28, fontSize:9, bold:true, color:C.red, fontFace:'Calibri', charSpacing:1 });
    s.addText(c.t,{ x:x+0.12, y:1.72, w:2.0, h:0.52, fontSize:12, bold:true, color:C.dark, fontFace:'Calibri' });

    s.addShape(pptx.ShapeType.rect,{ x:x+0.12, y:2.32, w:2.0, h:2.2, fill:{ color:'1A0F0F' }, line:{ color:'331515', width:1 } });
    const mocks = [
      ['╔══════════╗','Avg Score: 76.4','─────────────','🔥 3 Anomalies','Hallucination 18%','Intent Res. 61%','🇩🇪82  🇸🇦79  🇹🇷61'],
      ['╔══════════╗','Session #TK-847291','─────────────','❌ AED in TR ctx','Score: 38% acc','91% confidence','[Jira] [Retrain]'],
      ['╔══════════╗','Intent × Market','─────────────','Cuisine  DE:4% TR:31','Location DE:3% TR:28','Fix: ctx isolate','−11.4% impact →'],
      ['╔══════════╗','EXP-041 v2.1/v2.2','─────────────','Resolution +13.6%','Hallucin. −11.4%','p=0.031 ✓','🏆 v2.2 → Prod'],
    ];
    mocks[i].forEach((line,li)=>{
      s.addText(line,{
        x:x+0.15, y:2.35+li*0.275, w:1.95, h:0.26,
        fontSize:7.5, color: li===0?C.red:'6EE7B7',
        fontFace:'Courier New', bold: li===0
      });
    });

    s.addText(c.b,{ x:x+0.12, y:4.62, w:2.0, h:2.25, fontSize:9.5, color:'4B3535', fontFace:'Calibri' });
  });

  s.addText('Prototype: dh-agenteval-prototype.html · All 4 states interactive · Click-based navigation + live heatmap',{
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
  s.addText('PhonePe Ops Intelligence Proof — Applied to AI Quality at DH Scale',{
    x:0.35, y:0.62, w:9.3, h:0.55,
    fontSize:22, bold:true, color:C.white, fontFace:'Calibri'
  });

  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:1.35, w:4.5, h:0.38, fill:{ color:C.red } });
  s.addText('AI TEAM / QUALITY OPS IMPACT',{
    x:0.35, y:1.37, w:4.5, h:0.34,
    fontSize:10, bold:true, color:C.white, align:'center', charSpacing:1.5, fontFace:'Calibri'
  });

  s.addShape(pptx.ShapeType.rect,{ x:5.15, y:1.35, w:4.5, h:0.38, fill:{ color:'2C1A00' }, line:{ color:C.gold, width:1 } });
  s.addText('BUSINESS / PRODUCT ROI',{
    x:5.15, y:1.37, w:4.5, h:0.34,
    fontSize:10, bold:true, color:C.gold, align:'center', charSpacing:1.5, fontFace:'Calibri'
  });

  const leftMetrics = [
    { v:'−64%', l:'Hallucination rate (v2.1 → v2.2)', s:'Context isolation fix + AgentEval-guided retraining · Berlin pilot result' },
    { v:'4 min', l:'Anomaly detection time (vs manual)', s:'3σ classifier fires within 4 minutes of spike vs hours via user complaint' },
    { v:'−68%', l:'Escalation tickets (PhonePe parallel)', s:'Unified intelligence layer replacing silos · Direct architecture carryover' },
    { v:'30 sec', l:'Root cause triage time per incident', s:'AgentEval classifier + session replay vs 2–3 hour manual investigation' },
  ];

  const rightMetrics = [
    { v:'+13.6%', l:'Session resolution rate (v2.2 pilot)', s:'+7.3 quality score pts → more intents resolved → higher order conversion' },
    { v:'+7.1%', l:'Order conversion (Genie A/B winner)', s:'€1.8M GMV/month uplift at Berlin scale · Quality improvement = revenue' },
    { v:'×1', l:'Shared layer across 12+ agents', s:'One AgentEval platform vs 12 separate monitoring setups = massive leverage' },
    { v:'2 wks', l:'Avg fix-to-verify cycle', s:'AgentEval guides fix + auto-verifies regression · Previously undefined cycle' },
  ];

  leftMetrics.forEach((m,i)=>{
    const y = 1.9 + i*1.25;
    s.addShape(pptx.ShapeType.rect,{ x:0.35, y, w:4.5, h:1.15, fill:{ color:C.hero }, line:{ color:'331515', width:1 } });
    s.addText(m.v,{ x:0.5, y:y+0.08, w:1.6, h:0.55, fontSize:34, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(m.l,{ x:0.5, y:y+0.6, w:4.1, h:0.25, fontSize:12, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s,{ x:0.5, y:y+0.83, w:4.1, h:0.28, fontSize:9, color:C.muted, fontFace:'Calibri' });
  });

  rightMetrics.forEach((m,i)=>{
    const y = 1.9 + i*1.25;
    s.addShape(pptx.ShapeType.rect,{ x:5.15, y, w:4.5, h:1.15, fill:{ color:'1F1200' }, line:{ color:'4A3300', width:1 } });
    s.addText(m.v,{ x:5.3, y:y+0.08, w:1.6, h:0.55, fontSize:34, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(m.l,{ x:5.3, y:y+0.6, w:4.1, h:0.25, fontSize:12, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(m.s,{ x:5.3, y:y+0.83, w:4.1, h:0.28, fontSize:9, color:C.muted, fontFace:'Calibri' });
  });

  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:6.9, w:9.3, h:0.38, fill:{ color:'100505' }, line:{ color:C.red, width:1 } });
  s.addText('AgentEval is not overhead — it is the multiplier. Every 1% reduction in hallucination rate in a high-volume agent compounds: more resolved intents → more conversions → more GMV. Quality IS the growth lever.',{
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

  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:1.35, w:4.5, h:4.6, fill:{ color:C.dark } });
  s.addShape(pptx.ShapeType.rect,{ x:0.35, y:1.35, w:4.5, h:0.32, fill:{ color:C.red } });
  s.addText('What I shipped at PhonePe',{
    x:0.5, y:1.37, w:4.2, h:0.28,
    fontSize:10, bold:true, color:C.white, fontFace:'Calibri'
  });

  [
    { h:'Unified Ops Intelligence Layer', b:'Replaced 6 siloed ops workflows with single intelligence console. Auto-classification of breach type, severity triage, escalation routing → TAT 2d→4h, −68% escalation tickets' },
    { h:'A/B Testing at Scale (350M MAU)', b:'Ran ML model A/B tests (propensity scoring vs manual cohorts) with shared measurement framework → 32% marketing burn reduction, sustained GMV growth' },
    { h:'Personalized Offers Redesign', b:'Redesigned offers eligibility surfacing with ML scoring → 22% checkout conversion lift. Quality of recommendation = quality of conversion.' },
    { h:'Rewards Marketplace Intelligence', b:'Built brand self-serve + audience segmentation platform with measurement layer → 500+ partners, 11% YoY revenue, 26% engagement increase' },
  ].forEach((item,i)=>{
    const y = 1.82 + i*1.02;
    s.addShape(pptx.ShapeType.ellipse,{ x:0.48, y:y+0.05, w:0.22, h:0.22, fill:{ color:C.red } });
    s.addText(item.h,{ x:0.82, y:y, w:3.85, h:0.28, fontSize:11, bold:true, color:C.gold, fontFace:'Calibri' });
    s.addText(item.b,{ x:0.82, y:y+0.28, w:3.85, h:0.62, fontSize:9.5, color:C.white, fontFace:'Calibri' });
  });

  s.addShape(pptx.ShapeType.rect,{ x:5.15, y:1.35, w:4.5, h:4.6, fill:{ color:'FDE8E8' } });
  s.addShape(pptx.ShapeType.rect,{ x:5.15, y:1.35, w:4.5, h:0.32, fill:{ color:'F59090' } });
  s.addText('How it maps to AgentEval at DH',{
    x:5.3, y:1.37, w:4.2, h:0.28,
    fontSize:10, bold:true, color:C.dark, fontFace:'Calibri'
  });

  [
    '→ AgentEval Quality Console: unified layer replacing per-team monitoring silos. Same TAT-compression principle applied to AI quality ops.',
    '→ AgentEval A/B Test Runner: quality score per variant + conversion, shared measurement framework across all 12+ DH agents.',
    '→ AgentEval 3σ Anomaly Classifier: same anomaly-detection logic that flagged ops breaches now flags hallucination spikes in real time.',
    '→ AgentEval Session Inspector + root cause label: same "classify breach type" architecture applied to agent failure triage.',
  ].forEach((t,i)=>{
    const y = 1.82 + i*1.02;
    s.addShape(pptx.ShapeType.rect,{ x:5.22, y:y+0.04, w:0.05, h:0.88, fill:{ color:'B91C1C' } });
    s.addText(t,{ x:5.4, y:y+0.04, w:4.1, h:0.88, fontSize:10.5, color:'1F0A0A', fontFace:'Calibri' });
  });

  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:6.1, w:9.3, h:1.1,
    fill:{ color:C.dark }, line:{ color:C.red, width:1.5 }
  });
  s.addText('"AgentEval isn\'t a new idea — it\'s the ops intelligence architecture I shipped at PhonePe, translated from human-ops to AI-ops. I know it works because I built the original."',{
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
  s.addText('Three Phases — FoodGenie Pilot to Platform for All DH Agents',{
    x:0.35, y:0.62, w:9.3, h:0.55,
    fontSize:23, bold:true, color:C.white, fontFace:'Calibri'
  });

  const phases = [
    {
      ph:'Phase 1',
      timeline:'Month 1–2',
      sub:'FoodGenie · Berlin',
      body:'Deploy AgentEval quality scoring on FoodGenie only. Activate 3σ anomaly detection for hallucination + intent resolution. Ship session inspector and root cause classifier. Run EXP-041 (v2.1 vs v2.2) with AgentEval as measurement layer. Measure: quality score δ, triage time δ vs baseline.'
    },
    {
      ph:'Phase 2',
      timeline:'Month 3–4',
      sub:'3 Agents · 5 Markets',
      body:'Extend AgentEval to OrderAgent + SupportBot. Add intent heatmap — cross-agent failure pattern visibility for first time. Launch A/B test runner as self-serve tool for AI PMs. Begin cross-agent quality benchmarking. Target: all active anomalies resolved within 48h of detection.'
    },
    {
      ph:'Phase 3',
      timeline:'Month 5–6',
      sub:'All 12+ Agents · All Markets',
      body:'AgentEval becomes the shared quality layer for every DH AI product. CEO dashboard: global quality score + market rankings. Integrate with model deployment pipeline — no agent ships without AgentEval baseline passing. Platform becomes DH\'s quality moat.'
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
    s.addText(p.ph,{ x:x+0.15, y:1.52, w:2.7, h:0.35, fontSize:18, bold:true, color:C.white, fontFace:'Calibri' });
    s.addText(p.timeline,{ x:x+0.15, y:1.88, w:2.7, h:0.28, fontSize:11, bold:true, color: i===0?C.red:i===1?C.gold:C.green, fontFace:'Calibri' });
    s.addText(p.sub,{ x:x+0.15, y:2.14, w:2.7, h:0.28, fontSize:11, color:C.muted, fontFace:'Calibri' });
    s.addShape(pptx.ShapeType.line,{ x:x+0.15, y:2.47, w:2.6, h:0, line:{ color:'331515', width:0.8 } });
    s.addText(p.body,{ x:x+0.15, y:2.55, w:2.7, h:3.4, fontSize:10.5, color:C.white, fontFace:'Calibri' });
  });

  s.addShape(pptx.ShapeType.rect,{
    x:0.35, y:6.2, w:9.3, h:0.9,
    fill:{ color:'1F1200' }, line:{ color:C.gold, width:1.5 }
  });
  s.addText('What I need to ship this:',{
    x:0.55, y:6.27, w:3, h:0.28,
    fontSize:11, bold:true, color:C.gold, fontFace:'Calibri'
  });
  s.addText('Agent response log access (anonymized)  ·  1 ML engineer (quality scoring + 3σ classifier)  ·  Alignment with AI Platform team + 2 agent PMs',{
    x:0.55, y:6.55, w:9.0, h:0.42,
    fontSize:10, color:C.white, fontFace:'Calibri'
  });

  s.addText('Ajay Avaghade  |  avaghadeajay009@gmail.com  |  +91 9561558439  |  linkedin.com/in/ajay-avaghade',{
    x:0.35, y:7.2, w:9.3, h:0.22,
    fontSize:8.5, color:C.muted, align:'center', fontFace:'Calibri'
  });
}

// ─── SAVE ─────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/dh-agenteval-deck.pptx' })
  .then(()=>console.log('✅  dh-agenteval-deck.pptx saved'))
  .catch(e=>{ console.error(e); process.exit(1); });
