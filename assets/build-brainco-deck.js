const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0B1120',
  hero:   '1E293B',
  bright: '0EA5E9',
  orange: 'F59E0B',
  white:  'F8FAFC',
  lgray:  'F1F5F9',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// SLIDE 1: Cover
let s1 = pptx.addSlide();
s1.background = { color: C.dark };
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.2, fill: { color: C.bright } });
s1.addText('BrainTrust', { x: 0.5, y: 2.2, w: 9, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
s1.addText('Bridging the Gap Between AI Models and Enterprise Trust', { x: 0.5, y: 3.2, w: 9, fontSize: 24, color: C.bright, fontFace: 'Arial' });
s1.addText('Ajay Avaghade · Product Manager', { x: 0.5, y: 6.5, w: 5, fontSize: 14, color: C.white, fontFace: 'Arial' });
s1.addText('Brain Co. Case Study', { x: 0.5, y: 6.8, w: 5, fontSize: 12, color: C.bright, fontFace: 'Arial', bold: true });
s1.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: 1.5, w: 3, h: 3, fill: { color: C.hero } });
s1.addText('94%', { x: 6.5, y: 2.2, w: 3, fontSize: 44, bold: true, color: C.bright, align: 'center', fontFace: 'Arial' });
s1.addText('Compliance Goal', { x: 6.5, y: 3.2, w: 3, fontSize: 14, color: C.white, align: 'center', fontFace: 'Arial' });

// SLIDE 2: The Problem
let s2 = pptx.addSlide();
s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: C.muted || '64748B', bold: true, charSpacing: 2 });
s2.addText('Black-Box AI Models Are a Liability for Mission-Critical Clients', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });
s2.addText('Client hesitation isn\'t about AI capabilities—it\'s about explainability and regulatory risk.', { x: 0.5, y: 1.8, w: 9, fontSize: 16, color: C.white });

s2.addText('80%', { x: 0.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Adoption Drop-off', { x: 0.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Government clients resist using models they cannot explain to auditors.', { x: 0.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('2 Weeks', { x: 3.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.bright });
s2.addText('Compliance Lag', { x: 3.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Every model output requires manual checking against legal guidelines.', { x: 3.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('0%', { x: 6.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Stakeholder Trust', { x: 6.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Without transparent metrics, trust in AI systems rapidly deteriorates.', { x: 6.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addShape(pptx.ShapeType.rect, { x: 0.5, y: 6.0, w: 9, h: 1.0, fill: { color: C.hero }, shadow: makeShadow() });
s2.addText('Insight: Adoption accelerates when we shift from "Trust the AI" to "Verify the Rules".', { x: 0.7, y: 6.2, w: 8.6, fontSize: 14, bold: true, color: C.bright });

// SLIDE 3: The Insight
let s3 = pptx.addSlide();
s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s3.addText('From Black-Box Engineering to Plain-English Compliance', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.2, w: 4.2, h: 3.5, fill: { color: 'E2E8F0' }, shadow: makeShadow() });
s3.addText('Current State', { x: 0.8, y: 2.5, w: 3.6, fontSize: 18, bold: true, color: C.dark });
s3.addText('❌ Raw JSON decision logs\n❌ Heavy reliance on Data Scientists to explain outputs\n❌ Compliance is a reactive, manual process', { x: 0.8, y: 3.2, w: 3.6, fontSize: 14, color: C.dark, lineSpacing: 28 });

s3.addShape(pptx.ShapeType.ellipse, { x: 4.6, y: 3.5, w: 0.8, h: 0.8, fill: { color: C.hero }, shadow: makeShadow() });
s3.addText('VS', { x: 4.6, y: 3.7, w: 0.8, fontSize: 14, bold: true, color: C.white, align: 'center' });

s3.addShape(pptx.ShapeType.rect, { x: 5.3, y: 2.2, w: 4.2, h: 3.5, fill: { color: C.white }, shadow: makeShadow(), line: { color: C.bright, width: 2 } });
s3.addText('Proposed State (BrainTrust)', { x: 5.6, y: 2.5, w: 3.6, fontSize: 18, bold: true, color: C.bright });
s3.addText('✅ ML logic translated to Plain English\n✅ Non-technical officials can self-serve audits\n✅ Compliance checks are automated in the loop', { x: 5.6, y: 3.2, w: 3.6, fontSize: 14, color: C.dark, lineSpacing: 28 });

// SLIDE 4: The Mechanic
let s4 = pptx.addSlide();
s4.background = { color: C.dark };
s4.addText('THE MECHANIC', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s4.addText('How BrainTrust Secures the Deployment Pipeline', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });

s4.addShape(pptx.ShapeType.line, { x: 1.0, y: 3.5, w: 8.0, h: 0, line: { color: C.bright, width: 2, dashType: 'dash' } });

[
  { x: 0.5, t: '1. Model Executes', d: 'Core AI makes a prediction or classification.' },
  { x: 2.5, t: '2. Confidence Capture', d: 'System logs feature weights and certainty.' },
  { x: 4.5, t: '3. English Translation', d: 'ML factors converted to natural language.' },
  { x: 6.5, t: '4. Compliance Check', d: 'Output vetted against rigid legal constraints.' },
  { x: 8.5, t: '5. Exec Sign-Off', d: 'Clear audit trail delivered to stakeholders.' }
].forEach((st, i) => {
  s4.addShape(pptx.ShapeType.ellipse, { x: st.x+0.2, y: 3.2, w: 0.6, h: 0.6, fill: { color: C.bright } });
  s4.addText((i+1)+'', { x: st.x+0.2, y: 3.3, w: 0.6, fontSize: 14, bold: true, color: C.dark, align: 'center' });
  s4.addText(st.t, { x: st.x-0.2, y: 4.0, w: 1.4, fontSize: 12, bold: true, color: C.white, align: 'center' });
  s4.addText(st.d, { x: st.x-0.4, y: 4.5, w: 1.8, fontSize: 10, color: '94A3B8', align: 'center' });
});

s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 6.0, w: 9, h: 0.8, fill: { color: C.hero } });
s4.addText('Validation: We tested this framework on PhonePe\'s marketplace, reducing ops escalations by 40%.', { x: 0.7, y: 6.2, w: 8.6, fontSize: 12, color: C.white });

// SLIDE 5: The Product
let s5 = pptx.addSlide();
s5.background = { color: C.lgray };
s5.addText('THE PRODUCT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s5.addText('BrainTrust UX Overview', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

[
  { x: 0.5, y: 2.0, t: '01. Deployment Dashboard', d: 'High-level view of model health and compliance rates.' },
  { x: 5.0, y: 2.0, t: '02. Explainer Interface', d: 'Translates raw JSON logs into plain-English justifications.' },
  { x: 0.5, y: 4.5, t: '03. Compliance Audit', d: 'Checks model decisions against government regulations.' },
  { x: 5.0, y: 4.5, t: '04. Stakeholder Sign-Off', d: 'Generates non-technical reports for executive approval.' }
].forEach(c => {
  s5.addShape(pptx.ShapeType.rect, { x: c.x, y: c.y, w: 4.5, h: 2.0, fill: { color: C.white }, shadow: makeShadow() });
  s5.addText(c.t, { x: c.x+0.2, y: c.y+0.2, w: 4.1, fontSize: 16, bold: true, color: C.dark });
  s5.addText(c.d, { x: c.x+0.2, y: c.y+0.7, w: 4.1, fontSize: 12, color: '64748B' });
});

// SLIDE 6: Impact & ROI
let s6 = pptx.addSlide();
s6.background = { color: C.dark };
s6.addText('IMPACT & ROI', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s6.addText('Accelerating Adoption, Reducing Enterprise Friction', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });

s6.addText('Client Impact (Government / Enterprise)', { x: 0.5, y: 2.0, w: 4.5, fontSize: 16, bold: true, color: C.bright });
s6.addText('99% Audit Pass Rate', { x: 0.5, y: 2.5, w: 4.0, fontSize: 24, bold: true, color: C.white });
s6.addText('Automated guardrails ensure continuous compliance.', { x: 0.5, y: 3.0, w: 4.0, fontSize: 12, color: '94A3B8' });

s6.addText('10x Review Velocity', { x: 0.5, y: 4.0, w: 4.0, fontSize: 24, bold: true, color: C.white });
s6.addText('Plain-English explanations eliminate DS bottleneck.', { x: 0.5, y: 4.5, w: 4.0, fontSize: 12, color: '94A3B8' });

s6.addText('Brain Co. Impact', { x: 5.5, y: 2.0, w: 4.5, fontSize: 16, bold: true, color: C.orange });
s6.addText('Faster Deployment Lifecycle', { x: 5.5, y: 2.5, w: 4.0, fontSize: 24, bold: true, color: C.white });
s6.addText('Self-serve audits drastically cut post-sales integration time.', { x: 5.5, y: 3.0, w: 4.0, fontSize: 12, color: '94A3B8' });

s6.addText('Higher NRR / Renewals', { x: 5.5, y: 4.0, w: 4.0, fontSize: 24, bold: true, color: C.white });
s6.addText('Trusted AI leads to expanded client scopes and seats.', { x: 5.5, y: 4.5, w: 4.0, fontSize: 12, color: '94A3B8' });

// SLIDE 7: Proof of Work
let s7 = pptx.addSlide();
s7.background = { color: C.white };
s7.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s7.addText('Why Ajay Is Equipped to Build This', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s7.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.dark } });
s7.addText('What I Built at PhonePe', { x: 0.8, y: 2.3, w: 3.7, fontSize: 16, bold: true, color: C.white });
s7.addText('• Scaled Propensity-to-Transact ML to 350M+ users.\n• Translated complex DS models into actionable marketing rules.\n• Built zero-to-one B2B SaaS workflow tools, onboarding 5,000+ enterprise merchants.', { x: 0.8, y: 3.0, w: 3.7, fontSize: 13, color: '94A3B8', lineSpacing: 24 });

s7.addShape(pptx.ShapeType.rect, { x: 5.2, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.lgray }, shadow: makeShadow() });
s7.addText('How It Maps to Brain Co.', { x: 5.5, y: 2.3, w: 3.7, fontSize: 16, bold: true, color: C.dark });
s7.addText('• Native understanding of enterprise stakeholder management.\n• High technical fluency with ML deployments and data pipelines.\n• Ability to ship agile product delivery cycles for non-technical enterprise end-users.', { x: 5.5, y: 3.0, w: 3.7, fontSize: 13, color: C.dark, lineSpacing: 24 });

// SLIDE 8: Rollout Plan
let s8 = pptx.addSlide();
s8.background = { color: C.dark };
s8.addText('ROLLOUT PLAN', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s8.addText('Next Steps & Execution', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });

s8.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 2.8, h: 2.5, fill: { color: C.hero } });
s8.addText('Phase 1: Discovery', { x: 0.7, y: 2.2, w: 2.4, fontSize: 14, bold: true, color: C.white });
s8.addText('Audit existing models and interview GCC public sector clients on compliance blockers.', { x: 0.7, y: 2.8, w: 2.4, fontSize: 12, color: '94A3B8' });

s8.addShape(pptx.ShapeType.rect, { x: 3.6, y: 2.0, w: 2.8, h: 2.5, fill: { color: C.hero } });
s8.addText('Phase 2: MVP', { x: 3.8, y: 2.2, w: 2.4, fontSize: 14, bold: true, color: C.white });
s8.addText('Deploy the Plain-English Translation engine on one low-risk NLP model as a pilot.', { x: 3.8, y: 2.8, w: 2.4, fontSize: 12, color: '94A3B8' });

s8.addShape(pptx.ShapeType.rect, { x: 6.7, y: 2.0, w: 2.8, h: 2.5, fill: { color: C.hero } });
s8.addText('Phase 3: Scale', { x: 6.9, y: 2.2, w: 2.4, fontSize: 14, bold: true, color: C.bright });
s8.addText('Standardize the BrainTrust dashboard for all incoming government deployments.', { x: 6.9, y: 2.8, w: 2.4, fontSize: 12, color: '94A3B8' });

s8.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.0, w: 9, h: 1.2, fill: { color: C.white } });
s8.addText('Contact: Ajay Avaghade | avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 5.4, w: 9, fontSize: 14, bold: true, color: C.dark, align: 'center' });

pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/brainco-deck.pptx' });
