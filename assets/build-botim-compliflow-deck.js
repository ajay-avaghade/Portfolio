(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.defineLayout({ name:'LAYOUT_16x9', width:10, height:7.5 });
prs.layout = 'LAYOUT_16x9';

const C = {
  dark:   '0A1628', hero:   '0F1E3A',
  purple: '7C3AED', lavender:'A78BFA',
  green:  '00C853', orange: 'FF6B35',
  amber:  'F59E0B', white:  'FFFFFF',
  lgray:  'F0F4F8', muted:  '8b92a5',
};

// S1 COVER
const s1 = prs.addSlide(); s1.background = { color: C.dark };
s1.addText('BOTIM FOR BUSINESS · KYB & AML ONBOARDING', { x:.3, y:.4, w:9.4, h:.3, fontSize:9, bold:true, color:C.muted, align:'center' });
s1.addText('CompliFlow', { x:.5, y:1.2, w:9, h:1.5, fontSize:52, bold:true, color:C.white, align:'center' });
s1.addText('Frictionless KYB & AML Onboarding — 5–7 Days to Under 2 Hours', { x:.5, y:2.8, w:9, h:.8, fontSize:20, color:C.lavender, align:'center' });
s1.addText('Ajay Avaghade · Product Manager', { x:.5, y:3.8, w:9, h:.4, fontSize:12, color:C.muted, align:'center' });
s1.addShape('rect', { x:7.5, y:5.2, w:2, h:.1, fill:{ color:C.purple }, line:{ type:'none' } });
s1.addText('60% drop-off eliminated', { x:7.0, y:5.4, w:2.8, h:.8, fontSize:16, bold:true, color:C.lavender, align:'right' });

// S2 PROBLEM
const s2 = prs.addSlide(); s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s2.addText('60% of SMEs Drop Off Before First Transaction Due to 5–7 Day Manual KYB — Compliance Is Killing Conversion', { x:.5, y:.8, w:9, h:1.3, fontSize:26, bold:true, color:C.white });
[{ num:'5–7 days', label:'Manual KYB TAT', sub:'trade license review + AML check + sanctions screening' },
 { num:'60%+', label:'Drop-off rate', sub:'SMEs abandon before completing KYB onboarding' },
 { num:'85%', label:'Low-risk businesses', sub:'treated identically to high-risk — wasting compliance resources' },
].forEach((s, i) => {
  const x = .5 + i * 3.1;
  s2.addText(s.num, { x, y:2.3, w:2.8, h:.6, fontSize:28, bold:true, color:C.lavender, align:'center' });
  s2.addText(s.label, { x, y:3.0, w:2.8, h:.4, fontSize:12, bold:true, color:C.white, align:'center' });
  s2.addText(s.sub, { x, y:3.5, w:2.8, h:.8, fontSize:10, color:C.muted, align:'center' });
});
s2.addShape('rect', { x:.5, y:5.2, w:9, h:1.7, fill:{ color:C.hero }, line:{ color:C.purple, width:1 } });
s2.addText('Root cause: Every business goes through the same full manual KYB regardless of actual risk. Compliance teams spend 85% of time on businesses that could be auto-approved in minutes. High-risk businesses that need scrutiny wait behind a queue of low-risk ones.', { x:.7, y:5.35, w:8.6, h:1.4, fontSize:10, color:C.white });

// S3 INSIGHT
const s3 = prs.addSlide(); s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s3.addText('Risk-Tiering Compliance Turns a Conversion Killer into a Competitive Advantage', { x:.5, y:.8, w:9, h:1.0, fontSize:26, bold:true, color:C.dark });
s3.addShape('rect', { x:.5, y:2.0, w:4.2, h:3.8, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
s3.addText('❌ Today: Flat KYB', { x:.7, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.orange });
['All businesses: same 5–7 day process','Manual document review per case','AML + sanctions checked in serial','60%+ drop-off before first payment','Compliance team overwhelmed'].forEach((t,i) => s3.addText(t, { x:.9, y:2.8+i*.58, w:3.6, h:.5, fontSize:10, color:C.dark }));
s3.addShape('rect', { x:5.3, y:2.0, w:4.2, h:3.8, fill:{ color:C.hero }, line:{ color:C.purple, width:1 } });
s3.addText('✅ CompliFlow: Tiered', { x:5.5, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.lavender });
['LOW (84%): instant approval <2 hrs','MED (11%): human review 24h','HIGH (5%): full manual 5–7 days','OCR auto-extracts doc fields','Compliance team handles only high-risk'].forEach((t,i) => s3.addText(t, { x:5.7, y:2.8+i*.58, w:3.6, h:.5, fontSize:10, color:C.white }));
s3.addShape('ellipse', { x:4.6, y:3.5, w:.8, h:.8, fill:{ color:C.dark }, line:{ color:C.purple, width:2 } });
s3.addText('VS', { x:4.6, y:3.55, w:.8, h:.7, fontSize:14, bold:true, color:C.lavender, align:'center' });

// S4 MECHANIC
const s4 = prs.addSlide(); s4.background = { color: C.dark };
s4.addText('THE MECHANIC', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s4.addText('OCR Extraction → ML Risk Score → Tier Routing → Straight-Through or Human Review', { x:.5, y:.8, w:9, h:.5, fontSize:20, bold:true, color:C.white });
[{ n:'1', t:'Document Upload', d:'Trade license + MOA uploaded. OCR extracts all fields automatically.' },
 { n:'2', t:'Auto-Checks', d:'Parallel: OFAC, UN, UAE Central Bank sanctions. UBO PEP screening. License validity.' },
 { n:'3', t:'Risk Scoring', d:'ML model outputs score 0–100. LOW (<30): straight-through. MED (30–70): human review. HIGH (>70): full KYB.' },
 { n:'4', t:'Tier Routing', d:'LOW: instant approval <2h. MED: 24h queue. HIGH: 5–7 days with escalation path.' },
 { n:'5', t:'Continuous Learning', d:'Approval/rejection data retrains model quarterly. Threshold calibrated per market.' },
].forEach((s, i) => {
  const y = 1.6 + i * .84;
  s4.addShape('ellipse', { x:.5, y, w:.4, h:.4, fill:{ color:C.purple }, line:{ type:'none' } });
  s4.addText(s.n, { x:.5, y, w:.4, h:.4, fontSize:14, bold:true, color:C.white, align:'center' });
  s4.addText(s.t, { x:1.1, y, w:2.3, h:.22, fontSize:10, bold:true, color:C.white });
  s4.addText(s.d, { x:1.1, y:y+.25, w:2.3, h:.28, fontSize:8, color:C.muted });
  if (i < 4) s4.addShape('line', { x:.7, y:y+.45, w:0, h:.24, line:{ color:C.purple, width:2, dashType:'dash' } });
});
s4.addShape('rect', { x:3.8, y:1.6, w:5.7, h:4.3, fill:{ color:C.hero }, line:{ color:C.purple, width:1 } });
s4.addText('Proof: PhonePe Self-Serve Merchant Platform (5K merchants)', { x:4.0, y:1.8, w:5.3, h:.4, fontSize:10, bold:true, color:C.lavender });
['Built zero-to-one self-serve B2B merchant onboarding platform','Eliminated manual ops bottleneck (6 weeks → 2 days onboarding TAT)','5,000+ merchants acquired with 23% lower CAC, zero eng dependency','Standardised onboarding scaled across 25+ cities','CompliFlow: same "eliminate the manual bottleneck" architecture for KYB'].forEach((t,i) => s4.addText(t, { x:4.2, y:2.4+i*.62, w:5.1, h:.56, fontSize:9, color:C.white }));

// S5 PRODUCT
const s5 = prs.addSlide(); s5.background = { color: C.lgray };
s5.addText('THE PRODUCT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s5.addText('4 Screens — Onboarding to Compliance Intelligence', { x:.5, y:.8, w:9, h:.5, fontSize:24, bold:true, color:C.dark });
[{ n:'01', t:'Onboarding Wizard', d:'Trade license upload (OCR auto-fills all fields). UBO declaration. Expected volume. CompliFlow check initiated in one tap.' },
 { n:'02', t:'Risk Assessment', d:'Risk score 0–100. Checks completed: OFAC, UN, CBUAE, UBO PEP. GREEN = instant. AMBER = human review. RED = escalation.' },
 { n:'03', t:'Instant Approval', d:'Confetti success. "Approved in <2 hours." Payment limits unlocked. Corridors available. Next step CTA.' },
 { n:'04', t:'Compliance Ops Console', d:'Straight-through rate (84%), avg TAT (1.8h), human review queue (11%), daily KYB pipeline. PEP match alerts.' },
].forEach((s, i) => {
  const y = 1.6 + i * 1.35;
  s5.addShape('rect', { x:.5, y, w:9, h:1.2, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
  s5.addText(`${s.n} — ${s.t}`, { x:.7, y:y+.15, w:3.5, h:.3, fontSize:11, bold:true, color:C.dark });
  s5.addText(s.d, { x:.7, y:y+.5, w:8.6, h:.55, fontSize:9, color:C.dark });
});

// S6 ROI
const s6 = prs.addSlide(); s6.background = { color: C.dark };
s6.addText('IMPACT & ROI', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s6.addText('Onboarding Conversion & Compliance Efficiency', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ ico:'🚀', v:'<2 hrs', l:'LOW risk approval TAT', s:'vs 5–7 days manual (96% faster for 84% of SMEs)' },
 { ico:'💰', v:'−60%', l:'Drop-off rate eliminated', s:'straight-through conversion vs current 60%+ abandonment' },
 { ico:'👮', v:'84%', l:'Straight-through rate', s:'of businesses auto-approved — no compliance team time' },
 { ico:'📈', v:'+40%', l:'Compliance capacity freed', s:'team redirected from low-risk to high-risk cases' },
].forEach((m, i) => {
  const x = .5 + (i%2)*4.6, y = 1.6 + Math.floor(i/2)*2;
  s6.addShape('rect', { x, y, w:4, h:1.8, fill:{ color:C.hero }, line:{ color:C.purple, width:1 } });
  s6.addText(m.ico, { x:x+.2, y:y+.1, w:3.6, h:.4, fontSize:22, align:'left' });
  s6.addText(m.v, { x:x+.2, y:y+.52, w:3.6, h:.38, fontSize:20, bold:true, color:C.lavender, align:'left' });
  s6.addText(m.l, { x:x+.2, y:y+.96, w:3.6, h:.28, fontSize:10, bold:true, color:C.white, align:'left' });
  s6.addText(m.s, { x:x+.2, y:y+1.28, w:3.6, h:.38, fontSize:8, color:C.muted, align:'left' });
});

// S7 PROOF
const s7 = prs.addSlide(); s7.background = { color: C.lgray };
s7.addText('PROOF OF WORK', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s7.addText('PhonePe Self-Serve Platform: Eliminate the Manual Bottleneck', { x:.5, y:.8, w:9, h:.6, fontSize:24, bold:true, color:C.dark });
s7.addShape('rect', { x:.5, y:1.6, w:4.3, h:4.4, fill:{ color:C.dark }, line:{ color:C.purple, width:1 } });
s7.addText('PhonePe: What I Built', { x:.7, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.lavender });
['B2B merchant platform: manual ops was the bottleneck','Built self-serve PG integration (Instant Discount + EMI)','Eliminated back-and-forth with ops team per merchant','5,000+ merchants, 6 weeks → 2 days, −23% CAC','Scaled zero engineering overhead across 25+ cities'].forEach((t,i) => s7.addText(t, { x:.9, y:2.3+i*.58, w:3.7, h:.52, fontSize:9, color:C.white }));
s7.addShape('rect', { x:5.2, y:1.6, w:4.3, h:4.4, fill:{ color:'FFFFFF' }, line:{ color:C.purple, width:1 } });
s7.addText('CompliFlow: Application', { x:5.4, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.purple });
['KYB manual review was the SME onboarding bottleneck','OCR + ML risk scoring eliminates manual per-case review','84% of SMEs get straight-through: <2 hours, zero ops','Compliance team redirected to the 5% who truly need them','Architecture: same "self-serve removes the bottleneck" pattern'].forEach((t,i) => s7.addText(t, { x:5.4, y:2.3+i*.58, w:3.9, h:.52, fontSize:9, color:C.dark }));

// S8 ROLLOUT
const s8 = prs.addSlide(); s8.background = { color: C.dark };
s8.addText('ROLLOUT PLAN', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s8.addText('Three Phases — UAE Pilot to Full MENA KYB Coverage', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ t:'Phase 1 — UAE Pilot (Month 1–2)', b:'Deploy CompliFlow for UAE-registered businesses only. Measure: straight-through rate, TAT, false-positive rate on LOW/MED/HIGH classification. Compare vs manual baseline.' },
 { t:'Phase 2 — MENA Expansion (Month 3–4)', b:'Extend CompliFlow to Saudi Arabia, Egypt, Qatar registered entities. Add local regulatory lists per market. Compliance team SLA calibrated per tier.' },
 { t:'Phase 3 — Full Scale (Month 5–6)', b:'All cross-border business onboarding runs through CompliFlow. Annual re-verification automated. Model retraining quarterly on outcomes data.' },
].forEach((p, i) => {
  const y = 1.5 + i * 1.75;
  s8.addShape('rect', { x:.5, y, w:9, h:1.6, fill:{ color:C.hero }, line:{ color:C.purple, width:1 } });
  s8.addText(p.t, { x:.7, y:y+.15, w:8.6, h:.3, fontSize:11, bold:true, color:C.lavender });
  s8.addText(p.b, { x:.9, y:y+.52, w:8.2, h:.85, fontSize:9, color:C.white });
});
s8.addShape('rect', { x:.5, y:6.5, w:9, h:.8, fill:{ color:C.hero }, line:{ color:C.orange, width:2 } });
s8.addText('What I Need: Sanctions list API access (OFAC, UN, CBUAE) · OCR vendor (AWS Textract or equivalent) · ML engineer (risk scoring model) · Compliance team sign-off on tier thresholds', { x:.7, y:6.6, w:8.6, h:.6, fontSize:10, bold:true, color:C.orange });
s8.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x:.5, y:7.15, w:9, h:.25, fontSize:9, color:C.muted, align:'center' });

await prs.writeFile({ fileName: 'assets/botim-compliflow-deck.pptx' });
console.log('✓ botim-compliflow-deck.pptx created');
})();
