(async () => {
const PptxGenJS = require('pptxgenjs');
const prs = new PptxGenJS();
prs.defineLayout({ name:'LAYOUT_16x9', width:10, height:7.5 });
prs.layout = 'LAYOUT_16x9';
const C = { dark:'0F0A1E', hero:'1A1030', bright:'10B981', accent:'6EE7B7', indigo:'6366F1', pink:'EC4899', orange:'F97316', white:'FFFFFF', lgray:'F8F9FC', muted:'8b92a5' };

// S1 COVER
const s1 = prs.addSlide(); s1.background = { color:C.dark };
s1.addText('PENCIL · EDITOR SUITE · CASE STUDY 03', { x:.3, y:.4, w:9.4, h:.3, fontSize:9, bold:true, color:C.muted, align:'center' });
s1.addText('PrimitivesKit', { x:.5, y:1.1, w:9, h:1.5, fontSize:48, bold:true, color:C.white, align:'center' });
s1.addText('Shared Component Library — Define Once, Render Across All Formats', { x:.5, y:2.7, w:9, h:.8, fontSize:20, color:C.accent, align:'center' });
s1.addText('Ajay Avaghade · Product Manager', { x:.5, y:3.7, w:9, h:.4, fontSize:12, color:C.muted, align:'center' });
s1.addShape('rect', { x:7.5, y:5.1, w:2, h:.1, fill:{ color:C.bright }, line:{ type:'none' } });
s1.addText('−62% duplicate eng work', { x:6.8, y:5.3, w:2.9, h:.8, fontSize:15, bold:true, color:C.accent, align:'right' });

// S2 PROBLEM
const s2 = prs.addSlide(); s2.background = { color:C.dark };
s2.addText('THE PROBLEM', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s2.addText('Four Editor Formats. Four Implementations of the Same Button. Users Relearn. Eng Duplicates. Platform Fragments.', { x:.5, y:.8, w:9, h:1.2, fontSize:26, bold:true, color:C.white });
[{ num:'4×', label:'implementations of CTA Button', sub:'Email, Display, Print, Digital — each different, none shared' },
 { num:'30%', label:'of UX is relearned per format', sub:'Email editor pro must relearn Display editor interactions' },
 { num:'4×', label:'the design change cost', sub:'update a button style: 4 separate engineering sprints required' },
].forEach((s,i) => {
  const x=.5+i*3.1;
  s2.addText(s.num, { x, y:2.2, w:2.8, h:.6, fontSize:32, bold:true, color:C.accent, align:'center' });
  s2.addText(s.label, { x, y:2.9, w:2.8, h:.4, fontSize:12, bold:true, color:C.white, align:'center' });
  s2.addText(s.sub, { x, y:3.4, w:2.8, h:.8, fontSize:10, color:C.muted, align:'center' });
});
s2.addShape('rect', { x:.5, y:5.1, w:9, h:1.7, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s2.addText('Root cause: Each format editor was built independently. Text blocks, buttons, image containers — every primitive is reimplemented per format. Users experience fragmentation. Engineering pays 4× the cost for every change. Brand token updates require coordinated multi-team releases. The editor suite feels like 4 tools, not 1.', { x:.7, y:5.3, w:8.6, h:1.3, fontSize:10, color:C.white });

// S3 INSIGHT
const s3 = prs.addSlide(); s3.background = { color:C.lgray };
s3.addText('THE INSIGHT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s3.addText('Format Differences Are Rendering Rules — Not Separate Products', { x:.5, y:.8, w:9, h:.9, fontSize:28, bold:true, color:C.dark });
s3.addShape('rect', { x:.5, y:2.0, w:4.2, h:3.8, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
s3.addText('❌ Today: 4 Products', { x:.7, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.orange });
['Email editor has its own TextBlock','Display editor has its own TextBlock','Print editor has its own TextBlock','User moves from Email to Display: "Why is text alignment different?"','Eng: 4 sprint tickets for one button change','Design: 4 Figma files to update'].forEach((t,i) => s3.addText(t, { x:.9, y:2.8+i*.48, w:3.6, h:.44, fontSize:10, color:C.dark }));
s3.addShape('rect', { x:5.3, y:2.0, w:4.2, h:3.8, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s3.addText('✅ PrimitivesKit: 1 Product', { x:5.5, y:2.2, w:3.8, h:.4, fontSize:12, bold:true, color:C.accent });
['TextBlock defined once with format-aware rendering','Email: 16px; Display: 14px; Print: 12pt — all from 1 definition','User learns TextBlock once; works the same everywhere','Eng: 1 sprint ticket for one button change, all formats update','Design: 1 token file. Update once, propagates globally.'].forEach((t,i) => s3.addText(t, { x:5.7, y:2.8+i*.48, w:3.6, h:.44, fontSize:10, color:C.white }));
s3.addShape('ellipse', { x:4.6, y:3.5, w:.8, h:.8, fill:{ color:C.dark }, line:{ color:C.bright, width:2 } });
s3.addText('VS', { x:4.6, y:3.55, w:.8, h:.7, fontSize:14, bold:true, color:C.accent, align:'center' });

// S4 MECHANIC
const s4 = prs.addSlide(); s4.background = { color:C.dark };
s4.addText('THE MECHANIC', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s4.addText('5 Primitives × 4 Formats × 1 Shared Definition', { x:.5, y:.8, w:9, h:.5, fontSize:24, bold:true, color:C.white });
[{ n:'1', t:'Primitive Contract', d:'PM defines component API: props, constraints, accessibility rules. One spec, all formats.' },
 { n:'2', t:'Format Rendering Rules', d:'Engineering implements once per primitive: Email (44px touch, mobile reflow), Display (scale), Print (SVG/CMYK).' },
 { n:'3', t:'Brand Token Layer', d:'Colour, font, spacing tokens apply globally. Change --brand-primary once → all primitives in all formats update.' },
 { n:'4', t:'Designer Workflow', d:'Designer uses one component library. Selects TextBlock + chooses format preview. No format-specific file switching.' },
 { n:'5', t:'Engineering Handoff', d:'Spec sheet per primitive: props, format constraints, accessibility, performance budget. Build once, maintain once.' },
].forEach((s,i) => {
  const y=1.6+i*.84;
  s4.addShape('ellipse', { x:.5, y, w:.4, h:.4, fill:{ color:C.bright }, line:{ type:'none' } });
  s4.addText(s.n, { x:.5, y, w:.4, h:.4, fontSize:14, bold:true, color:C.dark, align:'center' });
  s4.addText(s.t, { x:1.1, y, w:2.3, h:.22, fontSize:10, bold:true, color:C.white });
  s4.addText(s.d, { x:1.1, y:y+.25, w:2.3, h:.28, fontSize:8, color:C.muted });
  if(i<4) s4.addShape('line', { x:.7, y:y+.45, w:0, h:.24, line:{ color:C.bright, width:2, dashType:'dash' } });
});
s4.addShape('rect', { x:3.8, y:1.6, w:5.7, h:4.3, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
s4.addText('Proof: PhonePe Rewards Marketplace — 500 Partners on One Primitive Set', { x:4.0, y:1.8, w:5.3, h:.4, fontSize:10, bold:true, color:C.accent });
['Rebuilt ₹100Cr static offer portfolio: one primitive library for 500+ brand partners','Offer card primitives: headline slot, image slot, CTA, audience targeting block','Each partner launched independently using the same component vocabulary','Changing offer card CTA style: 1 template update → all 500 partner campaigns updated','PrimitivesKit is that system applied to editor components instead of offer cards'].forEach((t,i) => s4.addText(t, { x:4.2, y:2.4+i*.62, w:5.1, h:.56, fontSize:9, color:C.white }));

// S5 PRODUCT
const s5 = prs.addSlide(); s5.background = { color:C.lgray };
s5.addText('THE PRODUCT', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s5.addText('4 Screens — Library to Handoff', { x:.5, y:.8, w:9, h:.5, fontSize:26, bold:true, color:C.dark });
[{ n:'01', t:'Component Library', d:'5 primitives (TextBlock, CTAButton, ImageContainer, Spacer, BrandToken). Each with format preview swatches. One vocabulary, all formats.' },
 { n:'02', t:'Format-Aware Rendering', d:'Same CTAButton primitive shown in Email (44px, full-width), Display (auto-scale), Print (SVG vector). User sees constraints are handled — not their problem.' },
 { n:'03', t:'Brand Token System', d:'Colour, typography, spacing tokens defined once. Change --brand-primary → all components in all formats update instantly. Single source of truth.' },
 { n:'04', t:'Engineering Handoff', d:'Primitive spec sheet: props, format constraints, accessibility rules, performance budget. Implement once per primitive, not once per format.' },
].forEach((s,i) => {
  const y=1.6+i*1.35;
  s5.addShape('rect', { x:.5, y, w:9, h:1.2, fill:{ color:'FFFFFF' }, line:{ color:C.muted, width:1 } });
  s5.addText(`${s.n} — ${s.t}`, { x:.7, y:y+.15, w:3.5, h:.3, fontSize:11, bold:true, color:C.dark });
  s5.addText(s.d, { x:.7, y:y+.5, w:8.6, h:.55, fontSize:9, color:C.dark });
});

// S6 ROI
const s6 = prs.addSlide(); s6.background = { color:C.dark };
s6.addText('IMPACT & ROI', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s6.addText('Engineering Velocity, UX Consistency & Platform Cohesion', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ ico:'⚡', v:'−62%', l:'Duplicate engineering work', s:'5 primitives × 4 formats: build 5× not 20×' },
 { ico:'🎨', v:'1 update', l:'Brand token change propagation', s:'Update --brand-primary once → all 5 primitives × 4 formats update' },
 { ico:'📚', v:'−30%', l:'User relearning between formats', s:'Same vocabulary across Email, Display, Print, Digital' },
 { ico:'🚀', v:'4×', l:'Design change velocity', s:'Figma: 1 component file vs 4. Design update ships to all formats simultaneously.' },
].forEach((m,i) => {
  const x=.5+(i%2)*4.6, y=1.6+Math.floor(i/2)*2;
  s6.addShape('rect', { x, y, w:4, h:1.8, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
  s6.addText(m.ico, { x:x+.2, y:y+.1, w:3.6, h:.4, fontSize:22, align:'left' });
  s6.addText(m.v, { x:x+.2, y:y+.52, w:3.6, h:.38, fontSize:20, bold:true, color:C.accent, align:'left' });
  s6.addText(m.l, { x:x+.2, y:y+.96, w:3.6, h:.28, fontSize:10, bold:true, color:C.white, align:'left' });
  s6.addText(m.s, { x:x+.2, y:y+1.28, w:3.6, h:.38, fontSize:8, color:C.muted, align:'left' });
});

// S7 PROOF
const s7 = prs.addSlide(); s7.background = { color:C.lgray };
s7.addText('PROOF OF WORK', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s7.addText('PhonePe Rewards Marketplace: 500 Partners, One Primitive Set', { x:.5, y:.8, w:9, h:.6, fontSize:22, bold:true, color:C.dark });
s7.addShape('rect', { x:.5, y:1.6, w:4.3, h:4.4, fill:{ color:C.dark }, line:{ color:C.bright, width:1 } });
s7.addText('PhonePe: What I Built', { x:.7, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.accent });
['Rebuilt ₹100Cr static reward portfolio: one shared offer card system','500+ brand partners launched campaigns using same component vocabulary','Headline slot, image slot, CTA button, audience targeting — defined once','Changing CTA style: 1 template update → all 500 partner campaigns updated','Designed the primitive contracts; Engineering implemented per primitive, not per partner'].forEach((t,i) => s7.addText(t, { x:.9, y:2.3+i*.58, w:3.7, h:.52, fontSize:9, color:C.white }));
s7.addShape('rect', { x:5.2, y:1.6, w:4.3, h:4.4, fill:{ color:'FFFFFF' }, line:{ color:C.bright, width:1 } });
s7.addText('PrimitivesKit: Application', { x:5.4, y:1.8, w:3.9, h:.3, fontSize:11, bold:true, color:C.bright });
['5 primitives replace 20 format-specific implementations','Each primitive has format-aware rendering rules','Brand token change propagates instantly across all formats','Users learn one vocabulary used everywhere in the editor suite','Engineering: implement once per primitive, maintain once, ship everywhere'].forEach((t,i) => s7.addText(t, { x:5.4, y:2.3+i*.58, w:3.9, h:.52, fontSize:9, color:C.dark }));

// S8 ROLLOUT
const s8 = prs.addSlide(); s8.background = { color:C.dark };
s8.addText('ROLLOUT PLAN', { x:.3, y:.3, w:9.4, h:.25, fontSize:9, bold:true, color:C.muted });
s8.addText('Three Phases — CTAButton Pilot to Full 5-Primitive Coverage', { x:.5, y:.8, w:9, h:.5, fontSize:22, bold:true, color:C.white });
[{ t:'Phase 1 — CTAButton Primitive (Month 1–2)', b:'Define CTAButton as the first shared primitive. Implement Email + Display rendering rules. A/B: shared primitive vs current implementations. Measure: bug rate, design change velocity, user confusion reports.' },
 { t:'Phase 2 — TextBlock + ImageContainer (Month 3–4)', b:'Migrate TextBlock (typography token integration) and ImageContainer (crop zone rules) to primitive model. Brand token layer MVP: --brand-primary, --font-heading live across both formats.' },
 { t:'Phase 3 — Full PrimitivesKit (Month 5–6)', b:'All 5 primitives migrated. Print format rendering rules added. Token system covers colour, typography, spacing. Engineering handoff spec standardised across all future primitives.' },
].forEach((p,i) => {
  const y=1.5+i*1.75;
  s8.addShape('rect', { x:.5, y, w:9, h:1.6, fill:{ color:C.hero }, line:{ color:C.bright, width:1 } });
  s8.addText(p.t, { x:.7, y:y+.15, w:8.6, h:.3, fontSize:11, bold:true, color:C.accent });
  s8.addText(p.b, { x:.9, y:y+.52, w:8.2, h:.85, fontSize:9, color:C.white });
});
s8.addShape('rect', { x:.5, y:6.5, w:9, h:.8, fill:{ color:C.hero }, line:{ color:C.orange, width:2 } });
s8.addText('What I Need: Design audit of current format-specific components · Eng audit of duplicate implementations · Token naming convention agreed · 1 eng sprint per primitive', { x:.7, y:6.6, w:8.6, h:.6, fontSize:10, bold:true, color:C.orange });
s8.addText('Ajay Avaghade | avaghadeajay009@gmail.com | +91 9561558439', { x:.5, y:7.15, w:9, h:.25, fontSize:9, color:C.muted, align:'center' });

await prs.writeFile({ fileName:'assets/pencil-primitiveskit-deck.pptx' });
console.log('✓ pencil-primitiveskit-deck.pptx created');
})();
