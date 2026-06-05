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
  green:  '10B981',
};

const makeShadow = () => ({ type:'outer', color:'000000', blur:3, offset:1, angle:45, opacity:0.12 });

// SLIDE 1: Cover
let s1 = pptx.addSlide();
s1.background = { color: C.dark };
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.2, fill: { color: C.bright } });
s1.addText('BrainSync', { x: 0.5, y: 2.2, w: 9, fontSize: 44, bold: true, color: C.white, fontFace: 'Arial' });
s1.addText('Enterprise Data Ingestion & PII Redaction Pipeline', { x: 0.5, y: 3.2, w: 9, fontSize: 24, color: C.bright, fontFace: 'Arial' });
s1.addText('Ajay Avaghade · Product Manager', { x: 0.5, y: 6.5, w: 5, fontSize: 14, color: C.white, fontFace: 'Arial' });
s1.addText('Brain Co. Case Study 2', { x: 0.5, y: 6.8, w: 5, fontSize: 12, color: C.bright, fontFace: 'Arial', bold: true });
s1.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: 1.5, w: 3, h: 3, fill: { color: C.hero } });
s1.addText('100%', { x: 6.5, y: 2.2, w: 3, fontSize: 44, bold: true, color: C.green, align: 'center', fontFace: 'Arial' });
s1.addText('Zero PII Leakage', { x: 6.5, y: 3.2, w: 3, fontSize: 14, color: C.white, align: 'center', fontFace: 'Arial' });

// SLIDE 2: The Problem
let s2 = pptx.addSlide();
s2.background = { color: C.dark };
s2.addText('THE PROBLEM', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s2.addText('Data Silos and Security Blockers', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.white });
s2.addText('AI models are useless if government clients can\'t securely feed them data.', { x: 0.5, y: 1.8, w: 9, fontSize: 16, color: C.white });

s2.addText('6 Months', { x: 0.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Security Review', { x: 0.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Infosec teams block data transfers due to massive PII liabilities.', { x: 0.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('Manual', { x: 3.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('ETL Pipelines', { x: 3.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Relying on engineers to custom-code redaction for every deployment.', { x: 3.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

s2.addText('Fragmented', { x: 6.5, y: 3.0, w: 2.5, fontSize: 40, bold: true, color: C.orange });
s2.addText('Legacy DBs', { x: 6.5, y: 3.8, w: 2.5, fontSize: 16, bold: true, color: C.white });
s2.addText('Government data is trapped across on-prem Postgres, S3, and mainframes.', { x: 6.5, y: 4.2, w: 2.5, fontSize: 13, color: '94A3B8' });

// SLIDE 3: The Insight
let s3 = pptx.addSlide();
s3.background = { color: C.lgray };
s3.addText('THE INSIGHT', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s3.addText('Empower IT Admins with a Visual "Clean Room"', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });
s3.addText('By building a self-serve B2B workflow, we abstract complex data engineering so enterprise clients can safely ingest their own data.', { x: 0.5, y: 1.8, w: 9, fontSize: 14, color: C.dark });

s3.addShape(pptx.ShapeType.rect, { x: 0.5, y: 3.0, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('1. Auto PII Detection', { x: 0.8, y: 3.5, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('The system automatically flags SSNs, Names, and Addresses in the schema.', { x: 0.8, y: 4.0, w: 3.4, fontSize: 13, color: '64748B' });

s3.addShape(pptx.ShapeType.rect, { x: 5.0, y: 3.0, w: 4.0, h: 2.5, fill: { color: C.white }, shadow: makeShadow() });
s3.addText('2. One-Click Redaction', { x: 5.3, y: 3.5, w: 3.4, fontSize: 16, bold: true, color: C.dark });
s3.addText('Apply tokenization or drop policies visually, without writing SQL or Python.', { x: 5.3, y: 4.0, w: 3.4, fontSize: 13, color: '64748B' });

// SLIDE 4: Proof of Work
let s4 = pptx.addSlide();
s4.background = { color: C.white };
s4.addText('PROOF OF WORK', { x: 0.5, y: 0.5, w: 3, fontSize: 10, color: '64748B', bold: true, charSpacing: 2 });
s4.addText('Why Ajay Is Equipped to Build This', { x: 0.5, y: 1.0, w: 9, fontSize: 28, bold: true, color: C.dark });

s4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.dark } });
s4.addText('B2B Workflows at PhonePe', { x: 0.8, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.white });
s4.addText('• Built a zero-to-one B2B SaaS workflow builder.\n• Enabled 5,000+ merchants to configure complex payment logic without coding.\n• Reduced integration time from 6 weeks to 2 days.', { x: 0.8, y: 3.2, w: 3.7, fontSize: 13, color: '94A3B8', lineSpacing: 24 });

s4.addShape(pptx.ShapeType.rect, { x: 5.2, y: 2.0, w: 4.3, h: 4.0, fill: { color: C.lgray } });
s4.addText('Mapping to Brain Co.', { x: 5.5, y: 2.5, w: 3.7, fontSize: 16, bold: true, color: C.dark });
s4.addText('• Proven ability to abstract backend engineering into intuitive B2B UX.\n• Understands how to navigate enterprise infosec blockers.\n• Accelerates deployment by empowering the customer.', { x: 5.5, y: 3.2, w: 3.7, fontSize: 13, color: C.dark, lineSpacing: 24 });

// SLIDE 5: Contact
let s5 = pptx.addSlide();
s5.background = { color: C.dark };
s5.addText('Brain Co. Case Study 2: BrainSync', { x: 0.5, y: 2.5, w: 9, fontSize: 24, bold: true, color: C.white, align: 'center' });
s5.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.5, w: 9, h: 1.2, fill: { color: C.white } });
s5.addText('Contact: Ajay Avaghade | avaghadeajay009@gmail.com | linkedin.com/in/ajay-avaghade', { x: 0.5, y: 4.9, w: 9, fontSize: 14, bold: true, color: C.dark, align: 'center' });

pptx.writeFile({ fileName: '/Users/ajayavaghade/Portfolio/assets/brainco-sync-deck.pptx' });
