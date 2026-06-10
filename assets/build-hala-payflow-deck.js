const pptxgen = require("pptxgenjs");
let pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";
pptx.defineSlideMaster({
  title: "MASTER_SLIDE",
  background: { color: "0F172A" },
  objects: [
    { rect: { x: 0, y: 0, w: "100%", h: "100%", fill: { color: "0F172A" } } }
  ]
});

const defaultProps = { fontFace: "Helvetica", color: "FFFFFF", valign: "top" };
const teal = "0D9488";
const lightGray = "94A3B8";

// Slide 1: Title
let slide1 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide1.addText("HALA PAYFLOW", { ...defaultProps, x: 0.5, y: 2.0, w: 9, h: 0.8, fontSize: 44, bold: true, color: teal });
slide1.addText("Unified Operations & Cash Flow Predictor", { ...defaultProps, x: 0.5, y: 3.0, w: 9, h: 0.5, fontSize: 24, color: "FFFFFF" });
slide1.addText("Prepared by: Ajay Avaghade", { ...defaultProps, x: 0.5, y: 4.5, w: 9, h: 0.4, fontSize: 16, color: lightGray });

// Slide 2: The Problem
let slide2 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide2.addText("THE RECONCILIATION NIGHTMARE", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide2.addText([
  { text: "1. Fragmented Ecosystem\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "SMEs use one system for POS sales, another for vendor payouts, and a third for accounting, leading to blind spots.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "2. Reactive Cash Flow Management\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Merchants discover they don't have enough capital for a supplier payment on the day it is due.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "3. Disconnected Financing\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "When deficits occur, securing emergency capital is a slow, manual process outside the core operational workflow.", options: { color: lightGray, fontSize: 16 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 3: The Solution
let slide3 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide3.addText("HALA PAYFLOW: THE COMMAND CENTER", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide3.addText([
  { text: "Unified Sales & Payouts\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "All incoming POS settlements and outgoing scheduled vendor payouts reside in one dashboard.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "Cash Flow Predictor\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "An algorithmic 7-day forecast matching projected incoming sales against scheduled outgoing payments.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "Actionable Smart Alerts\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "If Thursday's vendor payout exceeds projected balances, PayFlow triggers an alert on Monday.", options: { color: lightGray, fontSize: 16 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 4: Driving Lending Adoption
let slide4 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide4.addText("PAYFLOW AS A LENDING FUNNEL", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide4.addText([
  { text: "Context is King\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "A generic \"Get a Loan\" banner has low conversion. An alert saying \"You are short SAR 12,000 for Thursday's rent payment\" provides extreme context.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "One-Click Remediation\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Embedded directly inside the deficit alert is a button: \"Advance SAR 12,000 via Hala FlexiCap\".\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "Operational Layering\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "This transforms lending from a standalone product into a critical operational feature that saves the merchant's supply chain.", options: { color: lightGray, fontSize: 16 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 5: Success Metrics
let slide5 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide5.addText("MEASURING PLATFORM SUCCESS", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide5.addText([
  { text: "Primary Success Metrics:\n\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "• Dashboard MAU: Frequency of merchants logging into the PayFlow web portal weekly.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• Alert-to-Advance Conversion: % of merchants who draw down FlexiCap funds upon receiving a deficit alert.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• Payout Volume: Total volume of vendor payments processed through the platform.", options: { color: lightGray, fontSize: 18 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 6: Product Alignment
let slide6 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide6.addText("ALIGNMENT WITH HALA'S VISION", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide6.addText([
  { text: "The Future Bank of SMEs:\n\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "• \"Empowering SMEs to run and grow\": By automating reconciliation, we give merchants their time back to focus on growth.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• Synergistic Flywheel: PayFlow drives engagement, which drives POS volume, which fuels the data models for FlexiCap lending.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• Stickiness: A merchant using PayFlow to manage their entire supply chain payout schedule will never switch POS providers.", options: { color: lightGray, fontSize: 18 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

pptx.writeFile({ fileName: "hala-payflow-deck.pptx" }).then(() => {
    console.log("hala-payflow-deck.pptx created successfully!");
}).catch(err => {
    console.error("Error creating pptx:", err);
});
