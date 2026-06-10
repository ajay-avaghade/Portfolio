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
slide1.addText("HALA FLEXICAP", { ...defaultProps, x: 0.5, y: 2.0, w: 9, h: 0.8, fontSize: 44, bold: true, color: teal });
slide1.addText("Embedded Working Capital for MENAP SMEs", { ...defaultProps, x: 0.5, y: 3.0, w: 9, h: 0.5, fontSize: 24, color: "FFFFFF" });
slide1.addText("Prepared by: Ajay Avaghade", { ...defaultProps, x: 0.5, y: 4.5, w: 9, h: 0.4, fontSize: 16, color: lightGray });

// Slide 2: The Problem
let slide2 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide2.addText("THE SME WORKING CAPITAL GAP", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide2.addText([
  { text: "1. Underbanked Core\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Micro-businesses and freelancers lack the extensive credit history required for conventional bank loans.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "2. Operational Disconnect\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Traditional lending sits outside day-to-day operations, requiring manual applications and static monthly repayments regardless of daily sales.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "3. Cash Flow Volatility\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Retail SMEs experience massive weekly volatility; fixed EMI commitments create systemic default risk.", options: { color: lightGray, fontSize: 16 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 3: The Solution
let slide3 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide3.addText("HALA FLEXICAP: FINANCING AS A LAYER", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide3.addText([
  { text: "Dynamic Underwriting\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Credit limits are pre-approved instantly based on 30-day trailing digital payment volumes (POS/App sales).\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "Frictionless Drawdown\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Merchants can advance funds in a single click directly into their Hala merchant wallet.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "Volume-Linked Repayment\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "No fixed monthly EMI. Hala auto-deducts a flexible percentage (e.g., 10%) from daily settled POS sales until cleared.", options: { color: lightGray, fontSize: 16 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 4: Go-to-Market Strategy
let slide4 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide4.addText("GO-TO-MARKET & ADOPTION", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide4.addText([
  { text: "Phase 1: High-Volume Pilot\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Target the top 15% of Hala merchants by POS volume. Offer \"0% flat fee\" on the first advance to seed adoption.\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "Phase 2: Contextual Triggers\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Fire push notifications when POS sales spike. \"Your weekend volume is up 40%. Need SAR 10K to restock inventory?\"\n\n", options: { color: lightGray, fontSize: 16 } },
  { text: "Phase 3: Deep Embedded Logic\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "Integrate with Hala PayFlow to auto-suggest advances when upcoming vendor payouts exceed current wallet balance.", options: { color: lightGray, fontSize: 16 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 5: Success Metrics
let slide5 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide5.addText("MEASURING PRODUCT PERFORMANCE", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide5.addText([
  { text: "Primary Success Metrics:\n\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "• Drawdown Rate: % of eligible merchants who activate a FlexiCap advance.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• Repayment Velocity: Average days to clear an advance based on dynamic deductions.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• NPL (Non-Performing Loan) Ratio: Volume of defaults over 90 days vs. active book.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• POS Stickiness: Increase in POS usage and transaction volume post-advance activation.", options: { color: lightGray, fontSize: 18 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

// Slide 6: Product Alignment
let slide6 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide6.addText("ALIGNMENT WITH HALA'S VISION", { ...defaultProps, x: 0.5, y: 0.5, w: 9, h: 0.6, fontSize: 32, bold: true, color: teal });
slide6.addText([
  { text: "Why This Matters Now:\n\n", options: { bold: true, color: "FFFFFF", fontSize: 20 } },
  { text: "• \"Rethinking SME Financing\": FlexiCap explicitly abandons the \"conventional banking product\" model, embedding the loan entirely into daily POS activities.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• Risk Mitigation: By securing repayment directly from the source of revenue (Hala payments), risk is inherently minimized.\n", options: { color: lightGray, fontSize: 18 } },
  { text: "• Ecosystem Lock-in: Merchants are heavily incentivized to process 100% of their sales through Hala POS to access higher limits.", options: { color: lightGray, fontSize: 18 } }
], { ...defaultProps, x: 0.5, y: 1.5, w: 9, h: 3.5 });

pptx.writeFile({ fileName: "hala-flexicap-deck.pptx" }).then(() => {
    console.log("hala-flexicap-deck.pptx created successfully!");
}).catch(err => {
    console.error("Error creating pptx:", err);
});
