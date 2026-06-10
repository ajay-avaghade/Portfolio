const PptxGenJS = require("pptxgenjs");

let pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";
pptx.defineSlideMaster({
  title: "MASTER_SLIDE",
  background: { color: "F5F5F5" },
  objects: [
    { rect: { x: 0, y: 0, w: "100%", h: 0.8, fill: { color: "1E90FF" } } },
    { text: { text: "Mastercard Marketing AI • Ecosystem Acquire", options: { x: 0.5, y: 0.2, w: 6, h: 0.4, color: "FFFFFF", fontSize: 14, fontFace: "Arial", bold: true } } },
    { text: { text: "AJAY AVAGHADE", options: { x: 8, y: 0.2, w: 4, h: 0.4, color: "FFFFFF", fontSize: 12, align: "right", fontFace: "Arial" } } },
    { rect: { x: 0, y: 5.1, w: "100%", h: 0.5, fill: { color: "141414" } } },
    { text: { text: "CONFIDENTIAL & PROPRIETARY", options: { x: 0.5, y: 5.2, w: 4, h: 0.3, color: "FFFFFF", fontSize: 10, fontFace: "Arial" } } }
  ]
});

// SLIDE 1: Title
let slide1 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide1.addText("Mastercard Ecosystem Acquire", { x: 1, y: 2, w: 8, h: 1, fontSize: 36, bold: true, color: "141414", fontFace: "Arial" });
slide1.addText("B2B2C ML-Driven Partner Acquisition Engine", { x: 1, y: 3, w: 8, h: 0.5, fontSize: 18, color: "1E90FF", fontFace: "Arial" });
slide1.addText("Case Study for Manager, Product Management Marketing AI", { x: 1, y: 3.5, w: 8, h: 0.5, fontSize: 14, color: "666666", fontFace: "Arial" });

// SLIDE 2: Problem
let slide2 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide2.addText("THE ACQUISITION CHALLENGE", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide2.addText(
  "Current Challenge:\n• Scaling acquisition beyond traditional financial institutions is disjointed.\n• Huge overlapping audiences between Issuers and non-traditional segments (Telcos, large retailers).\n• Inability to cross-pollinate data securely to drive co-branded card adoption.",
  { x: 0.5, y: 1.8, w: 8, h: 2, fontSize: 16, color: "333333", bullet: true, fontFace: "Arial", lineSpacing: 24 }
);

// SLIDE 3: Solution
let slide3 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide3.addText("THE SOLUTION: ECOSYSTEM ACQUIRE", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide3.addText(
  "A machine learning-driven B2B2C acquisition platform:\n• Maps overlapping audiences (e.g., Telco subscriber vs Issuer cardholder).\n• Automates cross-acquisition campaigns across partner channels.\n• Triggers EMOB (Early Month On Book) activation nudges automatically.\n• Utilizes ML models to predict drop-off and trigger spend stimulation offers.",
  { x: 0.5, y: 1.8, w: 8, h: 2, fontSize: 16, color: "333333", bullet: true, fontFace: "Arial", lineSpacing: 24 }
);

// SLIDE 4: Workflow
let slide4 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide4.addText("AUTOMATED LIFECYCLE MANAGEMENT", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2, w: 2.5, h: 1.5, fill: "1E90FF", color: "FFFFFF", align: "center", bold: true, line: "141414", text: "1. Cross-Acquisition\n(Targeting overlapping Telco segment)"});
slide4.addShape(pptx.ShapeType.rightArrow, { x: 3.1, y: 2.5, w: 0.8, h: 0.5, fill: "CCCCCC" });
slide4.addShape(pptx.ShapeType.rect, { x: 4.0, y: 2, w: 2.5, h: 1.5, fill: "F79E1B", color: "FFFFFF", align: "center", bold: true, line: "141414", text: "2. EMOB Activation\n(Automated 30-day first-swipe nudges)"});
slide4.addShape(pptx.ShapeType.rightArrow, { x: 6.6, y: 2.5, w: 0.8, h: 0.5, fill: "CCCCCC" });
slide4.addShape(pptx.ShapeType.rect, { x: 7.5, y: 2, w: 2.5, h: 1.5, fill: "EB001B", color: "FFFFFF", align: "center", bold: true, line: "141414", text: "3. Spend Stimulation\n(ML-triggered retention discounts)"});

// SLIDE 5: Metrics & Impact
let slide5 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide5.addText("EXPECTED IMPACT", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide5.addText(
  "Business Outcomes:\n• Unlocks massive new scalable marketing plays for non-traditional partners.\n• Drives incremental growth for Mastercard clients in EEMEA.\n• +22% EMOB activation lift through automated, reward-based triggers.\n• Demonstrates product leadership in unifying the partner ecosystem.",
  { x: 0.5, y: 1.8, w: 8, h: 2, fontSize: 16, color: "333333", bullet: true, fontFace: "Arial", lineSpacing: 24 }
);

pptx.writeFile({ fileName: "mastercard-ecosystem-acquire-deck.pptx" }).then(fileName => {
    console.log(`Created file: ${fileName}`);
});
