const PptxGenJS = require("pptxgenjs");

let pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";
pptx.defineSlideMaster({
  title: "MASTER_SLIDE",
  background: { color: "F5F5F5" },
  objects: [
    { rect: { x: 0, y: 0, w: "100%", h: 0.8, fill: { color: "141414" } } },
    { text: { text: "Mastercard Marketing AI • AI-Gen Offer", options: { x: 0.5, y: 0.2, w: 5, h: 0.4, color: "FFFFFF", fontSize: 14, fontFace: "Arial", bold: true } } },
    { text: { text: "AJAY AVAGHADE", options: { x: 8, y: 0.2, w: 4, h: 0.4, color: "F79E1B", fontSize: 12, align: "right", fontFace: "Arial" } } },
    { rect: { x: 0, y: 5.1, w: "100%", h: 0.5, fill: { color: "141414" } } },
    { text: { text: "CONFIDENTIAL & PROPRIETARY", options: { x: 0.5, y: 5.2, w: 4, h: 0.3, color: "FFFFFF", fontSize: 10, fontFace: "Arial" } } }
  ]
});

// SLIDE 1: Title
let slide1 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide1.addText("Mastercard AI-Gen Offer Engine", { x: 1, y: 2, w: 8, h: 1, fontSize: 36, bold: true, color: "141414", fontFace: "Arial" });
slide1.addText("Hyper-Targeted Generative Campaigns at Scale", { x: 1, y: 3, w: 8, h: 0.5, fontSize: 18, color: "EB001B", fontFace: "Arial" });
slide1.addText("Case Study for Manager, Product Management Marketing AI", { x: 1, y: 3.5, w: 8, h: 0.5, fontSize: 14, color: "666666", fontFace: "Arial" });

// SLIDE 2: Problem
let slide2 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide2.addText("THE ENGAGEMENT PROBLEM", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide2.addText(
  "Current Challenge:\n• Issuers struggle to personalize marketing campaigns at scale.\n• Static, generic offers lead to low activation and spend stimulation.\n• Manual creative generation and segmentation takes weeks, driving up CAC.",
  { x: 0.5, y: 1.8, w: 8, h: 2, fontSize: 16, color: "333333", bullet: true, fontFace: "Arial", lineSpacing: 24 }
);

// SLIDE 3: Solution
let slide3 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide3.addText("THE SOLUTION: AI-GEN OFFER", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide3.addText(
  "A next-generation personalization engine:\n• Predictive ML Modeling: Identifies high-propensity segments.\n• Generative Creative: Instantly generates tailored images and copy.\n• Omnichannel Delivery: Automatically selects best channels (Push, SMS, Social).\n• Automated Execution: Compresses campaign deployment from weeks to minutes.",
  { x: 0.5, y: 1.8, w: 8, h: 2, fontSize: 16, color: "333333", bullet: true, fontFace: "Arial", lineSpacing: 24 }
);

// SLIDE 4: Workflow
let slide4 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide4.addText("HOW IT WORKS", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide4.addShape(pptx.ShapeType.rect, { x: 0.5, y: 2, w: 2.5, h: 1.5, fill: "EB001B", color: "FFFFFF", align: "center", bold: true, line: "141414", text: "1. AI Segmentation\n(Identify highest propensity audience)"});
slide4.addShape(pptx.ShapeType.rightArrow, { x: 3.1, y: 2.5, w: 0.8, h: 0.5, fill: "CCCCCC" });
slide4.addShape(pptx.ShapeType.rect, { x: 4.0, y: 2, w: 2.5, h: 1.5, fill: "F79E1B", color: "FFFFFF", align: "center", bold: true, line: "141414", text: "2. Gen-AI Creative\n(Dynamic images & tailored copy)"});
slide4.addShape(pptx.ShapeType.rightArrow, { x: 6.6, y: 2.5, w: 0.8, h: 0.5, fill: "CCCCCC" });
slide4.addShape(pptx.ShapeType.rect, { x: 7.5, y: 2, w: 2.5, h: 1.5, fill: "2196F3", color: "FFFFFF", align: "center", bold: true, line: "141414", text: "3. Omnichannel Delivery\n(Real-time push, API optimization)"});

// SLIDE 5: Metrics & Impact
let slide5 = pptx.addSlide({ masterName: "MASTER_SLIDE" });
slide5.addText("EXPECTED IMPACT", { x: 0.5, y: 1, w: 8, h: 0.5, fontSize: 24, bold: true, color: "141414", fontFace: "Arial" });
slide5.addText(
  "Based on previous implementation at PhonePe scale:\n• +22% Uplift in Checkout / Offer Conversions\n• -32% Reduction in overall Marketing Spend\n• 4h vs 3w Campaign Turnaround Time\n• Enables Mastercard to be the digital marketing partner of choice.",
  { x: 0.5, y: 1.8, w: 8, h: 2, fontSize: 16, color: "333333", bullet: true, fontFace: "Arial", lineSpacing: 24 }
);

pptx.writeFile({ fileName: "mastercard-aigen-offer-deck.pptx" }).then(fileName => {
    console.log(`Created file: ${fileName}`);
});
