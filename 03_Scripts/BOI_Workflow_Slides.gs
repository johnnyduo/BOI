/**
 * BOI WORKFLOW VISUALIZATION - Google Slides Generator
 *
 * Project: ASLAN AI GUARDIAN PLATFORM
 * Purpose: Generate visual workflow diagrams for BOI submission
 * Slide Format: 16:9 (960x540 points)
 *
 * Instructions:
 * 1. Open Google Slides and create a new presentation
 * 2. Go to Extensions > Apps Script
 * 3. Paste this entire script
 * 4. Run the function: createBOIWorkflowPresentation()
 * 5. Grant permissions when prompted
 */

// ============================================
// CONFIGURATION - 16:9 DIMENSIONS
// ============================================
const SLIDE = {
  WIDTH: 960,   // 16:9 standard width in points
  HEIGHT: 540,  // 16:9 standard height in points
  MARGIN: 40,   // Safe margin from edges
  CONTENT_WIDTH: 880,  // 960 - 80 (40 margin each side)
  CONTENT_HEIGHT: 460  // 540 - 80 (40 margin top/bottom)
};

const COLORS = {
  PRIMARY: '1565C0',
  SECONDARY: '2E7D32',
  ACCENT: 'F57C00',
  LIGHT_BLUE: 'E3F2FD',
  LIGHT_GREEN: 'E8F5E9',
  LIGHT_ORANGE: 'FFF3E0',
  LIGHT_PURPLE: 'F3E5F5',
  LIGHT_RED: 'FFCDD2',
  WHITE: 'FFFFFF',
  DARK: '212121',
  GRAY: '757575',
  DARK_RED: 'C62828',
  DARK_PURPLE: '7B1FA2'
};

// ============================================
// MAIN FUNCTION
// ============================================
function createBOIWorkflowPresentation() {
  const presentation = SlidesApp.create('ASLAN AI GUARDIAN - Technical Workflow');

  // Set page size to 16:9
  presentation.getPageWidth();
  presentation.getPageHeight();

  const slides = presentation.getSlides();
  if (slides.length > 0) {
    slides[0].remove();
  }

  // Create all slides
  createSlide01_Title(presentation);
  createSlide02_SubmissionFlow(presentation);
  createSlide03_DocumentChecklist(presentation);
  createSlide04_Timeline(presentation);
  createSlide05_MultiAgentArchitecture(presentation);
  createSlide06_BudgetAllocation(presentation);
  createSlide07_RDActivities(presentation);
  createSlide08_DigitalTransformation(presentation);
  createSlide09_KPIFlow(presentation);
  createSlide10_PhysicalAI(presentation);
  createSlide11_ComplianceMatrix(presentation);
  createSlide12_FundDisbursement(presentation);
  createSlide13_Summary(presentation);

  Logger.log('Presentation created: ' + presentation.getUrl());
  return presentation.getUrl();
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function addTitle(slide, text) {
  const title = slide.insertTextBox(text, SLIDE.MARGIN, 20, SLIDE.CONTENT_WIDTH, 40);
  title.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor(COLORS.PRIMARY);
  return title;
}

function addBox(slide, x, y, w, h, fillColor, borderColor) {
  const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h);
  box.getFill().setSolidFill(fillColor);
  if (borderColor) {
    box.getBorder().setWeight(2).setSolidFill(borderColor);
  } else {
    box.getBorder().setTransparent();
  }
  return box;
}

function addText(slide, text, x, y, w, h, fontSize, bold, color, align) {
  const textBox = slide.insertTextBox(text, x, y, w, h);
  const style = textBox.getText().getTextStyle();
  style.setFontSize(fontSize || 12);
  if (bold) style.setBold(true);
  if (color) style.setForegroundColor(color);
  if (align) textBox.getText().getParagraphStyle().setParagraphAlignment(align);
  return textBox;
}

function addArrow(slide, x1, y1, x2, y2) {
  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x1, y1, x2, y2);
  line.getLineFill().setSolidFill(COLORS.GRAY);
  line.setWeight(2);
  line.setEndArrow(SlidesApp.ArrowStyle.FILL_ARROW);
  return line;
}

// ============================================
// SLIDE 1: TITLE
// ============================================
function createSlide01_Title(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // Full background
  const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, SLIDE.WIDTH, SLIDE.HEIGHT);
  bg.getFill().setSolidFill(COLORS.PRIMARY);
  bg.getBorder().setTransparent();

  // Main title
  addText(slide, 'ASLAN AI GUARDIAN PLATFORM', 40, 160, 880, 50, 40, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Subtitle
  addText(slide, 'Technical Workflow & Implementation Guide', 40, 220, 880, 35, 22, false, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Investment badge
  addBox(slide, 330, 290, 300, 80, COLORS.WHITE, null);
  addText(slide, 'Total Investment\n85,000,000 THB', 330, 305, 300, 60, 20, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);

  // Category badge
  addText(slide, 'BOI Category 8.2.12 - Smart Electronics | Duration: 12 Months', 40, 420, 880, 30, 14, false, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Company
  addText(slide, 'ASLAN WEALTH COMPANY LIMITED', 40, 480, 880, 25, 12, false, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 2: SUBMISSION FLOW
// ============================================
function createSlide02_SubmissionFlow(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '1. BOI Document Submission Flow');

  // Three phase boxes at top
  const phases = [
    { label: 'Phase 1\nDocument Preparation', x: 60, color: COLORS.LIGHT_BLUE, border: COLORS.PRIMARY },
    { label: 'Phase 2\nBOI Submission', x: 360, color: COLORS.LIGHT_ORANGE, border: COLORS.ACCENT },
    { label: 'Phase 3\nEvaluation & Approval', x: 660, color: COLORS.LIGHT_GREEN, border: COLORS.SECONDARY }
  ];

  phases.forEach((p, i) => {
    addBox(slide, p.x, 80, 240, 55, p.color, p.border);
    addText(slide, p.label, p.x + 10, 88, 220, 45, 12, true, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
    if (i < 2) addArrow(slide, p.x + 240, 107, p.x + 260, 107);
  });

  // Phase 1 steps
  const p1Steps = ['Company Documents', 'Project Proposal', 'Attachment A (R&D)', 'Attachment C (Digital)', 'NDA Agreement (2 copies)'];
  p1Steps.forEach((step, i) => {
    addBox(slide, 70, 150 + i * 38, 220, 32, COLORS.WHITE, COLORS.GRAY);
    addText(slide, step, 80, 157 + i * 38, 200, 22, 10, false, COLORS.DARK, SlidesApp.ParagraphAlignment.START);
  });

  // Phase 2 steps
  const p2Steps = ['Submit to BOI Online', 'Document Review', 'Receive Confirmation'];
  p2Steps.forEach((step, i) => {
    addBox(slide, 370, 150 + i * 55, 220, 40, COLORS.WHITE, COLORS.GRAY);
    addText(slide, step, 380, 162 + i * 55, 200, 25, 11, false, COLORS.DARK, SlidesApp.ParagraphAlignment.START);
  });

  // Phase 3 steps
  const p3Steps = ['NSTDA: R&D Review', 'depa: Digital Review', 'BOI Approval', 'Certificate Issued'];
  p3Steps.forEach((step, i) => {
    const isLast = i === 3;
    addBox(slide, 670, 150 + i * 45, 220, 38, isLast ? COLORS.LIGHT_GREEN : COLORS.WHITE, isLast ? COLORS.SECONDARY : COLORS.GRAY);
    addText(slide, step, 680, 160 + i * 45, 200, 25, 10, isLast, COLORS.DARK, SlidesApp.ParagraphAlignment.START);
  });

  // Final outcome
  addBox(slide, 300, 380, 360, 50, COLORS.SECONDARY, null);
  addText(slide, 'Project Kickoff → 42.5M THB Support', 300, 393, 360, 30, 14, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Timeline note
  addText(slide, 'Estimated review period: 2-3 months after complete submission', 60, 450, 840, 25, 10, false, COLORS.GRAY, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 3: DOCUMENT CHECKLIST
// ============================================
function createSlide03_DocumentChecklist(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '2. Required Documents Checklist');

  // Core Documents
  addBox(slide, 40, 75, 280, 160, COLORS.LIGHT_BLUE, COLORS.PRIMARY);
  addText(slide, 'Core Documents', 40, 80, 280, 25, 14, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '1. Project Proposal\n   (ข้อเสนอโครงการ) 1 ชุด\n\n2. Attachments\n   (เอกสารแนบ) 1 ชุด', 55, 110, 250, 115, 11, false, COLORS.DARK, null);

  // Attachments
  addBox(slide, 340, 75, 280, 160, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'Attachments', 340, 80, 280, 25, 14, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, 'A: R&D Activities\n   Budget: 59M THB\n\nC: Digital Technology\n   Budget: 26M THB\n\n+ Process Flow Diagram', 355, 110, 250, 115, 11, false, COLORS.DARK, null);

  // Legal Documents
  addBox(slide, 640, 75, 280, 160, COLORS.LIGHT_ORANGE, COLORS.ACCENT);
  addText(slide, 'Legal Documents', 640, 80, 280, 25, 14, true, COLORS.ACCENT, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '3. NDA Agreement (2 copies)\n   บันทึกข้อตกลงฯ\n\n4. Power of Attorney\n   (if applicable)\n\n5. ID Copies (shareholders)', 655, 110, 250, 115, 11, false, COLORS.DARK, null);

  // File mapping table
  const tableY = 250;
  const headers = ['Document', 'File Location'];
  const rows = [
    ['Project Proposal', '01_Submission/BOI_COMPLETE_CONSOLIDATED.docx'],
    ['Attachment A', '02_Attachments/attachment_a_rd_COMPLETE.html'],
    ['Attachment C', '02_Attachments/attachment_c_digital_COMPLETE.html'],
    ['Attachment D', '02_Attachments/attachment_d_product_details.html'],
    ['NDA', '05_Government_Documents/บันทึกข้อตกลงฯ.pdf']
  ];

  // Header row
  addBox(slide, 40, tableY, 200, 28, COLORS.PRIMARY, null);
  addBox(slide, 240, tableY, 680, 28, COLORS.PRIMARY, null);
  addText(slide, 'Document', 50, tableY + 5, 180, 20, 11, true, COLORS.WHITE, null);
  addText(slide, 'File Location', 250, tableY + 5, 660, 20, 11, true, COLORS.WHITE, null);

  // Data rows
  rows.forEach((row, i) => {
    const y = tableY + 28 + i * 28;
    const bg = i % 2 === 0 ? COLORS.LIGHT_BLUE : COLORS.WHITE;
    addBox(slide, 40, y, 200, 28, bg, COLORS.GRAY);
    addBox(slide, 240, y, 680, 28, bg, COLORS.GRAY);
    addText(slide, row[0], 50, y + 6, 180, 20, 10, false, COLORS.DARK, null);
    addText(slide, row[1], 250, y + 6, 660, 20, 9, false, COLORS.DARK, null);
  });

  // Note
  addText(slide, 'Note: All documents must be signed by authorized personnel with company seal (if applicable)', 40, 430, 880, 25, 10, false, COLORS.GRAY, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 4: TIMELINE
// ============================================
function createSlide04_Timeline(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '3. Project Timeline (12 Months)');

  // Phase bars
  addBox(slide, 40, 80, 430, 35, COLORS.PRIMARY, null);
  addText(slide, 'Phase 1: Month 1-6 | Support: 21.25M THB', 50, 87, 410, 25, 13, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  addBox(slide, 490, 80, 430, 35, COLORS.SECONDARY, null);
  addText(slide, 'Phase 2: Month 7-12 | Support: 21.25M THB', 500, 87, 410, 25, 13, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Month markers
  for (let i = 0; i <= 12; i++) {
    const x = 70 + i * 70;
    addText(slide, 'M' + i, x - 10, 120, 30, 18, 9, false, COLORS.GRAY, SlidesApp.ParagraphAlignment.CENTER);

    const marker = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x - 4, 138, 8, 8);
    marker.getFill().setSolidFill(i === 6 ? COLORS.ACCENT : COLORS.GRAY);
    marker.getBorder().setTransparent();
  }

  // Activities Gantt
  const activities = [
    { name: 'R&D #1: Physical AI Agent (12.98M)', start: 0, dur: 6, color: COLORS.LIGHT_BLUE },
    { name: 'R&D #2: XAI Framework (11.24M)', start: 0, dur: 5, color: COLORS.LIGHT_GREEN },
    { name: 'R&D #3: Thai NLP Model (15.78M)', start: 1, dur: 5, color: COLORS.LIGHT_ORANGE },
    { name: 'R&D #4: Fraud Pattern (10.5M)', start: 5, dur: 4, color: COLORS.LIGHT_PURPLE },
    { name: 'R&D #5: Multi-Agent Integration (8.5M)', start: 6, dur: 4, color: COLORS.LIGHT_BLUE },
    { name: 'Digital Infrastructure (26M)', start: 7, dur: 4, color: COLORS.LIGHT_GREEN },
    { name: 'UAT & Testing', start: 10, dur: 2, color: COLORS.LIGHT_ORANGE }
  ];

  activities.forEach((act, i) => {
    const y = 155 + i * 32;
    const barX = 70 + act.start * 70;
    const barW = act.dur * 70;

    addBox(slide, barX, y, barW, 26, act.color, COLORS.GRAY);
    addText(slide, act.name, barX + 5, y + 5, barW - 10, 18, 9, false, COLORS.DARK, null);
  });

  // Milestones
  addBox(slide, 150, 390, 280, 45, COLORS.LIGHT_ORANGE, COLORS.ACCENT);
  addText(slide, 'Milestone 1 (Month 6)\nWorking Prototype - 3 Agents', 150, 397, 280, 35, 10, true, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  addBox(slide, 530, 390, 280, 45, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'Milestone 2 (Month 12)\nProduction System - 5 Agents', 530, 397, 280, 35, 10, true, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  // Diamonds for milestones
  const diamond1 = slide.insertShape(SlidesApp.ShapeType.DIAMOND, 475, 370, 20, 20);
  diamond1.getFill().setSolidFill(COLORS.ACCENT);
  diamond1.getBorder().setTransparent();

  const diamond2 = slide.insertShape(SlidesApp.ShapeType.DIAMOND, 895, 370, 20, 20);
  diamond2.getFill().setSolidFill(COLORS.SECONDARY);
  diamond2.getBorder().setTransparent();
}

// ============================================
// SLIDE 5: MULTI-AGENT ARCHITECTURE
// ============================================
function createSlide05_MultiAgentArchitecture(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '4. Multi-Agent AI Architecture');

  // Input Layer
  addBox(slide, 40, 80, 150, 200, COLORS.LIGHT_BLUE, COLORS.PRIMARY);
  addText(slide, 'INPUT', 40, 85, 150, 22, 13, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);

  const inputs = ['Voice Input\n(Thai Speech)', 'Text Input\n(Thai Text)', 'Transaction\nData', 'Behavioral\nData'];
  inputs.forEach((inp, i) => {
    addBox(slide, 50, 115 + i * 40, 130, 35, COLORS.WHITE, COLORS.GRAY);
    addText(slide, inp, 55, 118 + i * 40, 120, 30, 9, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
  });

  // Arrow to agents
  addArrow(slide, 190, 180, 230, 180);

  // 5 AI Agents
  addBox(slide, 230, 80, 400, 320, COLORS.LIGHT_PURPLE, COLORS.DARK_PURPLE);
  addText(slide, '5 AI AGENTS LAYER', 230, 85, 400, 22, 13, true, COLORS.DARK_PURPLE, SlidesApp.ParagraphAlignment.CENTER);

  const agents = [
    { name: 'Agent 1: NLP & Voice', desc: 'Thai Language Processing', color: 'E3F2FD' },
    { name: 'Agent 2: Quantitative Risk', desc: 'Financial Risk Scoring', color: 'FFF3E0' },
    { name: 'Agent 3: Behavioral', desc: 'Pattern & Anomaly Detection', color: 'E8F5E9' },
    { name: 'Agent 4: Fraud Pattern', desc: 'Scam Detection AI', color: 'FFEBEE' },
    { name: 'Agent 5: Meta-Decision', desc: 'Consensus & Explanation', color: 'F3E5F5' }
  ];

  agents.forEach((agent, i) => {
    addBox(slide, 245, 115 + i * 55, 370, 48, agent.color, COLORS.GRAY);
    addText(slide, agent.name, 255, 120 + i * 55, 350, 20, 11, true, COLORS.DARK, null);
    addText(slide, agent.desc, 255, 140 + i * 55, 350, 18, 9, false, COLORS.GRAY, null);
  });

  // Arrow to output
  addArrow(slide, 630, 180, 670, 180);

  // Output Layer
  addBox(slide, 670, 80, 250, 200, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'OUTPUT', 670, 85, 250, 22, 13, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);

  const outputs = ['Physical AI\n(Aslan Mascot)', 'Mobile App\n(Dashboard)', 'API Services\n(Enterprise)'];
  outputs.forEach((out, i) => {
    addBox(slide, 685, 115 + i * 55, 220, 45, COLORS.WHITE, COLORS.GRAY);
    addText(slide, out, 690, 122 + i * 55, 210, 35, 10, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
  });

  // Performance metrics
  addBox(slide, 670, 295, 250, 105, COLORS.WHITE, COLORS.GRAY);
  addText(slide, 'Performance Targets', 670, 300, 250, 20, 11, true, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '• Accuracy: ≥85% F1 Score\n• Response: <200ms\n• Uptime: ≥99.5%\n• Throughput: ≥500 TPS', 680, 322, 230, 70, 10, false, COLORS.DARK, null);

  // Footer
  addText(slide, 'All 5 agents work in coordination with explainable decision outputs', 40, 420, 880, 25, 10, false, COLORS.GRAY, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 6: BUDGET ALLOCATION
// ============================================
function createSlide06_BudgetAllocation(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '5. Budget Allocation (85M THB)');

  // R&D Section
  addBox(slide, 40, 75, 420, 300, COLORS.LIGHT_BLUE, COLORS.PRIMARY);
  addText(slide, 'R&D Activities: 59M THB (69.4%)', 40, 80, 420, 28, 15, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);

  const rdItems = [
    { name: 'Personnel', amt: '36.24M', pct: 61.4, color: '1565C0' },
    { name: 'Equipment', amt: '9M', pct: 15.3, color: '1976D2' },
    { name: 'Dataset', amt: '8.5M', pct: 14.4, color: '1E88E5' },
    { name: 'Other', amt: '5.26M', pct: 8.9, color: '42A5F5' }
  ];

  let rdY = 115;
  rdItems.forEach(item => {
    const barW = item.pct * 3.5;
    addBox(slide, 55, rdY, barW, 28, item.color, null);
    addText(slide, item.name + ' (' + item.amt + ') - ' + item.pct + '%', 60, rdY + 5, 350, 22, 10, false, COLORS.WHITE, null);
    rdY += 35;
  });

  // R&D Activities breakdown
  addText(slide, 'R&D Activities Breakdown:', 55, 265, 400, 20, 11, true, COLORS.PRIMARY, null);
  addText(slide, '#1 Physical AI Agent: 12.98M\n#2 XAI Framework: 11.24M\n#3 Thai NLP Model: 15.78M\n#4 Fraud Pattern: 10.5M\n#5 Multi-Agent: 8.5M', 55, 285, 390, 80, 10, false, COLORS.DARK, null);

  // Digital Section
  addBox(slide, 500, 75, 420, 300, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'Digital Technology: 26M THB (30.6%)', 500, 80, 420, 28, 15, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);

  const digItems = [
    { name: 'Cloud Infrastructure', amt: '5M', pct: 19.2, color: '2E7D32' },
    { name: 'Hardware (GPU)', amt: '8M', pct: 30.8, color: '388E3C' },
    { name: 'Software/AI Tools', amt: '3M', pct: 11.5, color: '43A047' },
    { name: 'Security Systems', amt: '2M', pct: 7.7, color: '66BB6A' },
    { name: 'Integration/API', amt: '8M', pct: 30.8, color: '81C784' }
  ];

  let digY = 115;
  digItems.forEach(item => {
    const barW = item.pct * 3.5;
    addBox(slide, 515, digY, barW, 25, item.color, null);
    addText(slide, item.name + ' (' + item.amt + ') - ' + item.pct + '%', 520, digY + 4, 380, 20, 9, false, COLORS.WHITE, null);
    digY += 30;
  });

  // Digital tech stack
  addText(slide, 'Technology Stack:', 515, 275, 400, 20, 11, true, COLORS.SECONDARY, null);
  addText(slide, '• AWS/Azure GPU Cloud\n• NVIDIA A100 Servers\n• TensorFlow, PyTorch\n• Kubernetes Cluster\n• Security & Monitoring', 515, 295, 390, 75, 10, false, COLORS.DARK, null);

  // Total summary
  addBox(slide, 250, 390, 460, 45, COLORS.ACCENT, null);
  addText(slide, 'BOI Support: 42,500,000 THB (50% of total investment)', 250, 402, 460, 30, 16, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 7: R&D ACTIVITIES
// ============================================
function createSlide07_RDActivities(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '6. R&D Activities Detail (59M THB)');

  const activities = [
    { num: '1', name: 'Physical AI Agent Development', budget: '12.98M', output: 'AI Mascot Prototype', color: COLORS.LIGHT_BLUE },
    { num: '2', name: 'Explainable AI (XAI) Framework', budget: '11.24M', output: 'XAI Module', color: COLORS.LIGHT_GREEN },
    { num: '3', name: 'Thai NLP Model Training', budget: '15.78M', output: 'Thai NLP Engine', color: COLORS.LIGHT_ORANGE },
    { num: '4', name: 'Fraud Pattern Recognition', budget: '10.5M', output: 'Fraud Detection AI', color: COLORS.LIGHT_PURPLE },
    { num: '5', name: 'Multi-Agent Integration', budget: '8.5M', output: 'Integrated Platform', color: COLORS.LIGHT_BLUE }
  ];

  activities.forEach((act, i) => {
    const y = 80 + i * 68;

    // Number circle
    const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 50, y + 10, 40, 40);
    circle.getFill().setSolidFill(COLORS.PRIMARY);
    circle.getBorder().setTransparent();
    addText(slide, act.num, 50, y + 18, 40, 30, 18, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

    // Activity box
    addBox(slide, 105, y, 380, 58, act.color, COLORS.GRAY);
    addText(slide, act.name, 115, y + 8, 360, 22, 13, true, COLORS.DARK, null);
    addText(slide, 'Budget: ' + act.budget + ' THB', 115, y + 32, 360, 20, 10, false, COLORS.GRAY, null);

    // Arrow
    addArrow(slide, 485, y + 29, 520, y + 29);

    // Output box
    addBox(slide, 520, y + 5, 160, 48, COLORS.WHITE, COLORS.SECONDARY);
    addText(slide, act.output, 525, y + 17, 150, 30, 11, true, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
  });

  // Final arrow and output
  const finalArrow = slide.insertShape(SlidesApp.ShapeType.RIGHT_ARROW, 700, 230, 50, 35);
  finalArrow.getFill().setSolidFill(COLORS.SECONDARY);
  finalArrow.getBorder().setTransparent();

  addBox(slide, 770, 150, 150, 200, COLORS.SECONDARY, null);
  addText(slide, 'PRODUCTION\nSYSTEM\n\n5 AI Agents\n\nReady for\nDeployment', 775, 165, 140, 175, 12, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Expected outputs
  addBox(slide, 50, 430, 860, 40, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'Expected Outputs: ≥2 Patents | ≥3 Publications | ≥10,000 Thai Scam Corpus | 5 AI Models', 60, 440, 840, 25, 12, true, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 8: DIGITAL TRANSFORMATION
// ============================================
function createSlide08_DigitalTransformation(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '7. Digital Transformation (Before vs After)');

  // Before
  addBox(slide, 40, 80, 350, 190, COLORS.LIGHT_RED, COLORS.DARK_RED);
  addText(slide, 'BEFORE', 40, 85, 350, 28, 16, true, COLORS.DARK_RED, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '✗ Manual Scam Detection\n✗ Human Review Only\n✗ Delayed Response (hours/days)\n✗ No Real-time Alerts\n✗ Foreign AI Solutions Only\n✗ Limited Thai NLP Capability', 55, 120, 320, 140, 12, false, COLORS.DARK, null);

  // Transform arrow
  const arrow = slide.insertShape(SlidesApp.ShapeType.RIGHT_ARROW, 410, 160, 70, 45);
  arrow.getFill().setSolidFill(COLORS.ACCENT);
  arrow.getBorder().setTransparent();
  addText(slide, '26M THB', 410, 210, 70, 25, 10, true, COLORS.ACCENT, SlidesApp.ParagraphAlignment.CENTER);

  // After
  addBox(slide, 500, 80, 420, 190, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'AFTER', 500, 85, 420, 28, 16, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '✓ AI-Powered Detection\n✓ Automated 24/7 Analysis\n✓ Response Time <200ms\n✓ Real-time Alerts & Notifications\n✓ Local Thai AI Solution\n✓ Native Thai NLP Processing', 515, 120, 390, 140, 12, false, COLORS.DARK, null);

  // Technology stack boxes
  addBox(slide, 40, 290, 880, 130, COLORS.LIGHT_BLUE, COLORS.PRIMARY);
  addText(slide, 'Digital Technology Stack (26M THB)', 40, 295, 880, 25, 14, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);

  const techItems = [
    { name: 'Cloud\nInfrastructure', amt: '5M', x: 60 },
    { name: 'GPU\nServers', amt: '8M', x: 235 },
    { name: 'AI/ML\nSoftware', amt: '3M', x: 410 },
    { name: 'Security\nSystems', amt: '2M', x: 585 },
    { name: 'Integration\n& API', amt: '8M', x: 760 }
  ];

  techItems.forEach(tech => {
    addBox(slide, tech.x, 325, 150, 80, COLORS.WHITE, COLORS.GRAY);
    addText(slide, tech.name + '\n' + tech.amt + ' THB', tech.x + 10, 335, 130, 60, 11, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
  });

  // Impact note
  addText(slide, 'Impact: Reduce scam losses by 70% (est. 52,500M THB/year savings nationally)', 40, 435, 880, 25, 11, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 9: KPI FLOW
// ============================================
function createSlide09_KPIFlow(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '8. KPI Verification Flow');

  // Phase 1 KPIs
  addBox(slide, 40, 75, 420, 200, COLORS.LIGHT_BLUE, COLORS.PRIMARY);
  addText(slide, 'Phase 1 KPIs (Month 1-6)', 40, 80, 420, 25, 14, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '☐ Working Prototype (3 AI Agents)\n☐ Thai NLP Model F1 Score ≥80%\n☐ Multi-Agent Message Passing ≥99%\n☐ Integration Testing Complete\n☐ Technical Documentation ≥80%\n☐ Budget Utilization 50% (±15%)\n☐ Personnel Training 15 people ≥80%', 55, 110, 390, 155, 11, false, COLORS.DARK, null);

  // Phase 2 KPIs
  addBox(slide, 500, 75, 420, 200, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'Phase 2 KPIs (Month 7-12)', 500, 80, 420, 25, 14, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '☐ Production System (5 AI Agents)\n☐ Accuracy ≥85%, Response <200ms\n☐ Patents Filed ≥2\n☐ Research Publications ≥3\n☐ Thai Scam Corpus ≥10,000 samples\n☐ UAT ≥50 users, Score ≥3.5/5\n☐ depa Certification Obtained', 515, 110, 390, 155, 11, false, COLORS.DARK, null);

  // Verification bodies
  addBox(slide, 40, 290, 880, 55, COLORS.LIGHT_ORANGE, COLORS.ACCENT);
  addText(slide, 'Verification Bodies', 40, 295, 880, 22, 13, true, COLORS.ACCENT, SlidesApp.ParagraphAlignment.CENTER);

  addBox(slide, 100, 318, 280, 22, COLORS.WHITE, COLORS.GRAY);
  addText(slide, 'NSTDA: R&D Activities Verification', 105, 320, 270, 18, 10, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  addBox(slide, 580, 318, 280, 22, COLORS.WHITE, COLORS.GRAY);
  addText(slide, 'depa: Digital Technology Verification', 585, 320, 270, 18, 10, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  // Disbursement boxes
  addBox(slide, 100, 365, 280, 55, COLORS.PRIMARY, null);
  addText(slide, 'Disbursement 1\n21.25M THB', 100, 375, 280, 40, 13, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  addBox(slide, 580, 365, 280, 55, COLORS.SECONDARY, null);
  addText(slide, 'Disbursement 2\n21.25M THB', 580, 375, 280, 40, 13, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Arrows
  addArrow(slide, 240, 275, 240, 290);
  addArrow(slide, 720, 275, 720, 290);
  addArrow(slide, 240, 345, 240, 365);
  addArrow(slide, 720, 345, 720, 365);

  // Total
  addText(slide, 'Total BOI Support: 42,500,000 THB (50% of 85M investment)', 40, 435, 880, 25, 12, true, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 10: PHYSICAL AI
// ============================================
function createSlide10_PhysicalAI(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '9. Physical AI Agent Architecture');

  // Hardware Layer
  addBox(slide, 40, 80, 250, 180, COLORS.LIGHT_ORANGE, COLORS.ACCENT);
  addText(slide, 'Hardware Layer', 40, 85, 250, 22, 13, true, COLORS.ACCENT, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '• Microphone Array\n  (Thai Voice Input)\n• Speaker System\n  (Thai Voice Output)\n• LED Display\n  (Visual Feedback)\n• IoT Sensors\n• Embedded AI Chip', 55, 112, 220, 140, 10, false, COLORS.DARK, null);

  // Connectivity Layer
  addBox(slide, 330, 80, 250, 180, COLORS.LIGHT_BLUE, COLORS.PRIMARY);
  addText(slide, 'Connectivity', 330, 85, 250, 22, 13, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '• WiFi Module\n  (Cloud Connection)\n• Bluetooth\n  (Mobile App Sync)\n• Embedded OS\n  (Real-time Processing)', 345, 112, 220, 140, 10, false, COLORS.DARK, null);

  // Cloud Layer
  addBox(slide, 620, 80, 300, 180, COLORS.LIGHT_PURPLE, COLORS.DARK_PURPLE);
  addText(slide, 'Cloud AI Platform', 620, 85, 300, 22, 13, true, COLORS.DARK_PURPLE, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '• Multi-Agent Platform\n  (5 AI Agents)\n• Thai NLP Engine\n  (Language Processing)\n• Fraud Detection\n  (Scam Protection)', 635, 112, 270, 140, 10, false, COLORS.DARK, null);

  // Arrows
  addArrow(slide, 290, 170, 330, 170);
  addArrow(slide, 580, 170, 620, 170);

  // BOI Compliance
  addBox(slide, 40, 280, 880, 140, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'BOI Category 8.2.12 Compliance (Smart Electronics)', 40, 285, 880, 25, 14, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);

  const criteria = [
    { num: '1', desc: 'Electronic sensors for data detection', feature: 'Microphone, Speaker, LED, IoT Sensors' },
    { num: '2', desc: 'Wireless connectivity', feature: 'WiFi Module, Bluetooth' },
    { num: '3', desc: 'Embedded processing system', feature: 'AI Chip, Real-time OS, Edge Processing' }
  ];

  criteria.forEach((c, i) => {
    const y = 315 + i * 32;

    const numCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 60, y, 25, 25);
    numCircle.getFill().setSolidFill(COLORS.PRIMARY);
    numCircle.getBorder().setTransparent();
    addText(slide, c.num, 60, y + 4, 25, 20, 12, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

    addText(slide, c.desc + ': ' + c.feature, 95, y + 4, 680, 22, 10, false, COLORS.DARK, null);
    addText(slide, '✓ PASS', 820, y + 4, 80, 22, 11, true, COLORS.SECONDARY, null);
  });

  // Result
  addText(slide, 'Result: 100% Compliant with BOI Category 8.2.12 Smart Electronics', 40, 435, 880, 25, 12, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 11: COMPLIANCE MATRIX
// ============================================
function createSlide11_ComplianceMatrix(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '10. BOI Compliance Matrix');

  const headers = ['Requirement', 'Threshold', 'This Project', 'Status'];
  const colWidths = [220, 180, 200, 100];
  const data = [
    ['Industry Category', '8.2.12 Smart Electronics', '3/3 criteria met', '✓ PASS'],
    ['Minimum Investment', '≥50M THB', '85M THB (170%)', '✓ PASS'],
    ['Thai Shareholding', '≥51%', '100%', '✓ PASS'],
    ['Project Duration', '≤12 months', '12 months', '✓ PASS'],
    ['Eligible Expenses', 'Per AOR3/2568', '100% compliant', '✓ PASS'],
    ['R&D Activities', 'Per Attachment A', '59M THB (5 activities)', '✓ PASS'],
    ['Digital Technology', 'Per Attachment C', '26M THB (5 categories)', '✓ PASS']
  ];

  // Headers
  let xPos = 80;
  headers.forEach((h, i) => {
    addBox(slide, xPos, 80, colWidths[i], 35, COLORS.PRIMARY, null);
    addText(slide, h, xPos + 10, 88, colWidths[i] - 20, 25, 12, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);
    xPos += colWidths[i];
  });

  // Data rows
  data.forEach((row, rowIdx) => {
    const y = 115 + rowIdx * 35;
    const bgColor = rowIdx % 2 === 0 ? COLORS.LIGHT_BLUE : COLORS.WHITE;

    xPos = 80;
    row.forEach((cell, colIdx) => {
      const cellBg = colIdx === 3 ? COLORS.LIGHT_GREEN : bgColor;
      addBox(slide, xPos, y, colWidths[colIdx], 35, cellBg, COLORS.GRAY);

      const textColor = colIdx === 3 ? COLORS.SECONDARY : COLORS.DARK;
      const isBold = colIdx === 3;
      addText(slide, cell, xPos + 5, y + 8, colWidths[colIdx] - 10, 25, 10, isBold, textColor, SlidesApp.ParagraphAlignment.CENTER);
      xPos += colWidths[colIdx];
    });
  });

  // Summary box
  addBox(slide, 80, 380, 700, 60, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'Compliance Summary', 80, 385, 700, 22, 13, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, 'All 7 requirements PASSED | Investment 170% of minimum | Approval Probability: 90-95%', 90, 410, 680, 25, 11, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 12: FUND DISBURSEMENT
// ============================================
function createSlide12_FundDisbursement(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitle(slide, '11. Fund Disbursement Flow');

  // Total badge
  addBox(slide, 320, 65, 320, 45, COLORS.PRIMARY, null);
  addText(slide, 'Total BOI Support: 42.5M THB', 320, 77, 320, 30, 16, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Phase 1
  addBox(slide, 40, 125, 420, 160, COLORS.LIGHT_BLUE, COLORS.PRIMARY);
  addText(slide, 'Phase 1 (Month 1-6)', 40, 130, 420, 25, 14, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '1. Submit R&D Progress Report → NSTDA\n2. Submit Digital Progress → depa\n3. Verification (50% completion)\n4. BOI Approval', 55, 160, 390, 90, 11, false, COLORS.DARK, null);

  addBox(slide, 120, 255, 260, 30, COLORS.PRIMARY, null);
  addText(slide, 'Disbursement: 21.25M THB', 120, 260, 260, 22, 12, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Arrow between phases
  addArrow(slide, 460, 200, 500, 200);

  // Phase 2
  addBox(slide, 500, 125, 420, 160, COLORS.LIGHT_GREEN, COLORS.SECONDARY);
  addText(slide, 'Phase 2 (Month 7-12)', 500, 130, 420, 25, 14, true, COLORS.SECONDARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, '1. Submit Final R&D Report → NSTDA\n2. Submit Final Digital Report → depa\n3. Verification (100% completion)\n4. Final BOI Approval', 515, 160, 390, 90, 11, false, COLORS.DARK, null);

  addBox(slide, 580, 255, 260, 30, COLORS.SECONDARY, null);
  addText(slide, 'Disbursement: 21.25M THB', 580, 260, 260, 22, 12, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Verification bodies
  addBox(slide, 40, 305, 880, 55, COLORS.LIGHT_ORANGE, COLORS.ACCENT);
  addText(slide, 'Verification Bodies', 40, 310, 880, 22, 12, true, COLORS.ACCENT, SlidesApp.ParagraphAlignment.CENTER);

  addBox(slide, 120, 333, 280, 22, COLORS.WHITE, COLORS.GRAY);
  addText(slide, 'NSTDA: R&D Verification', 125, 335, 270, 18, 10, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  addBox(slide, 560, 333, 280, 22, COLORS.WHITE, COLORS.GRAY);
  addText(slide, 'depa: Digital Verification', 565, 335, 270, 18, 10, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  // Bank info
  addBox(slide, 40, 380, 880, 50, COLORS.WHITE, COLORS.GRAY);
  addText(slide, 'Bank Account: 198-1-59068-2 | Kasikorn Bank, Sam Sen Nai 56 | ASLAN WEALTH COMPANY LIMITED', 50, 395, 860, 25, 11, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  // Timeline
  addText(slide, 'Timeline: Phase 1 disbursement at Month 6 | Phase 2 disbursement at Month 12', 40, 445, 880, 22, 10, false, COLORS.GRAY, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// SLIDE 13: SUMMARY
// ============================================
function createSlide13_Summary(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // Full background
  const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, SLIDE.WIDTH, SLIDE.HEIGHT);
  bg.getFill().setSolidFill(COLORS.PRIMARY);
  bg.getBorder().setTransparent();

  // Title
  addText(slide, 'Project Summary', 40, 30, 880, 40, 28, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Summary cards
  const cards = [
    { label: 'Total Investment', value: '85M THB' },
    { label: 'BOI Support', value: '42.5M THB' },
    { label: 'R&D Budget', value: '59M THB' },
    { label: 'Digital Budget', value: '26M THB' }
  ];

  cards.forEach((card, i) => {
    const x = 60 + i * 220;
    addBox(slide, x, 90, 200, 80, COLORS.WHITE, null);
    addText(slide, card.label, x, 100, 200, 22, 12, false, COLORS.GRAY, SlidesApp.ParagraphAlignment.CENTER);
    addText(slide, card.value, x, 125, 200, 40, 22, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);
  });

  // Key highlights box
  addBox(slide, 40, 190, 540, 170, '1565C0', COLORS.WHITE);
  addText(slide, 'Key Highlights', 50, 195, 520, 25, 14, true, COLORS.WHITE, null);
  addText(slide, '✓ BOI Category 8.2.12 - Smart Electronics (100% Compliant)\n✓ 5 AI Agents: NLP, Risk, Behavioral, Fraud, Meta-Decision\n✓ Physical AI Agent (Aslan Mascot) - Thai Financial Guardian\n✓ 12-Month Timeline with 2 Milestone Disbursements\n✓ Verified by NSTDA (R&D) and depa (Digital)\n✓ Approval Probability: 90-95%', 50, 222, 520, 130, 12, false, COLORS.WHITE, null);

  // Contact box
  addBox(slide, 600, 190, 320, 170, COLORS.WHITE, null);
  addText(slide, 'Contact Information', 610, 195, 300, 25, 14, true, COLORS.PRIMARY, SlidesApp.ParagraphAlignment.CENTER);
  addText(slide, 'ASLAN WEALTH\nCOMPANY LIMITED\n\nCEO: Surakiat Yawanopas\nTel: 082-619-5459\nEmail: admin@aslanwealth.com\n\n88/260 Sai Mai Road\nBangkok 10220', 610, 225, 300, 130, 10, false, COLORS.DARK, SlidesApp.ParagraphAlignment.CENTER);

  // Products
  addBox(slide, 40, 380, 880, 55, '1565C0', COLORS.WHITE);
  addText(slide, '5 Products: Physical AI Agent | Thai NLP Engine | Multi-Agent Platform | XAI Module | Mobile App', 50, 395, 860, 30, 13, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);

  // Footer
  addText(slide, 'Ready for BOI Submission', 40, 455, 880, 35, 18, true, COLORS.WHITE, SlidesApp.ParagraphAlignment.CENTER);
}

// ============================================
// MENU SETUP
// ============================================
function onOpen() {
  try {
    const ui = SlidesApp.getUi();
    ui.createMenu('BOI Workflow')
      .addItem('Generate Workflow Slides', 'createBOIWorkflowPresentation')
      .addToUi();
  } catch (e) {
    Logger.log('Menu not available: ' + e.message);
  }
}
