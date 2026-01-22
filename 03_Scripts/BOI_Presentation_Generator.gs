/**
 * BOI Presentation Generator - Google Apps Script
 * สร้าง Google Slides สรุปโครงการ BOI - ASLAN AI GUARDIAN PLATFORM
 *
 * 10 สไลด์ มืออาชีพ กระชับ เนื้อหาสำคัญครบถ้วน
 *
 * ⚠️ สำคัญ: หลังรันสคริปต์แล้ว ให้เปลี่ยนเป็น 16:9 ด้วยตัวเอง:
 *    1. เปิด Google Slides ที่สร้างใหม่
 *    2. ไปที่ File > Page setup
 *    3. เลือก "Widescreen 16:9"
 *    4. คลิก Apply
 *
 * วิธีใช้งาน:
 * 1. เปิด Google Apps Script (script.google.com)
 * 2. วางโค้ดนี้ทั้งหมด
 * 3. ใส่ URL รูป Aslan Lion ใน LOGO_URL (ถ้ามี)
 * 4. รัน function createBOIPresentation()
 * 5. อนุญาตสิทธิ์การเข้าถึง
 * 6. เปลี่ยนเป็น 16:9 ตามขั้นตอนด้านบน
 *
 * @author Claude AI Assistant
 * @version 2.1 (Professional Design)
 * @date 21 มกราคม 2569
 */

// ====================================
// CONFIGURATION
// ====================================

// ⭐ ใส่ URL รูป Aslan Lion ที่นี่ (ต้องเป็น URL สาธารณะ)
const LOGO_URL = '';  // ใส่ URL รูปที่นี่

// Slide dimensions - Google Slides default is 720 x 540 (4:3)
const SLIDE_WIDTH = 720;
const SLIDE_HEIGHT = 540;

// ====================================
// COLOR PALETTE - Executive Design
// ====================================
const COLORS = {
  primary: '#0D1B2A',      // Deep Navy
  secondary: '#1B3A5F',    // Navy Blue
  accent: '#C9A227',       // Gold
  accentLight: '#E8D591',  // Light Gold
  success: '#1A5F4A',      // Deep Green
  info: '#1565C0',         // Blue
  warning: '#E65100',      // Orange
  white: '#FFFFFF',
  light: '#F8F9FA',
  lightGray: '#E9ECEF',
  mediumGray: '#6C757D',
  dark: '#212529',
  rdColor: '#1A5F4A',
  rdLight: '#E8F5E9',
  digitalColor: '#1565C0',
  digitalLight: '#E3F2FD',
};

// ====================================
// PROJECT DATA
// ====================================
const PROJECT = {
  name: 'ASLAN AI GUARDIAN PLATFORM',
  nameTH: 'แพลตฟอร์มปัญญาประดิษฐ์เพื่อการปกป้องทางการเงิน',
  company: 'บริษัท อัสลาน เวลธ์ จำกัด',
  companyEN: 'ASLAN WEALTH COMPANY LIMITED',
  regNo: '0105567220030',
  industry: 'อุตสาหกรรมอิเล็กทรอนิกส์อัจฉริยะ (8.2.12)',
  totalInvestment: 85,
  rdInvestment: 59,
  digitalInvestment: 26,
  totalSupport: 42.5,
  rdSupport: 29.5,
  digitalSupport: 13,
  supportRate: 50,
  duration: '12 เดือน',
  startDate: '1 มิถุนายน 2569',
  endDate: '31 พฤษภาคม 2570',
  rdAgency: 'สวทช.',
  digitalAgency: 'depa'
};

// ====================================
// MAIN FUNCTION
// ====================================
function createBOIPresentation() {
  const presentation = SlidesApp.create('BOI_ASLAN_PRESENTATION_' + new Date().toISOString().split('T')[0]);

  const slides = presentation.getSlides();
  if (slides.length > 0) {
    slides[0].remove();
  }

  createSlide1_Cover(presentation);
  createSlide2_Overview(presentation);
  createSlide3_Problem(presentation);
  createSlide4_Solution(presentation);
  createSlide5_Technology(presentation);
  createSlide6_Budget(presentation);
  createSlide7_Timeline(presentation);
  createSlide8_KPI(presentation);
  createSlide9_Impact(presentation);
  createSlide10_Summary(presentation);

  Logger.log('✅ สร้าง Presentation สำเร็จ!');
  Logger.log('⚠️ สำคัญ: ไปที่ File > Page setup > เลือก Widescreen 16:9');
  Logger.log('🔗 URL: ' + presentation.getUrl());

  return presentation.getUrl();
}

// ====================================
// SLIDE 1: COVER PAGE
// ====================================
function createSlide1_Cover(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.primary);

  // Top accent line
  const topLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, SLIDE_WIDTH, 6);
  topLine.getFill().setSolidFill(COLORS.accent);
  topLine.getBorder().setTransparent();

  // Logo circle
  const logoSize = 90;
  const logoX = (SLIDE_WIDTH - logoSize) / 2;
  const logoY = 60;

  const logoCircle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, logoX, logoY, logoSize, logoSize);
  logoCircle.getFill().setSolidFill(COLORS.accent);
  logoCircle.getBorder().getLineFill().setSolidFill(COLORS.white);
  logoCircle.getBorder().setWeight(3);

  // Logo text or image
  if (LOGO_URL && LOGO_URL.trim() !== '') {
    try {
      slide.insertImage(LOGO_URL, logoX + 10, logoY + 10, logoSize - 20, logoSize - 20);
    } catch (e) {
      insertLogoText(slide, logoX, logoY, logoSize);
    }
  } else {
    insertLogoText(slide, logoX, logoY, logoSize);
  }

  // Main Title
  const title = slide.insertTextBox(PROJECT.name, 40, 170, SLIDE_WIDTH - 80, 40);
  title.getText().getTextStyle()
    .setFontSize(28)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');
  title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Subtitle
  const subtitle = slide.insertTextBox(PROJECT.nameTH, 40, 215, SLIDE_WIDTH - 80, 30);
  subtitle.getText().getTextStyle()
    .setFontSize(16)
    .setForegroundColor(COLORS.accent)
    .setFontFamily('Sarabun');
  subtitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Divider
  const divider = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 280, 260, 160, 2);
  divider.getFill().setSolidFill(COLORS.accent);
  divider.getBorder().setTransparent();

  // Company name
  const company = slide.insertTextBox(PROJECT.company, 40, 280, SLIDE_WIDTH - 80, 25);
  company.getText().getTextStyle()
    .setFontSize(14)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');
  company.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // BOI Badge
  const badgeWidth = 300;
  const badgeX = (SLIDE_WIDTH - badgeWidth) / 2;
  const badge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, badgeX, 330, badgeWidth, 45);
  badge.getFill().setSolidFill(COLORS.success);
  badge.getBorder().setTransparent();

  const badgeText = slide.insertTextBox('ขอรับการส่งเสริมจาก BOI', badgeX, 342, badgeWidth, 25);
  badgeText.getText().getTextStyle()
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');
  badgeText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Footer
  const footer = slide.insertTextBox('มาตรการสนับสนุนผู้ประกอบการไทยเพื่อเพิ่มขีดความสามารถในการแข่งขัน', 40, 410, SLIDE_WIDTH - 80, 20);
  footer.getText().getTextStyle()
    .setFontSize(10)
    .setForegroundColor(COLORS.mediumGray)
    .setFontFamily('Sarabun');
  footer.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Date
  const date = slide.insertTextBox('มกราคม 2569', 40, 430, SLIDE_WIDTH - 80, 18);
  date.getText().getTextStyle()
    .setFontSize(11)
    .setForegroundColor(COLORS.accentLight)
    .setFontFamily('Sarabun');
  date.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Bottom accent
  const bottomLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, SLIDE_HEIGHT - 6, SLIDE_WIDTH, 6);
  bottomLine.getFill().setSolidFill(COLORS.accent);
  bottomLine.getBorder().setTransparent();
}

function insertLogoText(slide, logoX, logoY, logoSize) {
  const logoText = slide.insertTextBox('ASLAN', logoX, logoY + 28, logoSize, 35);
  logoText.getText().getTextStyle()
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(COLORS.primary)
    .setFontFamily('Sarabun');
  logoText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ====================================
// SLIDE 2: PROJECT OVERVIEW
// ====================================
function createSlide2_Overview(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'ภาพรวมโครงการ', '2/10');

  // Company info
  const infoData = [
    { label: 'บริษัท', value: PROJECT.company },
    { label: 'ทะเบียน', value: PROJECT.regNo },
    { label: 'หมวด', value: PROJECT.industry },
    { label: 'ระยะเวลา', value: PROJECT.duration }
  ];

  let yPos = 80;
  infoData.forEach(function(info) {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, yPos, 320, 32);
    box.getFill().setSolidFill(COLORS.lightGray);
    box.getBorder().setTransparent();

    const text = slide.insertTextBox(info.label + ': ' + info.value, 40, yPos + 7, 300, 20);
    text.getText().getTextStyle()
      .setFontSize(11)
      .setForegroundColor(COLORS.dark)
      .setFontFamily('Sarabun');
    yPos += 38;
  });

  // Budget box
  const budgetBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 370, 80, 320, 150);
  budgetBox.getFill().setSolidFill(COLORS.primary);
  budgetBox.getBorder().setTransparent();

  const budgetTitle = slide.insertTextBox('งบประมาณและเงินสนับสนุน', 385, 92, 290, 22);
  budgetTitle.getText().getTextStyle()
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(COLORS.accent)
    .setFontFamily('Sarabun');

  const investText = slide.insertTextBox('งบลงทุน: 85 ล้านบาท', 385, 118, 150, 20);
  investText.getText().getTextStyle()
    .setFontSize(11)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  const supportText = slide.insertTextBox('สนับสนุน: 42.5 ล้านบาท', 385, 140, 180, 22);
  supportText.getText().getTextStyle()
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.accent)
    .setFontFamily('Sarabun');

  // Rate badge
  const rateBadge = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 620, 165, 50, 50);
  rateBadge.getFill().setSolidFill(COLORS.accent);
  rateBadge.getBorder().setTransparent();

  const rateText = slide.insertTextBox('50%', 620, 180, 50, 22);
  rateText.getText().getTextStyle()
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.primary)
    .setFontFamily('Sarabun');
  rateText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // R&D Box
  const rdBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 250, 330, 95);
  rdBox.getFill().setSolidFill(COLORS.rdLight);
  rdBox.getBorder().getLineFill().setSolidFill(COLORS.rdColor);
  rdBox.getBorder().setWeight(2);

  const rdTitle = slide.insertTextBox('R&D (สวทช.)', 45, 260, 300, 20);
  rdTitle.getText().getTextStyle()
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(COLORS.rdColor)
    .setFontFamily('Sarabun');

  const rdDetails = slide.insertTextBox('ลงทุน: 59M | สนับสนุน: 29.5M\n69% ของงบประมาณรวม', 45, 285, 300, 45);
  rdDetails.getText().getTextStyle()
    .setFontSize(11)
    .setForegroundColor(COLORS.dark)
    .setFontFamily('Sarabun');

  // Digital Box
  const digiBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 370, 250, 320, 95);
  digiBox.getFill().setSolidFill(COLORS.digitalLight);
  digiBox.getBorder().getLineFill().setSolidFill(COLORS.digitalColor);
  digiBox.getBorder().setWeight(2);

  const digiTitle = slide.insertTextBox('Digital (depa)', 385, 260, 290, 20);
  digiTitle.getText().getTextStyle()
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(COLORS.digitalColor)
    .setFontFamily('Sarabun');

  const digiDetails = slide.insertTextBox('ลงทุน: 26M | สนับสนุน: 13M\n31% ของงบประมาณรวม', 385, 285, 290, 45);
  digiDetails.getText().getTextStyle()
    .setFontSize(11)
    .setForegroundColor(COLORS.dark)
    .setFontFamily('Sarabun');

  // Timeline
  const timeBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 360, 660, 45);
  timeBox.getFill().setSolidFill(COLORS.lightGray);
  timeBox.getBorder().setTransparent();

  const timeText = slide.insertTextBox('ระยะเวลา: ' + PROJECT.startDate + ' - ' + PROJECT.endDate + '   (แบ่งเป็น 2 งวด งวดละ 6 เดือน)', 45, 372, 630, 22);
  timeText.getText().getTextStyle()
    .setFontSize(12)
    .setForegroundColor(COLORS.dark)
    .setFontFamily('Sarabun');
}

// ====================================
// SLIDE 3: PROBLEM
// ====================================
function createSlide3_Problem(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'ปัญหาที่ต้องการแก้ไข', '3/10');

  // Main stat
  const statBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 75, 660, 65);
  statBox.getFill().setSolidFill('#FFF3E0');
  statBox.getBorder().getLineFill().setSolidFill(COLORS.warning);
  statBox.getBorder().setWeight(2);

  const statText = slide.insertTextBox('ความเสียหายจาก Financial Scam ในไทย: มากกว่า 50,000 ล้านบาท/ปี', 50, 95, 620, 28);
  statText.getText().getTextStyle()
    .setFontSize(16)
    .setBold(true)
    .setForegroundColor(COLORS.warning)
    .setFontFamily('Sarabun');
  statText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Problem cards
  const problems = [
    { num: '1', title: 'ผู้สูงอายุ', desc: 'ตกเป็นเหยื่อมากที่สุด' },
    { num: '2', title: 'เยาวชน', desc: 'ขาด Financial Literacy' },
    { num: '3', title: 'กลโกง', desc: 'ใช้ AI หลอกลวง' },
    { num: '4', title: 'ตรวจจับ', desc: 'ไม่มี AI ภาษาไทย' }
  ];

  let xPos = 30;
  problems.forEach(function(p) {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, xPos, 155, 155, 100);
    card.getFill().setSolidFill(COLORS.lightGray);
    card.getBorder().setTransparent();

    const numBadge = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, xPos + 55, 165, 40, 40);
    numBadge.getFill().setSolidFill(COLORS.primary);
    numBadge.getBorder().setTransparent();

    const numText = slide.insertTextBox(p.num, xPos + 55, 175, 40, 22);
    numText.getText().getTextStyle()
      .setFontSize(14)
      .setBold(true)
      .setForegroundColor(COLORS.white)
      .setFontFamily('Sarabun');
    numText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const title = slide.insertTextBox(p.title, xPos + 5, 210, 145, 18);
    title.getText().getTextStyle()
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(COLORS.dark)
      .setFontFamily('Sarabun');
    title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const desc = slide.insertTextBox(p.desc, xPos + 5, 228, 145, 20);
    desc.getText().getTextStyle()
      .setFontSize(10)
      .setForegroundColor(COLORS.mediumGray)
      .setFontFamily('Sarabun');
    desc.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    xPos += 170;
  });

  // Solution box
  const solutionBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 270, 660, 75);
  solutionBox.getFill().setSolidFill(COLORS.primary);
  solutionBox.getBorder().setTransparent();

  const solutionTitle = slide.insertTextBox('ทางออก: ASLAN AI GUARDIAN PLATFORM', 50, 285, 500, 25);
  solutionTitle.getText().getTextStyle()
    .setFontSize(16)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  const solutionDesc = slide.insertTextBox('ระบบ AI หลายตัวแทนที่ออกแบบมาเพื่อปกป้องคนไทยจากภัยทางการเงิน', 50, 310, 500, 22);
  solutionDesc.getText().getTextStyle()
    .setFontSize(11)
    .setForegroundColor(COLORS.lightGray)
    .setFontFamily('Sarabun');

  // Target badge
  const targetBadge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 560, 280, 120, 55);
  targetBadge.getFill().setSolidFill(COLORS.accent);
  targetBadge.getBorder().setTransparent();

  const targetText = slide.insertTextBox('เป้าหมาย\n55,000 ผู้ใช้', 560, 290, 120, 40);
  targetText.getText().getTextStyle()
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(COLORS.primary)
    .setFontFamily('Sarabun');
  targetText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Target groups
  const groupText = slide.insertTextBox('กลุ่มเป้าหมาย: ครอบครัว 20,000 | ผู้สูงอายุ 30,000 | มืออาชีพ 5,000', 30, 360, 660, 22);
  groupText.getText().getTextStyle()
    .setFontSize(12)
    .setForegroundColor(COLORS.dark)
    .setFontFamily('Sarabun');
  groupText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ====================================
// SLIDE 4: SOLUTION / PRODUCTS
// ====================================
function createSlide4_Solution(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'ผลิตภัณฑ์และบริการ (5 Products)', '4/10');

  const products = [
    { id: '01', name: 'Physical AI Agent', desc: 'หุ่นยนต์ AI "Aslan"', metric: '10,000 หน่วย/ปี', color: '#E8F5E9' },
    { id: '02', name: 'Thai NLP Engine', desc: 'ประมวลผลภาษาไทย', metric: '100,000 คำขอ/วัน', color: '#E3F2FD' },
    { id: '03', name: 'Multi-Agent AI', desc: '5 AI Agents', metric: '50,000 Users', color: '#FFF3E0' },
    { id: '04', name: 'Explainable AI', desc: 'อธิบายการตัดสินใจ AI', metric: '100,000 วิเคราะห์/วัน', color: '#F3E5F5' },
    { id: '05', name: 'Mobile App', desc: 'iOS/Android', metric: '100,000 Active Users', color: '#FCE4EC' }
  ];

  // Row 1 - 3 products
  const cardW = 220;
  const cardH = 115;
  const gap = 10;

  for (let i = 0; i < 3; i++) {
    const p = products[i];
    const x = 30 + (i * (cardW + gap));

    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 75, cardW, cardH);
    card.getFill().setSolidFill(p.color);
    card.getBorder().setTransparent();

    const idBadge = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x + 10, 85, 35, 35);
    idBadge.getFill().setSolidFill(COLORS.primary);
    idBadge.getBorder().setTransparent();

    const idText = slide.insertTextBox(p.id, x + 10, 93, 35, 20);
    idText.getText().getTextStyle()
      .setFontSize(11)
      .setBold(true)
      .setForegroundColor(COLORS.white)
      .setFontFamily('Sarabun');
    idText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const name = slide.insertTextBox(p.name, x + 50, 87, 165, 18);
    name.getText().getTextStyle()
      .setFontSize(11)
      .setBold(true)
      .setForegroundColor(COLORS.dark)
      .setFontFamily('Sarabun');

    const desc = slide.insertTextBox(p.desc, x + 50, 105, 165, 16);
    desc.getText().getTextStyle()
      .setFontSize(9)
      .setForegroundColor(COLORS.mediumGray)
      .setFontFamily('Sarabun');

    const metricBg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 10, 130, cardW - 20, 28);
    metricBg.getFill().setSolidFill(COLORS.white);
    metricBg.getBorder().setTransparent();

    const metric = slide.insertTextBox(p.metric, x + 10, 135, cardW - 20, 20);
    metric.getText().getTextStyle()
      .setFontSize(10)
      .setBold(true)
      .setForegroundColor(COLORS.primary)
      .setFontFamily('Sarabun');
    metric.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  // Row 2 - 2 products
  const row2Y = 200;
  const row2StartX = 30 + ((cardW + gap) / 2);

  for (let i = 3; i < 5; i++) {
    const p = products[i];
    const x = row2StartX + ((i - 3) * (cardW + gap));

    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, row2Y, cardW, cardH);
    card.getFill().setSolidFill(p.color);
    card.getBorder().setTransparent();

    const idBadge = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x + 10, row2Y + 10, 35, 35);
    idBadge.getFill().setSolidFill(COLORS.primary);
    idBadge.getBorder().setTransparent();

    const idText = slide.insertTextBox(p.id, x + 10, row2Y + 18, 35, 20);
    idText.getText().getTextStyle()
      .setFontSize(11)
      .setBold(true)
      .setForegroundColor(COLORS.white)
      .setFontFamily('Sarabun');
    idText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const name = slide.insertTextBox(p.name, x + 50, row2Y + 12, 165, 18);
    name.getText().getTextStyle()
      .setFontSize(11)
      .setBold(true)
      .setForegroundColor(COLORS.dark)
      .setFontFamily('Sarabun');

    const desc = slide.insertTextBox(p.desc, x + 50, row2Y + 30, 165, 16);
    desc.getText().getTextStyle()
      .setFontSize(9)
      .setForegroundColor(COLORS.mediumGray)
      .setFontFamily('Sarabun');

    const metricBg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 10, row2Y + 55, cardW - 20, 28);
    metricBg.getFill().setSolidFill(COLORS.white);
    metricBg.getBorder().setTransparent();

    const metric = slide.insertTextBox(p.metric, x + 10, row2Y + 60, cardW - 20, 20);
    metric.getText().getTextStyle()
      .setFontSize(10)
      .setBold(true)
      .setForegroundColor(COLORS.primary)
      .setFontFamily('Sarabun');
    metric.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  // Innovation bar
  const highlightBar = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 330, 660, 45);
  highlightBar.getFill().setSolidFill(COLORS.primary);
  highlightBar.getBorder().setTransparent();

  const highlightText = slide.insertTextBox('นวัตกรรมครั้งแรกในประเทศไทยและอาเซียน', 50, 342, 400, 22);
  highlightText.getText().getTextStyle()
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  const highlightSub = slide.insertTextBox('Physical AI + Multi-Agent + Thai NLP', 50, 360, 400, 16);
  highlightSub.getText().getTextStyle()
    .setFontSize(10)
    .setForegroundColor(COLORS.accentLight)
    .setFontFamily('Sarabun');
}

// ====================================
// SLIDE 5: TECHNOLOGY
// ====================================
function createSlide5_Technology(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'เทคโนโลยีและกิจกรรม', '5/10');

  // R&D Section
  const rdHeader = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 75, 320, 28);
  rdHeader.getFill().setSolidFill(COLORS.rdColor);
  rdHeader.getBorder().setTransparent();

  const rdTitle = slide.insertTextBox('R&D Activities (59M)', 40, 80, 300, 20);
  rdTitle.getText().getTextStyle()
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  const rdActivities = [
    'Physical AI Agent - 13.0M (22%)',
    'Explainable AI - 11.2M (19%)',
    'Thai NLP - 15.8M (27%)',
    'Fraud Recognition - 10.5M (18%)',
    'Multi-Agent - 8.5M (14%)'
  ];

  let rdY = 108;
  rdActivities.forEach(function(act) {
    const actBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, rdY, 320, 24);
    actBox.getFill().setSolidFill(COLORS.rdLight);
    actBox.getBorder().setTransparent();

    const actText = slide.insertTextBox(act, 40, rdY + 4, 300, 18);
    actText.getText().getTextStyle()
      .setFontSize(10)
      .setForegroundColor(COLORS.dark)
      .setFontFamily('Sarabun');
    rdY += 28;
  });

  // Digital Section
  const digiHeader = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 370, 75, 320, 28);
  digiHeader.getFill().setSolidFill(COLORS.digitalColor);
  digiHeader.getBorder().setTransparent();

  const digiTitle = slide.insertTextBox('Digital Technology (26M)', 380, 80, 300, 20);
  digiTitle.getText().getTextStyle()
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  const digiCategories = [
    'Cloud Services - 5.0M (19%)',
    'On-premise Hardware - 8.0M (31%)',
    'Software Licenses - 3.0M (12%)',
    'Security Systems - 2.0M (8%)',
    'System Integration - 8.0M (31%)'
  ];

  let digiY = 108;
  digiCategories.forEach(function(cat) {
    const catBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 370, digiY, 320, 24);
    catBox.getFill().setSolidFill(COLORS.digitalLight);
    catBox.getBorder().setTransparent();

    const catText = slide.insertTextBox(cat, 380, digiY + 4, 300, 18);
    catText.getText().getTextStyle()
      .setFontSize(10)
      .setForegroundColor(COLORS.dark)
      .setFontFamily('Sarabun');
    digiY += 28;
  });

  // Performance metrics
  const perfBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 260, 660, 75);
  perfBox.getFill().setSolidFill(COLORS.primary);
  perfBox.getBorder().setTransparent();

  const perfTitle = slide.insertTextBox('Performance Targets', 45, 268, 200, 20);
  perfTitle.getText().getTextStyle()
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(COLORS.accent)
    .setFontFamily('Sarabun');

  const metrics = [
    { label: 'Uptime', value: '≥99.5%' },
    { label: 'Response', value: '<200ms' },
    { label: 'Users', value: '50,000' },
    { label: 'NLP F1', value: '≥85%' }
  ];

  let metricX = 45;
  metrics.forEach(function(m) {
    const metricCard = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, metricX, 292, 155, 35);
    metricCard.getFill().setSolidFill(COLORS.secondary);
    metricCard.getBorder().setTransparent();

    const metricText = slide.insertTextBox(m.label + ': ' + m.value, metricX + 10, 300, 135, 20);
    metricText.getText().getTextStyle()
      .setFontSize(11)
      .setBold(true)
      .setForegroundColor(COLORS.white)
      .setFontFamily('Sarabun');
    metricX += 165;
  });

  // Tech stack
  const techBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 350, 660, 40);
  techBox.getFill().setSolidFill(COLORS.lightGray);
  techBox.getBorder().setTransparent();

  const techText = slide.insertTextBox('Tech: PyTorch, TensorFlow | Cloud: AWS/GCP | Infra: Kubernetes | DB: PostgreSQL, MongoDB | Security: WAF, SIEM', 45, 362, 630, 18);
  techText.getText().getTextStyle()
    .setFontSize(9)
    .setForegroundColor(COLORS.mediumGray)
    .setFontFamily('Sarabun');
}

// ====================================
// SLIDE 6: BUDGET
// ====================================
function createSlide6_Budget(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'งบประมาณโครงการ', '6/10');

  // Main budget
  const mainBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 75, 660, 55);
  mainBox.getFill().setSolidFill(COLORS.primary);
  mainBox.getBorder().setTransparent();

  const mainText = slide.insertTextBox('งบลงทุนรวม: 85 ล้านบาท          เงินสนับสนุนสูงสุด: 42.5 ล้านบาท (50%)', 50, 92, 550, 25);
  mainText.getText().getTextStyle()
    .setFontSize(16)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  // Rate badge
  const rateBadge = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 630, 80, 45, 45);
  rateBadge.getFill().setSolidFill(COLORS.accent);
  rateBadge.getBorder().setTransparent();

  const rateText = slide.insertTextBox('50%', 630, 93, 45, 22);
  rateText.getText().getTextStyle()
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.primary)
    .setFontFamily('Sarabun');
  rateText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Table
  const tableData = [
    { category: 'R&D (สวทช.)', invest: '59.0', support: '29.5', color: COLORS.rdLight },
    { category: 'Digital (depa)', invest: '26.0', support: '13.0', color: COLORS.digitalLight },
    { category: 'รวมทั้งหมด', invest: '85.0', support: '42.5', color: COLORS.lightGray }
  ];

  // Header
  const headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 30, 145, 660, 28);
  headerBg.getFill().setSolidFill(COLORS.secondary);
  headerBg.getBorder().setTransparent();

  const h1 = slide.insertTextBox('หมวด', 40, 150, 280, 20);
  h1.getText().getTextStyle().setFontSize(11).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');

  const h2 = slide.insertTextBox('ลงทุน (M)', 340, 150, 120, 20);
  h2.getText().getTextStyle().setFontSize(11).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');
  h2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const h3 = slide.insertTextBox('สนับสนุน (M)', 480, 150, 140, 20);
  h3.getText().getTextStyle().setFontSize(11).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');
  h3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  let rowY = 173;
  tableData.forEach(function(row) {
    const rowBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 30, rowY, 660, 32);
    rowBg.getFill().setSolidFill(row.color);
    rowBg.getBorder().setTransparent();

    const c1 = slide.insertTextBox(row.category, 40, rowY + 7, 280, 20);
    c1.getText().getTextStyle().setFontSize(11).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

    const c2 = slide.insertTextBox(row.invest, 340, rowY + 7, 120, 20);
    c2.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');
    c2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const c3 = slide.insertTextBox(row.support, 480, rowY + 7, 140, 20);
    c3.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.rdColor).setFontFamily('Sarabun');
    c3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    rowY += 32;
  });

  // Hardware summary
  const hwTitle = slide.insertTextBox('Hardware ที่ได้รับการสนับสนุน', 30, 285, 300, 20);
  hwTitle.getText().getTextStyle()
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(COLORS.dark)
    .setFontFamily('Sarabun');

  const hwData = [
    { name: 'Hardware R&D', invest: '9.0M', support: '4.5M', color: COLORS.rdLight },
    { name: 'Hardware Digital', invest: '10.25M', support: '5.125M', color: COLORS.digitalLight },
    { name: 'รวม Hardware', invest: '19.25M', support: '9.625M', color: COLORS.primary }
  ];

  let hwX = 30;
  hwData.forEach(function(hw) {
    const hwBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, hwX, 310, 215, 85);
    hwBox.getFill().setSolidFill(hw.color);
    hwBox.getBorder().setTransparent();

    const isTotal = hw.name === 'รวม Hardware';
    const hwName = slide.insertTextBox(hw.name, hwX + 10, 320, 195, 18);
    hwName.getText().getTextStyle()
      .setFontSize(11)
      .setBold(true)
      .setForegroundColor(isTotal ? COLORS.white : COLORS.dark)
      .setFontFamily('Sarabun');

    const hwInvest = slide.insertTextBox('ลงทุน: ' + hw.invest, hwX + 10, 342, 195, 16);
    hwInvest.getText().getTextStyle()
      .setFontSize(10)
      .setForegroundColor(isTotal ? COLORS.lightGray : COLORS.mediumGray)
      .setFontFamily('Sarabun');

    const hwSupport = slide.insertTextBox('สนับสนุน: ' + hw.support, hwX + 10, 362, 195, 22);
    hwSupport.getText().getTextStyle()
      .setFontSize(13)
      .setBold(true)
      .setForegroundColor(isTotal ? COLORS.accent : COLORS.rdColor)
      .setFontFamily('Sarabun');

    hwX += 225;
  });
}

// ====================================
// SLIDE 7: TIMELINE
// ====================================
function createSlide7_Timeline(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'แผนการดำเนินงานและเบิกจ่าย', '7/10');

  // Timeline bar
  const timelineBar = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 80, 85, 560, 6);
  timelineBar.getFill().setSolidFill(COLORS.lightGray);
  timelineBar.getBorder().setTransparent();

  // Phase markers
  const phase1Dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 180, 75, 26, 26);
  phase1Dot.getFill().setSolidFill(COLORS.rdColor);
  phase1Dot.getBorder().getLineFill().setSolidFill(COLORS.white);
  phase1Dot.getBorder().setWeight(2);

  const phase1Num = slide.insertTextBox('1', 180, 80, 26, 18);
  phase1Num.getText().getTextStyle().setFontSize(11).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');
  phase1Num.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const phase2Dot = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 510, 75, 26, 26);
  phase2Dot.getFill().setSolidFill(COLORS.digitalColor);
  phase2Dot.getBorder().getLineFill().setSolidFill(COLORS.white);
  phase2Dot.getBorder().setWeight(2);

  const phase2Num = slide.insertTextBox('2', 510, 80, 26, 18);
  phase2Num.getText().getTextStyle().setFontSize(11).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');
  phase2Num.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Phase 1 box
  const phase1Box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 115, 330, 175);
  phase1Box.getFill().setSolidFill(COLORS.rdLight);
  phase1Box.getBorder().getLineFill().setSolidFill(COLORS.rdColor);
  phase1Box.getBorder().setWeight(2);

  const phase1Title = slide.insertTextBox('งวดที่ 1 (6 เดือน)', 45, 122, 300, 18);
  phase1Title.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.rdColor).setFontFamily('Sarabun');

  const phase1Budget = slide.insertTextBox('ลงทุน: 42.5M | สนับสนุน: 21.25M', 45, 142, 300, 16);
  phase1Budget.getText().getTextStyle().setFontSize(10).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

  const phase1KPI = slide.insertTextBox(
    'R&D:\n• Prototype 3 AI Agents\n• NLP F1 Score ≥ 80%\n\nDigital:\n• Hardware 50%\n• Cloud Infrastructure พร้อม',
    45, 165, 300, 115);
  phase1KPI.getText().getTextStyle().setFontSize(9).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

  // Phase 2 box
  const phase2Box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 370, 115, 320, 175);
  phase2Box.getFill().setSolidFill(COLORS.digitalLight);
  phase2Box.getBorder().getLineFill().setSolidFill(COLORS.digitalColor);
  phase2Box.getBorder().setWeight(2);

  const phase2Title = slide.insertTextBox('งวดที่ 2 (12 เดือน)', 385, 122, 290, 18);
  phase2Title.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.digitalColor).setFontFamily('Sarabun');

  const phase2Budget = slide.insertTextBox('ลงทุน: 42.5M | สนับสนุน: 21.25M', 385, 142, 290, 16);
  phase2Budget.getText().getTextStyle().setFontSize(10).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

  const phase2KPI = slide.insertTextBox(
    'R&D:\n• 5 AI Agents Production\n• NLP F1 ≥ 85%, สิทธิบัตร ≥2\n\nDigital:\n• Hardware 100%\n• Security Audit ผ่าน, Go Live',
    385, 165, 290, 115);
  phase2KPI.getText().getTextStyle().setFontSize(9).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

  // Summary bar
  const summaryBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 305, 660, 85);
  summaryBox.getFill().setSolidFill(COLORS.primary);
  summaryBox.getBorder().setTransparent();

  const summaryTitle = slide.insertTextBox('สรุปการเบิกจ่าย', 45, 315, 200, 18);
  summaryTitle.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.accent).setFontFamily('Sarabun');

  const summaryGrid = [
    { label: 'รวมลงทุน', value: '85M' },
    { label: 'รวมสนับสนุน', value: '42.5M' },
    { label: 'งวดที่ 1', value: '21.25M' },
    { label: 'งวดที่ 2', value: '21.25M' }
  ];

  let sumX = 45;
  summaryGrid.forEach(function(s) {
    const sumCard = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, sumX, 340, 150, 40);
    sumCard.getFill().setSolidFill(COLORS.secondary);
    sumCard.getBorder().setTransparent();

    const sumLabel = slide.insertTextBox(s.label, sumX + 10, 345, 70, 14);
    sumLabel.getText().getTextStyle().setFontSize(9).setForegroundColor(COLORS.accentLight).setFontFamily('Sarabun');

    const sumValue = slide.insertTextBox(s.value, sumX + 10, 358, 130, 20);
    sumValue.getText().getTextStyle().setFontSize(14).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');

    sumX += 160;
  });
}

// ====================================
// SLIDE 8: KPIs
// ====================================
function createSlide8_KPI(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'ตัวชี้วัดความสำเร็จ (KPIs)', '8/10');

  const kpiCategories = [
    {
      title: 'R&D Output',
      color: COLORS.rdLight,
      borderColor: COLORS.rdColor,
      kpis: ['NLP F1 ≥ 85%', 'สิทธิบัตร ≥ 2', 'บทความ ≥ 3', 'Dataset ≥ 10,000']
    },
    {
      title: 'Performance',
      color: COLORS.digitalLight,
      borderColor: COLORS.digitalColor,
      kpis: ['Uptime ≥ 99.5%', 'Response < 200ms', '≥ 500 TPS', '50,000 Users']
    },
    {
      title: 'Product',
      color: '#FFF3E0',
      borderColor: COLORS.warning,
      kpis: ['5 AI Agents', 'Mobile App Live', 'Physical AI', 'Security Audit']
    },
    {
      title: 'Validation',
      color: '#FCE4EC',
      borderColor: '#E91E63',
      kpis: ['UAT ≥ 50 Users', 'Satisfaction ≥ 3.5', '100,000 Target', 'depa Cert']
    }
  ];

  const cardW = 330;
  const cardH = 145;
  const gap = 10;

  const positions = [
    { x: 30, y: 75 },
    { x: 30 + cardW + gap, y: 75 },
    { x: 30, y: 75 + cardH + gap },
    { x: 30 + cardW + gap, y: 75 + cardH + gap }
  ];

  kpiCategories.forEach(function(cat, i) {
    const pos = positions[i];

    const catBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, pos.x, pos.y, cardW, cardH);
    catBox.getFill().setSolidFill(cat.color);
    catBox.getBorder().getLineFill().setSolidFill(cat.borderColor);
    catBox.getBorder().setWeight(2);

    const catTitle = slide.insertTextBox(cat.title, pos.x + 10, pos.y + 8, 310, 18);
    catTitle.getText().getTextStyle()
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(cat.borderColor)
      .setFontFamily('Sarabun');

    let kpiY = pos.y + 32;
    cat.kpis.forEach(function(kpi) {
      const kpiBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, pos.x + 10, kpiY, cardW - 20, 24);
      kpiBox.getFill().setSolidFill(COLORS.white);
      kpiBox.getBorder().setTransparent();

      const kpiText = slide.insertTextBox('✓ ' + kpi, pos.x + 18, kpiY + 4, cardW - 36, 18);
      kpiText.getText().getTextStyle()
        .setFontSize(10)
        .setForegroundColor(COLORS.dark)
        .setFontFamily('Sarabun');

      kpiY += 27;
    });
  });

  // Footer
  const footerNote = slide.insertTextBox('* ต้องได้รับหนังสือรับรองจาก สวทช. และ depa', 30, 385, 660, 18);
  footerNote.getText().getTextStyle()
    .setFontSize(9)
    .setItalic(true)
    .setForegroundColor(COLORS.mediumGray)
    .setFontFamily('Sarabun');
}

// ====================================
// SLIDE 9: IMPACT
// ====================================
function createSlide9_Impact(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.white);
  addSlideHeader(slide, 'ผลกระทบเชิงบวกต่อประเทศ', '9/10');

  // Direct Benefits
  const directBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 75, 330, 140);
  directBox.getFill().setSolidFill(COLORS.rdLight);
  directBox.getBorder().getLineFill().setSolidFill(COLORS.rdColor);
  directBox.getBorder().setWeight(2);

  const directTitle = slide.insertTextBox('ผลประโยชน์ทางตรง', 45, 82, 300, 18);
  directTitle.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.rdColor).setFontFamily('Sarabun');

  const directText = slide.insertTextBox(
    '• จ้างงานคนไทย 15 ตำแหน่ง\n• พัฒนาทักษะ AI/ML 15 ราย\n• รายได้เพิ่ม 150-200M/ปี\n• ใช้วัตถุดิบในประเทศ 65M\n• พัฒนา Supply Chain 5-10 ราย',
    45, 102, 300, 100);
  directText.getText().getTextStyle().setFontSize(10).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

  // Indirect Benefits
  const indirectBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 370, 75, 320, 140);
  indirectBox.getFill().setSolidFill(COLORS.digitalLight);
  indirectBox.getBorder().getLineFill().setSolidFill(COLORS.digitalColor);
  indirectBox.getBorder().setWeight(2);

  const indirectTitle = slide.insertTextBox('ผลประโยชน์ทางอ้อม', 385, 82, 290, 18);
  indirectTitle.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.digitalColor).setFontFamily('Sarabun');

  const indirectText = slide.insertTextBox(
    '• ลดความเสียหาย Scam 500-1,000M/ปี\n• ผู้ใช้ปีแรก 55,000 ราย\n• Thai Scam Dataset เปิดให้ใช้ฟรี\n• เพิ่ม Financial Literacy\n• ยกระดับมาตรฐาน AI ไทย',
    385, 102, 290, 100);
  indirectText.getText().getTextStyle().setFontSize(10).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

  // Innovation box
  const innovationBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 225, 660, 70);
  innovationBox.getFill().setSolidFill(COLORS.primary);
  innovationBox.getBorder().setTransparent();

  const innovationBadge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 45, 238, 85, 45);
  innovationBadge.getFill().setSolidFill(COLORS.accent);
  innovationBadge.getBorder().setTransparent();

  const innovationBadgeText = slide.insertTextBox('First in\nASEAN', 45, 245, 85, 35);
  innovationBadgeText.getText().getTextStyle().setFontSize(10).setBold(true).setForegroundColor(COLORS.primary).setFontFamily('Sarabun');
  innovationBadgeText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const innovationTitle = slide.insertTextBox('นวัตกรรมครั้งแรกในประเทศไทยและอาเซียน', 145, 235, 530, 20);
  innovationTitle.getText().getTextStyle().setFontSize(13).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');

  const innovationList = slide.insertTextBox(
    'Physical AI Agent | Multi-Agent AI | Explainable AI | Thai Scam Dataset',
    145, 258, 530, 30);
  innovationList.getText().getTextStyle().setFontSize(10).setForegroundColor(COLORS.accentLight).setFontFamily('Sarabun');

  // Research outputs
  const researchTitle = slide.insertTextBox('ผลงานวิจัย', 30, 305, 200, 18);
  researchTitle.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

  const researchItems = [
    { num: '01', text: 'สิทธิบัตร ≥ 2' },
    { num: '02', text: 'บทความ ≥ 3' },
    { num: '03', text: 'Dataset ≥ 10,000' },
    { num: '04', text: 'UAT ≥ 50 คน' }
  ];

  let resX = 30;
  researchItems.forEach(function(r) {
    const resBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, resX, 328, 160, 50);
    resBox.getFill().setSolidFill(COLORS.lightGray);
    resBox.getBorder().setTransparent();

    const resNum = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, resX + 10, 335, 28, 28);
    resNum.getFill().setSolidFill(COLORS.primary);
    resNum.getBorder().setTransparent();

    const resNumText = slide.insertTextBox(r.num, resX + 10, 341, 28, 18);
    resNumText.getText().getTextStyle().setFontSize(9).setBold(true).setForegroundColor(COLORS.white).setFontFamily('Sarabun');
    resNumText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const resText = slide.insertTextBox(r.text, resX + 45, 343, 110, 28);
    resText.getText().getTextStyle().setFontSize(10).setForegroundColor(COLORS.dark).setFontFamily('Sarabun');

    resX += 170;
  });
}

// ====================================
// SLIDE 10: SUMMARY
// ====================================
function createSlide10_Summary(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(COLORS.primary);

  // Top accent
  const topAccent = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, SLIDE_WIDTH, 6);
  topAccent.getFill().setSolidFill(COLORS.accent);
  topAccent.getBorder().setTransparent();

  // Title
  const title = slide.insertTextBox('สรุปการขอรับการส่งเสริม', 40, 25, SLIDE_WIDTH - 80, 30);
  title.getText().getTextStyle()
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');
  title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Summary cards
  const summaryData = [
    { label: 'งบลงทุน', value: '85M' },
    { label: 'สนับสนุน', value: '42.5M' },
    { label: 'อัตรา', value: '50%' },
    { label: 'ระยะเวลา', value: '12 เดือน' }
  ];

  let sumX = 30;
  summaryData.forEach(function(s) {
    const sumBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, sumX, 65, 160, 55);
    sumBox.getFill().setSolidFill(COLORS.white);
    sumBox.getBorder().setTransparent();

    const sumValue = slide.insertTextBox(s.value, sumX + 10, 72, 140, 22);
    sumValue.getText().getTextStyle()
      .setFontSize(16)
      .setBold(true)
      .setForegroundColor(COLORS.primary)
      .setFontFamily('Sarabun');

    const sumLabel = slide.insertTextBox(s.label, sumX + 10, 95, 140, 18);
    sumLabel.getText().getTextStyle()
      .setFontSize(10)
      .setForegroundColor(COLORS.mediumGray)
      .setFontFamily('Sarabun');

    sumX += 170;
  });

  // Highlights box
  const highlightBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 135, 660, 120);
  highlightBox.getFill().setSolidFill(COLORS.secondary);
  highlightBox.getBorder().getLineFill().setSolidFill(COLORS.accent);
  highlightBox.getBorder().setWeight(2);

  const highlightTitle = slide.insertTextBox('จุดเด่นของโครงการ', 45, 145, 300, 18);
  highlightTitle.getText().getTextStyle()
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(COLORS.accent)
    .setFontFamily('Sarabun');

  const highlights = slide.insertTextBox(
    '✓ นวัตกรรม AI ครั้งแรกในอาเซียน - Physical AI + Multi-Agent\n' +
    '✓ ปกป้องคนไทยจาก Scam - ลดความเสียหาย 500-1,000M/ปี\n' +
    '✓ พัฒนาบุคลากรไทย - 15 ตำแหน่ง AI, ML, NLP\n' +
    '✓ สร้างงานวิจัย - สิทธิบัตร ≥2, บทความ ≥3, Dataset ≥10,000',
    45, 168, 630, 80);
  highlights.getText().getTextStyle()
    .setFontSize(11)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  // Agency box
  const agencyBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 270, 660, 40);
  agencyBox.getFill().setSolidFill(COLORS.white);
  agencyBox.getBorder().setTransparent();

  const agencyText = slide.insertTextBox('หน่วยรับรอง R&D: สวทช.          หน่วยรับรอง Digital: depa', 50, 280, 620, 22);
  agencyText.getText().getTextStyle()
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(COLORS.primary)
    .setFontFamily('Sarabun');
  agencyText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Contact box
  const contactBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 325, 660, 70);
  contactBox.getFill().setSolidFill(COLORS.accent);
  contactBox.getBorder().setTransparent();

  const contactLogo = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 45, 338, 45, 45);
  contactLogo.getFill().setSolidFill(COLORS.primary);
  contactLogo.getBorder().setTransparent();

  const contactLogoText = slide.insertTextBox('AW', 45, 352, 45, 22);
  contactLogoText.getText().getTextStyle()
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.accent)
    .setFontFamily('Sarabun');
  contactLogoText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const contactCompany = slide.insertTextBox(PROJECT.company, 105, 335, 400, 22);
  contactCompany.getText().getTextStyle()
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.primary)
    .setFontFamily('Sarabun');

  const contactDetails = slide.insertTextBox('นายสุรเกียรติ ยาวะโนภาส | 082-619-5459 | admin@aslanwealth.com', 105, 358, 550, 20);
  contactDetails.getText().getTextStyle()
    .setFontSize(11)
    .setForegroundColor(COLORS.dark)
    .setFontFamily('Sarabun');

  // Slide number
  const slideNum = slide.insertTextBox('10/10', SLIDE_WIDTH - 60, SLIDE_HEIGHT - 25, 45, 18);
  slideNum.getText().getTextStyle()
    .setFontSize(10)
    .setForegroundColor(COLORS.accentLight)
    .setFontFamily('Sarabun');

  // Bottom accent
  const bottomAccent = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, SLIDE_HEIGHT - 6, SLIDE_WIDTH, 6);
  bottomAccent.getFill().setSolidFill(COLORS.accent);
  bottomAccent.getBorder().setTransparent();
}

// ====================================
// HELPER FUNCTIONS
// ====================================
function addSlideHeader(slide, title, slideNum) {
  const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, SLIDE_WIDTH, 55);
  header.getFill().setSolidFill(COLORS.primary);
  header.getBorder().setTransparent();

  const accentLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 55, SLIDE_WIDTH, 3);
  accentLine.getFill().setSolidFill(COLORS.accent);
  accentLine.getBorder().setTransparent();

  const headerText = slide.insertTextBox(title, 30, 15, 550, 28);
  headerText.getText().getTextStyle()
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(COLORS.white)
    .setFontFamily('Sarabun');

  const slideNumText = slide.insertTextBox(slideNum, SLIDE_WIDTH - 60, 18, 45, 22);
  slideNumText.getText().getTextStyle()
    .setFontSize(12)
    .setForegroundColor(COLORS.accentLight)
    .setFontFamily('Sarabun');
  slideNumText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
}

// ====================================
// MENU FUNCTION
// ====================================
function onOpen() {
  const ui = SlidesApp.getUi();
  ui.createMenu('BOI Tools')
    .addItem('สร้าง BOI Presentation', 'createBOIPresentation')
    .addToUi();
}
