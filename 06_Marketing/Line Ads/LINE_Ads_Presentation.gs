/**
 * LINE Ads Campaign Presentation Generator
 * Aslan Wealth Channel
 * งบประมาณ: 100,000 บาท | ระยะเวลา: 4 สัปดาห์
 *
 * วิธีใช้งาน:
 * 1. เปิด Google Apps Script (script.google.com)
 * 2. สร้างโปรเจกต์ใหม่
 * 3. วาง Code นี้ทั้งหมด
 * 4. รัน function: createLINEAdsCampaignPresentation()
 * 5. อนุญาต permissions ที่ขอ
 * 6. Presentation จะสร้างใน Google Drive อัตโนมัติ
 *
 * Slide Ratio: 16:9 (720x405 points)
 */

// ==================== CONFIGURATION ====================
const CONFIG = {
  BRAND_NAME: 'Aslan Wealth Channel',
  WEBSITE: 'https://aslanwealth.com',
  LINE_OA: '@aslanwealth',
  EMAIL: 'contact@aslanwealth.com',
  BUDGET: '100,000',
  DURATION: '4 สัปดาห์',
  CURRENCY: 'THB',

  // Slide Dimensions (16:9 ratio)
  SLIDE: {
    WIDTH: 720,
    HEIGHT: 405,
    MARGIN: 25,
    HEADER_HEIGHT: 55
  },

  // Brand Colors
  COLORS: {
    PRIMARY: '#1A365D',      // Navy Blue
    SECONDARY: '#2B6CB0',    // Blue
    ACCENT: '#38A169',       // Green
    WARNING: '#DD6B20',      // Orange
    GOLD: '#D69E2E',         // Gold
    WHITE: '#FFFFFF',
    BLACK: '#1A202C',
    LIGHT_GRAY: '#F7FAFC',
    GRAY: '#718096',
    DARK_GRAY: '#4A5568'
  }
};

// ==================== ROYALTY-FREE IMAGES (Unsplash) ====================
const IMAGES = {
  // Finance & Investment
  INVESTMENT: 'https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=1200&q=80',
  CHART_UP: 'https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=1200&q=80',
  ANALYTICS: 'https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=1200&q=80',
  MONEY_GROW: 'https://images.unsplash.com/photo-1579621970563-ebec7560ff3e?w=1200&q=80',
  TRADING: 'https://images.unsplash.com/photo-1642790106117-e829e14a795f?w=1200&q=80',

  // Technology & AI
  AI_TECH: 'https://images.unsplash.com/photo-1677442136019-21780ecad995?w=1200&q=80',
  AI_ROBOT: 'https://images.unsplash.com/photo-1485827404703-89b55fcc595e?w=1200&q=80',
  DIGITAL: 'https://images.unsplash.com/photo-1518770660439-4636190af475?w=1200&q=80',

  // Mobile & LINE
  MOBILE_APP: 'https://images.unsplash.com/photo-1512941937669-90a1b58e7e9c?w=1200&q=80',
  SMARTPHONE: 'https://images.unsplash.com/photo-1556656793-08538906a9f8?w=1200&q=80',

  // Business & Success
  SUCCESS: 'https://images.unsplash.com/photo-1553729459-efe14ef6055d?w=1200&q=80',
  TEAM: 'https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=1200&q=80',
  MEETING: 'https://images.unsplash.com/photo-1542744173-8e7e53415bb0?w=1200&q=80',
  TARGET: 'https://images.unsplash.com/photo-1533750349088-cd871a92f312?w=1200&q=80',

  // Thailand
  THAILAND_CITY: 'https://images.unsplash.com/photo-1508009603885-50cf7c579365?w=1200&q=80',
  BANGKOK: 'https://images.unsplash.com/photo-1563492065599-3520f775eeed?w=1200&q=80',

  // Education & Learning
  LEARNING: 'https://images.unsplash.com/photo-1434030216411-0b793f4b4173?w=1200&q=80',
  EDUCATION: 'https://images.unsplash.com/photo-1523240795612-9a054b0db644?w=1200&q=80'
};

// ==================== MAIN FUNCTION ====================
function createLINEAdsCampaignPresentation() {
  // สร้าง Presentation ใหม่
  const presentation = SlidesApp.create('LINE Ads Campaign - Aslan Wealth Channel (4 Weeks)');

  // ตั้งค่า Slide Size เป็น 16:9
  // Google Slides default ใช้ points (1 point = 1/72 inch)
  // 720 x 405 points = 10" x 5.625" = 16:9 ratio

  const slides = presentation.getSlides();

  // ลบ slide แรกที่ว่างเปล่า
  if (slides.length > 0) {
    slides[0].remove();
  }

  Logger.log('เริ่มสร้าง Presentation (16:9 ratio)...');

  // สร้างทุก Slides (21 slides รวม Ad Copy ใหม่)
  createSlide01_Title(presentation);
  createSlide02_Agenda(presentation);
  createSlide03_Overview(presentation);
  createSlide04_WhyLINE(presentation);
  createSlide05_Objectives(presentation);
  createSlide06_CampaignStructure(presentation);
  createSlide07_BudgetAllocation(presentation);
  createSlide08_TargetAudience(presentation);
  createSlide09_AudienceDetails(presentation);
  createSlide10_CreativeSpecs(presentation);
  createSlide11_AdExamples(presentation);
  createSlide12_AdCopyAI(presentation);        // NEW: AI Learning + Aslan Agent
  createSlide13_Placements(presentation);
  createSlide14_ContentCalendar(presentation);
  createSlide15_KPIs(presentation);
  createSlide16_Timeline(presentation);
  createSlide17_RichMenu(presentation);
  createSlide18_Compliance(presentation);
  createSlide19_ExpectedResults(presentation);
  createSlide20_NextSteps(presentation);
  createSlide21_ThankYou(presentation);

  Logger.log('สร้าง Presentation สำเร็จ!');
  Logger.log('URL: ' + presentation.getUrl());
  Logger.log('จำนวน Slides: 21');

  return presentation.getUrl();
}

// ==================== SLIDE 1: TITLE ====================
function createSlide01_Title(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;
  const H = CONFIG.SLIDE.HEIGHT;

  // Background - Gradient effect with shape overlay
  slide.getBackground().setSolidFill(CONFIG.COLORS.PRIMARY);

  // Decorative diagonal shape
  const diag = slide.insertShape(SlidesApp.ShapeType.PARALLELOGRAM, W * 0.6, 0, W * 0.5, H);
  diag.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  diag.setRotation(0);
  diag.getBorder().setTransparent();
  diag.sendToBack();

  // Logo Box
  const logoBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 20, 120, 35);
  logoBox.getFill().setSolidFill(CONFIG.COLORS.WHITE);
  logoBox.getBorder().setTransparent();

  const logoText = logoBox.getText();
  logoText.setText('ASLAN');
  logoText.getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);
  logoText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Main Title
  const title = slide.insertTextBox('LINE Ads Campaign', 30, 100, 450, 50);
  title.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(36)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // Subtitle
  const subtitle = slide.insertTextBox('แผนการตลาดออนไลน์ 4 สัปดาห์', 30, 150, 450, 35);
  subtitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(20)
    .setForegroundColor(CONFIG.COLORS.ACCENT);

  // Brand Name
  const brand = slide.insertTextBox('Aslan Wealth Channel', 30, 195, 450, 30);
  brand.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(18)
    .setForegroundColor(CONFIG.COLORS.LIGHT_GRAY);

  // Budget Box
  const budgetBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 240, 180, 50);
  budgetBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  budgetBox.getBorder().setTransparent();

  const budgetText = slide.insertTextBox('฿100,000', 30, 250, 180, 35);
  budgetText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(24)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  budgetText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Duration Box
  const durationBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 220, 240, 120, 50);
  durationBox.getFill().setSolidFill(CONFIG.COLORS.WARNING);
  durationBox.getBorder().setTransparent();

  const durationText = slide.insertTextBox('4 สัปดาห์', 220, 250, 120, 35);
  durationText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  durationText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Footer Info
  const footer = slide.insertTextBox('aslanwealth.com  |  @aslanwealth  |  contact@aslanwealth.com', 30, 370, 400, 20);
  footer.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.LIGHT_GRAY);

  // Date
  const date = slide.insertTextBox('มกราคม 2025', W - 130, 370, 100, 20);
  date.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.GRAY);
  date.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
}

// ==================== SLIDE 2: AGENDA ====================
function createSlide02_Agenda(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;
  const H = CONFIG.SLIDE.HEIGHT;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  // Header
  addSlideHeader(slide, 'Agenda', '📋');

  // Agenda Items - 2 columns
  const itemsLeft = [
    '1. ภาพรวมและวัตถุประสงค์',
    '2. ทำไมต้อง LINE Ads?',
    '3. โครงสร้างแคมเปญ',
    '4. การจัดสรรงบประมาณ',
    '5. กลุ่มเป้าหมาย'
  ];

  const itemsRight = [
    '6. ชิ้นงานโฆษณา + AI Copy',
    '7. ปฏิทินคอนเทนต์ 4 สัปดาห์',
    '8. ตัวชี้วัด KPIs',
    '9. ผลลัพธ์ที่คาดหวัง',
    '10. ขั้นตอนถัดไป'
  ];

  let yPos = 75;
  itemsLeft.forEach((item, i) => {
    const text = slide.insertTextBox(item, 40, yPos, 300, 28);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(14)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 32;
  });

  yPos = 75;
  itemsRight.forEach((item, i) => {
    const text = slide.insertTextBox(item, 380, yPos, 300, 28);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(14)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 32;
  });

  // Highlight Box
  const highlightBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 250, W - 80, 50);
  highlightBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  highlightBox.getBorder().setTransparent();

  const highlightText = slide.insertTextBox('💡 ใหม่! เพิ่ม Ad Copy สไตล์ AI Learning + Aslan Agent เพื่อนคู่คิด', 50, 260, W - 100, 30);
  highlightText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  highlightText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Footer
  addSlideFooter(slide, 2);
}

// ==================== SLIDE 3: OVERVIEW ====================
function createSlide03_Overview(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;
  const H = CONFIG.SLIDE.HEIGHT;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ภาพรวม Aslan Wealth Channel', '🎯');

  // Left Content - About
  const aboutTitle = slide.insertTextBox('เกี่ยวกับเรา', 25, 70, 200, 22);
  aboutTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const aboutText = slide.insertTextBox(
    'แหล่งความรู้ด้านการลงทุนสำหรับคนไทย\n\n' +
    '✅ คอร์สเรียนการลงทุน\n' +
    '✅ AI Assistant (Aslan Agent)\n' +
    '✅ บทความและข่าวสารการเงิน\n' +
    '✅ Webinar และ Workshop',
    25, 95, 220, 130
  );
  aboutText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setForegroundColor(CONFIG.COLORS.BLACK);

  // Contact Box
  const contactBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 235, 220, 70);
  contactBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  contactBox.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.SECONDARY);

  const contactText = slide.insertTextBox(
    '🌐 aslanwealth.com\n' +
    '📱 LINE: @aslanwealth\n' +
    '📧 contact@aslanwealth.com',
    35, 245, 200, 55
  );
  contactText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.BLACK);

  // Right - Stats Grid
  const statsTitle = slide.insertTextBox('ตัวเลขสำคัญ', 270, 70, 200, 22);
  statsTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  // Stats - 2x2 Grid
  createStatBox(slide, 270, 95, '฿100K', 'งบประมาณ', CONFIG.COLORS.ACCENT, 100, 60);
  createStatBox(slide, 380, 95, '4 Wks', 'ระยะเวลา', CONFIG.COLORS.WARNING, 100, 60);
  createStatBox(slide, 270, 165, '4', 'แคมเปญ', CONFIG.COLORS.SECONDARY, 100, 60);
  createStatBox(slide, 380, 165, '3', 'กลุ่มเป้าหมาย', CONFIG.COLORS.PRIMARY, 100, 60);

  // Unique Selling Point
  const uspBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 270, 240, 210, 65);
  uspBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  uspBox.getBorder().setTransparent();

  const uspText = slide.insertTextBox('🤖 Aslan Agent\nAI เพื่อนคู่คิดการลงทุน', 280, 250, 190, 45);
  uspText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  uspText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Right Side - Decorative
  const sideBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 500, 60, 200, 250);
  sideBar.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  sideBar.getBorder().setTransparent();

  const sideText = slide.insertTextBox('🎯\n\nเรียนรู้\nลงทุน\nกับ AI', 510, 100, 180, 180);
  sideText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  sideText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  addSlideFooter(slide, 3);
}

// ==================== SLIDE 4: WHY LINE ====================
function createSlide04_WhyLINE(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ทำไมต้อง LINE Ads?', '📱');

  // Stats Row
  const stats = [
    { number: '54M+', label: 'ผู้ใช้/เดือน', color: CONFIG.COLORS.PRIMARY },
    { number: '#2', label: 'แอปยอดนิยม', color: CONFIG.COLORS.SECONDARY },
    { number: '91%', label: 'ใช้ทุกวัน', color: CONFIG.COLORS.ACCENT },
    { number: '฿15', label: 'CPM เฉลี่ย', color: CONFIG.COLORS.WARNING }
  ];

  let xPos = 25;
  stats.forEach((stat) => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, xPos, 70, 160, 75);
    box.getFill().setSolidFill(stat.color);
    box.getBorder().setTransparent();

    const numText = slide.insertTextBox(stat.number, xPos, 78, 160, 35);
    numText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(24)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    numText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const labelText = slide.insertTextBox(stat.label, xPos, 115, 160, 22);
    labelText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(11)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    labelText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    xPos += 170;
  });

  // Benefits Section
  const benefitsTitle = slide.insertTextBox('ข้อดีของ LINE Ads สำหรับธุรกิจการเงิน', 25, 160, 400, 22);
  benefitsTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const benefits = [
    '✅ เข้าถึงกลุ่มการเงินผ่าน LINE Today, Wallet',
    '✅ Gain Friends - ฟีเจอร์เฉพาะ LINE',
    '✅ Retargeting คนที่ Block ได้',
    '✅ ต้นทุนต่ำกว่า Facebook 7 เท่า'
  ];

  let yPos = 185;
  benefits.forEach(benefit => {
    const text = slide.insertTextBox(benefit, 25, yPos, 350, 22);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(11)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 25;
  });

  // Comparison Box
  const compBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 400, 160, 290, 130);
  compBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  compBox.getBorder().setTransparent();

  const compText = slide.insertTextBox(
    '🆚 เทียบ Facebook\n\n' +
    'LINE CPM: ฿15\n' +
    'Facebook CPM: ฿110\n\n' +
    '✅ LINE ถูกกว่า 7 เท่า!',
    415, 168, 260, 110
  );
  compText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 4);
}

// ==================== SLIDE 5: OBJECTIVES ====================
function createSlide05_Objectives(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'วัตถุประสงค์แคมเปญ 4 สัปดาห์', '🎯');

  // Primary Objectives - Left
  const primaryTitle = slide.insertTextBox('เป้าหมายหลัก', 25, 70, 200, 20);
  primaryTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const objectives = [
    { icon: '👥', text: 'เพิ่มเพื่อน LINE OA', target: '1,500-2,500 คน' },
    { icon: '📺', text: 'Video Views', target: '50,000+ views' },
    { icon: '🔗', text: 'Traffic เข้าเว็บ', target: '8,000+ clicks' },
    { icon: '📝', text: 'Leads คุณภาพ', target: '200+ รายชื่อ' }
  ];

  let yPos = 95;
  objectives.forEach(obj => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, yPos, 220, 42);
    box.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    box.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.SECONDARY);

    const text = slide.insertTextBox(obj.icon + ' ' + obj.text + '\n🎯 ' + obj.target, 32, yPos + 5, 205, 35);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(10)
      .setForegroundColor(CONFIG.COLORS.BLACK);

    yPos += 48;
  });

  // Secondary - Right
  const secondaryTitle = slide.insertTextBox('เป้าหมายรอง', 270, 70, 200, 20);
  secondaryTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.SECONDARY);

  const secondary = [
    '📊 สร้าง Brand Awareness',
    '👥 สร้าง Community นักลงทุน',
    '📈 เก็บ Data สำหรับ Retargeting',
    '🤖 โปรโมท Aslan Agent AI'
  ];

  yPos = 95;
  secondary.forEach(item => {
    const text = slide.insertTextBox(item, 270, yPos, 200, 22);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(11)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 28;
  });

  // Success Metrics Box
  const metricsBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 270, 210, 200, 85);
  metricsBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  metricsBox.getBorder().setTransparent();

  const metricsText = slide.insertTextBox(
    '📊 เมตริกความสำเร็จ\n' +
    '• CPF < 40 บาท\n' +
    '• CTR > 1%\n' +
    '• ROAS > 2x',
    280, 218, 180, 70
  );
  metricsText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // AI Highlight
  const aiBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 490, 70, 200, 225);
  aiBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  aiBox.getBorder().setTransparent();

  const aiText = slide.insertTextBox(
    '🤖\n\nAslan Agent\n\nAI เพื่อนคู่คิด\nการลงทุน\n\nเรียนรู้ได้\n24 ชั่วโมง',
    500, 80, 180, 200
  );
  aiText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  aiText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  addSlideFooter(slide, 5);
}

// ==================== SLIDE 6: CAMPAIGN STRUCTURE ====================
function createSlide06_CampaignStructure(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'โครงสร้าง 4 แคมเปญ', '📊');

  const campaigns = [
    { name: 'เพิ่มเพื่อน', obj: 'Gain Friends', budget: '40K', pct: '40%', color: CONFIG.COLORS.PRIMARY, icon: '👥' },
    { name: 'Video Views', obj: 'Awareness', budget: '25K', pct: '25%', color: CONFIG.COLORS.SECONDARY, icon: '📺' },
    { name: 'Traffic', obj: 'Website', budget: '20K', pct: '20%', color: CONFIG.COLORS.ACCENT, icon: '🔗' },
    { name: 'Retarget', obj: 'Conversions', budget: '15K', pct: '15%', color: CONFIG.COLORS.WARNING, icon: '🎯' }
  ];

  let xPos = 20;
  campaigns.forEach((c) => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, xPos, 70, 165, 145);
    box.getFill().setSolidFill(c.color);
    box.getBorder().setTransparent();

    // Icon
    const icon = slide.insertTextBox(c.icon, xPos, 78, 165, 30);
    icon.getText().getTextStyle().setFontSize(22);
    icon.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Name
    const name = slide.insertTextBox(c.name, xPos, 108, 165, 22);
    name.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    name.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Objective
    const obj = slide.insertTextBox(c.obj, xPos, 128, 165, 18);
    obj.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.LIGHT_GRAY);
    obj.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Budget
    const budget = slide.insertTextBox('฿' + c.budget, xPos, 150, 165, 28);
    budget.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(18)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    budget.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Percentage
    const pct = slide.insertTextBox(c.pct, xPos, 180, 165, 22);
    pct.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(14)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    pct.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    xPos += 175;
  });

  // Total Bar
  const totalBar = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 20, 225, W - 40, 40);
  totalBar.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  totalBar.getBorder().setTransparent();

  const totalText = slide.insertTextBox('💰 งบประมาณรวม: ฿100,000 บาท  |  ⏱️ 4 สัปดาห์  |  📅 งบต่อวัน: ~฿3,500', 20, 233, W - 40, 28);
  totalText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(13)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  totalText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Note
  const note = slide.insertTextBox('* งบต่อวันที่สูงขึ้นทำให้เห็นผลเร็วกว่าแคมเปญ 8 สัปดาห์', 20, 275, W - 40, 18);
  note.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.GRAY);
  note.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  addSlideFooter(slide, 6);
}

// ==================== SLIDE 7: BUDGET ALLOCATION ====================
function createSlide07_BudgetAllocation(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'การจัดสรรงบประมาณ', '💰');

  // Table Header
  const headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 25, 70, W - 50, 28);
  headerBg.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  headerBg.getBorder().setTransparent();

  const header = slide.insertTextBox('แคมเปญ              งบประมาณ        เป้าหมาย                ต้นทุน', 35, 75, W - 70, 20);
  header.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // Table Rows
  const rows = [
    { campaign: '1. เพิ่มเพื่อน', budget: '฿40,000', target: '1,000-1,600 คน', cost: 'CPF 25-40฿' },
    { campaign: '2. Video Views', budget: '฿25,000', target: '50,000+ views', cost: 'CPV 0.50฿' },
    { campaign: '3. Traffic', budget: '฿20,000', target: '4,000-5,000 clicks', cost: 'CPC 4-5฿' },
    { campaign: '4. Retargeting', budget: '฿15,000', target: '100-200 leads', cost: 'CPA 75-150฿' }
  ];

  let yPos = 100;
  rows.forEach((row, i) => {
    const bgColor = i % 2 === 0 ? CONFIG.COLORS.LIGHT_GRAY : CONFIG.COLORS.WHITE;
    const rowBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 25, yPos, W - 50, 26);
    rowBg.getFill().setSolidFill(bgColor);
    rowBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.GRAY);

    const rowText = slide.insertTextBox(
      row.campaign + '          ' + row.budget + '         ' + row.target + '          ' + row.cost,
      35, yPos + 5, W - 70, 18
    );
    rowText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(10)
      .setForegroundColor(CONFIG.COLORS.BLACK);

    yPos += 28;
  });

  // Benchmarks
  const benchTitle = slide.insertTextBox('📊 Benchmarks ไทย', 25, 220, 180, 18);
  benchTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const benchmarks = ['CPM: ~15฿', 'CPV: ~1.22฿', 'CPC: 3-10฿', 'CPF: 15-50฿'];
  yPos = 240;
  benchmarks.forEach(b => {
    const text = slide.insertTextBox('• ' + b, 25, yPos, 150, 16);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 18;
  });

  // 4 Week vs 8 Week Comparison
  const compBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 200, 218, 250, 80);
  compBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  compBox.getBorder().setTransparent();

  const compText = slide.insertTextBox(
    '⏱️ ทำไม 4 สัปดาห์?\n\n' +
    '• งบ/วัน สูงขึ้น = เห็นผลเร็ว\n' +
    '• เหมาะกับการ Test & Learn\n' +
    '• ปรับกลยุทธ์ได้รวดเร็ว',
    210, 225, 230, 68
  );
  compText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // Summary Box
  const summaryBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 470, 218, 200, 80);
  summaryBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  summaryBox.getBorder().setTransparent();

  const summaryText = slide.insertTextBox(
    '💰 สรุป\n\n' +
    'งบรวม: ฿100,000\n' +
    'ระยะเวลา: 4 สัปดาห์\n' +
    'งบ/วัน: ~฿3,500',
    480, 225, 180, 68
  );
  summaryText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 7);
}

// ==================== SLIDE 8: TARGET AUDIENCE ====================
function createSlide08_TargetAudience(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'กลุ่มเป้าหมาย 3 กลุ่มหลัก', '👥');

  const audiences = [
    {
      name: 'นักลงทุนมือใหม่-กลาง',
      age: '25-45 ปี',
      income: '25K-80K',
      interests: 'หุ้น, กองทุน, คริปโต',
      color: CONFIG.COLORS.PRIMARY
    },
    {
      name: 'High Net Worth',
      age: '35-55 ปี',
      income: '100K+',
      interests: 'Luxury, อสังหาฯ',
      color: CONFIG.COLORS.SECONDARY
    },
    {
      name: 'คนทำงานรุ่นใหม่',
      age: '22-32 ปี',
      income: '15K-40K',
      interests: 'Side Hustle, AI',
      color: CONFIG.COLORS.ACCENT
    }
  ];

  let xPos = 20;
  audiences.forEach((aud, i) => {
    // Card
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, xPos, 70, 220, 155);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().getLineFill().setSolidFill(aud.color);
    card.getBorder().setWeight(2);

    // Header
    const header = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, xPos, 70, 220, 32);
    header.getFill().setSolidFill(aud.color);
    header.getBorder().setTransparent();

    const headerText = slide.insertTextBox('กลุ่ม ' + String.fromCharCode(65 + i), xPos, 77, 220, 20);
    headerText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    headerText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Name
    const name = slide.insertTextBox(aud.name, xPos + 10, 108, 200, 20);
    name.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(11)
      .setBold(true)
      .setForegroundColor(aud.color);

    // Details
    const details = slide.insertTextBox(
      '👤 ' + aud.age + '\n' +
      '💰 ' + aud.income + '/เดือน\n' +
      '❤️ ' + aud.interests,
      xPos + 10, 130, 200, 60
    );
    details.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(10)
      .setForegroundColor(CONFIG.COLORS.BLACK);

    // AI Interest Badge
    if (i === 2) {
      const aiBadge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, xPos + 10, 190, 80, 22);
      aiBadge.getFill().setSolidFill(CONFIG.COLORS.WARNING);
      aiBadge.getBorder().setTransparent();

      const aiText = slide.insertTextBox('🤖 AI Ready', xPos + 15, 193, 70, 16);
      aiText.getText().getTextStyle()
        .setFontFamily('Arial')
        .setFontSize(8)
        .setBold(true)
        .setForegroundColor(CONFIG.COLORS.WHITE);
    }

    xPos += 230;
  });

  // Location Info
  const locBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 20, 235, W - 40, 55);
  locBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  locBox.getBorder().setTransparent();

  const locText = slide.insertTextBox(
    '📍 พื้นที่: กรุงเทพฯ, นนทบุรี, ปทุมธานี, ชลบุรี, เชียงใหม่, ภูเก็ต\n' +
    '📱 อุปกรณ์: iOS + Android  |  💼 อาชีพ: พนักงานบริษัท, ข้าราชการ, ผู้ประกอบการ SME',
    30, 245, W - 60, 40
  );
  locText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.BLACK);

  addSlideFooter(slide, 8);
}

// ==================== SLIDE 9: AUDIENCE DETAILS ====================
function createSlide09_AudienceDetails(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'Targeting Options', '🎯');

  // Targeting Types - Left
  const targetTitle = slide.insertTextBox('ประเภท Targeting', 25, 70, 200, 18);
  targetTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const targeting = [
    '📊 Demographics: เพศ, อายุ, พื้นที่',
    '❤️ Interests: การเงิน, ลงทุน, AI',
    '📱 Behavior: อ่าน LINE Today',
    '🔄 Retargeting: Website Visitors',
    '👥 Lookalike: จาก Converters'
  ];

  let yPos = 92;
  targeting.forEach(t => {
    const text = slide.insertTextBox(t, 25, yPos, 220, 18);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(10)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 22;
  });

  // Custom Audiences - Right
  const customTitle = slide.insertTextBox('Custom Audiences', 260, 70, 200, 18);
  customTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.SECONDARY);

  const custom = [
    '• Website Visitors (LINE Tag)',
    '• LINE OA Followers',
    '• Inactive Friends (30+ วัน)',
    '• Cart Abandoners',
    '• Upload จาก CRM'
  ];

  yPos = 92;
  custom.forEach(c => {
    const text = slide.insertTextBox(c, 260, yPos, 200, 18);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(10)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 22;
  });

  // Interest Categories Box
  const interestBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 210, W - 50, 85);
  interestBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  interestBox.getBorder().setTransparent();

  const interestTitle2 = slide.insertTextBox('❤️ Interest Categories ที่เลือก', 35, 218, 300, 18);
  interestTitle2.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  const interestText = slide.insertTextBox(
    '💰 การเงิน: หุ้น, กองทุน, Cryptocurrency, Forex\n' +
    '💼 ธุรกิจ: ข่าวธุรกิจ, Startup, ผู้ประกอบการ\n' +
    '🤖 เทคโนโลยี: AI, Machine Learning, Automation',
    35, 240, W - 80, 50
  );
  interestText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // AI Sidebar
  const aiSide = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 480, 60, 210, 140);
  aiSide.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  aiSide.getBorder().setTransparent();

  const aiSideText = slide.insertTextBox(
    '🤖 AI Targeting\n\n' +
    'กลุ่มที่สนใจ:\n' +
    '• ChatGPT / AI Tools\n' +
    '• Automation\n' +
    '• Tech News\n\n' +
    '→ Perfect for\n   Aslan Agent',
    490, 68, 190, 125
  );
  aiSideText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 9);
}

// ==================== SLIDE 10: CREATIVE SPECS ====================
function createSlide10_CreativeSpecs(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'สเปคชิ้นงานโฆษณา', '🎨');

  // Image Specs - Left
  const imageTitle = slide.insertTextBox('📷 Image Ads', 25, 70, 150, 18);
  imageTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const imageSpecs = [
    'Card: 1200 x 628 px',
    'Square: 1080 x 1080 px',
    'Small: 600 x 400 px',
    'Carousel: 1080x1080 x10',
    'File: JPG/PNG <10MB'
  ];

  let yPos = 92;
  imageSpecs.forEach(spec => {
    const text = slide.insertTextBox('• ' + spec, 25, yPos, 180, 16);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 18;
  });

  // Video Specs - Middle
  const videoTitle = slide.insertTextBox('🎬 Video Ads', 220, 70, 150, 18);
  videoTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.SECONDARY);

  const videoSpecs = [
    'Horizontal: 1920x1080',
    'Square: 1080x1080',
    'Vertical: 1080x1920',
    'Duration: 15-60 วินาที',
    'File: MP4 <1GB'
  ];

  yPos = 92;
  videoSpecs.forEach(spec => {
    const text = slide.insertTextBox('• ' + spec, 220, yPos, 180, 16);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 18;
  });

  // Character Limits - Right
  const charTitle = slide.insertTextBox('📝 ตัวอักษร', 415, 70, 150, 18);
  charTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.ACCENT);

  const charSpecs = [
    'Title: 20 ตัวอักษร',
    'Long Title 1: 20',
    'Long Title 2: 17',
    'Description: 75',
    'CTA: กำหนดไว้'
  ];

  yPos = 92;
  charSpecs.forEach(spec => {
    const text = slide.insertTextBox('• ' + spec, 415, yPos, 170, 16);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 18;
  });

  // Best Practices Box
  const practicesBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 195, W - 50, 50);
  practicesBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  practicesBox.getBorder().setTransparent();

  const practicesText = slide.insertTextBox(
    '✅ Best Practices: ภาพชัดบนมือถือ | ภาษาไทยเข้าใจง่าย | CTA ชัดเจน | ใส่คำเตือนความเสี่ยง | A/B Test',
    35, 205, W - 70, 30
  );
  practicesText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.BLACK);

  // AI Content Note
  const aiBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 255, W - 50, 40);
  aiBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  aiBox.getBorder().setTransparent();

  const aiNote = slide.insertTextBox(
    '🤖 สำหรับ Ad Copy สไตล์ AI: เน้นคำว่า "AI เพื่อนคู่คิด" "เรียนรู้กับ AI" "Aslan Agent" ไม่ชี้นำการลงทุน',
    35, 262, W - 70, 28
  );
  aiNote.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 10);
}

// ==================== SLIDE 11: AD EXAMPLES ====================
function createSlide11_AdExamples(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ตัวอย่าง Ad Copy - พื้นฐานการลงทุน', '📱');

  // Ad 1 - Educational
  createAdMockupCompact(slide, 25, 70,
    'เรียนลงทุนฟรี',
    '🎯 อยากเริ่มลงทุน?\n\n' +
    'Aslan ช่วยคุณได้!\n' +
    '✅ เรียนรู้ฟรี\n' +
    '✅ เข้าใจง่าย\n' +
    '✅ ไม่ชี้นำ',
    'แอดเพื่อน',
    CONFIG.COLORS.PRIMARY
  );

  // Ad 2 - Problem/Solution
  createAdMockupCompact(slide, 195, 70,
    'เงินหมดก่อนสิ้นเดือน?',
    '💸 ทำงานหนัก\nไม่มีเงินเก็บ?\n\n' +
    'เรามีทางออก!\n' +
    '📌 วางแผนการเงิน\n' +
    '📌 เริ่มออมง่ายๆ',
    'แอดเพื่อน',
    CONFIG.COLORS.SECONDARY
  );

  // Ad 3 - Lead Magnet
  createAdMockupCompact(slide, 365, 70,
    'E-Book ฟรี!',
    '📚 คู่มือเริ่มต้นลงทุน\n\n' +
    '✅ 5 วิธีลงทุน 1,000฿\n' +
    '✅ ความผิดพลาด\n' +
    '     ที่ต้องเลี่ยง\n' +
    '✅ พื้นฐานครบ',
    'ดาวน์โหลดฟรี',
    CONFIG.COLORS.ACCENT
  );

  // Ad 4 - Retargeting
  createAdMockupCompact(slide, 535, 70,
    'กลับมานะ!',
    '👋 คุณเคยสนใจ\nAslan Wealth\n\n' +
    '🎁 มีของพิเศษ!\n' +
    '⏰ 48 ชม.\n' +
    '     เท่านั้น',
    'รับสิทธิ์เลย',
    CONFIG.COLORS.WARNING
  );

  // Disclaimer
  const disclaimer = slide.insertTextBox(
    '⚠️ ทุกชิ้นงานต้องมี: "การลงทุนมีความเสี่ยง ผู้ลงทุนควรศึกษาข้อมูลก่อนการตัดสินใจลงทุน"',
    25, 280, W - 50, 20
  );
  disclaimer.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.WARNING);

  addSlideFooter(slide, 11);
}

// ==================== SLIDE 12: AD COPY AI (NEW) ====================
function createSlide12_AdCopyAI(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'Ad Copy - AI Learning + Aslan Agent', '🤖');

  // Ad 1 - AI Friend
  createAdMockupCompact(slide, 25, 70,
    '🤖 AI เพื่อนคู่คิด',
    'ไม่รู้จะถามใคร?\n\n' +
    'Aslan Agent\n' +
    'AI ที่ตอบได้ 24 ชม.\n\n' +
    '✅ ถามได้ทุกเรื่อง\n' +
    '✅ ไม่ชี้นำ',
    'ลองคุยกับ AI',
    CONFIG.COLORS.PRIMARY
  );

  // Ad 2 - Learning Journey
  createAdMockupCompact(slide, 195, 70,
    'เรียนรู้กับ AI',
    '📚 อยากเข้าใจการลงทุน\nแต่ไม่รู้จะเริ่มยังไง?\n\n' +
    'Aslan Agent\n' +
    'สอนแบบเข้าใจง่าย\n' +
    'ไม่มีศัพท์ยาก',
    'เริ่มเรียนฟรี',
    CONFIG.COLORS.ACCENT
  );

  // Ad 3 - Compare/Decide
  createAdMockupCompact(slide, 365, 70,
    'ไม่แน่ใจ เลือกอะไรดี?',
    '🤔 กองทุน หรือ หุ้น?\nออมทรัพย์ หรือ ฝากประจำ?\n\n' +
    'ถาม Aslan Agent\n' +
    'ได้ข้อมูลครบ\n' +
    'ตัดสินใจเอง',
    'ถาม AI เลย',
    CONFIG.COLORS.SECONDARY
  );

  // Ad 4 - Safe Learning
  createAdMockupCompact(slide, 535, 70,
    'เรียนรู้อย่างปลอดภัย',
    '🛡️ ไม่มีการชี้นำ\nไม่มีการขายคอร์ส\n\n' +
    'แค่ให้ความรู้\n' +
    'คุณตัดสินใจเอง\n\n' +
    '100% FREE',
    'แอดเพื่อน',
    CONFIG.COLORS.WARNING
  );

  // Key Messages Box
  const keyBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 275, W - 50, 25);
  keyBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  keyBox.getBorder().setTransparent();

  const keyText = slide.insertTextBox(
    '💡 Key Message: "AI เพื่อนคู่คิด ไม่ชี้นำ ให้ความรู้พื้นฐาน ตัดสินใจเอง"',
    35, 279, W - 70, 18
  );
  keyText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 12);
}

// ==================== SLIDE 13: PLACEMENTS ====================
function createSlide13_Placements(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ตำแหน่งโฆษณา (Placements)', '📍');

  // Table Header
  const headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 25, 70, W - 50, 25);
  headerBg.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  headerBg.getBorder().setTransparent();

  const header = slide.insertTextBox('ตำแหน่ง              การเข้าถึง       เหมาะกับ                  Priority', 35, 74, W - 70, 18);
  header.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  const placements = [
    { name: 'Smart Channel', reach: 'สูงสุด', best: 'เพิ่มเพื่อน', priority: '⭐⭐⭐' },
    { name: 'LINE Home Tab', reach: 'สูงมาก', best: 'Brand Awareness', priority: '⭐⭐⭐' },
    { name: 'LINE Today', reach: 'สูง', best: 'กลุ่มอ่านข่าวการเงิน', priority: '⭐⭐⭐' },
    { name: 'LINE Wallet', reach: 'ปานกลาง', best: 'กลุ่มการเงิน + AI', priority: '⭐⭐⭐' },
    { name: 'LINE VOOM', reach: 'สูง', best: 'Video Content', priority: '⭐⭐' }
  ];

  let yPos = 97;
  placements.forEach((p, i) => {
    const bgColor = i % 2 === 0 ? CONFIG.COLORS.LIGHT_GRAY : CONFIG.COLORS.WHITE;
    const rowBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 25, yPos, W - 50, 22);
    rowBg.getFill().setSolidFill(bgColor);
    rowBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.GRAY);

    const rowText = slide.insertTextBox(
      p.name + '           ' + p.reach + '          ' + p.best + '          ' + p.priority,
      35, yPos + 4, W - 70, 16
    );
    rowText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);

    yPos += 24;
  });

  // Recommendation Box
  const recBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 225, W - 50, 70);
  recBox.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  recBox.getBorder().setTransparent();

  const recText = slide.insertTextBox(
    '💡 คำแนะนำสำหรับแต่ละแคมเปญ\n\n' +
    'เพิ่มเพื่อน → Smart Channel + Home Tab  |  Video → LINE VOOM + Today\n' +
    'Traffic → LINE Today + Wallet  |  Retargeting → ทุกตำแหน่ง (เตือนความจำ)',
    35, 232, W - 70, 55
  );
  recText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 13);
}

// ==================== SLIDE 14: CONTENT CALENDAR (4 WEEKS) ====================
function createSlide14_ContentCalendar(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ปฏิทินคอนเทนต์ 4 สัปดาห์', '📅');

  const weeks = [
    { week: 'Week 1', phase: 'Launch', activities: 'เปิด Campaign 1&2\nA/B Test Creative', color: CONFIG.COLORS.PRIMARY },
    { week: 'Week 2', phase: 'Optimize', activities: 'เพิ่ม Campaign 3\nปรับ Targeting', color: CONFIG.COLORS.SECONDARY },
    { week: 'Week 3', phase: 'Scale', activities: 'เพิ่มงบ Winners\nเปิด Retargeting', color: CONFIG.COLORS.ACCENT },
    { week: 'Week 4', phase: 'Final Push', activities: 'Push All Campaigns\nReport & Learn', color: CONFIG.COLORS.WARNING }
  ];

  let xPos = 20;
  weeks.forEach((w) => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, xPos, 70, 165, 95);
    box.getFill().setSolidFill(w.color);
    box.getBorder().setTransparent();

    const weekText = slide.insertTextBox(w.week, xPos, 78, 165, 20);
    weekText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    weekText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const phaseText = slide.insertTextBox(w.phase, xPos, 98, 165, 18);
    phaseText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(10)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    phaseText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const actText = slide.insertTextBox(w.activities, xPos + 10, 118, 145, 42);
    actText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(8)
      .setForegroundColor(CONFIG.COLORS.LIGHT_GRAY);
    actText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    xPos += 175;
  });

  // Content Themes
  const themesTitle = slide.insertTextBox('📝 ธีมคอนเทนต์รายสัปดาห์', 25, 175, 350, 18);
  themesTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const themes = [
    'Week 1: "ทำไมต้องลงทุน?" - พื้นฐาน, เงินเฟ้อ, ดอกเบี้ยทบต้น',
    'Week 2: "เริ่มต้นลงทุน" - หุ้น, กองทุน, เงินเริ่มต้นเท่าไหร่',
    'Week 3: "AI เพื่อนคู่คิด" - Aslan Agent, ถามได้ 24 ชม.',
    'Week 4: "สรุปความรู้" - Checklist, ทบทวน, Call to Action'
  ];

  let yPos = 195;
  themes.forEach(theme => {
    const text = slide.insertTextBox('• ' + theme, 25, yPos, W - 50, 18);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 20;
  });

  // Schedule Note
  const scheduleBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 280, W - 50, 20);
  scheduleBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  scheduleBox.getBorder().setTransparent();

  const scheduleText = slide.insertTextBox(
    '📱 LINE OA Post: จันทร์ 10:00 (บทความ) | พุธ 12:00 (Tips) | ศุกร์ 18:00 (สรุป)',
    35, 283, W - 70, 14
  );
  scheduleText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.DARK_GRAY);

  addSlideFooter(slide, 14);
}

// ==================== SLIDE 15: KPIs ====================
function createSlide15_KPIs(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ตัวชี้วัด KPIs', '📊');

  // Primary KPIs - Left
  const primaryTitle = slide.insertTextBox('Primary KPIs', 25, 70, 150, 18);
  primaryTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const primaryKPIs = [
    { metric: 'เพื่อนใหม่ LINE OA', target: '1,500-2,500 คน' },
    { metric: 'Cost per Friend', target: '< 40 บาท' },
    { metric: 'ROAS', target: '> 2x' },
    { metric: 'Conversion Rate', target: '3-5%' }
  ];

  let yPos = 92;
  primaryKPIs.forEach(kpi => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, yPos, 200, 28);
    box.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    box.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.PRIMARY);

    const text = slide.insertTextBox(kpi.metric + ': ' + kpi.target, 32, yPos + 6, 190, 18);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);

    yPos += 32;
  });

  // Secondary KPIs - Right
  const secondaryTitle = slide.insertTextBox('Secondary KPIs', 250, 70, 150, 18);
  secondaryTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.SECONDARY);

  const secondaryKPIs = [
    { metric: 'Impressions', target: '3-5M' },
    { metric: 'CTR', target: '> 1%' },
    { metric: 'Video Views', target: '50,000+' },
    { metric: 'Website Clicks', target: '8,000+' }
  ];

  yPos = 92;
  secondaryKPIs.forEach(kpi => {
    const text = slide.insertTextBox('• ' + kpi.metric + ': ' + kpi.target, 250, yPos, 180, 18);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(10)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 24;
  });

  // Benchmark Table
  const benchBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 230, W - 50, 65);
  benchBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  benchBox.getBorder().setTransparent();

  const benchTitle = slide.insertTextBox('📈 Performance Benchmarks', 35, 238, 250, 16);
  benchTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  const benchText = slide.insertTextBox(
    'Metric          ต่ำ 🔴          ปานกลาง 🟡          ดี 🟢          ดีมาก 🌟\n' +
    'CPF              >50฿            35-50฿               25-35฿         <25฿\n' +
    'CTR              <0.5%          0.5-1%               1-2%            >2%',
    35, 258, W - 70, 32
  );
  benchText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // AI KPI Sidebar
  const aiKPI = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 450, 70, 220, 150);
  aiKPI.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  aiKPI.getBorder().setTransparent();

  const aiKPIText = slide.insertTextBox(
    '🤖 AI Content KPIs\n\n' +
    '• Engagement Rate: >3%\n' +
    '• Chat Opens: >500\n' +
    '• AI Questions: >200\n' +
    '• Positive Sentiment: >80%',
    460, 78, 200, 135
  );
  aiKPIText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 15);
}

// ==================== SLIDE 16: TIMELINE ====================
function createSlide16_Timeline(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ไทม์ไลน์การดำเนินงาน', '⏱️');

  // Timeline Bar
  const barBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, 150, 600, 6);
  barBg.getFill().setSolidFill(CONFIG.COLORS.GRAY);
  barBg.getBorder().setTransparent();

  // Timeline Points
  const points = [
    { day: 'Day 1-2', task: 'Setup\nLINE OA', x: 80 },
    { day: 'Day 3-4', task: 'Creative\nDesign', x: 200 },
    { day: 'Day 5', task: 'LINE Tag\nSetup', x: 320 },
    { day: 'Day 6-7', task: 'Campaign\nConfig', x: 440 },
    { day: 'Day 8', task: 'Launch!\n🚀', x: 560 }
  ];

  points.forEach((p) => {
    // Circle
    const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, p.x - 8, 146, 16, 16);
    circle.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
    circle.getBorder().setTransparent();

    // Day Label
    const day = slide.insertTextBox(p.day, p.x - 35, 125, 70, 16);
    day.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.PRIMARY);
    day.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Task
    const task = slide.insertTextBox(p.task, p.x - 40, 170, 80, 40);
    task.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(8)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    task.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  });

  // Checklist Box
  const checkBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 220, W - 50, 75);
  checkBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  checkBox.getBorder().setTransparent();

  const checkTitle = slide.insertTextBox('📋 สิ่งที่ต้องเตรียมก่อนเริ่ม', 35, 228, 300, 16);
  checkTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const checkText = slide.insertTextBox(
    '□ สร้าง/ยืนยัน LINE OA          □ ซื้อ Premium ID (@aslanwealth)          □ ออกแบบ Rich Menu\n' +
    '□ ตั้งค่า Chatbot + AI               □ ติดตั้ง LINE Tag บน aslanwealth.com    □ สร้าง Creative 5-10 แบบ',
    35, 248, W - 70, 40
  );
  checkText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.BLACK);

  addSlideFooter(slide, 16);
}

// ==================== SLIDE 17: RICH MENU ====================
function createSlide17_RichMenu(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'Rich Menu Design', '📱');

  // Rich Menu Mock (6 cells)
  const menuBg = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 70, 260, 160);
  menuBg.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  menuBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  menuBg.getBorder().setWeight(2);

  const menuItems = [
    { icon: '📰', text: 'ข่าวลงทุน', row: 0, col: 0 },
    { icon: '📚', text: 'คอร์สเรียน', row: 0, col: 1 },
    { icon: '🤖', text: 'Aslan AI', row: 0, col: 2 },
    { icon: '📖', text: 'บทความฟรี', row: 1, col: 0 },
    { icon: '🎁', text: 'โปรโมชั่น', row: 1, col: 1 },
    { icon: 'ℹ️', text: 'เกี่ยวกับเรา', row: 1, col: 2 }
  ];

  menuItems.forEach(item => {
    const x = 30 + (item.col * 85);
    const y = 75 + (item.row * 78);

    const cell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, 80, 73);
    cell.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    cell.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.GRAY);

    const iconText = slide.insertTextBox(item.icon, x, y + 12, 80, 25);
    iconText.getText().getTextStyle().setFontSize(18);
    iconText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const labelText = slide.insertTextBox(item.text, x, y + 42, 80, 20);
    labelText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    labelText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  });

  // Specs
  const specsTitle = slide.insertTextBox('📐 Specifications', 310, 70, 180, 18);
  specsTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.PRIMARY);

  const specs = [
    'ขนาด: 2500 x 1686 px',
    'Layout: 6 ช่อง (3x2)',
    'File: JPEG/PNG <1MB',
    'Actions: URL + Message',
    '🤖 Aslan AI Button'
  ];

  let yPos = 92;
  specs.forEach(spec => {
    const text = slide.insertTextBox('• ' + spec, 310, yPos, 180, 16);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 18;
  });

  // Welcome Message Box
  const welcomeBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 310, 185, 360, 60);
  welcomeBox.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  welcomeBox.getBorder().setTransparent();

  const welcomeText = slide.insertTextBox(
    '💬 ข้อความต้อนรับ\n' +
    'สวัสดี! ยินดีต้อนรับสู่ Aslan Wealth 🙏\n' +
    'เรียนรู้การลงทุนกับ AI ได้เลย → กดเมนู "Aslan AI"',
    320, 192, 340, 48
  );
  welcomeText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // AI Highlight
  const aiHighlight = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 245, 645, 50);
  aiHighlight.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  aiHighlight.getBorder().setTransparent();

  const aiText = slide.insertTextBox(
    '🤖 Aslan Agent Integration: ปุ่ม "Aslan AI" จะพาผู้ใช้ไปคุยกับ AI เพื่อนคู่คิดโดยตรง ถามได้ทุกเรื่องการเงิน 24 ชม.',
    35, 255, 625, 35
  );
  aiText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  addSlideFooter(slide, 17);
}

// ==================== SLIDE 18: COMPLIANCE ====================
function createSlide18_Compliance(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ข้อกำหนดและกฎหมาย', '⚖️');

  // Must Have - Left
  const mustTitle = slide.insertTextBox('✅ สิ่งที่ต้องมี', 25, 70, 150, 18);
  mustTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.ACCENT);

  const mustItems = [
    'คำเตือนความเสี่ยงทุกชิ้น',
    'ข้อมูลบริษัทที่ชัดเจน',
    'ช่องทางติดต่อจริง',
    'ระบุว่า AI ไม่ใช่ที่ปรึกษา',
    'ไม่ชี้นำการลงทุน'
  ];

  let yPos = 92;
  mustItems.forEach(item => {
    const text = slide.insertTextBox('• ' + item, 25, yPos, 180, 16);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 18;
  });

  // Must Not - Right
  const mustNotTitle = slide.insertTextBox('❌ สิ่งที่ห้ามทำ', 220, 70, 150, 18);
  mustNotTitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WARNING);

  const mustNotItems = [
    'รับประกันผลตอบแทน',
    'อ้างผลตอบแทนเกินจริง',
    'ใช้ข้อมูลเท็จ/หลอกลวง',
    'AI แนะนำหุ้นเฉพาะตัว',
    'บังคับซื้อ/กดดัน'
  ];

  yPos = 92;
  mustNotItems.forEach(item => {
    const text = slide.insertTextBox('• ' + item, 220, yPos, 180, 16);
    text.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);
    yPos += 18;
  });

  // Warning Box
  const warnBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 195, W - 50, 45);
  warnBox.getFill().setSolidFill(CONFIG.COLORS.WARNING);
  warnBox.getBorder().setTransparent();

  const warnText = slide.insertTextBox(
    '⚠️ คำเตือนมาตรฐาน (ต้องใส่ทุกชิ้นงาน)\n' +
    '"การลงทุนมีความเสี่ยง ผู้ลงทุนควรศึกษาข้อมูลก่อนการตัดสินใจลงทุน"',
    35, 202, W - 70, 32
  );
  warnText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // AI Disclaimer
  const aiDiscBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 250, W - 50, 45);
  aiDiscBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  aiDiscBox.getBorder().setTransparent();

  const aiDiscText = slide.insertTextBox(
    '🤖 AI Disclaimer (สำหรับ Aslan Agent)\n' +
    '"Aslan Agent เป็นเครื่องมือให้ความรู้เท่านั้น ไม่ใช่ที่ปรึกษาการลงทุน การตัดสินใจลงทุนเป็นความรับผิดชอบของผู้ใช้"',
    35, 257, W - 70, 32
  );
  aiDiscText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // Regulatory
  const regText = slide.insertTextBox(
    '🏛️ หน่วยงาน: ก.ล.ต. | ธปท. | คปภ. | สคบ.',
    25, 300, W - 50, 14
  );
  regText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.GRAY);

  addSlideFooter(slide, 18);
}

// ==================== SLIDE 19: EXPECTED RESULTS ====================
function createSlide19_ExpectedResults(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ผลลัพธ์ที่คาดหวัง (4 สัปดาห์)', '📈');

  // Table Header
  const headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 25, 70, W - 50, 25);
  headerBg.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  headerBg.getBorder().setTransparent();

  const header = slide.insertTextBox('ตัวชี้วัด                          ต่ำ                  กลาง               สูง', 35, 74, W - 70, 18);
  header.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  const results = [
    { metric: 'Impressions', low: '2 ล้าน', mid: '3.5 ล้าน', high: '5 ล้าน' },
    { metric: 'เพื่อนใหม่ LINE OA', low: '1,000 คน', mid: '1,800 คน', high: '2,500 คน' },
    { metric: 'Video Views', low: '30,000', mid: '50,000', high: '70,000' },
    { metric: 'Website Clicks', low: '5,000', mid: '8,000', high: '12,000' },
    { metric: 'Leads', low: '100 คน', mid: '200 คน', high: '350 คน' }
  ];

  let yPos = 97;
  results.forEach((r, i) => {
    const bgColor = i % 2 === 0 ? CONFIG.COLORS.LIGHT_GRAY : CONFIG.COLORS.WHITE;
    const rowBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 25, yPos, W - 50, 22);
    rowBg.getFill().setSolidFill(bgColor);
    rowBg.getBorder().getLineFill().setSolidFill(CONFIG.COLORS.GRAY);

    const rowText = slide.insertTextBox(
      r.metric + '                ' + r.low + '            ' + r.mid + '           ' + r.high,
      35, yPos + 4, W - 70, 16
    );
    rowText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setForegroundColor(CONFIG.COLORS.BLACK);

    yPos += 24;
  });

  // ROI Box
  const roiBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 25, 225, 220, 70);
  roiBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  roiBox.getBorder().setTransparent();

  const roiText = slide.insertTextBox(
    '💰 ROI คาดการณ์\n\n' +
    'งบโฆษณา: ฿100,000\n' +
    'ลูกค้าคาดการณ์: 30-50 คน\n' +
    'ROAS: 2-3x ✅',
    35, 232, 200, 58
  );
  roiText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // AI Engagement Box
  const aiEngBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 260, 225, 220, 70);
  aiEngBox.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  aiEngBox.getBorder().setTransparent();

  const aiEngText = slide.insertTextBox(
    '🤖 AI Engagement\n\n' +
    'AI Conversations: 500+\n' +
    'Questions Asked: 1,500+\n' +
    'Satisfaction: >80%',
    270, 232, 200, 58
  );
  aiEngText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);

  // Note
  const note = slide.insertTextBox(
    '* ตัวเลขประมาณการจาก Benchmark และงบประมาณ อาจแตกต่างตามคุณภาพ Creative และการปรับ Optimize',
    260, 290, 400, 18
  );
  note.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(8)
    .setForegroundColor(CONFIG.COLORS.GRAY);

  addSlideFooter(slide, 19);
}

// ==================== SLIDE 20: NEXT STEPS ====================
function createSlide20_NextSteps(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;

  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addSlideHeader(slide, 'ขั้นตอนถัดไป', '🚀');

  const steps = [
    { num: '1', title: 'อนุมัติแผน', timeline: 'ทันที' },
    { num: '2', title: 'Setup LINE OA + Premium ID', timeline: 'Day 1-2' },
    { num: '3', title: 'Design Creative + Rich Menu', timeline: 'Day 3-4' },
    { num: '4', title: 'LINE Tag + AI Chatbot Setup', timeline: 'Day 4-5' },
    { num: '5', title: 'Campaign Configuration', timeline: 'Day 6-7' },
    { num: '6', title: 'Launch! 🚀', timeline: 'Day 8' }
  ];

  let yPos = 72;
  steps.forEach((step, i) => {
    // Number Circle
    const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 25, yPos, 28, 28);
    circle.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
    circle.getBorder().setTransparent();

    const numText = slide.insertTextBox(step.num, 25, yPos + 5, 28, 18);
    numText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    numText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Title
    const titleText = slide.insertTextBox(step.title, 60, yPos + 2, 250, 18);
    titleText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(11)
      .setForegroundColor(CONFIG.COLORS.BLACK);

    // Timeline Badge
    const badge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 320, yPos + 2, 70, 22);
    badge.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
    badge.getBorder().setTransparent();

    const badgeText = slide.insertTextBox(step.timeline, 320, yPos + 5, 70, 16);
    badgeText.getText().getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(9)
      .setBold(true)
      .setForegroundColor(CONFIG.COLORS.WHITE);
    badgeText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Connector
    if (i < steps.length - 1) {
      const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 38, yPos + 30, 3, 12);
      line.getFill().setSolidFill(CONFIG.COLORS.GRAY);
      line.getBorder().setTransparent();
    }

    yPos += 38;
  });

  // CTA Box
  const ctaBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 420, 70, 270, 135);
  ctaBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  ctaBox.getBorder().setTransparent();

  const ctaText = slide.insertTextBox(
    '✅ พร้อมเริ่มต้น!\n\n' +
    'กรุณาอนุมัติเพื่อ\n' +
    'ดำเนินการตามแผน\n\n' +
    '📅 4 สัปดาห์\n' +
    '💰 ฿100,000',
    430, 78, 250, 120
  );
  ctaText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  ctaText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // AI Ready Badge
  const aiBadge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 420, 215, 270, 35);
  aiBadge.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  aiBadge.getBorder().setTransparent();

  const aiBadgeText = slide.insertTextBox('🤖 พร้อมโปรโมท Aslan Agent AI!', 430, 222, 250, 22);
  aiBadgeText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  aiBadgeText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  addSlideFooter(slide, 20);
}

// ==================== SLIDE 21: THANK YOU ====================
function createSlide21_ThankYou(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;
  const H = CONFIG.SLIDE.HEIGHT;

  slide.getBackground().setSolidFill(CONFIG.COLORS.PRIMARY);

  // Thank You
  const thankYou = slide.insertTextBox('ขอบคุณครับ/ค่ะ', 30, 100, W - 60, 55);
  thankYou.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(40)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  thankYou.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Subtitle
  const subtitle = slide.insertTextBox('พร้อมตอบคำถามและรับฟังความคิดเห็น', 30, 155, W - 60, 30);
  subtitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(16)
    .setForegroundColor(CONFIG.COLORS.ACCENT);
  subtitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Contact Box
  const contactBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 150, 200, 420, 90);
  contactBox.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  contactBox.getBorder().setTransparent();

  const contactText = slide.insertTextBox(
    '📞 ติดต่อ Aslan Wealth Channel\n\n' +
    '🌐 aslanwealth.com  |  📱 @aslanwealth  |  📧 contact@aslanwealth.com',
    160, 212, 400, 68
  );
  contactText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  contactText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // AI Badge
  const aiBadge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 250, 300, 220, 30);
  aiBadge.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  aiBadge.getBorder().setTransparent();

  const aiBadgeText = slide.insertTextBox('🤖 Powered by Aslan Agent AI', 260, 306, 200, 20);
  aiBadgeText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  aiBadgeText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Footer
  const footer = slide.insertTextBox('LINE Ads Campaign | งบประมาณ ฿100,000 | 4 สัปดาห์ | มกราคม 2025', 30, 370, W - 60, 18);
  footer.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor(CONFIG.COLORS.GRAY);
  footer.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ==================== HELPER FUNCTIONS ====================

function addSlideHeader(slide, title, icon) {
  const W = CONFIG.SLIDE.WIDTH;

  const headerBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, W, CONFIG.SLIDE.HEADER_HEIGHT);
  headerBar.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  headerBar.getBorder().setTransparent();

  const titleText = slide.insertTextBox((icon ? icon + ' ' : '') + title, 20, 15, W - 40, 30);
  titleText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
}

function addSlideFooter(slide, pageNum) {
  const W = CONFIG.SLIDE.WIDTH;
  const H = CONFIG.SLIDE.HEIGHT;

  // Brand
  const brand = slide.insertTextBox('Aslan Wealth Channel', 20, H - 25, 150, 15);
  brand.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(8)
    .setForegroundColor(CONFIG.COLORS.GRAY);

  // Page Number
  const page = slide.insertTextBox(pageNum + ' / 21', W - 60, H - 25, 40, 15);
  page.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(8)
    .setForegroundColor(CONFIG.COLORS.GRAY);
  page.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
}

function createStatBox(slide, x, y, value, label, color, width, height) {
  width = width || 100;
  height = height || 70;

  const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, width, height);
  box.getFill().setSolidFill(color);
  box.getBorder().setTransparent();

  const valueText = slide.insertTextBox(value, x, y + 10, width, 28);
  valueText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  valueText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const labelText = slide.insertTextBox(label, x, y + 38, width, 20);
  labelText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  labelText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function createAdMockupCompact(slide, x, y, headline, body, cta, color) {
  // Ad Frame
  const frame = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, 160, 195);
  frame.getFill().setSolidFill(CONFIG.COLORS.WHITE);
  frame.getBorder().getLineFill().setSolidFill(color);
  frame.getBorder().setWeight(2);

  // Image Placeholder
  const imgPlaceholder = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x + 5, y + 5, 150, 55);
  imgPlaceholder.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  imgPlaceholder.getBorder().setTransparent();

  const imgText = slide.insertTextBox('📷', x + 5, y + 18, 150, 30);
  imgText.getText().getTextStyle().setFontSize(22);
  imgText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Headline
  const headlineText = slide.insertTextBox(headline, x + 8, y + 62, 144, 22);
  headlineText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(true)
    .setForegroundColor(color);

  // Body
  const bodyText = slide.insertTextBox(body, x + 8, y + 85, 144, 75);
  bodyText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(8)
    .setForegroundColor(CONFIG.COLORS.BLACK);

  // CTA Button
  const ctaBtn = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 20, y + 163, 120, 24);
  ctaBtn.getFill().setSolidFill(color);
  ctaBtn.getBorder().setTransparent();

  const ctaText = slide.insertTextBox(cta, x + 20, y + 167, 120, 18);
  ctaText.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(true)
    .setForegroundColor(CONFIG.COLORS.WHITE);
  ctaText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ==================== RUN ====================
// เรียกใช้ function นี้เพื่อสร้าง Presentation
// createLINEAdsCampaignPresentation();
