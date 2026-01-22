/**
 * YouTube Ads Campaign Presentation Generator
 * Aslan Wealth Channel
 * งบประมาณ: 100,000 บาท | ระยะเวลา: 4 สัปดาห์
 *
 * วิธีใช้งาน:
 * 1. เปิด Google Apps Script (script.google.com)
 * 2. สร้างโปรเจกต์ใหม่
 * 3. วาง Code นี้ทั้งหมด
 * 4. รัน function: createYouTubeAdsCampaignPresentation()
 * 5. อนุญาต permissions ที่ขอ
 * 6. Presentation จะสร้างใน Google Drive อัตโนมัติ
 *
 * Slide Ratio: 16:9 (720x405 points)
 */

// ==================== CONFIGURATION ====================
const CONFIG = {
  BRAND_NAME: 'Aslan Wealth Channel',
  WEBSITE: 'https://aslanwealth.com',
  YOUTUBE: '@AslanWealthChannel',
  EMAIL: 'contact@aslanwealth.com',
  BUDGET: '100,000',
  DURATION: '4 สัปดาห์',
  CURRENCY: 'THB',

  SLIDE: {
    WIDTH: 720,
    HEIGHT: 405,
    MARGIN: 25,
    HEADER_HEIGHT: 55
  },

  COLORS: {
    PRIMARY: '#FF0000',       // YouTube Red
    SECONDARY: '#282828',     // YouTube Dark
    ACCENT: '#065FD4',        // YouTube Blue
    SUCCESS: '#2E7D32',       // Green
    WARNING: '#F9A825',       // Yellow
    WHITE: '#FFFFFF',
    BLACK: '#0F0F0F',
    LIGHT_GRAY: '#F9F9F9',
    GRAY: '#606060',
    DARK_GRAY: '#3D3D3D'
  }
};

// ==================== IMAGES ====================
const IMAGES = {
  YOUTUBE_LOGO: 'https://images.unsplash.com/photo-1611162616475-46b635cb6868?w=800&q=80',
  VIDEO_CREATOR: 'https://images.unsplash.com/photo-1598550476439-6847785fcea6?w=800&q=80',
  ANALYTICS: 'https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=800&q=80',
  INVESTMENT: 'https://images.unsplash.com/photo-1611974789855-9c2a0a7236a3?w=800&q=80',
  SUCCESS: 'https://images.unsplash.com/photo-1553729459-efe14ef6055d?w=800&q=80',
  MOBILE: 'https://images.unsplash.com/photo-1512941937669-90a1b58e7e9c?w=800&q=80',
  TEAM: 'https://images.unsplash.com/photo-1522071820081-009f0129c71c?w=800&q=80'
};

// ==================== MAIN FUNCTION ====================
function createYouTubeAdsCampaignPresentation() {
  const presentation = SlidesApp.create('YouTube Ads Campaign - Aslan Wealth Channel');
  const slides = presentation.getSlides();

  if (slides.length > 0) {
    slides[0].remove();
  }

  Logger.log('เริ่มสร้าง YouTube Ads Presentation...');

  // สร้างทุก Slides
  createSlide01_Title(presentation);
  createSlide02_Agenda(presentation);
  createSlide03_WhyYouTube(presentation);
  createSlide04_Objectives(presentation);
  createSlide05_CampaignStructure(presentation);
  createSlide06_BudgetAllocation(presentation);
  createSlide07_AdFormats(presentation);
  createSlide08_TargetAudience(presentation);
  createSlide09_ContentCalendar(presentation);
  createSlide10_CreativeSpecs(presentation);
  createSlide11_GrowthHacks(presentation);
  createSlide12_KPIs(presentation);
  createSlide13_Timeline(presentation);
  createSlide14_ExpectedResults(presentation);
  createSlide15_ThankYou(presentation);

  Logger.log('สร้าง Presentation สำเร็จ!');
  Logger.log('URL: ' + presentation.getUrl());

  return presentation.getUrl();
}

// ==================== SLIDE 1: TITLE ====================
function createSlide01_Title(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;
  const H = CONFIG.SLIDE.HEIGHT;

  slide.getBackground().setSolidFill(CONFIG.COLORS.SECONDARY);

  // YouTube Red accent bar
  const redBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, W, 8);
  redBar.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  redBar.getBorder().setTransparent();

  // Play button icon (triangle shape)
  const playBg = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, W - 100, 30, 60, 60);
  playBg.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  playBg.getBorder().setTransparent();

  // Logo
  const logoBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 30, 140, 40);
  logoBox.getFill().setSolidFill(CONFIG.COLORS.WHITE);
  logoBox.getBorder().setTransparent();
  const logoText = logoBox.getText();
  logoText.setText('ASLAN');
  logoText.getTextStyle().setFontFamily('Arial').setFontSize(20).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);
  logoText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Main Title
  const title = slide.insertTextBox('YouTube Ads Campaign', 30, 120, 500, 60);
  title.getText().getTextStyle().setFontFamily('Arial').setFontSize(38).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  // Subtitle
  const subtitle = slide.insertTextBox('แผนโฆษณา YouTube 4 สัปดาห์ | Aslan Wealth Channel', 30, 180, 500, 30);
  subtitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setForegroundColor(CONFIG.COLORS.GRAY);

  // Budget Box
  const budgetBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 230, 150, 45);
  budgetBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  budgetBox.getBorder().setTransparent();
  const budgetText = slide.insertTextBox('฿100,000', 30, 240, 150, 30);
  budgetText.getText().getTextStyle().setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
  budgetText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Duration Box
  const durBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 190, 230, 100, 45);
  durBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  durBox.getBorder().setTransparent();
  const durText = slide.insertTextBox('4 สัปดาห์', 190, 240, 100, 30);
  durText.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
  durText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Stats
  const stats = slide.insertTextBox('🎬 2.6B+ Monthly Users | 🇹🇭 40M+ Thai Users', 30, 350, 400, 25);
  stats.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setForegroundColor(CONFIG.COLORS.GRAY);

  // Date
  const date = slide.insertTextBox('มกราคม 2025', W - 120, H - 30, 100, 20);
  date.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
}

// ==================== SLIDE 2: AGENDA ====================
function createSlide02_Agenda(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Agenda', 'สิ่งที่จะนำเสนอ');

  const agendaItems = [
    ['1', 'ทำไมต้อง YouTube Ads', 'ข้อมูลตลาดและโอกาส'],
    ['2', 'วัตถุประสงค์แคมเปญ', 'เป้าหมายที่ต้องการ'],
    ['3', 'โครงสร้างแคมเปญ', '4 แคมเปญหลัก'],
    ['4', 'งบประมาณ', 'การจัดสรร 100,000 บาท'],
    ['5', 'Ad Formats', 'รูปแบบโฆษณาทั้งหมด'],
    ['6', 'กลุ่มเป้าหมาย', '4 กลุ่มหลัก'],
    ['7', 'Content Calendar', 'แผนคอนเทนต์ 4 สัปดาห์'],
    ['8', 'Growth Hacks', 'เทคนิคเพิ่ม Subscribers'],
    ['9', 'KPIs & Timeline', 'ตัวชี้วัดและไทม์ไลน์']
  ];

  let y = 75;
  agendaItems.forEach((item, i) => {
    const numBox = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 40, y, 25, 25);
    numBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
    numBox.getBorder().setTransparent();
    const numText = slide.insertTextBox(item[0], 40, y + 3, 25, 20);
    numText.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    numText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const titleText = slide.insertTextBox(item[1], 75, y, 200, 20);
    titleText.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const descText = slide.insertTextBox(item[2], 280, y, 200, 20);
    descText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 35;
  });
}

// ==================== SLIDE 3: WHY YOUTUBE ====================
function createSlide03_WhyYouTube(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'ทำไมต้อง YouTube Ads?', 'โอกาสในตลาดไทย');

  const stats = [
    ['2.6B+', 'ผู้ใช้ต่อเดือนทั่วโลก', CONFIG.COLORS.PRIMARY],
    ['40M+', 'ผู้ใช้ในไทย (62%)', CONFIG.COLORS.ACCENT],
    ['1B+', 'ชั่วโมงวิดีโอ/วัน', CONFIG.COLORS.SUCCESS],
    ['#2', 'Search Engine โลก', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  stats.forEach(stat => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 150, 80);
    box.getFill().setSolidFill(stat[2]);
    box.getBorder().setTransparent();

    const numText = slide.insertTextBox(stat[0], x, 90, 150, 35);
    numText.getText().getTextStyle().setFontFamily('Arial').setFontSize(24).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    numText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const descText = slide.insertTextBox(stat[1], x, 125, 150, 25);
    descText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);
    descText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Benefits
  const benefits = [
    '✅ เข้าถึงกลุ่มเป้าหมายที่สนใจการเงิน/การลงทุน',
    '✅ วิดีโอสร้างความน่าเชื่อถือมากกว่าภาพนิ่ง',
    '✅ Remarketing ไปยังคนที่ดูวิดีโอได้',
    '✅ YouTube Shorts เข้าถึงคนรุ่นใหม่',
    '✅ ราคาถูกกว่า TV Ads 10 เท่า'
  ];

  let y = 180;
  benefits.forEach(benefit => {
    const text = slide.insertTextBox(benefit, 40, y, 400, 25);
    text.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    y += 30;
  });

  // Insight box
  const insightBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 460, 180, 230, 120);
  insightBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  insightBox.getBorder().setTransparent();

  const insightTitle = slide.insertTextBox('💡 Finance Content', 460, 190, 230, 25);
  insightTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);
  insightTitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const insightText = slide.insertTextBox('คนไทยค้นหา:\n• "ลงทุนอะไรดี" 50K/เดือน\n• "หุ้นเด่น" 30K/เดือน\n• "วิธีออมเงิน" 40K/เดือน', 470, 220, 210, 70);
  insightText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
}

// ==================== SLIDE 4: OBJECTIVES ====================
function createSlide04_Objectives(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'วัตถุประสงค์แคมเปญ', 'Campaign Objectives');

  const objectives = [
    ['🎯', 'Subscriber Growth', '+1,500 Subscribers', '35%', CONFIG.COLORS.PRIMARY],
    ['📢', 'Brand Awareness', '2M+ Impressions', '25%', CONFIG.COLORS.ACCENT],
    ['🌐', 'Website Traffic', '8,000+ Clicks', '25%', CONFIG.COLORS.SUCCESS],
    ['🔄', 'Remarketing', '3-5% Conversion', '15%', CONFIG.COLORS.WARNING]
  ];

  let y = 85;
  objectives.forEach(obj => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, y, 640, 65);
    card.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    card.getBorder().setTransparent();

    const iconBox = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 50, y + 12, 40, 40);
    iconBox.getFill().setSolidFill(obj[4]);
    iconBox.getBorder().setTransparent();

    const icon = slide.insertTextBox(obj[0], 50, y + 18, 40, 30);
    icon.getText().getTextStyle().setFontFamily('Arial').setFontSize(18);
    icon.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const title = slide.insertTextBox(obj[1], 100, y + 10, 180, 25);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const target = slide.insertTextBox(obj[2], 100, y + 35, 180, 20);
    target.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.GRAY);

    const budgetBar = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 520, y + 15, 80, 35);
    budgetBar.getFill().setSolidFill(obj[4]);
    budgetBar.getBorder().setTransparent();

    const budgetText = slide.insertTextBox(obj[3], 520, y + 22, 80, 25);
    budgetText.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    budgetText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    y += 75;
  });
}

// ==================== SLIDE 5: CAMPAIGN STRUCTURE ====================
function createSlide05_CampaignStructure(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'โครงสร้างแคมเปญ', '4 แคมเปญหลัก');

  const campaigns = [
    ['Subscriber Growth', '35,000 บาท', 'In-feed + Shorts Ads', CONFIG.COLORS.PRIMARY],
    ['Brand Awareness', '25,000 บาท', 'Skippable + Bumper', CONFIG.COLORS.ACCENT],
    ['Website Traffic', '25,000 บาท', 'In-feed + In-stream', CONFIG.COLORS.SUCCESS],
    ['Remarketing', '15,000 บาท', 'In-stream + Bumper', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  campaigns.forEach(camp => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 180);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().setWeight(2).getLineFill().setSolidFill(camp[3]);

    const topBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, 80, 155, 8);
    topBar.getFill().setSolidFill(camp[3]);
    topBar.getBorder().setTransparent();

    const title = slide.insertTextBox(camp[0], x + 10, 100, 135, 40);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const budget = slide.insertTextBox(camp[1], x + 10, 145, 135, 25);
    budget.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(camp[3]);

    const format = slide.insertTextBox(camp[2], x + 10, 175, 135, 40);
    format.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);

    x += 165;
  });

  // Weekly breakdown
  const weeklyTitle = slide.insertTextBox('📅 งบประมาณรายสัปดาห์', 40, 280, 200, 25);
  weeklyTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const weeks = [
    ['Week 1', '20,000 บาท', 'Testing'],
    ['Week 2', '25,000 บาท', 'Optimize'],
    ['Week 3', '30,000 บาท', 'Scale'],
    ['Week 4', '25,000 บาท', 'Maximize']
  ];

  let wx = 40;
  weeks.forEach(week => {
    const weekBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, wx, 310, 155, 55);
    weekBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    weekBox.getBorder().setTransparent();

    const weekTitle = slide.insertTextBox(week[0] + ': ' + week[1], wx + 10, 318, 135, 20);
    weekTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const weekDesc = slide.insertTextBox(week[2], wx + 10, 338, 135, 20);
    weekDesc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    wx += 165;
  });
}

// ==================== SLIDE 6: BUDGET ALLOCATION ====================
function createSlide06_BudgetAllocation(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'การจัดสรรงบประมาณ', 'Budget: ฿100,000');

  // Pie chart representation
  const pieData = [
    ['Subscriber Growth', '35%', '35,000', CONFIG.COLORS.PRIMARY, 0],
    ['Brand Awareness', '25%', '25,000', CONFIG.COLORS.ACCENT, 126],
    ['Website Traffic', '25%', '25,000', CONFIG.COLORS.SUCCESS, 216],
    ['Remarketing', '15%', '15,000', CONFIG.COLORS.WARNING, 306]
  ];

  // Draw pie segments as colored boxes (simplified)
  let y = 90;
  pieData.forEach(item => {
    const colorBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, y, 30, 30);
    colorBox.getFill().setSolidFill(item[3]);
    colorBox.getBorder().setTransparent();

    const label = slide.insertTextBox(item[0], 90, y + 5, 150, 25);
    label.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const percent = slide.insertTextBox(item[1], 250, y + 5, 50, 25);
    percent.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.GRAY);

    const amount = slide.insertTextBox('฿' + item[2], 310, y + 5, 80, 25);
    amount.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(item[3]);

    y += 40;
  });

  // Cost benchmarks
  const benchTitle = slide.insertTextBox('📊 Cost Benchmarks (ไทย)', 420, 80, 250, 25);
  benchTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const benchmarks = [
    ['CPV (Skippable)', '฿0.20-0.80'],
    ['CPV (Shorts)', '฿0.05-0.30'],
    ['CPM (Bumper)', '฿40-100'],
    ['CPC', '฿2-8'],
    ['Cost/Subscriber', '฿23-35']
  ];

  y = 110;
  benchmarks.forEach(bench => {
    const row = slide.insertTextBox(bench[0] + ': ' + bench[1], 420, y, 250, 20);
    row.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    y += 25;
  });

  // ROI Box
  const roiBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 420, 250, 250, 100);
  roiBox.getFill().setSolidFill(CONFIG.COLORS.SUCCESS);
  roiBox.getBorder().setTransparent();

  const roiTitle = slide.insertTextBox('💰 Expected ROI', 430, 260, 230, 25);
  roiTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  const roiText = slide.insertTextBox('• Moderate: +25% ROI\n• Optimistic: +135% ROI\n• Brand Value: Priceless', 430, 290, 230, 55);
  roiText.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 7: AD FORMATS ====================
function createSlide07_AdFormats(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'รูปแบบโฆษณา YouTube', 'Ad Formats');

  const formats = [
    ['▶️ Skippable In-stream', 'ข้าม Skip ได้หลัง 5 วินาที', '15-60 วินาที', 'CPV ฿0.40'],
    ['⏸️ Non-Skippable', 'ข้ามไม่ได้ ดูจนจบ', '15-20 วินาที', 'CPM ฿120'],
    ['💨 Bumper Ads', 'สั้น กระชับ จำง่าย', '6 วินาที', 'CPM ฿70'],
    ['📋 In-feed (Discovery)', 'แสดงในหน้า Search/Home', 'Thumbnail+Title', 'CPV ฿2.00'],
    ['📱 Shorts Ads', 'เข้าถึง Gen Z สูง', '15-60 วินาที', 'CPV ฿0.15']
  ];

  let y = 80;
  formats.forEach((format, i) => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, y, 640, 55);
    card.getFill().setSolidFill(i % 2 === 0 ? CONFIG.COLORS.LIGHT_GRAY : CONFIG.COLORS.WHITE);
    card.getBorder().setTransparent();

    const name = slide.insertTextBox(format[0], 50, y + 8, 180, 20);
    name.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const desc = slide.insertTextBox(format[1], 50, y + 28, 180, 20);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    const length = slide.insertTextBox(format[2], 250, y + 15, 120, 25);
    length.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);

    const costBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 550, y + 10, 100, 35);
    costBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
    costBox.getBorder().setTransparent();

    const cost = slide.insertTextBox(format[3], 550, y + 17, 100, 25);
    cost.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    cost.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    y += 60;
  });

  // Recommendation
  const recBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 385, 640, 10);
  recBox.getFill().setSolidFill(CONFIG.COLORS.SUCCESS);
  recBox.getBorder().setTransparent();
}

// ==================== SLIDE 8: TARGET AUDIENCE ====================
function createSlide08_TargetAudience(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'กลุ่มเป้าหมาย', 'Target Audiences');

  const audiences = [
    ['👨‍💼', 'นักลงทุนมือใหม่', '25-45 ปี, สนใจหุ้น/กองทุน', '40%', CONFIG.COLORS.PRIMARY],
    ['👴', 'ผู้สูงอายุ', '45-65 ปี, ระวังมิจฉาชีพ', '25%', CONFIG.COLORS.WARNING],
    ['👨‍🎓', 'Gen Z/คนทำงานใหม่', '18-30 ปี, เริ่มต้นออมเงิน', '25%', CONFIG.COLORS.ACCENT],
    ['💼', 'ผู้ประกอบการ SME', '30-55 ปี, บริหารการเงิน', '10%', CONFIG.COLORS.SUCCESS]
  ];

  let x = 40;
  audiences.forEach(aud => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 150);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().setWeight(2).getLineFill().setSolidFill(aud[4]);

    const icon = slide.insertTextBox(aud[0], x, 90, 155, 35);
    icon.getText().getTextStyle().setFontFamily('Arial').setFontSize(28);
    icon.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const name = slide.insertTextBox(aud[1], x + 10, 125, 135, 25);
    name.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);
    name.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const desc = slide.insertTextBox(aud[2], x + 10, 150, 135, 35);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(8).setForegroundColor(CONFIG.COLORS.GRAY);
    desc.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const budgetBadge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 50, 195, 55, 25);
    budgetBadge.getFill().setSolidFill(aud[4]);
    budgetBadge.getBorder().setTransparent();

    const budgetText = slide.insertTextBox(aud[3], x + 50, 200, 55, 20);
    budgetText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    budgetText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Targeting methods
  const targetTitle = slide.insertTextBox('🎯 วิธี Targeting', 40, 250, 200, 25);
  targetTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const methods = [
    'Interest: Finance, Investing, Stocks',
    'In-market: Financial Services',
    'Custom Intent: คำค้นหาการเงิน',
    'Remarketing: Video viewers 50%+',
    'Similar Audiences: Lookalike'
  ];

  let y = 280;
  methods.forEach(method => {
    const text = slide.insertTextBox('• ' + method, 50, y, 300, 20);
    text.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    y += 22;
  });

  // Location
  const locBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 400, 250, 280, 130);
  locBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  locBox.getBorder().setTransparent();

  const locTitle = slide.insertTextBox('📍 พื้นที่เป้าหมาย', 410, 260, 260, 25);
  locTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const locText = slide.insertTextBox('• กรุงเทพฯ และปริมณฑล (60%)\n• เชียงใหม่, ขอนแก่น, สงขลา (20%)\n• จังหวัดอื่นๆ ทั่วประเทศ (20%)\n\n🗣️ ภาษา: ไทย', 410, 285, 260, 85);
  locText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
}

// ==================== SLIDE 9: CONTENT CALENDAR ====================
function createSlide09_ContentCalendar(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Content Calendar', 'แผนคอนเทนต์ 4 สัปดาห์');

  // Week headers
  const weeks = ['Week 1: Launch', 'Week 2: Optimize', 'Week 3: Scale', 'Week 4: Maximize'];
  let x = 40;
  weeks.forEach((week, i) => {
    const colors = [CONFIG.COLORS.PRIMARY, CONFIG.COLORS.ACCENT, CONFIG.COLORS.SUCCESS, CONFIG.COLORS.WARNING];
    const header = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 75, 155, 30);
    header.getFill().setSolidFill(colors[i]);
    header.getBorder().setTransparent();

    const headerText = slide.insertTextBox(week, x, 80, 155, 25);
    headerText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    headerText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Content types
  const content = [
    ['Long-form', '2 วิดีโอ', '2 วิดีโอ', '2 วิดีโอ', '2 วิดีโอ'],
    ['Shorts', '7 Shorts', '7 Shorts', '7 Shorts', '7 Shorts'],
    ['Ads', '4 Creatives', '2 A/B Test', '2 New', '1 Best'],
    ['Community', 'ทุกวัน', 'ทุกวัน', 'ทุกวัน', 'ทุกวัน']
  ];

  let y = 115;
  content.forEach(row => {
    const labelBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, y, 155, 35);
    labelBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    labelBox.getBorder().setTransparent();

    const label = slide.insertTextBox(row[0], 50, y + 8, 145, 25);
    label.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    x = 205;
    for (let i = 1; i < row.length; i++) {
      const cell = slide.insertTextBox(row[i], x, y + 8, 155, 25);
      cell.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
      cell.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      x += 165;
    }
    y += 40;
  });

  // Content Pillars
  const pillarsTitle = slide.insertTextBox('📚 Content Pillars', 40, 285, 200, 25);
  pillarsTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const pillars = [
    ['Education', '40%', CONFIG.COLORS.PRIMARY],
    ['Protection', '25%', CONFIG.COLORS.WARNING],
    ['Inspiration', '20%', CONFIG.COLORS.SUCCESS],
    ['Entertainment', '15%', CONFIG.COLORS.ACCENT]
  ];

  x = 40;
  pillars.forEach(pillar => {
    const pBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 315, 155, 50);
    pBox.getFill().setSolidFill(pillar[2]);
    pBox.getBorder().setTransparent();

    const pText = slide.insertTextBox(pillar[0] + '\n' + pillar[1], x, 325, 155, 40);
    pText.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    pText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });
}

// ==================== SLIDE 10: CREATIVE SPECS ====================
function createSlide10_CreativeSpecs(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Creative Specifications', 'สเปควิดีโอและ Thumbnail');

  // Video specs
  const videoTitle = slide.insertTextBox('🎬 Video Specs', 40, 75, 200, 25);
  videoTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const videoSpecs = [
    ['Resolution', '1920 x 1080 px (แนะนำ)'],
    ['Aspect Ratio', '16:9, 1:1, 9:16 (Shorts)'],
    ['Format', 'MP4, MOV'],
    ['File Size', 'Max 256 GB'],
    ['Frame Rate', '30 fps (แนะนำ)']
  ];

  let y = 100;
  videoSpecs.forEach(spec => {
    const row = slide.insertTextBox(spec[0] + ': ' + spec[1], 50, y, 280, 20);
    row.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    y += 22;
  });

  // Thumbnail specs
  const thumbTitle = slide.insertTextBox('🖼️ Thumbnail Specs', 380, 75, 200, 25);
  thumbTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const thumbSpecs = [
    ['Size', '1280 x 720 px'],
    ['Aspect', '16:9'],
    ['Format', 'JPG, PNG, GIF'],
    ['Max Size', '2 MB'],
    ['Text', 'น้อยกว่า 20% พื้นที่']
  ];

  y = 100;
  thumbSpecs.forEach(spec => {
    const row = slide.insertTextBox(spec[0] + ': ' + spec[1], 390, y, 280, 20);
    row.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    y += 22;
  });

  // Video structure
  const structTitle = slide.insertTextBox('📝 โครงสร้างวิดีโอ Ad (30 วินาที)', 40, 230, 300, 25);
  structTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const structure = [
    ['0:00-0:03', 'Hook (ก่อน Skip)', CONFIG.COLORS.PRIMARY],
    ['0:03-0:05', 'Problem', CONFIG.COLORS.WARNING],
    ['0:05-0:15', 'Solution', CONFIG.COLORS.ACCENT],
    ['0:15-0:25', 'Benefits', CONFIG.COLORS.SUCCESS],
    ['0:25-0:30', 'CTA', CONFIG.COLORS.PRIMARY]
  ];

  let sx = 40;
  structure.forEach(item => {
    const sBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, sx, 260, 125, 55);
    sBox.getFill().setSolidFill(item[2]);
    sBox.getBorder().setTransparent();

    const sText = slide.insertTextBox(item[0] + '\n' + item[1], sx + 5, 268, 115, 40);
    sText.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    sText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    sx += 130;
  });

  // Tips
  const tipsBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 330, 640, 55);
  tipsBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  tipsBox.getBorder().setTransparent();

  const tips = slide.insertTextBox('💡 Tips: Hook ใน 3 วินาทีแรกสำคัญที่สุด! | ใส่ Subtitle ภาษาไทย | CTA ชัดเจน | Thumbnail ต้องโดดเด่น', 50, 345, 620, 25);
  tips.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
}

// ==================== SLIDE 11: GROWTH HACKS ====================
function createSlide11_GrowthHacks(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Subscriber Growth Hacks', '10 วิธีเพิ่ม Subscribers เร็ว');

  const hacks = [
    ['1', 'Shorts Strategy', 'โพสต์ 3-5 Shorts/วัน', CONFIG.COLORS.PRIMARY],
    ['2', 'Comment First', 'ตอบ Comment ใน 1 ชม.แรก', CONFIG.COLORS.ACCENT],
    ['3', 'End Screen', 'ใส่ทุกวิดีโอ ขอ Subscribe', CONFIG.COLORS.SUCCESS],
    ['4', 'Collaboration', 'ร่วมกับ Creator อื่น', CONFIG.COLORS.WARNING],
    ['5', 'Consistency', '2 Long + 7 Shorts/สัปดาห์', CONFIG.COLORS.PRIMARY]
  ];

  const hacks2 = [
    ['6', 'Trending Topics', 'ทำคอนเทนต์ตาม Trend', CONFIG.COLORS.ACCENT],
    ['7', 'Playlist', 'จัดกลุ่มวิดีโอเป็น Playlist', CONFIG.COLORS.SUCCESS],
    ['8', 'Community Post', 'โพสต์ทุกวัน เพิ่ม Engage', CONFIG.COLORS.WARNING],
    ['9', 'Cross-Promote', 'แชร์ไป LINE, X, TikTok', CONFIG.COLORS.PRIMARY],
    ['10', 'Quality > Quantity', 'วิดีโอคุณภาพ ไม่ใช่แค่ปริมาณ', CONFIG.COLORS.SUCCESS]
  ];

  // Column 1
  let y = 80;
  hacks.forEach(hack => {
    const numBox = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 40, y, 28, 28);
    numBox.getFill().setSolidFill(hack[3]);
    numBox.getBorder().setTransparent();

    const num = slide.insertTextBox(hack[0], 40, y + 5, 28, 20);
    num.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    num.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const title = slide.insertTextBox(hack[1], 75, y, 120, 18);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const desc = slide.insertTextBox(hack[2], 75, y + 16, 230, 18);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 55;
  });

  // Column 2
  y = 80;
  hacks2.forEach(hack => {
    const numBox = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 370, y, 28, 28);
    numBox.getFill().setSolidFill(hack[3]);
    numBox.getBorder().setTransparent();

    const num = slide.insertTextBox(hack[0], 370, y + 5, 28, 20);
    num.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    num.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const title = slide.insertTextBox(hack[1], 405, y, 120, 18);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

    const desc = slide.insertTextBox(hack[2], 405, y + 16, 250, 18);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 55;
  });

  // Bottom tip
  const tipBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 360, 640, 35);
  tipBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  tipBox.getBorder().setTransparent();

  const tip = slide.insertTextBox('🚀 เป้าหมาย: +1,500 Subscribers ใน 4 สัปดาห์ ด้วย Organic + Paid Strategy', 50, 368, 620, 25);
  tip.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 12: KPIs ====================
function createSlide12_KPIs(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'KPIs & Metrics', 'ตัวชี้วัดความสำเร็จ');

  const kpis = [
    ['Subscribers', '+1,500', 'คน', CONFIG.COLORS.PRIMARY],
    ['Video Views', '220,000+', 'views', CONFIG.COLORS.ACCENT],
    ['Watch Time', '15,000+', 'ชั่วโมง', CONFIG.COLORS.SUCCESS],
    ['Website Clicks', '8,000+', 'clicks', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  kpis.forEach(kpi => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 90);
    card.getFill().setSolidFill(kpi[3]);
    card.getBorder().setTransparent();

    const value = slide.insertTextBox(kpi[1], x, 95, 155, 35);
    value.getText().getTextStyle().setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    value.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const label = slide.insertTextBox(kpi[0] + ' ' + kpi[2], x, 130, 155, 25);
    label.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);
    label.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Secondary metrics
  const secTitle = slide.insertTextBox('📊 Secondary Metrics', 40, 185, 200, 25);
  secTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const metrics = [
    ['View Rate', '> 25%'],
    ['CTR', '> 2%'],
    ['Avg. View Duration', '> 50%'],
    ['Engagement Rate', '> 5%'],
    ['Cost per Subscriber', '< ฿35']
  ];

  let y = 215;
  metrics.forEach(metric => {
    const row = slide.insertTextBox('• ' + metric[0] + ': ' + metric[1], 50, y, 280, 20);
    row.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    y += 25;
  });

  // Tracking tools
  const toolsBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 380, 185, 300, 160);
  toolsBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  toolsBox.getBorder().setTransparent();

  const toolsTitle = slide.insertTextBox('🔧 เครื่องมือติดตาม', 390, 195, 280, 25);
  toolsTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);

  const tools = [
    '• YouTube Studio Analytics',
    '• Google Ads Dashboard',
    '• Google Analytics 4',
    '• Google Tag Manager',
    '• Data Studio (Reports)'
  ];

  y = 220;
  tools.forEach(tool => {
    const t = slide.insertTextBox(tool, 390, y, 280, 20);
    t.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    y += 22;
  });

  // Bottom note
  const note = slide.insertTextBox('📈 รายงานผล: Weekly + Monthly Report | ปรับกลยุทธ์ทุกสัปดาห์', 40, 365, 640, 25);
  note.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
}

// ==================== SLIDE 13: TIMELINE ====================
function createSlide13_Timeline(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Timeline', 'ไทม์ไลน์ 4 สัปดาห์');

  // Timeline line
  const line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 170, 640, 4);
  line.getFill().setSolidFill(CONFIG.COLORS.GRAY);
  line.getBorder().setTransparent();

  const phases = [
    ['Week 1', 'Launch', '• Setup Campaigns\n• Test 4 Creatives\n• งบ 20,000 บาท', CONFIG.COLORS.PRIMARY],
    ['Week 2', 'Optimize', '• A/B Testing\n• เริ่ม Remarketing\n• งบ 25,000 บาท', CONFIG.COLORS.ACCENT],
    ['Week 3', 'Scale', '• Scale Winners\n• New Audiences\n• งบ 30,000 บาท', CONFIG.COLORS.SUCCESS],
    ['Week 4', 'Maximize', '• Best Performers\n• Max Conversions\n• งบ 25,000 บาท', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  phases.forEach((phase, i) => {
    // Circle on timeline
    const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x + 70, 160, 25, 25);
    circle.getFill().setSolidFill(phase[3]);
    circle.getBorder().setTransparent();

    // Week label
    const weekLabel = slide.insertTextBox(phase[0], x, 120, 155, 25);
    weekLabel.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(phase[3]);
    weekLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Phase name
    const phaseName = slide.insertTextBox(phase[1], x, 195, 155, 25);
    phaseName.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.BLACK);
    phaseName.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Details box
    const detailBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 225, 155, 100);
    detailBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    detailBox.getBorder().setTransparent();

    const details = slide.insertTextBox(phase[2], x + 10, 235, 135, 85);
    details.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.DARK_GRAY);

    x += 165;
  });

  // Success criteria
  const successBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 340, 640, 50);
  successBox.getFill().setSolidFill(CONFIG.COLORS.SUCCESS);
  successBox.getBorder().setTransparent();

  const success = slide.insertTextBox('✅ เกณฑ์ความสำเร็จ: CTR > 2% | View Rate > 25% | Cost/Sub < ฿35 | ROAS > 1.5x', 50, 355, 620, 25);
  success.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 14: EXPECTED RESULTS ====================
function createSlide14_ExpectedResults(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Expected Results', 'ผลลัพธ์ที่คาดหวัง');

  const scenarios = [
    ['Conservative', '800 Subs\n150K Views\n4K Clicks', '-76% ROI', CONFIG.COLORS.WARNING],
    ['Moderate', '1,500 Subs\n200K Views\n6K Clicks', '+25% ROI', CONFIG.COLORS.ACCENT],
    ['Optimistic', '2,500 Subs\n300K Views\n10K Clicks', '+135% ROI', CONFIG.COLORS.SUCCESS]
  ];

  let x = 80;
  scenarios.forEach(scenario => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 180, 150);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().setWeight(3).getLineFill().setSolidFill(scenario[3]);

    const title = slide.insertTextBox(scenario[0], x, 90, 180, 25);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(scenario[3]);
    title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const metrics = slide.insertTextBox(scenario[1], x + 10, 120, 160, 60);
    metrics.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.DARK_GRAY);
    metrics.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const roi = slide.insertTextBox(scenario[2], x, 190, 180, 25);
    roi.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(scenario[3]);
    roi.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 200;
  });

  // Long-term value
  const ltBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 250, 640, 100);
  ltBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  ltBox.getBorder().setTransparent();

  const ltTitle = slide.insertTextBox('💎 Long-term Value (ไม่ได้วัดเป็นตัวเลข)', 50, 260, 620, 25);
  ltTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  const ltText = slide.insertTextBox('• Brand Awareness ที่สะสมต่อเนื่อง\n• YouTube SEO - วิดีโอติดอันดับตลอดไป\n• Content Asset ที่ใช้ได้ตลอด\n• Community ที่เติบโตแบบ Organic', 50, 285, 620, 60);
  ltText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);

  // Recommendation
  const rec = slide.insertTextBox('📌 แนะนำ: ใช้ Moderate Scenario เป็นเกณฑ์ | ปรับกลยุทธ์ตามผลจริงทุกสัปดาห์', 40, 365, 640, 25);
  rec.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
}

// ==================== SLIDE 15: THANK YOU ====================
function createSlide15_ThankYou(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.SECONDARY);

  // Red accent bar
  const redBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, CONFIG.SLIDE.HEIGHT - 8, CONFIG.SLIDE.WIDTH, 8);
  redBar.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  redBar.getBorder().setTransparent();

  // Thank you text
  const thanks = slide.insertTextBox('ขอบคุณครับ', 0, 100, CONFIG.SLIDE.WIDTH, 50);
  thanks.getText().getTextStyle().setFontFamily('Arial').setFontSize(40).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
  thanks.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const subtitle = slide.insertTextBox('YouTube Ads Campaign | Aslan Wealth Channel', 0, 155, CONFIG.SLIDE.WIDTH, 30);
  subtitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setForegroundColor(CONFIG.COLORS.GRAY);
  subtitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Contact info box
  const contactBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 200, 200, 320, 120);
  contactBox.getFill().setSolidFill(CONFIG.COLORS.DARK_GRAY);
  contactBox.getBorder().setTransparent();

  const contactText = slide.insertTextBox('📧 contact@aslanwealth.com\n🌐 https://aslanwealth.com\n📺 @AslanWealthChannel\n💬 LINE: @aslanwealth', 220, 215, 280, 90);
  contactText.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setForegroundColor(CONFIG.COLORS.WHITE);

  // Budget summary
  const budgetSummary = slide.insertTextBox('งบประมาณ: ฿100,000 | ระยะเวลา: 4 สัปดาห์ | เป้าหมาย: +1,500 Subscribers', 0, 340, CONFIG.SLIDE.WIDTH, 25);
  budgetSummary.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.GRAY);
  budgetSummary.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // CTA
  const cta = slide.insertTextBox('🚀 พร้อมเริ่มแคมเปญ!', 0, 370, CONFIG.SLIDE.WIDTH, 25);
  cta.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);
  cta.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ==================== HELPER FUNCTIONS ====================
function addHeader(slide, title, subtitle) {
  const headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, CONFIG.SLIDE.WIDTH, CONFIG.SLIDE.HEADER_HEIGHT);
  headerBg.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  headerBg.getBorder().setTransparent();

  const redLine = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, CONFIG.SLIDE.HEADER_HEIGHT - 4, CONFIG.SLIDE.WIDTH, 4);
  redLine.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  redLine.getBorder().setTransparent();

  const titleText = slide.insertTextBox(title, 25, 12, 400, 28);
  titleText.getText().getTextStyle().setFontFamily('Arial').setFontSize(20).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  if (subtitle) {
    const subtitleText = slide.insertTextBox(subtitle, 25, 35, 400, 18);
    subtitleText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
  }

  const logo = slide.insertTextBox('ASLAN', CONFIG.SLIDE.WIDTH - 80, 18, 60, 20);
  logo.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const pageNum = slide.insertTextBox('', CONFIG.SLIDE.WIDTH - 40, CONFIG.SLIDE.HEIGHT - 20, 30, 15);
  pageNum.getText().getTextStyle().setFontFamily('Arial').setFontSize(8).setForegroundColor(CONFIG.COLORS.GRAY);
}
