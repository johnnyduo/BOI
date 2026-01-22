/**
 * X (Twitter) Ads Campaign Presentation Generator
 * Aslan Wealth Channel
 * งบประมาณ: 100,000 บาท | ระยะเวลา: 4 สัปดาห์
 *
 * วิธีใช้งาน:
 * 1. เปิด Google Apps Script (script.google.com)
 * 2. สร้างโปรเจกต์ใหม่
 * 3. วาง Code นี้ทั้งหมด
 * 4. รัน function: createXAdsCampaignPresentation()
 * 5. อนุญาต permissions ที่ขอ
 * 6. Presentation จะสร้างใน Google Drive อัตโนมัติ
 *
 * Slide Ratio: 16:9 (720x405 points)
 */

// ==================== CONFIGURATION ====================
const CONFIG = {
  BRAND_NAME: 'Aslan Wealth Channel',
  WEBSITE: 'https://aslanwealth.com',
  X_HANDLE: '@AslanWealth',
  EMAIL: 'contact@aslanwealth.com',
  BUDGET: '100,000',
  DURATION: '4 สัปดาห์',

  SLIDE: {
    WIDTH: 720,
    HEIGHT: 405,
    MARGIN: 25,
    HEADER_HEIGHT: 55
  },

  COLORS: {
    PRIMARY: '#000000',       // X Black
    SECONDARY: '#14171A',     // X Dark
    ACCENT: '#1DA1F2',        // Twitter Blue (legacy)
    SUCCESS: '#17BF63',       // Green
    WARNING: '#FFAD1F',       // Yellow
    PREMIUM: '#FFD700',       // Gold (Premium)
    WHITE: '#FFFFFF',
    LIGHT_GRAY: '#F5F8FA',
    GRAY: '#657786',
    DARK_GRAY: '#AAB8C2'
  }
};

// ==================== MAIN FUNCTION ====================
function createXAdsCampaignPresentation() {
  const presentation = SlidesApp.create('X (Twitter) Ads Campaign - Aslan Wealth Channel');
  const slides = presentation.getSlides();

  if (slides.length > 0) {
    slides[0].remove();
  }

  Logger.log('เริ่มสร้าง X Ads Presentation...');

  // สร้างทุก Slides
  createSlide01_Title(presentation);
  createSlide02_Agenda(presentation);
  createSlide03_WhyX(presentation);
  createSlide04_AlgorithmHacks(presentation);
  createSlide05_Objectives(presentation);
  createSlide06_CampaignStructure(presentation);
  createSlide07_BudgetAllocation(presentation);
  createSlide08_AdFormats(presentation);
  createSlide09_TargetAudience(presentation);
  createSlide10_ContentStrategy(presentation);
  createSlide11_ViralGrowth(presentation);
  createSlide12_ThreadMastery(presentation);
  createSlide13_ContentCalendar(presentation);
  createSlide14_KPIs(presentation);
  createSlide15_ExpectedResults(presentation);
  createSlide16_ThankYou(presentation);

  Logger.log('สร้าง Presentation สำเร็จ!');
  Logger.log('URL: ' + presentation.getUrl());

  return presentation.getUrl();
}

// ==================== SLIDE 1: TITLE ====================
function createSlide01_Title(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const W = CONFIG.SLIDE.WIDTH;
  const H = CONFIG.SLIDE.HEIGHT;

  slide.getBackground().setSolidFill(CONFIG.COLORS.PRIMARY);

  // X Logo (large)
  const xLogo = slide.insertTextBox('𝕏', W - 150, 30, 100, 100);
  xLogo.getText().getTextStyle().setFontFamily('Arial').setFontSize(72).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  // Logo box
  const logoBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 30, 140, 40);
  logoBox.getFill().setSolidFill(CONFIG.COLORS.WHITE);
  logoBox.getBorder().setTransparent();
  const logoText = logoBox.getText();
  logoText.setText('ASLAN');
  logoText.getTextStyle().setFontFamily('Arial').setFontSize(20).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);
  logoText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Main Title
  const title = slide.insertTextBox('X (Twitter) Ads Campaign', 30, 120, 500, 60);
  title.getText().getTextStyle().setFontFamily('Arial').setFontSize(36).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  // Subtitle
  const subtitle = slide.insertTextBox('แผนโฆษณา X 4 สัปดาห์ | Aslan Wealth Channel', 30, 180, 500, 30);
  subtitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setForegroundColor(CONFIG.COLORS.GRAY);

  // Budget Box
  const budgetBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 30, 230, 150, 45);
  budgetBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  budgetBox.getBorder().setTransparent();
  const budgetText = slide.insertTextBox('฿100,000', 30, 240, 150, 30);
  budgetText.getText().getTextStyle().setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
  budgetText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Duration Box
  const durBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 190, 230, 100, 45);
  durBox.getFill().setSolidFill(CONFIG.COLORS.SUCCESS);
  durBox.getBorder().setTransparent();
  const durText = slide.insertTextBox('4 สัปดาห์', 190, 240, 100, 30);
  durText.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
  durText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Premium badge
  const premiumBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 300, 230, 120, 45);
  premiumBox.getFill().setSolidFill(CONFIG.COLORS.PREMIUM);
  premiumBox.getBorder().setTransparent();
  const premiumText = slide.insertTextBox('✓ Premium', 300, 240, 120, 30);
  premiumText.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);
  premiumText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Stats
  const stats = slide.insertTextBox('𝕏 550M+ Active Users | 🇹🇭 9M+ Thai Users | Real-time Engagement', 30, 350, 500, 25);
  stats.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.GRAY);

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
    ['1', 'ทำไมต้อง X Ads', 'ข้อมูลตลาดและโอกาส'],
    ['2', 'Algorithm Hacks', 'เคล็ดลับจาก Open Source'],
    ['3', 'วัตถุประสงค์', 'เป้าหมายที่ต้องการ'],
    ['4', 'โครงสร้างแคมเปญ', '4 แคมเปญหลัก'],
    ['5', 'Ad Formats', 'รูปแบบโฆษณาทั้งหมด'],
    ['6', 'Viral Growth', '10 วิธีเพิ่ม Followers'],
    ['7', 'Thread Mastery', 'เทคนิค Thread ให้ Viral'],
    ['8', 'Content Calendar', 'แผนโพสต์ 4 สัปดาห์'],
    ['9', 'KPIs & Results', 'ตัวชี้วัดและผลลัพธ์']
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
    titleText.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const descText = slide.insertTextBox(item[2], 280, y, 200, 20);
    descText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 35;
  });
}

// ==================== SLIDE 3: WHY X ====================
function createSlide03_WhyX(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'ทำไมต้อง X (Twitter)?', 'โอกาสในตลาดไทย');

  const stats = [
    ['550M+', 'Active Users/เดือน', CONFIG.COLORS.PRIMARY],
    ['9M+', 'ผู้ใช้ในไทย', CONFIG.COLORS.ACCENT],
    ['#1', 'Real-time News', CONFIG.COLORS.SUCCESS],
    ['500M', 'Tweets/วัน', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  stats.forEach(stat => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 75);
    box.getFill().setSolidFill(stat[2]);
    box.getBorder().setTransparent();

    const numText = slide.insertTextBox(stat[0], x, 90, 155, 32);
    numText.getText().getTextStyle().setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    numText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const descText = slide.insertTextBox(stat[1], x, 122, 155, 25);
    descText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);
    descText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Benefits
  const benefits = [
    '✅ Real-time engagement กับกลุ่มเป้าหมาย',
    '✅ Finance community ที่แข็งแกร่งมากในไทย',
    '✅ Thread format เหมาะกับ Educational content',
    '✅ Viral potential สูงกว่า platform อื่น',
    '✅ Premium features ช่วย boost content'
  ];

  let y = 175;
  benefits.forEach(benefit => {
    const text = slide.insertTextBox(benefit, 40, y, 400, 25);
    text.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.SECONDARY);
    y += 28;
  });

  // Premium box
  const premBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 460, 175, 230, 130);
  premBox.getFill().setSolidFill(CONFIG.COLORS.PREMIUM);
  premBox.getBorder().setTransparent();

  const premTitle = slide.insertTextBox('✓ X Premium Benefits', 470, 185, 210, 25);
  premTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const premText = slide.insertTextBox('• Blue checkmark\n• 25,000 ตัวอักษร/โพสต์\n• 60 นาที วิดีโอ\n• Algorithm boost 2-4x\n• Reply priority', 470, 210, 210, 90);
  premText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.PRIMARY);

  // Price
  const priceBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 520, 320, 120, 35);
  priceBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  priceBox.getBorder().setTransparent();

  const price = slide.insertTextBox('฿299/เดือน', 520, 328, 120, 25);
  price.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PREMIUM);
  price.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ==================== SLIDE 4: ALGORITHM HACKS ====================
function createSlide04_AlgorithmHacks(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'X Algorithm Hacks', 'เคล็ดลับจาก Open Source Code');

  // Source note
  const sourceBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 70, 640, 30);
  sourceBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  sourceBox.getBorder().setTransparent();

  const source = slide.insertTextBox('📂 Source: github.com/twitter/the-algorithm (Open Source มีนาคม 2023)', 50, 75, 620, 25);
  source.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);

  const hacks = [
    ['Replies ได้ Boost', '13.5x', 'ตอบ comment = engagement สูงสุด', CONFIG.COLORS.PRIMARY],
    ['Retweet + Quote', '20x', 'Quote RT ดีกว่า Retweet ธรรมดา', CONFIG.COLORS.ACCENT],
    ['Like คะแนนต่ำสุด', '0.5x', 'Like อย่างเดียวไม่พอ ต้องมี action อื่น', CONFIG.COLORS.WARNING],
    ['ภาพ/วิดีโอ Boost', '2x', 'Media content ได้คะแนนมากกว่า text', CONFIG.COLORS.SUCCESS]
  ];

  let y = 110;
  hacks.forEach(hack => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, y, 310, 55);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().setWeight(2).getLineFill().setSolidFill(hack[3]);

    const title = slide.insertTextBox(hack[0], 50, y + 8, 180, 20);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const boost = slide.insertTextBox(hack[1], 240, y + 8, 60, 20);
    boost.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(hack[3]);

    const desc = slide.insertTextBox(hack[2], 50, y + 30, 290, 20);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 60;
  });

  // Right column - Key insights
  const insightBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 370, 110, 310, 245);
  insightBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  insightBox.getBorder().setTransparent();

  const insightTitle = slide.insertTextBox('🔥 Key Algorithm Insights', 380, 120, 290, 25);
  insightTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  const insights = [
    '1. โพสต์ใน 15 นาทีแรกสำคัญที่สุด',
    '2. Engagement ใน 1 ชม.แรก = ตัดสินชะตา',
    '3. External links ถูก penalize',
    '4. Thread > Single tweet',
    '5. Blue checkmark boost 2-4x',
    '6. Consistency > Sporadic posts',
    '7. Reply หา big accounts = exposure',
    '8. Negative words = lower reach'
  ];

  let iy = 150;
  insights.forEach(insight => {
    const iText = slide.insertTextBox(insight, 385, iy, 285, 22);
    iText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);
    iy += 26;
  });

  // Bottom tip
  const tipBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 365, 640, 30);
  tipBox.getFill().setSolidFill(CONFIG.COLORS.ACCENT);
  tipBox.getBorder().setTransparent();

  const tip = slide.insertTextBox('💡 Tip: หลีกเลี่ยง External Links | ใส่ลิงก์ใน Reply แทน | ใช้ Thread เสมอ', 50, 372, 620, 20);
  tip.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 5: OBJECTIVES ====================
function createSlide05_Objectives(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'วัตถุประสงค์แคมเปญ', 'Campaign Objectives');

  const objectives = [
    ['🎯', 'Follower Growth', '+2,000-3,000 Followers', '30%', CONFIG.COLORS.PRIMARY],
    ['👁️', 'Impressions', '5M+ Impressions', '25%', CONFIG.COLORS.ACCENT],
    ['💬', 'Engagement', '100K+ Engagements', '25%', CONFIG.COLORS.SUCCESS],
    ['🌐', 'Website Traffic', '10,000+ Clicks', '20%', CONFIG.COLORS.WARNING]
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
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

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

// ==================== SLIDE 6: CAMPAIGN STRUCTURE ====================
function createSlide06_CampaignStructure(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'โครงสร้างแคมเปญ', '4 แคมเปญหลัก');

  const campaigns = [
    ['Follower Ads', '30,000 บาท', 'เพิ่ม Followers ตรงเป้า', CONFIG.COLORS.PRIMARY],
    ['Promoted Posts', '30,000 บาท', 'Boost คอนเทนต์คุณภาพ', CONFIG.COLORS.ACCENT],
    ['Video Views', '20,000 บาท', 'วิดีโอ Educational', CONFIG.COLORS.SUCCESS],
    ['Website Traffic', '20,000 บาท', 'Drive Traffic to Website', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  campaigns.forEach(camp => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 170);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().setWeight(2).getLineFill().setSolidFill(camp[3]);

    const topBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, 80, 155, 8);
    topBar.getFill().setSolidFill(camp[3]);
    topBar.getBorder().setTransparent();

    const title = slide.insertTextBox(camp[0], x + 10, 100, 135, 35);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const budget = slide.insertTextBox(camp[1], x + 10, 140, 135, 25);
    budget.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(camp[3]);

    const desc = slide.insertTextBox(camp[2], x + 10, 170, 135, 40);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    x += 165;
  });

  // Weekly breakdown
  const weeklyTitle = slide.insertTextBox('📅 เป้าหมายรายสัปดาห์', 40, 270, 250, 25);
  weeklyTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const weeks = [
    ['Week 1', '+300', 'Foundation'],
    ['Week 2', '+500', 'Scale'],
    ['Week 3', '+700', 'Viral'],
    ['Week 4', '+500', 'Maintain']
  ];

  let wx = 40;
  weeks.forEach(week => {
    const weekBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, wx, 300, 155, 65);
    weekBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    weekBox.getBorder().setTransparent();

    const weekTitle = slide.insertTextBox(week[0], wx + 10, 308, 135, 18);
    weekTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const weekNum = slide.insertTextBox(week[1] + ' Followers', wx + 10, 328, 135, 18);
    weekNum.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.ACCENT);

    const weekDesc = slide.insertTextBox(week[2], wx + 10, 346, 135, 15);
    weekDesc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    wx += 165;
  });
}

// ==================== SLIDE 7: BUDGET ALLOCATION ====================
function createSlide07_BudgetAllocation(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'การจัดสรรงบประมาณ', 'Budget: ฿100,000');

  const budgetData = [
    ['Follower Ads', '30%', '30,000', CONFIG.COLORS.PRIMARY],
    ['Promoted Posts', '30%', '30,000', CONFIG.COLORS.ACCENT],
    ['Video Views', '20%', '20,000', CONFIG.COLORS.SUCCESS],
    ['Website Traffic', '20%', '20,000', CONFIG.COLORS.WARNING]
  ];

  let y = 85;
  budgetData.forEach(item => {
    const colorBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, y, 30, 30);
    colorBox.getFill().setSolidFill(item[3]);
    colorBox.getBorder().setTransparent();

    const label = slide.insertTextBox(item[0], 90, y + 5, 150, 25);
    label.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const percent = slide.insertTextBox(item[1], 250, y + 5, 50, 25);
    percent.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.GRAY);

    const amount = slide.insertTextBox('฿' + item[2], 310, y + 5, 80, 25);
    amount.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(item[3]);

    y += 40;
  });

  // Cost benchmarks
  const benchTitle = slide.insertTextBox('📊 Cost Benchmarks (ไทย)', 420, 80, 250, 25);
  benchTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const benchmarks = [
    ['Cost per Follow', '฿10-30'],
    ['CPE (Engagement)', '฿2-8'],
    ['CPC', '฿3-10'],
    ['CPV (6 sec)', '฿1-3'],
    ['CPM', '฿30-80']
  ];

  y = 110;
  benchmarks.forEach(bench => {
    const row = slide.insertTextBox(bench[0] + ': ' + bench[1], 420, y, 250, 20);
    row.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
    y += 25;
  });

  // Additional costs
  const addBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 260, 320, 100);
  addBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  addBox.getBorder().setTransparent();

  const addTitle = slide.insertTextBox('💰 ค่าใช้จ่ายเพิ่มเติม (ไม่รวมในงบ Ads)', 50, 270, 300, 25);
  addTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const addText = slide.insertTextBox('• X Premium: ฿299/เดือน = ฿299\n• Creator สร้าง Content: ฿0-15,000\n• Design/Video: ฿0-10,000', 50, 295, 300, 60);
  addText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);

  // Expected box
  const expBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 380, 260, 300, 100);
  expBox.getFill().setSolidFill(CONFIG.COLORS.SUCCESS);
  expBox.getBorder().setTransparent();

  const expTitle = slide.insertTextBox('📈 Expected Return', 390, 270, 280, 25);
  expTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  const expText = slide.insertTextBox('• 2,000-3,000 Followers ใหม่\n• 5M+ Impressions\n• 100K+ Engagements\n• Brand Value: Priceless', 390, 295, 280, 60);
  expText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 8: AD FORMATS ====================
function createSlide08_AdFormats(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'รูปแบบโฆษณา X', 'Ad Formats & Specs');

  const formats = [
    ['📱 Promoted Ads (Image)', '1200x675 หรือ 1200x1200', 'CPE ฿2-8', CONFIG.COLORS.PRIMARY],
    ['🎬 Promoted Ads (Video)', '1920x1080, Max 2:20', 'CPV ฿1-3', CONFIG.COLORS.ACCENT],
    ['👥 Follower Ads', 'Profile card + Follow button', '฿10-30/follow', CONFIG.COLORS.SUCCESS],
    ['🔗 Website Cards', '800x418 หรือ 800x800', 'CPC ฿3-10', CONFIG.COLORS.WARNING],
    ['🎠 Carousel Ads', '2-6 cards, swipeable', 'CPE ฿2-8', CONFIG.COLORS.PRIMARY]
  ];

  let y = 80;
  formats.forEach((format, i) => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, y, 640, 52);
    card.getFill().setSolidFill(i % 2 === 0 ? CONFIG.COLORS.LIGHT_GRAY : CONFIG.COLORS.WHITE);
    card.getBorder().setTransparent();

    const name = slide.insertTextBox(format[0], 50, y + 8, 200, 20);
    name.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const spec = slide.insertTextBox(format[1], 50, y + 28, 200, 18);
    spec.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    const costBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 550, y + 10, 100, 32);
    costBox.getFill().setSolidFill(format[3]);
    costBox.getBorder().setTransparent();

    const cost = slide.insertTextBox(format[2], 550, y + 16, 100, 22);
    cost.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    cost.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    y += 57;
  });

  // Tips
  const tipBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 365, 640, 30);
  tipBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  tipBox.getBorder().setTransparent();

  const tip = slide.insertTextBox('💡 แนะนำ: ใช้ Follower Ads + Promoted Posts ร่วมกัน | Video Ads สำหรับ Brand Awareness', 50, 372, 620, 20);
  tip.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 9: TARGET AUDIENCE ====================
function createSlide09_TargetAudience(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'กลุ่มเป้าหมาย', 'Target Audiences');

  const audiences = [
    ['👨‍💼', 'นักลงทุน/FinTwit', '25-45 ปี, ติดตามข่าวหุ้น', '40%', CONFIG.COLORS.PRIMARY],
    ['📰', 'ผู้ติดตามข่าว', '30-55 ปี, อ่านข่าวทุกวัน', '25%', CONFIG.COLORS.ACCENT],
    ['🎓', 'คนรุ่นใหม่', '18-30 ปี, สนใจการเงิน', '25%', CONFIG.COLORS.SUCCESS],
    ['💼', 'ผู้ประกอบการ', '30-55 ปี, ธุรกิจ SME', '10%', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  audiences.forEach(aud => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 140);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().setWeight(2).getLineFill().setSolidFill(aud[4]);

    const icon = slide.insertTextBox(aud[0], x, 90, 155, 32);
    icon.getText().getTextStyle().setFontFamily('Arial').setFontSize(24);
    icon.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const name = slide.insertTextBox(aud[1], x + 10, 122, 135, 22);
    name.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);
    name.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const desc = slide.insertTextBox(aud[2], x + 10, 144, 135, 30);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(8).setForegroundColor(CONFIG.COLORS.GRAY);
    desc.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const budgetBadge = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 50, 180, 55, 25);
    budgetBadge.getFill().setSolidFill(aud[4]);
    budgetBadge.getBorder().setTransparent();

    const budgetText = slide.insertTextBox(aud[3], x + 50, 185, 55, 18);
    budgetText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    budgetText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Targeting methods
  const methodTitle = slide.insertTextBox('🎯 Targeting Methods', 40, 235, 200, 25);
  methodTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const methods = [
    '• Interest: Finance, Investing, Cryptocurrency',
    '• Follower Lookalikes: @thunhoon, @finnomena',
    '• Keywords: หุ้น, ลงทุน, การเงิน',
    '• Conversation Topics: Stock Market TH'
  ];

  let y = 260;
  methods.forEach(method => {
    const m = slide.insertTextBox(method, 50, y, 300, 20);
    m.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
    y += 22;
  });

  // Accounts to target
  const accBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 380, 235, 300, 120);
  accBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  accBox.getBorder().setTransparent();

  const accTitle = slide.insertTextBox('📌 Accounts ให้ Reply หา', 390, 245, 280, 25);
  accTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const accounts = slide.insertTextBox('@thunhoon, @finnaborjar\n@settrade, @jitta_invest\n@money_channel, @thestandard_th\n@workloads, @pptv36', 390, 270, 280, 80);
  accounts.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
}

// ==================== SLIDE 10: CONTENT STRATEGY ====================
function createSlide10_ContentStrategy(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Content Strategy', 'กลยุทธ์คอนเทนต์');

  // Content pillars
  const pillars = [
    ['Education', '40%', 'ความรู้การเงิน/ลงทุน', CONFIG.COLORS.PRIMARY],
    ['News', '25%', 'ข่าวและวิเคราะห์', CONFIG.COLORS.ACCENT],
    ['Engagement', '20%', 'Poll, Q&A, Hot Takes', CONFIG.COLORS.SUCCESS],
    ['Promotion', '15%', 'CTA, คอร์ส, เว็บไซต์', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  pillars.forEach(pillar => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 70);
    box.getFill().setSolidFill(pillar[3]);
    box.getBorder().setTransparent();

    const title = slide.insertTextBox(pillar[0] + ' (' + pillar[1] + ')', x, 90, 155, 25);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const desc = slide.insertTextBox(pillar[2], x, 115, 155, 25);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.WHITE);
    desc.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Post templates
  const tempTitle = slide.insertTextBox('📝 Post Templates', 40, 165, 200, 25);
  tempTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const templates = [
    ['💡 Quick Tip', 'เคล็ดลับสั้นๆ + Save ไว้อ่านอีก'],
    ['📊 Stat/Fact', 'ตัวเลขที่น่าตกใจ + วิเคราะห์'],
    ['❓ Question', 'คำถามกระตุ้น Reply'],
    ['🔥 Hot Take', 'Unpopular opinion + debate'],
    ['📰 News', 'ข่าว + ความเห็นสั้นๆ']
  ];

  let y = 190;
  templates.forEach(temp => {
    const tRow = slide.insertTextBox(temp[0] + ' - ' + temp[1], 50, y, 280, 22);
    tRow.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
    y += 25;
  });

  // Posting schedule
  const schedBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 360, 165, 320, 180);
  schedBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  schedBox.getBorder().setTransparent();

  const schedTitle = slide.insertTextBox('⏰ Daily Posting Schedule', 370, 175, 300, 25);
  schedTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const schedule = [
    ['07:00', 'News Commentary (ข่าวเช้า)'],
    ['12:00', 'Quick Tip (พักกลางวัน)'],
    ['18:00', 'Thread / Educational'],
    ['21:00', 'Poll / Engagement']
  ];

  y = 205;
  schedule.forEach(sched => {
    const time = slide.insertTextBox(sched[0], 380, y, 50, 20);
    time.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const content = slide.insertTextBox(sched[1], 440, y, 230, 20);
    content.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 30;
  });

  const schedNote = slide.insertTextBox('📌 เป้าหมาย: 5-10 โพสต์/วัน', 380, 320, 280, 20);
  schedNote.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.ACCENT);

  // Bottom tip
  const tipBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 365, 640, 30);
  tipBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  tipBox.getBorder().setTransparent();

  const tip = slide.insertTextBox('💡 Golden Rule: 80% Value, 20% Promotion | ให้ก่อนขาย | สร้าง Trust ก่อน Convert', 50, 372, 620, 20);
  tip.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 11: VIRAL GROWTH ====================
function createSlide11_ViralGrowth(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Viral Growth Hacks', '10 วิธีเพิ่ม Followers เร็ว');

  const hacks = [
    ['1', 'Blue Checkmark', 'Algorithm Boost 2-4x', CONFIG.COLORS.PREMIUM],
    ['2', 'Thread Strategy', '10-15 tweets/thread', CONFIG.COLORS.PRIMARY],
    ['3', 'Reply Game', '20 replies/วัน หา big accounts', CONFIG.COLORS.ACCENT],
    ['4', 'Trend Hijacking', 'ดู Trending + สร้างเนื้อหา', CONFIG.COLORS.SUCCESS],
    ['5', 'Consistent Posting', '5-10 โพสต์/วัน', CONFIG.COLORS.WARNING]
  ];

  const hacks2 = [
    ['6', 'Cross-Promote', 'แชร์ไป LINE, YouTube', CONFIG.COLORS.PRIMARY],
    ['7', 'Engagement First', 'ตอบทุก Reply', CONFIG.COLORS.ACCENT],
    ['8', 'Quote > Retweet', 'Quote RT + เพิ่มความเห็น', CONFIG.COLORS.SUCCESS],
    ['9', 'No External Links', 'ใส่ลิงก์ใน Reply', CONFIG.COLORS.WARNING],
    ['10', 'Collaborate', 'Space, Quote กับคนอื่น', CONFIG.COLORS.PREMIUM]
  ];

  // Column 1
  let y = 80;
  hacks.forEach(hack => {
    const numBox = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 40, y, 25, 25);
    numBox.getFill().setSolidFill(hack[3]);
    numBox.getBorder().setTransparent();

    const num = slide.insertTextBox(hack[0], 40, y + 3, 25, 20);
    num.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    num.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const title = slide.insertTextBox(hack[1], 72, y, 120, 15);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const desc = slide.insertTextBox(hack[2], 72, y + 14, 230, 15);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 50;
  });

  // Column 2
  y = 80;
  hacks2.forEach(hack => {
    const numBox = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 370, y, 25, 25);
    numBox.getFill().setSolidFill(hack[3]);
    numBox.getBorder().setTransparent();

    const num = slide.insertTextBox(hack[0], 370, y + 3, 25, 20);
    num.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    num.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const title = slide.insertTextBox(hack[1], 402, y, 120, 15);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    const desc = slide.insertTextBox(hack[2], 402, y + 14, 250, 15);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 50;
  });

  // Target box
  const targetBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 340, 640, 50);
  targetBox.getFill().setSolidFill(CONFIG.COLORS.SUCCESS);
  targetBox.getBorder().setTransparent();

  const target = slide.insertTextBox('🎯 เป้าหมาย: +2,000-3,000 Followers ใน 4 สัปดาห์ | Organic + Paid Strategy', 50, 355, 620, 25);
  target.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
}

// ==================== SLIDE 12: THREAD MASTERY ====================
function createSlide12_ThreadMastery(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Thread Mastery', 'เทคนิค Thread ให้ Viral');

  // Thread structure
  const structTitle = slide.insertTextBox('📝 โครงสร้าง Thread ที่ได้ผล', 40, 75, 300, 25);
  structTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const structure = [
    ['Tweet 1', 'HOOK - คำถาม/ปัญหาที่น่าสนใจ', CONFIG.COLORS.PRIMARY],
    ['Tweet 2-9', 'CONTENT - บทเรียน/ข้อมูล + ตัวอย่าง', CONFIG.COLORS.ACCENT],
    ['Tweet 10', 'SUMMARY - สรุป + CTA', CONFIG.COLORS.SUCCESS]
  ];

  let y = 100;
  structure.forEach(item => {
    const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 50, y, 280, 45);
    box.getFill().setSolidFill(item[2]);
    box.getBorder().setTransparent();

    const label = slide.insertTextBox(item[0], 60, y + 5, 80, 18);
    label.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

    const desc = slide.insertTextBox(item[1], 60, y + 23, 260, 18);
    desc.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.WHITE);

    y += 50;
  });

  // Thread templates
  const tempTitle = slide.insertTextBox('🔥 Thread Templates', 380, 75, 200, 25);
  tempTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const templates = [
    '"X สิ่งที่ต้องรู้เกี่ยวกับ [หัวข้อ]"',
    '"ผมศึกษาเรื่อง [X] มา [เวลา] นี่คือบทเรียน"',
    '"เมื่อ [เวลา] ที่แล้ว [เหตุการณ์] นี่คือเรื่องราว"',
    '"ทุกคนเข้าใจผิดเรื่องนี้... ความจริงคือ"',
    '"วิธีทำ [X] ใน [เวลา] ง่ายๆ แค่ [จำนวน] ขั้นตอน"'
  ];

  y = 100;
  templates.forEach(temp => {
    const tBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 360, y, 320, 35);
    tBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
    tBox.getBorder().setTransparent();

    const tText = slide.insertTextBox(temp, 370, y + 8, 300, 25);
    tText.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);

    y += 40;
  });

  // CTA template
  const ctaBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 270, 640, 85);
  ctaBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  ctaBox.getBorder().setTransparent();

  const ctaTitle = slide.insertTextBox('✅ CTA ท้าย Thread (Tweet สุดท้าย)', 50, 280, 620, 25);
  ctaTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  const ctaText = slide.insertTextBox('สรุป:\n✅ [Point 1] ✅ [Point 2] ✅ [Point 3]\n\nถ้าชอบ Thread นี้:\n• Like Tweet แรก\n• Repost ให้เพื่อน\n• Follow @AslanWealth', 50, 300, 620, 50);
  ctaText.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.WHITE);

  // Tips
  const tips = slide.insertTextBox('💡 Thread Tips: 10-15 tweets ต่อ Thread | ใส่ภาพทุก 2-3 tweets | Hook ต้องแรง | โพสต์ 1-2 ครั้ง/สัปดาห์', 40, 370, 640, 25);
  tips.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
}

// ==================== SLIDE 13: CONTENT CALENDAR ====================
function createSlide13_ContentCalendar(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Content Calendar', 'ปฏิทินโพสต์ 1 สัปดาห์');

  // Schedule table
  const days = ['จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.', 'อา.'];
  const times = ['7am', '12pm', '6pm', '9pm'];
  const content = [
    ['News', 'News', 'News', 'News', 'News', 'Tip', 'Tip'],
    ['Tip', 'Tip', 'Tip', 'Tip', 'Tip', 'Engage', 'Engage'],
    ['Thread', 'Engage', 'Video', 'Engage', 'Thread', 'Thread', 'Summary'],
    ['Poll', 'Quote', 'Poll', 'Quote', 'Poll', 'Poll', 'Poll']
  ];

  // Day headers
  let x = 100;
  days.forEach(day => {
    const dayBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 75, 80, 25);
    dayBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
    dayBox.getBorder().setTransparent();

    const dayText = slide.insertTextBox(day, x, 80, 80, 20);
    dayText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    dayText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 85;
  });

  // Time labels and content
  let y = 105;
  times.forEach((time, ti) => {
    const timeText = slide.insertTextBox(time, 45, y + 8, 50, 20);
    timeText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

    x = 100;
    content[ti].forEach((cell, ci) => {
      const cellBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, 80, 35);

      let color = CONFIG.COLORS.LIGHT_GRAY;
      if (cell === 'Thread') color = CONFIG.COLORS.PRIMARY;
      else if (cell === 'News') color = CONFIG.COLORS.ACCENT;
      else if (cell === 'Poll' || cell === 'Engage') color = CONFIG.COLORS.SUCCESS;
      else if (cell === 'Video') color = CONFIG.COLORS.WARNING;

      cellBox.getFill().setSolidFill(color);
      cellBox.getBorder().setTransparent();

      const cellText = slide.insertTextBox(cell, x, y + 8, 80, 25);
      const textColor = (cell === 'Tip' || cell === 'Quote' || cell === 'Summary') ? CONFIG.COLORS.GRAY : CONFIG.COLORS.WHITE;
      cellText.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(textColor);
      cellText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      x += 85;
    });
    y += 40;
  });

  // Legend
  const legend = [
    ['Thread', CONFIG.COLORS.PRIMARY],
    ['News', CONFIG.COLORS.ACCENT],
    ['Poll/Engage', CONFIG.COLORS.SUCCESS],
    ['Video', CONFIG.COLORS.WARNING]
  ];

  x = 100;
  legend.forEach(item => {
    const legBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 280, 80, 22);
    legBox.getFill().setSolidFill(item[1]);
    legBox.getBorder().setTransparent();

    const legText = slide.insertTextBox(item[0], x, 284, 80, 18);
    legText.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.WHITE);
    legText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 150;
  });

  // Daily checklist
  const checkTitle = slide.insertTextBox('✅ Daily Checklist', 40, 315, 200, 25);
  checkTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const checks = [
    '□ โพสต์ 5-10 tweets',
    '□ Reply 20 accounts ใหญ่',
    '□ Retweet 2-3 โพสต์ดี',
    '□ ตอบ comments ทุกอัน'
  ];

  x = 40;
  checks.forEach(check => {
    const cText = slide.insertTextBox(check, x, 340, 160, 50);
    cText.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(CONFIG.COLORS.GRAY);
    x += 165;
  });
}

// ==================== SLIDE 14: KPIs ====================
function createSlide14_KPIs(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'KPIs & Metrics', 'ตัวชี้วัดความสำเร็จ');

  const kpis = [
    ['Followers', '+2,500', 'คน', CONFIG.COLORS.PRIMARY],
    ['Impressions', '5M+', 'views', CONFIG.COLORS.ACCENT],
    ['Engagements', '100K+', 'actions', CONFIG.COLORS.SUCCESS],
    ['Website Clicks', '10K+', 'clicks', CONFIG.COLORS.WARNING]
  ];

  let x = 40;
  kpis.forEach(kpi => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 155, 85);
    card.getFill().setSolidFill(kpi[3]);
    card.getBorder().setTransparent();

    const value = slide.insertTextBox(kpi[1], x, 92, 155, 32);
    value.getText().getTextStyle().setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
    value.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const label = slide.insertTextBox(kpi[0] + ' ' + kpi[2], x, 125, 155, 25);
    label.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);
    label.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 165;
  });

  // Secondary metrics
  const secTitle = slide.insertTextBox('📊 Secondary Metrics', 40, 180, 200, 25);
  secTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const metrics = [
    ['Engagement Rate', '> 3%'],
    ['Profile Visits', '> 50K'],
    ['Cost per Follow', '< ฿30'],
    ['Thread Completion', '> 60%']
  ];

  let y = 205;
  metrics.forEach(metric => {
    const row = slide.insertTextBox('• ' + metric[0] + ': ' + metric[1], 50, y, 250, 20);
    row.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
    y += 25;
  });

  // Tools
  const toolsBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 380, 180, 300, 130);
  toolsBox.getFill().setSolidFill(CONFIG.COLORS.LIGHT_GRAY);
  toolsBox.getBorder().setTransparent();

  const toolsTitle = slide.insertTextBox('🔧 เครื่องมือติดตาม', 390, 190, 280, 25);
  toolsTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor(CONFIG.COLORS.PRIMARY);

  const tools = [
    '• X Analytics (Premium)',
    '• X Ads Manager',
    '• TweetDeck',
    '• Buffer/Hootsuite',
    '• Google Analytics'
  ];

  y = 215;
  tools.forEach(tool => {
    const t = slide.insertTextBox(tool, 390, y, 280, 20);
    t.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
    y += 22;
  });

  // Report note
  const note = slide.insertTextBox('📈 รายงานผล: Weekly + Monthly Report | ปรับกลยุทธ์ทุกสัปดาห์ตามข้อมูล', 40, 330, 640, 25);
  note.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
}

// ==================== SLIDE 15: EXPECTED RESULTS ====================
function createSlide15_ExpectedResults(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.WHITE);

  addHeader(slide, 'Expected Results', 'ผลลัพธ์ที่คาดหวัง');

  const scenarios = [
    ['Conservative', '1,500 Followers\n3M Impressions\n50K Engagements', CONFIG.COLORS.WARNING],
    ['Moderate', '2,500 Followers\n5M Impressions\n100K Engagements', CONFIG.COLORS.ACCENT],
    ['Optimistic', '4,000 Followers\n8M Impressions\n150K Engagements', CONFIG.COLORS.SUCCESS]
  ];

  let x = 60;
  scenarios.forEach(scenario => {
    const card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 190, 140);
    card.getFill().setSolidFill(CONFIG.COLORS.WHITE);
    card.getBorder().setWeight(3).getLineFill().setSolidFill(scenario[2]);

    const title = slide.insertTextBox(scenario[0], x, 90, 190, 25);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(scenario[2]);
    title.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    const metrics = slide.insertTextBox(scenario[1], x + 15, 120, 160, 90);
    metrics.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.GRAY);
    metrics.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    x += 210;
  });

  // Long-term value
  const ltBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 240, 640, 100);
  ltBox.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  ltBox.getBorder().setTransparent();

  const ltTitle = slide.insertTextBox('💎 Long-term Value', 50, 250, 620, 25);
  ltTitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(12).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  const ltText = slide.insertTextBox('• สร้าง Authority ในวงการ Finance Twitter (FinTwit)\n• Network กับ Influencers และนักลงทุน\n• Content ที่ Viral ต่อเนื่อง (Evergreen threads)\n• Community ที่ Engage และ Convert ได้ระยะยาว', 50, 275, 620, 60);
  ltText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.WHITE);

  // Recommendation
  const rec = slide.insertTextBox('📌 แนะนำ: เริ่มต้นด้วย Moderate target | วัดผลทุกสัปดาห์ | ปรับกลยุทธ์ตามข้อมูลจริง', 40, 360, 640, 25);
  rec.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
}

// ==================== SLIDE 16: THANK YOU ====================
function createSlide16_ThankYou(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(CONFIG.COLORS.PRIMARY);

  // X Logo
  const xLogo = slide.insertTextBox('𝕏', CONFIG.SLIDE.WIDTH / 2 - 40, 40, 80, 80);
  xLogo.getText().getTextStyle().setFontFamily('Arial').setFontSize(60).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
  xLogo.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Thank you
  const thanks = slide.insertTextBox('ขอบคุณครับ', 0, 130, CONFIG.SLIDE.WIDTH, 45);
  thanks.getText().getTextStyle().setFontFamily('Arial').setFontSize(36).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);
  thanks.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const subtitle = slide.insertTextBox('X (Twitter) Ads Campaign | Aslan Wealth Channel', 0, 175, CONFIG.SLIDE.WIDTH, 30);
  subtitle.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setForegroundColor(CONFIG.COLORS.GRAY);
  subtitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Contact box
  const contactBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 200, 210, 320, 100);
  contactBox.getFill().setSolidFill(CONFIG.COLORS.SECONDARY);
  contactBox.getBorder().setTransparent();

  const contactText = slide.insertTextBox('📧 contact@aslanwealth.com\n🌐 https://aslanwealth.com\n𝕏  @AslanWealth\n💬 LINE: @aslanwealth', 220, 225, 280, 80);
  contactText.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.WHITE);

  // Budget summary
  const summary = slide.insertTextBox('งบประมาณ: ฿100,000 | ระยะเวลา: 4 สัปดาห์ | เป้าหมาย: +2,500 Followers', 0, 330, CONFIG.SLIDE.WIDTH, 25);
  summary.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(CONFIG.COLORS.GRAY);
  summary.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // CTA
  const cta = slide.insertTextBox('🚀 พร้อมเริ่มแคมเปญ!', 0, 360, CONFIG.SLIDE.WIDTH, 30);
  cta.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(CONFIG.COLORS.PREMIUM);
  cta.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

// ==================== HELPER FUNCTIONS ====================
function addHeader(slide, title, subtitle) {
  const headerBg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, CONFIG.SLIDE.WIDTH, CONFIG.SLIDE.HEADER_HEIGHT);
  headerBg.getFill().setSolidFill(CONFIG.COLORS.PRIMARY);
  headerBg.getBorder().setTransparent();

  const titleText = slide.insertTextBox(title, 25, 12, 400, 28);
  titleText.getText().getTextStyle().setFontFamily('Arial').setFontSize(20).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  if (subtitle) {
    const subtitleText = slide.insertTextBox(subtitle, 25, 35, 400, 18);
    subtitleText.getText().getTextStyle().setFontFamily('Arial').setFontSize(10).setForegroundColor(CONFIG.COLORS.GRAY);
  }

  // X logo in header
  const xLogo = slide.insertTextBox('𝕏', CONFIG.SLIDE.WIDTH - 50, 15, 30, 30);
  xLogo.getText().getTextStyle().setFontFamily('Arial').setFontSize(20).setBold(true).setForegroundColor(CONFIG.COLORS.WHITE);

  const pageNum = slide.insertTextBox('', CONFIG.SLIDE.WIDTH - 40, CONFIG.SLIDE.HEIGHT - 20, 30, 15);
  pageNum.getText().getTextStyle().setFontFamily('Arial').setFontSize(8).setForegroundColor(CONFIG.COLORS.GRAY);
}
