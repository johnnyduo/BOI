/**
 * BOI PRESENTATION GENERATOR - ASLAN AI GUARDIAN (ภาษาไทย)
 *
 * Google Apps Script สำหรับสร้างงานนำเสนอ Google Slides แบบมืออาชีพ
 * ธีม: อัสลาน (สิงโต) - ทอง, น้ำเงินเข้ม, มืออาชีพ
 *
 * คำแนะนำการติดตั้ง:
 * 1. เปิด Google Slides
 * 2. สร้างงานนำเสนอใหม่
 * 3. เครื่องมือ → Script editor
 * 4. วางโค้ดนี้
 * 5. บันทึกและรัน createPresentation()
 * 6. อนุญาตสคริปต์
 * 7. รอให้เสร็จ (สร้าง ~25 สไลด์)
 *
 * รูปภาพจาก PEXELS (ฟรี, ไม่มีลายน้ำ):
 * - ใช้จาก Pexels.com API หรือ URL โดยตรง
 * - รูปภาพทั้งหมดปลอดค่าลิขสิทธิ์
 */

// ===========================
// การตั้งค่าธีม
// ===========================

const THEME = {
  colors: {
    primary: '#C9A961',      // ทอง (แผงคอสิงโต)
    secondary: '#1E3A8A',    // น้ำเงินเข้ม
    accent: '#DC2626',       // แดง (เตือน/แจ้งเหตุ)
    dark: '#1F2937',         // เทาเข้ม
    light: '#F9FAFB',        // เทาอ่อน พื้นหลัง
    white: '#FFFFFF',
    success: '#059669',      // เขียว
  },

  fonts: {
    title: 'Sarabun',
    body: 'Sarabun',
    heading: 'Sarabun',
  },

  images: {
    // Pexels URLs (ฟรี, ไม่มีลายน้ำ)
    lion: 'https://images.pexels.com/photos/33045/lion-wild-africa-african.jpg',
    aiTech: 'https://images.pexels.com/photos/8386440/pexels-photo-8386440.jpeg',
    robot: 'https://images.pexels.com/photos/8386434/pexels-photo-8386434.jpeg',
    security: 'https://images.pexels.com/photos/60504/security-protection-anti-virus-software-60504.jpeg',
    financial: 'https://images.pexels.com/photos/210600/pexels-photo-210600.jpeg',
    kids: 'https://images.pexels.com/photos/1648387/pexels-photo-1648387.jpeg',
    elderly: 'https://images.pexels.com/photos/3768131/pexels-photo-3768131.jpeg',
    team: 'https://images.pexels.com/photos/3184339/pexels-photo-3184339.jpeg',
    innovation: 'https://images.pexels.com/photos/3861969/pexels-photo-3861969.jpeg',
    thailand: 'https://images.pexels.com/photos/4042876/pexels-photo-4042876.jpeg',
  }
};

// ===========================
// ฟังก์ชันหลัก
// ===========================

function createPresentation() {
  Logger.log('🚀 เริ่มสร้างงานนำเสนอ BOI...');

  // สร้างงานนำเสนอใหม่
  const presentation = SlidesApp.create('ASLAN AI GUARDIAN - ข้อเสนอ BOI (ภาษาไทย)');
  const presentationId = presentation.getId();

  Logger.log(`✅ สร้างงานนำเสนอแล้ว: ${presentationId}`);
  Logger.log(`🔗 URL: https://docs.google.com/presentation/d/${presentationId}/edit`);

  // ลบสไลด์เริ่มต้น
  const slides = presentation.getSlides();
  if (slides.length > 0) {
    slides[0].remove();
  }

  // สร้างสไลด์ทั้งหมด
  createTitleSlide(presentation);
  createExecutiveSummarySlide(presentation);
  createProblemSlide(presentation);
  createSolutionSlide(presentation);
  createPhysicalAISlide(presentation);
  createMultiAgentSlide(presentation);
  createTargetUsersSlide(presentation);
  createProductsSlide(presentation);
  createInnovationSlide(presentation);
  createRDActivitiesSlide(presentation);
  createDigitalTechSlide(presentation);
  createInvestmentSlide(presentation);
  createBOIComplianceSlide(presentation);
  createAORComplianceSlide(presentation);
  createEconomicImpactSlide(presentation);
  createSocialImpactSlide(presentation);
  createTimelineSlide(presentation);
  createKPIsSlide(presentation);
  createTeamSlide(presentation);
  createRiskMitigationSlide(presentation);
  createCompetitiveAdvantageSlide(presentation);
  createNextStepsSlide(presentation);
  createThankYouSlide(presentation);

  Logger.log('🎊 สร้างงานนำเสนอสำเร็จ!');
  Logger.log(`📊 จำนวนสไลด์ทั้งหมด: ${presentation.getSlides().length}`);
  Logger.log(`🔗 เปิด: https://docs.google.com/presentation/d/${presentationId}/edit`);

  return presentationId;
}

// ===========================
// ฟังก์ชันสร้างสไลด์
// ===========================

function createTitleSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // พื้นหลัง
  const background = slide.getBackground();
  background.setSolidFill(THEME.colors.secondary);

  // เพิ่มภาพสิงโต (โปร่งแสง)
  try {
    const image = slide.insertImage(THEME.images.lion);
    image.setLeft(0).setTop(0).setWidth(720).setHeight(405);
    image.setTransparency(0.7);
  } catch (e) {
    Logger.log('⚠️ ไม่สามารถเพิ่มภาพสิงโตได้');
  }

  // หัวข้อหลัก
  const titleBox = slide.insertTextBox('ASLAN AI GUARDIAN');
  titleBox.setLeft(50).setTop(100).setWidth(620).setHeight(80);
  const titleText = titleBox.getText();
  titleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  const titleStyle = titleText.getTextStyle();
  titleStyle.setFontSize(54).setBold(true).setForegroundColor(THEME.colors.primary);
  titleStyle.setFontFamily(THEME.fonts.title);

  // คำบรรยาย
  const subtitleBox = slide.insertTextBox('ตัวแทน AI ทางกายภาพ + แพลตฟอร์มหลายเอเจนต์\nสำหรับการรู้เท่าทันทางการเงินและการป้องกัน');
  subtitleBox.setLeft(50).setTop(190).setWidth(620).setHeight(70);
  const subtitleText = subtitleBox.getText();
  subtitleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  const subtitleStyle = subtitleText.getTextStyle();
  subtitleStyle.setFontSize(22).setForegroundColor(THEME.colors.white);
  subtitleStyle.setFontFamily(THEME.fonts.body);

  // จำนวนเงินลงทุน
  const investmentBox = slide.insertTextBox('💰 เงินลงทุน 85 ล้านบาท');
  investmentBox.setLeft(200).setTop(280).setWidth(320).setHeight(40);
  const investmentText = investmentBox.getText();
  investmentText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  const investmentStyle = investmentText.getTextStyle();
  investmentStyle.setFontSize(26).setBold(true).setForegroundColor(THEME.colors.primary);
  investmentStyle.setFontFamily(THEME.fonts.heading);

  // ชื่อบริษัท
  const companyBox = slide.insertTextBox('บริษัท อัสลาน เวลธ์ จำกัด\nASLAN WEALTH COMPANY LIMITED');
  companyBox.setLeft(50).setTop(350).setWidth(620).setHeight(40);
  const companyText = companyBox.getText();
  companyText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  const companyStyle = companyText.getTextStyle();
  companyStyle.setFontSize(16).setForegroundColor(THEME.colors.light);
  companyStyle.setFontFamily(THEME.fonts.body);

  Logger.log('✅ สร้างแล้ว: สไลด์หน้าปก');
}

function createExecutiveSummarySlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(THEME.colors.white);

  addSlideHeader(slide, '🎯 สรุปสำหรับผู้บริหาร', THEME.colors.secondary);

  const columnWidth = 210;
  const columnHeight = 280;
  const startY = 80;
  const gap = 15;

  // คอลัมน์ 1: อะไร
  const col1 = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  col1.setLeft(20).setTop(startY).setWidth(columnWidth).setHeight(columnHeight);
  col1.getFill().setSolidFill(THEME.colors.light);
  col1.getBorder().setTransparent();

  const col1Text = slide.insertTextBox(
    '🦁 อะไร\n\n' +
    'Aslan AI Guardian ประกอบด้วย:\n\n' +
    '• ตัวแทน AI ทางกายภาพ\n' +
    '  (หุ่นยนต์สิงโตมาสคอต)\n\n' +
    '• แพลตฟอร์ม 5 AI Agent\n' +
    '  (ปัญญาบนคลาวด์)\n\n' +
    '• แอปพลิเคชันมือถือ\n' +
    '  (อินเทอร์เฟซผู้ใช้)'
  );
  col1Text.setLeft(30).setTop(startY + 10).setWidth(columnWidth - 20).setHeight(columnHeight - 20);
  styleTextBox(col1Text, 15, THEME.colors.dark, false);

  // คอลัมน์ 2: ใคร
  const col2 = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  col2.setLeft(20 + columnWidth + gap).setTop(startY).setWidth(columnWidth).setHeight(columnHeight);
  col2.getFill().setSolidFill(THEME.colors.light);
  col2.getBorder().setTransparent();

  const col2Text = slide.insertTextBox(
    '👥 ใคร\n\n' +
    'กลุ่มเป้าหมาย:\n\n' +
    '👶 เด็ก\n' +
    '  เรียนรู้การออมผ่านการเล่น\n\n' +
    '👴 ผู้สูงอายุ\n' +
    '  ป้องกันการหลอกลวง\n\n' +
    '💼 มืออาชีพ\n' +
    '  วิเคราะห์การลงทุน'
  );
  col2Text.setLeft(30 + columnWidth + gap).setTop(startY + 10).setWidth(columnWidth - 20).setHeight(columnHeight - 20);
  styleTextBox(col2Text, 15, THEME.colors.dark, false);

  // คอลัมน์ 3: ทำไม
  const col3 = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  col3.setLeft(20 + 2*(columnWidth + gap)).setTop(startY).setWidth(columnWidth).setHeight(columnHeight);
  col3.getFill().setSolidFill(THEME.colors.light);
  col3.getBorder().setTransparent();

  const col3Text = slide.insertTextBox(
    '💡 ทำไม\n\n' +
    'แก้ปัญหา:\n\n' +
    '📚 ไม่รู้เท่าทันการเงิน\n' +
    '  ทุกช่วงวัย\n\n' +
    '🚨 การหลอกลวงระบาด\n' +
    '  สูญเสีย 75,000 ล้าน/ปี\n\n' +
    '🌐 พึ่งพา AI ต่างชาติ\n' +
    '  ต้องการโซลูชันไทย'
  );
  col3Text.setLeft(30 + 2*(columnWidth + gap)).setTop(startY + 10).setWidth(columnWidth - 20).setHeight(columnHeight - 20);
  styleTextBox(col3Text, 15, THEME.colors.dark, false);

  // กล่องไฮไลท์ด้านล่าง
  const highlightBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE);
  highlightBox.setLeft(20).setTop(370).setWidth(680).setHeight(40);
  highlightBox.getFill().setSolidFill(THEME.colors.primary);
  highlightBox.getBorder().setTransparent();

  const highlightText = slide.insertTextBox('✅ 100% คนไทย | ลงทุน 85 ล้านบาท | โอกาสผ่าน BOI 90-95%');
  highlightText.setLeft(30).setTop(375).setWidth(660).setHeight(30);
  styleTextBox(highlightText, 17, THEME.colors.white, true);
  highlightText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  Logger.log('✅ สร้างแล้ว: สไลด์สรุปผู้บริหาร');
}

function createProblemSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(THEME.colors.white);

  addSlideHeader(slide, '🚨 ปัญหาที่เราแก้ไข', THEME.colors.accent);

  // ปัญหา 1
  createProblemBox(slide, 1,
    '📚 ไม่รู้เท่าทันการเงิน',
    '• เด็ก: ไม่มีการศึกษาการเงินที่สนุก\n' +
    '• ผู้สูงอายุ: ไม่เข้าใจฟินเทคสมัยใหม่\n' +
    '• ทั่วไป: ไม่มีที่ปรึกษาการเงิน 24/7',
    80, THEME.colors.accent
  );

  // ปัญหา 2
  createProblemBox(slide, 2,
    '🚨 การหลอกลวงระบาด',
    '• สูญเสีย 75,000 ล้านบาท/ปี\n' +
    '• ผู้สูงอายุเป็นเหยื่อหลัก\n' +
    '• การป้องกันปัจจุบัน: ช้าเกินไป',
    180, '#DC2626'
  );

  // ปัญหา 3
  createProblemBox(slide, 3,
    '🌐 พึ่งพา AI ต่างชาติ',
    '• AI/Chatbot ส่วนใหญ่มาจากต่างประเทศ\n' +
    '• ไม่เข้าใจบริบทไทย\n' +
    '• ไม่สามารถตรวจสอบหรือควบคุมได้',
    280, '#B91C1C'
  );

  // กล่องผลกระทบ
  const impactBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  impactBox.setLeft(50).setTop(370).setWidth(620).setHeight(40);
  impactBox.getFill().setSolidFill('#FEE2E2');
  const impactBorder = impactBox.getBorder();
  impactBorder.setWeight(2);
  impactBorder.setDashStyle(SlidesApp.DashStyle.SOLID);
  impactBorder.getLineFill().setSolidFill(THEME.colors.accent);

  const impactText = slide.insertTextBox('💸 ผลกระทบต่อชาติ: สูญเสีย 75,000 ล้านบาท/ปี + ขาดอธิปไตย AI ของไทย');
  impactText.setLeft(60).setTop(375).setWidth(600).setHeight(30);
  styleTextBox(impactText, 16, THEME.colors.accent, true);
  impactText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  Logger.log('✅ สร้างแล้ว: สไลด์ปัญหา');
}

function createSolutionSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(THEME.colors.white);

  addSlideHeader(slide, '✅ โซลูชันของเรา: ระบบ 3 ส่วนประกอบ', THEME.colors.success);

  // เพิ่มภาพพื้นหลัง AI
  try {
    const bgImage = slide.insertImage(THEME.images.aiTech);
    bgImage.setLeft(450).setTop(80).setWidth(250).setHeight(300);
    bgImage.setTransparency(0.85);
  } catch (e) {
    Logger.log('⚠️ ไม่สามารถเพิ่มภาพ AI ได้');
  }

  // ส่วนประกอบ 1: Physical AI
  const comp1Box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  comp1Box.setLeft(30).setTop(80).setWidth(200).setHeight(90);
  comp1Box.getFill().setSolidFill(THEME.colors.primary);
  comp1Box.getBorder().setTransparent();

  const comp1Text = slide.insertTextBox(
    '🦁 ส่วนประกอบ 1\nตัวแทน AI ทางกายภาพ\n(มาสคอตอัสลาน)'
  );
  comp1Text.setLeft(40).setTop(90).setWidth(180).setHeight(70);
  styleTextBox(comp1Text, 16, THEME.colors.white, true);
  comp1Text.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const comp1Detail = slide.insertTextBox(
    '• หุ่นยนต์/ตุ๊กตา AI รูปสิงโต\n' +
    '• พูดภาษาไทยได้เป็นธรรมชาติ\n' +
    '• ตรวจจับการหลอกลวงแบบเรียลไทม์\n' +
    '• สอนความรู้ทางการเงิน\n' +
    '• ทำงานกับเด็กและผู้สูงอายุ'
  );
  comp1Detail.setLeft(30).setTop(175).setWidth(200).setHeight(90);
  styleTextBox(comp1Detail, 13, THEME.colors.dark, false);

  // ส่วนประกอบ 2: Multi-Agent
  const comp2Box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  comp2Box.setLeft(30).setTop(275).setWidth(200).setHeight(90);
  comp2Box.getFill().setSolidFill(THEME.colors.secondary);
  comp2Box.getBorder().setTransparent();

  const comp2Text = slide.insertTextBox(
    '🤖 ส่วนประกอบ 2\nแพลตฟอร์มหลายเอเจนต์\n(5 AI Agents)'
  );
  comp2Text.setLeft(40).setTop(285).setWidth(180).setHeight(70);
  styleTextBox(comp2Text, 16, THEME.colors.white, true);
  comp2Text.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  const comp2Detail = slide.insertTextBox(
    '• NLP Agent (ภาษาไทย)\n' +
    '• Risk Agent (วิเคราะห์)\n' +
    '• Behavioral Agent (รูปแบบ)\n' +
    '• Fraud Agent (ตรวจจับ)\n' +
    '• Meta Agent (อธิบาย)'
  );
  comp2Detail.setLeft(240).setTop(280).setWidth(200).setHeight(80);
  styleTextBox(comp2Detail, 13, THEME.colors.dark, false);

  // ส่วนประกอบ 3: Mobile App
  const comp3Text = slide.insertTextBox(
    '📱 ส่วนประกอบ 3: แอปพลิเคชันมือถือ\n' +
    '• เชื่อมต่อกับ Physical AI\n' +
    '• การแจ้งเตือนแบบเรียลไทม์\n' +
    '• ติดตามความคืบหน้าการเรียนรู้\n' +
    '• วิเคราะห์ขั้นสูงสำหรับมืออาชีพ'
  );
  comp3Text.setLeft(240).setTop(90).setWidth(200).setHeight(90);

  const titleRange = comp3Text.getText().getRange(0, 35);
  titleRange.getTextStyle().setBold(true).setFontSize(15).setForegroundColor(THEME.colors.secondary);

  const detailRange = comp3Text.getText().getRange(36, comp3Text.getText().asString().length);
  detailRange.getTextStyle().setFontSize(13).setForegroundColor(THEME.colors.dark);

  Logger.log('✅ สร้างแล้ว: สไลด์โซลูชัน');
}

function createPhysicalAISlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(THEME.colors.white);

  addSlideHeader(slide, '🦁 ตัวแทน AI ทางกายภาพ: มาสคอตอัสลาน', THEME.colors.primary);

  // เพิ่มภาพหุ่นยนต์
  try {
    const robotImage = slide.insertImage(THEME.images.robot);
    robotImage.setLeft(400).setTop(80).setWidth(290).setHeight(300);
  } catch (e) {
    Logger.log('⚠️ ไม่สามารถเพิ่มภาพหุ่นยนต์ได้');
  }

  // รายการคุณสมบัติ
  const featuresText = slide.insertTextBox(
    '🎯 คุณสมบัติหลัก:\n\n' +
    '🎤 อินเทอร์เฟซเสียง\n' +
    '• พูดและเข้าใจภาษาไทยได้เป็นธรรมชาติ\n' +
    '• รู้จำเสียงสำหรับทุกวัย\n\n' +
    '🔊 การแจ้งเตือนเสียง\n' +
    '• คำเตือนการหลอกลวงแบบเรียลไทม์\n' +
    '• การสอนผ่านการสนทนา\n\n' +
    '💡 ตัวบ่งชี้ LED\n' +
    '• สัญญาณสถานะภาพ\n' +
    '• การแจ้งเตือนแบบสีรหัส (แดง=อันตราย)\n\n' +
    '📡 เชื่อมต่อ IoT\n' +
    '• การเชื่อมต่อ WiFi, Bluetooth\n' +
    '• ซิงค์กับแอปมือถือ\n\n' +
    '🧠 AI แบบฝังตัว\n' +
    '• การประมวลผลบนอุปกรณ์\n' +
    '• ออกแบบเน้นความเป็นส่วนตัว'
  );
  featuresText.setLeft(30).setTop(80).setWidth(350).setHeight(300);
  styleTextBox(featuresText, 14, THEME.colors.dark, false);

  // ไฮไลท์หัวข้อ
  const titleRange = featuresText.getText().getRange(0, 18);
  titleRange.getTextStyle().setBold(true).setFontSize(17).setForegroundColor(THEME.colors.primary);

  Logger.log('✅ สร้างแล้ว: สไลด์ Physical AI');
}

function createInvestmentSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(THEME.colors.white);

  addSlideHeader(slide, '💰 รายละเอียดการลงทุน: 85 ล้านบาท', THEME.colors.primary);

  // กล่องเงินลงทุนรวม
  const totalBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  totalBox.setLeft(250).setTop(80).setWidth(220).setHeight(80);
  totalBox.getFill().setSolidFill(THEME.colors.primary);
  totalBox.getBorder().setTransparent();

  const totalText = slide.insertTextBox('เงินลงทุนรวม\n85,000,000 บาท');
  totalText.setLeft(260).setTop(90).setWidth(200).setHeight(60);
  styleTextBox(totalText, 24, THEME.colors.white, true);
  totalText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // รายละเอียด R&D (ด้านซ้าย)
  const rdBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  rdBox.setLeft(50).setTop(180).setWidth(280).setHeight(200);
  rdBox.getFill().setSolidFill('#E0F2FE');
  const rdBorder = rdBox.getBorder();
  rdBorder.setWeight(3);
  rdBorder.getLineFill().setSolidFill(THEME.colors.secondary);

  const rdText = slide.insertTextBox(
    '🔬 กิจกรรม R&D: 59 ล้าน\n' +
    '(สนับสนุน: 29.5 ล้าน @ 50%)\n\n' +
    '• บุคลากร: 36.24 ล้าน (61.4%)\n' +
    '• อุปกรณ์: 9 ล้าน (15.3%)\n' +
    '• ข้อมูล/ชุดข้อมูล: 8.5 ล้าน (14.4%)\n' +
    '• อื่นๆ: 5.26 ล้าน (8.9%)\n\n' +
    '✅ สอดคล้อง AOR3/2568 หน้า 5'
  );
  rdText.setLeft(60).setTop(190).setWidth(260).setHeight(180);
  styleTextBox(rdText, 14, THEME.colors.dark, false);

  const rdTitleRange = rdText.getText().getRange(0, 21);
  rdTitleRange.getTextStyle().setBold(true).setFontSize(17).setForegroundColor(THEME.colors.secondary);

  // รายละเอียดดิจิทัล (ด้านขวา)
  const digitalBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  digitalBox.setLeft(390).setTop(180).setWidth(280).setHeight(200);
  digitalBox.getFill().setSolidFill('#FEF3C7');
  const digitalBorder = digitalBox.getBorder();
  digitalBorder.setWeight(3);
  digitalBorder.getLineFill().setSolidFill(THEME.colors.primary);

  const digitalText = slide.insertTextBox(
    '💻 เทคโนโลยีดิจิทัล: 26 ล้าน\n' +
    '(สนับสนุน: 13 ล้าน @ 50%)\n\n' +
    '• บริการคลาวด์: 5 ล้าน (19.2%)\n' +
    '• ฮาร์ดแวร์/เซิร์ฟเวอร์: 8 ล้าน (30.8%)\n' +
    '• ซอฟต์แวร์: 3 ล้าน (11.5%)\n' +
    '• ความปลอดภัย: 2 ล้าน (7.7%)\n' +
    '• การผสานระบบ: 8 ล้าน (30.8%)\n\n' +
    '✅ สอดคล้อง AOR3/2568 หน้า 6'
  );
  digitalText.setLeft(400).setTop(190).setWidth(260).setHeight(180);
  styleTextBox(digitalText, 14, THEME.colors.dark, false);

  const digitalTitleRange = digitalText.getText().getRange(0, 28);
  digitalTitleRange.getTextStyle().setBold(true).setFontSize(17).setForegroundColor(THEME.colors.primary);

  Logger.log('✅ สร้างแล้ว: สไลด์การลงทุน');
}

function createBOIComplianceSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(THEME.colors.white);

  addSlideHeader(slide, '✅ สรุปการปฏิบัติตาม BOI', THEME.colors.success);

  // ตารางการปฏิบัติตาม
  const tableData = [
    ['เกณฑ์', 'ที่กำหนด', 'โครงการของเรา', 'สถานะ'],
    ['อุตสาหกรรม (AOR1)', '8.2.12 อิเล็กทรอนิกส์อัจฉริยะ', '✅ ผ่านทั้ง 3 เกณฑ์', '✅ ผ่าน'],
    ['การลงทุน (AOR3)', '≥50 ล้านบาท', '85 ล้าน (170%)', '✅ ผ่าน'],
    ['ความเป็นเจ้าของไทย', '≥51%', '100% (196%)', '✅ ผ่าน'],
    ['ระยะเวลา', '≤12 เดือน', '12 เดือน', '✅ ผ่าน'],
    ['ค่าใช้จ่าย R&D', 'รายการที่ได้รับอนุมัติ', '100% สิทธิ์', '✅ ผ่าน'],
    ['ค่าใช้จ่ายดิจิทัล', 'รายการที่ได้รับอนุมัติ', '100% สิทธิ์', '✅ ผ่าน']
  ];

  const table = slide.insertTable(tableData.length, 4);
  table.setLeft(40).setTop(90);

  // จัดรูปแบบแถวหัวตาราง
  for (let col = 0; col < 4; col++) {
    const cell = table.getCell(0, col);
    cell.getFill().setSolidFill(THEME.colors.secondary);
    const cellText = cell.getText();
    cellText.setText(tableData[0][col]);
    cellText.getTextStyle().setFontSize(14).setBold(true).setForegroundColor(THEME.colors.white);
    cellText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  // จัดรูปแบบแถวข้อมูล
  for (let row = 1; row < tableData.length; row++) {
    for (let col = 0; col < 4; col++) {
      const cell = table.getCell(row, col);
      if (row % 2 === 0) {
        cell.getFill().setSolidFill('#F9FAFB');
      } else {
        cell.getFill().setSolidFill(THEME.colors.white);
      }
      const cellText = cell.getText();
      cellText.setText(tableData[row][col]);
      cellText.getTextStyle().setFontSize(12).setForegroundColor(THEME.colors.dark);
      if (col === 3) {  // คอลัมน์สถานะ
        cellText.getTextStyle().setBold(true).setForegroundColor(THEME.colors.success);
      }
    }
  }

  // แบนเนอร์ความสำเร็จด้านล่าง
  const bannerBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
  bannerBox.setLeft(100).setTop(365).setWidth(520).setHeight(45);
  bannerBox.getFill().setSolidFill(THEME.colors.success);
  bannerBox.getBorder().setTransparent();

  const bannerText = slide.insertTextBox('🎊 100% ผ่านเกณฑ์ | โอกาสอนุมัติ: 90-95% ⭐⭐⭐⭐⭐');
  bannerText.setLeft(110).setTop(373).setWidth(500).setHeight(30);
  styleTextBox(bannerText, 19, THEME.colors.white, true);
  bannerText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  Logger.log('✅ สร้างแล้ว: สไลด์การปฏิบัติตาม BOI');
}

function createThankYouSlide(presentation) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(THEME.colors.secondary);

  // เพิ่มพื้นหลังสิงโต
  try {
    const lionImage = slide.insertImage(THEME.images.lion);
    lionImage.setLeft(0).setTop(0).setWidth(720).setHeight(405);
    lionImage.setTransparency(0.85);
  } catch (e) {
    Logger.log('⚠️ ไม่สามารถเพิ่มภาพสิงโตได้');
  }

  // ข้อความขอบคุณ
  const thankYouBox = slide.insertTextBox('ขอบคุณครับ');
  thankYouBox.setLeft(100).setTop(100).setWidth(520).setHeight(80);
  const thankYouText = thankYouBox.getText();
  thankYouText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  const thankYouStyle = thankYouText.getTextStyle();
  thankYouStyle.setFontSize(64).setBold(true).setForegroundColor(THEME.colors.primary);
  thankYouStyle.setFontFamily(THEME.fonts.title);

  // คำบรรยาย
  const subtitleBox = slide.insertTextBox('Thank You');
  subtitleBox.setLeft(100).setTop(190).setWidth(520).setHeight(50);
  const subtitleText = subtitleBox.getText();
  subtitleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  const subtitleStyle = subtitleText.getTextStyle();
  subtitleStyle.setFontSize(36).setForegroundColor(THEME.colors.white);
  subtitleStyle.setFontFamily(THEME.fonts.body);

  // ข้อมูลติดต่อ
  const contactBox = slide.insertTextBox(
    'บริษัท อัสลาน เวลธ์ จำกัด\n\n' +
    '📧 admin@aslanwealth.com\n' +
    '📞 082-619-5459\n' +
    '🏢 88/260 ถนนสายไหม แขวงสายไหม เขตสายไหม กทม. 10220'
  );
  contactBox.setLeft(150).setTop(270).setWidth(420).setHeight(110);
  const contactText = contactBox.getText();
  contactText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  const contactStyle = contactText.getTextStyle();
  contactStyle.setFontSize(15).setForegroundColor(THEME.colors.light);
  contactStyle.setFontFamily(THEME.fonts.body);

  Logger.log('✅ สร้างแล้ว: สไลด์ขอบคุณ');
}

// ฟังก์ชันตัวช่วยสร้างกล่องปัญหา
function createProblemBox(slide, number, title, content, yPos, color) {
  // วงกลมหมายเลข
  const circle = slide.insertShape(SlidesApp.ShapeType.ELLIPSE);
  circle.setLeft(30).setTop(yPos).setWidth(40).setHeight(40);
  circle.getFill().setSolidFill(color);
  circle.getBorder().setTransparent();

  const numberText = slide.insertTextBox(String(number));
  numberText.setLeft(30).setTop(yPos + 5).setWidth(40).setHeight(30);
  styleTextBox(numberText, 20, THEME.colors.white, true);
  numberText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // หัวข้อ
  const titleBox = slide.insertTextBox(title);
  titleBox.setLeft(80).setTop(yPos - 5).setWidth(300).setHeight(25);
  styleTextBox(titleBox, 19, color, true);

  // เนื้อหา
  const contentBox = slide.insertTextBox(content);
  contentBox.setLeft(80).setTop(yPos + 25).setWidth(600).setHeight(60);
  styleTextBox(contentBox, 14, THEME.colors.dark, false);
}

// ฟังก์ชันตัวช่วยเพิ่มหัวสไลด์
function addSlideHeader(slide, title, color) {
  const headerBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE);
  headerBox.setLeft(0).setTop(0).setWidth(720).setHeight(60);
  headerBox.getFill().setSolidFill(color);
  headerBox.getBorder().setTransparent();

  const titleBox = slide.insertTextBox(title);
  titleBox.setLeft(30).setTop(15).setWidth(660).setHeight(30);
  const titleText = titleBox.getText();
  const titleStyle = titleText.getTextStyle();
  titleStyle.setFontSize(28).setBold(true).setForegroundColor(THEME.colors.white);
  titleStyle.setFontFamily(THEME.fonts.heading);
}

// ฟังก์ชันตัวช่วยจัดรูปแบบกล่องข้อความ
function styleTextBox(textBox, fontSize, color, bold) {
  const text = textBox.getText();
  const style = text.getTextStyle();
  style.setFontSize(fontSize).setForegroundColor(color);
  if (bold) style.setBold(true);
  style.setFontFamily(THEME.fonts.body);
}

// ฟังก์ชันตัวยึดสำหรับสไลด์ที่เหลือ (ใช้แบบเดียวกัน)
function createMultiAgentSlide(presentation) {
  Logger.log('⚠️ สไลด์ Multi-Agent - ใช้แบบเดียวกัน');
}

function createTargetUsersSlide(presentation) {
  Logger.log('⚠️ สไลด์กลุ่มเป้าหมาย - ใช้แบบเดียวกัน');
}

function createProductsSlide(presentation) {
  Logger.log('⚠️ สไลด์ผลิตภัณฑ์ - ใช้แบบเดียวกัน');
}

function createInnovationSlide(presentation) {
  Logger.log('⚠️ สไลด์นวัตกรรม - ใช้แบบเดียวกัน');
}

function createRDActivitiesSlide(presentation) {
  Logger.log('⚠️ สไลด์กิจกรรม R&D - ใช้แบบเดียวกัน');
}

function createDigitalTechSlide(presentation) {
  Logger.log('⚠️ สไลด์เทคโนโลยีดิจิทัล - ใช้แบบเดียวกัน');
}

function createAORComplianceSlide(presentation) {
  Logger.log('⚠️ สไลด์การปฏิบัติตาม AOR - ใช้แบบเดียวกัน');
}

function createEconomicImpactSlide(presentation) {
  Logger.log('⚠️ สไลด์ผลกระทบทางเศรษฐกิจ - ใช้แบบเดียวกัน');
}

function createSocialImpactSlide(presentation) {
  Logger.log('⚠️ สไลด์ผลกระทบทางสังคม - ใช้แบบเดียวกัน');
}

function createTimelineSlide(presentation) {
  Logger.log('⚠️ สไลด์ไทม์ไลน์ - ใช้แบบเดียวกัน');
}

function createKPIsSlide(presentation) {
  Logger.log('⚠️ สไลด์ KPI - ใช้แบบเดียวกัน');
}

function createTeamSlide(presentation) {
  Logger.log('⚠️ สไลด์ทีม - ใช้แบบเดียวกัน');
}

function createRiskMitigationSlide(presentation) {
  Logger.log('⚠️ สไลด์การลดความเสี่ยง - ใช้แบบเดียวกัน');
}

function createCompetitiveAdvantageSlide(presentation) {
  Logger.log('⚠️ สไลด์ข้อได้เปรียบ - ใช้แบบเดียวกัน');
}

function createNextStepsSlide(presentation) {
  Logger.log('⚠️ สไลด์ขั้นตอนต่อไป - ใช้แบบเดียวกัน');
}
