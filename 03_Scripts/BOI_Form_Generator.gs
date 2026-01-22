/**
 * BOI Form Generator - Google Apps Script
 * สร้าง Google Doc ที่ตรงตามรูปแบบฟอร์ม BOI (form_support.1.pdf)
 *
 * บริษัท: อัสลาน เวลธ์ จำกัด (ASLAN WEALTH COMPANY LIMITED)
 * โครงการ: ASLAN AI GUARDIAN PLATFORM
 *
 * วิธีใช้งาน:
 * 1. เปิด Google Apps Script (script.google.com)
 * 2. วางโค้ดนี้ทั้งหมด
 * 3. รัน function createBOIForm()
 * 4. อนุญาตสิทธิ์การเข้าถึง
 * 5. ได้ Google Doc พร้อมใช้งาน
 *
 * @author Claude AI Assistant
 * @version 1.0
 * @date 21 มกราคม 2569
 */

// ====================================
// ข้อมูลบริษัทและโครงการ - แก้ไขได้ตามต้องการ
// ====================================
const COMPANY_DATA = {
  // ข้อมูลบริษัท
  nameTH: 'บริษัท อัสลาน เวลธ์ จำกัด',
  nameEN: 'ASLAN WEALTH COMPANY LIMITED',
  registrationNo: '0105567220030',
  registrationDate: '25 ตุลาคม 2567',
  registeredCapital: '1,000,000',
  paidUpCapital: '1,000,000',
  address: '88/260 ถนนสายไหม แขวงสายไหม เขตสายไหม กรุงเทพมหานคร 10220',
  phone: '082-619-5459',
  mobile: '082-619-5459',
  email: 'admin@aslanwealth.com',
  contactPerson: 'นายสุรเกียรติ ยาวะโนภาส',
  contactPosition: 'กรรมการผู้จัดการ',

  // ข้อมูลธนาคาร
  bankName: 'ธนาคารกรุงไทย',
  bankBranch: 'สายไหม',
  bankAccountNo: '',  // กรอกเลขบัญชีจริง
  bankAccountName: 'บริษัท อัสลาน เวลธ์ จำกัด',

  // ผู้ถือหุ้น
  shareholders: [
    {
      name: 'นายสุรเกียรติ ยาวะโนภาส',
      nationality: 'ไทย',
      idNo: '1-1014-01139-71-7',
      percentage: '50',
      amount: '500,000'
    },
    {
      name: 'นายณัฐพัทธ์ ปิยะมาลาไชย',
      nationality: 'ไทย',
      idNo: '1-5799-00656-77-0',
      percentage: '50',
      amount: '500,000'
    }
  ],
  thaiShareholderPercent: '100',
  foreignShareholderPercent: '0'
};

const PROJECT_DATA = {
  // ข้อมูลโครงการ
  nameEN: 'ASLAN AI GUARDIAN PLATFORM',
  nameTH: 'การวิจัยและพัฒนาระบบปัญญาประดิษฐ์หลายตัวแทนเพื่อการปกป้องทางการเงิน',
  industry: 'อุตสาหกรรมอิเล็กทรอนิกส์อัจฉริยะ',
  businessType: '8.2.12',
  businessDescription: 'กิจการพัฒนาซอฟต์แวร์ที่สร้างมูลค่าเพิ่มสูง (Smart Electronics)',

  // ที่ตั้งโครงการ
  location: {
    industrialEstate: '-',
    subDistrict: 'สายไหม',
    district: 'สายไหม',
    province: 'กรุงเทพมหานคร'
  },
  area: '428 ตารางเมตร',

  // ระยะเวลา
  startDate: '1 มิถุนายน 2569',
  endDate: '31 พฤษภาคม 2570',
  duration: '12 เดือน',

  // งบประมาณ
  totalInvestment: 85000000,
  rdInvestment: 59000000,
  digitalInvestment: 26000000,
  rdSupport: 29500000,
  digitalSupport: 13000000,
  totalSupport: 42500000,
  supportRate: 50
};

const PRODUCTS = [
  {
    name: 'Physical AI Agent (Aslan Mascot)',
    capacity: '10,000',
    unit: 'Units',
    hoursPerDay: '24',
    daysPerYear: '365'
  },
  {
    name: 'Thai Natural Language Processing Engine',
    capacity: '100,000',
    unit: 'Requests/day',
    hoursPerDay: '24',
    daysPerYear: '365'
  },
  {
    name: 'Multi-Agent AI Platform',
    capacity: '50,000',
    unit: 'Concurrent Users',
    hoursPerDay: '24',
    daysPerYear: '365'
  },
  {
    name: 'Explainable AI (XAI) Module',
    capacity: '100,000',
    unit: 'Analyses/day',
    hoursPerDay: '24',
    daysPerYear: '365'
  },
  {
    name: 'Mobile Application',
    capacity: '100,000',
    unit: 'Active Users',
    hoursPerDay: '24',
    daysPerYear: '365'
  }
];

const PROCESSES = [
  {
    step: 'AI Model Training & Multi-Agent Integration',
    equipment: 'GPU Servers (NVIDIA A100), Cloud Infrastructure'
  },
  {
    step: 'Deep Learning Training บน Thai Corpus',
    equipment: 'GPU Servers, NLP Development Tools'
  },
  {
    step: 'XAI Algorithm Development & Testing',
    equipment: 'Development Servers, XAI Frameworks'
  },
  {
    step: 'Pattern Mining & Machine Learning',
    equipment: 'Big Data Cluster, ML Infrastructure'
  },
  {
    step: 'System Integration & API Development',
    equipment: 'Production Servers, Kubernetes Cluster'
  }
];

const PERSONNEL = [
  {
    position: 'Project Manager',
    education: "Master's in CS/AI",
    thaiCount: 1,
    foreignCount: 0,
    totalCount: 1,
    annualSalary: '1,200,000'
  },
  {
    position: 'Researcher/Developer (AI Scientists)',
    education: 'PhD in AI/ML/CS',
    thaiCount: 2,
    foreignCount: 0,
    totalCount: 2,
    annualSalary: '3,600,000'
  },
  {
    position: 'Designer (Senior Engineers)',
    education: "Master's in CS/Engineering",
    thaiCount: 4,
    foreignCount: 0,
    totalCount: 4,
    annualSalary: '4,800,000'
  },
  {
    position: 'ผู้ช่วยนักวิจัย (AI/Software Dev)',
    education: "Bachelor's in CS",
    thaiCount: 8,
    foreignCount: 0,
    totalCount: 8,
    annualSalary: '14,400,000'
  }
];

const SECTION6_DATA = {
  directQuantitative: [
    'จำนวนการจ้างงานบุคลากรไทย 15 ราย',
    'จำนวนบุคลากรไทยที่ได้รับการถ่ายทอดเทคโนโลยีในโครงการ 15 ราย',
    'รายได้และกำไรของธุรกิจส่วนที่เพิ่มขึ้นจากการลงทุนตามโครงการ 150-200 ล้านบาท/ปี',
    'มูลค่าการจัดซื้อวัตถุดิบในประเทศ 65 ล้านบาท (คิดเป็น 76% ของเงินลงทุน)',
    'จำนวนธุรกิจ Supply Chain ที่ตามมา 5-10 ราย',
    'สร้างรายได้จากการขายสินค้าของโครงการประมาณ 2 ล้านบาท/ปี'
  ],
  directQualitative: [
    'พัฒนาทักษะและความรู้ของบุคลากรไทย 15 ราย ในสาขา AI, Machine Learning, NLP, Multi-Agent Systems และเทคโนโลยีดิจิทัลขั้นสูง ยกระดับขีดความสามารถของบุคลากรไทยให้สามารถแข่งขันในตลาดโลกได้',
    'ยกระดับคุณภาพผลิตภัณฑ์และบริการด้วยเทคโนโลยี AI ที่ทันสมัย สร้างมาตรฐานใหม่ในอุตสาหกรรม Smart Electronics และ FinTech ในประเทศไทย',
    'เพิ่มศักยภาพในการแข่งขันของธุรกิจในตลาดโลก ด้วยนวัตกรรม Physical AI Agent และ Multi-Agent Platform ที่เป็นครั้งแรกในอาเซียน',
    'สร้างความเข้มแข็งให้กับซัพพลายเชนและผู้ประกอบการไทย โดยใช้บริการและอุปกรณ์จากผู้ผลิตในประเทศ'
  ],
  indirectQuantitative: [
    'ลดความเสียหายจากการถูกหลอกลวงทางการเงิน (Scam losses) คาดว่าจะช่วยลดความเสียหายได้ประมาณ 500-1,000 ล้านบาท/ปี',
    'สร้างงานวิจัยคุณภาพสูง สิทธิบัตร ≥2 ฉบับ และบทความวิชาการ ≥3 ฉบับ',
    'เป้าหมายผู้ใช้บริการปีแรก 55,000 ราย (20,000 ครอบครัว + 30,000 ผู้สูงอายุ + 5,000 มืออาชีพ)',
    'Thai Scam Corpus Dataset ≥10,000 ตัวอย่าง เปิดให้ชุมชนวิจัยในประเทศใช้ฟรี'
  ],
  indirectQualitative: [
    'เพิ่มความรู้ทางการเงินให้กับเด็กและเยาวชนไทย (Financial Literacy) ผ่าน Physical AI Agent ที่เป็นมิตร',
    'ปกป้องผู้สูงอายุและกลุ่มเปราะบางจากการถูกหลอกลวงทางการเงิน ด้วยระบบ AI ที่ตรวจจับและเตือนภัยแบบ Real-time',
    'สร้างนวัตกรรม AI ที่เป็นครั้งแรกในประเทศไทย และภูมิภาคอาเซียน ในสาขา Multi-Agent AI for Financial Protection',
    'พัฒนาและเผยแพร่ฐานข้อมูล Thai Language Corpus สำหรับวิจัย NLP เปิดให้ชุมชนวิจัยไทยและนานาชาติใช้ฟรี',
    'สร้างมาตรฐานใหม่ด้าน Explainable AI (XAI) ที่โปร่งใสและอธิบายได้ตามหลักการ Digital Governance'
  ]
};

const KPI_PHASE1 = [
  'Working Prototype with 3 AI Agents ทำงานร่วมกันได้',
  'Thai NLP Model Training: F1 Score ≥80%',
  'Multi-Agent Communication Framework: Message passing ≥99%',
  'Integration Testing เสร็จสมบูรณ์',
  'Technical Documentation V1 ครบถ้วน ≥80%',
  'R&D Budget Utilization: ใช้งบ 50%, ความคลาดเคลื่อน <15%',
  'Personnel Development Progress: 15 คน เข้าฝึกอบรม ≥80%'
];

const KPI_PHASE2 = [
  'Production System: ครบ 5 AI Agents, Uptime ≥99.5%',
  'Model Accuracy: ≥85% F1 Score, False Positive ≤10%',
  'Response Time: <200ms, Throughput ≥500 TPS',
  'ยื่นสิทธิบัตร ≥2 ฉบับ',
  'ตีพิมพ์วิชาการ ≥3 ฉบับ',
  'Personnel Development: ฝึกอบรม ≥15 คน, อัตราสำเร็จ ≥80%',
  'Thai Scam Dataset ≥10,000 ตัวอย่าง, คุณภาพ ≥75%',
  'UAT กับ ≥50 คน, ความพึงพอใจ ≥3.5/5',
  'Complete Platform Delivery ส่งมอบครบทุกส่วน',
  'Digital Technology Certification จาก depa'
];

// ====================================
// MAIN FUNCTION - สร้างฟอร์ม BOI
// ====================================
function createBOIForm() {
  // สร้าง Google Doc ใหม่
  const doc = DocumentApp.create('BOI_Form_ASLAN_WEALTH_' + new Date().toISOString().split('T')[0]);
  const body = doc.getBody();

  // ตั้งค่า Font และ Margin
  body.setMarginTop(72);  // 1 inch = 72 points
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);

  // สร้างฟอร์มทั้งหมด
  createPage1(body);  // หน้า 1/7: ผู้ขอรับการส่งเสริม
  createPage2(body);  // หน้า 2/7: โครงสร้างผู้ถือหุ้น
  createPage3(body);  // หน้า 3/7: รายละเอียดโครงการปัจจุบัน
  createPage4(body);  // หน้า 4/7: รายละเอียดโครงการภายหลังเปลี่ยนผ่าน
  createPage5(body);  // หน้า 5/7: การดำเนินการเพื่อเพิ่มขีดความสามารถ
  createPage6(body);  // หน้า 6/7: การตรวจสอบการดำเนินการ
  createPage7(body);  // หน้า 7/7: ตัวชี้วัดและผลกระทบเชิงบวก
  createAttachmentA(body);  // เอกสารแนบ A: R&D
  createAttachmentC(body);  // เอกสารแนบ C: Digital Technology

  // บันทึกและแสดงผล
  doc.saveAndClose();

  Logger.log('สร้างฟอร์ม BOI สำเร็จ!');
  Logger.log('URL: ' + doc.getUrl());

  return doc.getUrl();
}

// ====================================
// หน้า 1/7: ผู้ขอรับการส่งเสริม
// ====================================
function createPage1(body) {
  // หัวเรื่อง
  addPageHeader(body, 'หน้า 1/7');

  const title = body.appendParagraph('ข้อเสนอโครงการ');
  setTitleStyle(title);

  const subtitle1 = body.appendParagraph('ภายใต้มาตรการสนับสนุนผู้ประกอบการไทยเพื่อเพิ่มขีดความสามารถในการแข่งขัน');
  setSubtitleStyle(subtitle1);

  const subtitle2 = body.appendParagraph('ตามพระราชบัญญัติการเพิ่มขีดความสามารถในการแข่งขันของประเทศ');
  setSubtitleStyle(subtitle2);

  const subtitle3 = body.appendParagraph('สำหรับอุตสาหกรรมเป้าหมาย พ.ศ. 2560');
  setSubtitleStyle(subtitle3);

  body.appendParagraph('');

  // ข้อ 1
  const section1 = body.appendParagraph('1. ผู้ขอรับการส่งเสริม');
  setSectionHeaderStyle(section1);

  // ข้อมูลบริษัท
  addFormField(body, 'ชื่อนิติบุคคล', COMPANY_DATA.nameTH + ' (' + COMPANY_DATA.nameEN + ')');
  addFormField(body, 'เลขทะเบียนนิติบุคคล', COMPANY_DATA.registrationNo);
  addFormField(body, 'วันที่จดทะเบียน', COMPANY_DATA.registrationDate);
  addFormField(body, 'ทุนจดทะเบียน', COMPANY_DATA.registeredCapital + ' บาท');
  addFormField(body, 'เรียกชำระแล้ว', COMPANY_DATA.paidUpCapital + ' บาท (100%)');
  addFormField(body, 'สำนักงานตั้งอยู่ที่', COMPANY_DATA.address);
  addFormField(body, 'โทรศัพท์', COMPANY_DATA.phone);
  addFormField(body, 'มือถือ', COMPANY_DATA.mobile);
  addFormField(body, 'E-mail', COMPANY_DATA.email);
  addFormField(body, 'ชื่อตัวแทนในการติดต่อ', COMPANY_DATA.contactPerson);
  addFormField(body, 'ตำแหน่ง', COMPANY_DATA.contactPosition);
  addFormField(body, 'ที่อยู่ในการรับเอกสาร', COMPANY_DATA.address);

  body.appendParagraph('');

  // ข้อมูลธนาคาร
  const bankInfo = body.appendParagraph('ขอรับเงินสนับสนุนจากกองทุนเพิ่มขีดความสามารถในการแข่งขันของประเทศ โดยขอให้สำนักงานดำเนินการโอนเงินสนับสนุน (หากได้รับอนุมัติ) เข้าบัญชีของบริษัทฯ ดังนี้');
  setBodyStyle(bankInfo);

  addFormField(body, 'บัญชี', COMPANY_DATA.bankName + ' สาขา ' + COMPANY_DATA.bankBranch);
  addFormField(body, 'เลขที่บัญชี', COMPANY_DATA.bankAccountNo || '(กรุณากรอกเลขบัญชี)');
  addFormField(body, 'ชื่อบัญชี', COMPANY_DATA.bankAccountName);

  body.appendParagraph('');

  // คำรับรอง
  const declarations = [
    '☑ ข้าพเจ้าขอรับรองว่าข้อมูลที่จะให้ดังต่อไปนี้ตรงกับความเป็นจริง หรือเป็นประมาณการที่ดีที่สุดในความเห็นของข้าพเจ้า ทั้งนี้ ข้าพเจ้าจะไม่เปิดเผยข้อมูลในเอกสารนี้ หรือข้อมูลใด ๆ ที่เกี่ยวข้องกับกระบวนการส่งเสริมการลงทุนแก่บุคคลภายนอก โดยมิได้รับอนุญาตจากสำนักงานคณะกรรมการส่งเสริมการลงทุนอย่างเด็ดขาด',
    '☑ ข้าพเจ้าขอรับรองว่าไม่ได้รับเงินสนับสนุนจากหน่วยงานภาครัฐอื่น ๆ สำหรับโครงการที่ขอรับการส่งเสริมนี้ ตลอดช่วงระยะเวลาที่ได้รับเงินสนับสนุนจากกองทุนเพิ่มขีดความสามารถในการแข่งขันของประเทศสำหรับอุตสาหกรรมเป้าหมาย',
    '☑ ข้าพเจ้ารับทราบว่าจะไม่นำเงินลงทุนและ/หรือค่าใช้จ่ายเพื่อเพิ่มขีดความสามารถในการแข่งขันภายใต้มาตรการนี้ไปขอรับสิทธิจากหน่วยงานภาครัฐอื่น และจากคณะกรรมการส่งเสริมการลงทุน',
    '☑ ข้าพเจ้ายินยอมให้สำนักงานเปิดเผยข้อมูลและเงื่อนไขการให้สิทธิและประโยชน์เท่าที่จำเป็นแก่หน่วยงานที่เกี่ยวข้อง'
  ];

  declarations.forEach(function(decl) {
    const p = body.appendParagraph(decl);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // ลายเซ็น
  addSignatureBlock(body, 'ผู้ขอรับการส่งเสริม', COMPANY_DATA.contactPerson);

  addPageBreak(body);
}

// ====================================
// หน้า 2/7: โครงสร้างผู้ถือหุ้น
// ====================================
function createPage2(body) {
  addPageHeader(body, 'หน้า 2/7');

  const section2 = body.appendParagraph('2. โครงสร้างผู้ถือหุ้นและประเภทนิติบุคคล');
  setSectionHeaderStyle(section2);

  // 2.1 สัดส่วนการถือหุ้น
  const sub21 = body.appendParagraph('2.1 สัดส่วนการถือหุ้นตามสัญชาติ');
  setSubSectionStyle(sub21);

  const shareholderTable = body.appendTable([
    ['สัญชาติผู้ถือหุ้น', 'สัดส่วน (ร้อยละ)'],
    ['บุคคลหรือนิติบุคคลสัญชาติไทย', COMPANY_DATA.thaiShareholderPercent],
    ['บุคคลหรือนิติบุคคลต่างด้าว', COMPANY_DATA.foreignShareholderPercent]
  ]);
  setTableStyle(shareholderTable);

  body.appendParagraph('');

  // 2.2 รายชื่อผู้ถือหุ้น
  const sub22 = body.appendParagraph('2.2 รายชื่อผู้ถือหุ้น/ผู้ร่วมลงทุน เรียงลำดับการถือหุ้นจากมากไปน้อย');
  setSubSectionStyle(sub22);

  const shareholderListData = [['ลำดับ', 'ชื่อผู้ถือหุ้น/ผู้ร่วมลงทุน', 'สัญชาติ', 'เลขบัตรประชาชน', 'สัดส่วน (%)', 'มูลค่า (บาท)']];
  COMPANY_DATA.shareholders.forEach(function(sh, index) {
    shareholderListData.push([
      (index + 1).toString(),
      sh.name,
      sh.nationality,
      sh.idNo,
      sh.percentage,
      sh.amount
    ]);
  });

  const shareholderListTable = body.appendTable(shareholderListData);
  setTableStyle(shareholderListTable);

  body.appendParagraph('');

  // 2.3 ประเภทนิติบุคคล
  const sub23 = body.appendParagraph('2.3 ประเภทนิติบุคคล');
  setSubSectionStyle(sub23);

  const entityType = body.appendParagraph('☑ กรณีอื่น ๆ (ต้องมีเงินลงทุนไม่น้อยกว่า 50 ล้านบาท โดยไม่รวมค่าที่ดินและทุนหมุนเวียน)');
  setBodyStyle(entityType);

  body.appendParagraph('');

  // 3. ข้อมูลโครงการ
  const section3 = body.appendParagraph('3. ข้อมูลโครงการที่ขอรับการส่งเสริมตามมาตรการนี้');
  setSectionHeaderStyle(section3);

  // 3.1 อุตสาหกรรมเป้าหมาย
  const sub31 = body.appendParagraph('3.1 อุตสาหกรรมเป้าหมาย');
  setSubSectionStyle(sub31);

  const industryNote = body.appendParagraph('(อ้างอิงจากประกาศคณะกรรมการนโยบายเพิ่มขีดความสามารถในการแข่งขันของประเทศสำหรับอุตสาหกรรมเป้าหมาย ที่ 3/2568 ลงวันที่ 19 พฤศจิกายน 2568)');
  industryNote.setFontSize(10);
  industryNote.setItalic(true);

  const industrySelection = body.appendParagraph('☑ ' + PROJECT_DATA.industry);
  setBodyStyle(industrySelection);
  industrySelection.setBold(true);

  body.appendParagraph('');

  // 3.2 ประเภทกิจการ
  const sub32 = body.appendParagraph('3.2 ประเภทกิจการที่ขอรับการส่งเสริม');
  setSubSectionStyle(sub32);

  const businessTypeNote = body.appendParagraph('(อ้างอิงจากประกาศสำนักงานคณะกรรมการส่งเสริมการลงทุน ที่ อ.1/2568 ลงวันที่ 26 พฤศจิกายน 2568)');
  businessTypeNote.setFontSize(10);
  businessTypeNote.setItalic(true);

  const businessTable = body.appendTable([
    ['ประเภทกิจการ', 'กิจการ'],
    [PROJECT_DATA.businessType, PROJECT_DATA.businessDescription]
  ]);
  setTableStyle(businessTable);

  addPageBreak(body);
}

// ====================================
// หน้า 3/7: รายละเอียดโครงการปัจจุบัน
// ====================================
function createPage3(body) {
  addPageHeader(body, 'หน้า 3/7');

  // 3.3 สถานภาพโครงการ
  const sub33 = body.appendParagraph('3.3 สถานภาพโครงการ');
  setSubSectionStyle(sub33);

  const statusSelection = body.appendParagraph('☑ ไม่ได้รับการส่งเสริม');
  setBodyStyle(statusSelection);

  body.appendParagraph('');

  // 3.4 รายละเอียดโครงการในปัจจุบัน
  const sub34 = body.appendParagraph('3.4 รายละเอียดโครงการในปัจจุบัน');
  setSubSectionStyle(sub34);

  // 1) ผลิตภัณฑ์และกำลังการผลิต
  const productHeader = body.appendParagraph('1) ผลิตภัณฑ์และกำลังการผลิต');
  setBodyStyle(productHeader);
  productHeader.setBold(true);

  const productTableData = [['ผลิตภัณฑ์', 'กำลังการผลิตเต็มที่ต่อปี', 'หน่วย', 'ชั่วโมง/วัน', 'วัน/ปี']];
  PRODUCTS.forEach(function(p) {
    productTableData.push([p.name, p.capacity, p.unit, p.hoursPerDay, p.daysPerYear]);
  });

  const productTable = body.appendTable(productTableData);
  setTableStyle(productTable);

  body.appendParagraph('');

  // 2) รายละเอียดของผลิตภัณฑ์
  const productDetailHeader = body.appendParagraph('2) รายละเอียดของผลิตภัณฑ์');
  setBodyStyle(productDetailHeader);
  productDetailHeader.setBold(true);

  const productDetails = [
    '1. Physical AI Agent (Aslan Mascot): หุ่นยนต์ AI รูปสิงโต "Aslan" พร้อมระบบเซนเซอร์อิเล็กทรอนิกส์ (Microphone, Speaker, LED, Touch sensors) และระบบ Wireless connectivity (WiFi, Bluetooth) สำหรับสอน Financial Literacy และป้องกัน Scam',
    '2. Thai Natural Language Processing Engine: โมเดล Deep Learning สำหรับประมวลผลภาษาไทยในบริบททางการเงิน พร้อม Thai Scam Corpus Dataset',
    '3. Multi-Agent AI Platform: ระบบ 5 AI Agents ทำงานร่วมกันแบบ collaborative และ autonomous สำหรับ Financial Protection',
    '4. Explainable AI (XAI) Module: โมดูลอธิบายการตัดสินใจของ AI เป็นภาษาไทยที่เข้าใจง่าย ตามหลักการ Digital Governance',
    '5. Mobile Application: แอปพลิเคชันมือถือ iOS/Android เชื่อมต่อกับ Physical AI Agent พร้อม Dashboard และ Learning Center'
  ];

  productDetails.forEach(function(detail) {
    const p = body.appendParagraph(detail);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // 3) กรรมวิธีการผลิต
  const processHeader = body.appendParagraph('3) กรรมวิธีการผลิต');
  setBodyStyle(processHeader);
  processHeader.setBold(true);

  const processTableData = [['ขั้นตอนการผลิต', 'เครื่องจักรหลัก']];
  PROCESSES.forEach(function(p, index) {
    processTableData.push([(index + 1) + '. ' + p.step, p.equipment]);
  });

  const processTable = body.appendTable(processTableData);
  setTableStyle(processTable);

  addPageBreak(body);
}

// ====================================
// หน้า 4/7: รายละเอียดโครงการภายหลังเปลี่ยนผ่าน
// ====================================
function createPage4(body) {
  addPageHeader(body, 'หน้า 4/7');

  // 3.5 (ไม่ใช้ - เฉพาะกรณีเปลี่ยนผ่าน)
  const sub35 = body.appendParagraph('3.5 รายละเอียดโครงการภายหลังการเปลี่ยนผ่านสู่อุตสาหกรรมใหม่');
  setSubSectionStyle(sub35);

  const notApplicable = body.appendParagraph('(ไม่ใช้ - โครงการนี้ไม่ใช่กรณีเปลี่ยนผ่านไปสู่อุตสาหกรรมใหม่)');
  setBodyStyle(notApplicable);
  notApplicable.setItalic(true);

  body.appendParagraph('');

  // 3.6 ที่ตั้งโครงการ
  const sub36 = body.appendParagraph('3.6 ที่ตั้งโครงการ');
  setSubSectionStyle(sub36);

  const locationTable = body.appendTable([
    ['ลำดับที่', 'นิคม/เขตอุตสาหกรรม', 'แขวง/ตำบล', 'เขต/อำเภอ', 'จังหวัด'],
    ['1', PROJECT_DATA.location.industrialEstate, PROJECT_DATA.location.subDistrict, PROJECT_DATA.location.district, PROJECT_DATA.location.province]
  ]);
  setTableStyle(locationTable);

  body.appendParagraph('');

  addFormField(body, 'พื้นที่โครงการ', PROJECT_DATA.area);
  addFormField(body, 'วันเริ่มโครงการ', PROJECT_DATA.startDate);
  addFormField(body, 'วันสิ้นสุดโครงการ', PROJECT_DATA.endDate);
  addFormField(body, 'ระยะเวลา', PROJECT_DATA.duration);

  body.appendParagraph('');

  // ประเภทการส่งเสริม
  const promotionType = body.appendParagraph('ประเภทการส่งเสริมที่ขอรับ:');
  setBodyStyle(promotionType);
  promotionType.setBold(true);

  const type1 = body.appendParagraph('☑ กิจกรรมวิจัยและพัฒนา (R&D) → เอกสารแนบ A');
  setBodyStyle(type1);

  const type2 = body.appendParagraph('☑ การใช้เทคโนโลยีดิจิทัล → เอกสารแนบ C');
  setBodyStyle(type2);

  addPageBreak(body);
}

// ====================================
// หน้า 5/7: การดำเนินการเพื่อเพิ่มขีดความสามารถ
// ====================================
function createPage5(body) {
  addPageHeader(body, 'หน้า 5/7');

  // 3.7 การดำเนินการเพื่อเพิ่มขีดความสามารถ
  const sub37 = body.appendParagraph('3.7 การดำเนินการเพื่อเพิ่มขีดความสามารถในการแข่งขันและประมาณการลงทุน');
  setSubSectionStyle(sub37);

  const investmentTable = body.appendTable([
    ['ขอบข่ายการดำเนินการ', 'ประมาณการลงทุน (ล้านบาท)', '', 'สัดส่วนการสนับสนุน (%)', 'มูลค่าการสนับสนุน (ล้านบาท)'],
    ['', 'ในประเทศ', 'ต่างประเทศ', '', ''],
    ['1. ด้านการวิจัยและพัฒนา (เอกสารแนบ A)', '59', '-', '50', '29.5'],
    ['2. ด้านการใช้เทคโนโลยีดิจิทัล (เอกสารแนบ C)', '26', '-', '50', '13'],
    ['รวม', '85', '-', '-', '42.5']
  ]);
  setTableStyle(investmentTable);

  body.appendParagraph('');

  // รายละเอียดงบประมาณ R&D
  const rdBudgetHeader = body.appendParagraph('รายละเอียดงบประมาณ R&D (59 ล้านบาท):');
  setBodyStyle(rdBudgetHeader);
  rdBudgetHeader.setBold(true);

  const rdBudgetTable = body.appendTable([
    ['หมวด', 'งบประมาณ (บาท)', '%'],
    ['บุคลากรวิจัย (Personnel)', '36,240,000', '61.4%'],
    ['อุปกรณ์และเครื่องมือ (Equipment)', '9,000,000', '15.3%'],
    ['ข้อมูลและ Dataset', '8,500,000', '14.4%'],
    ['อื่นๆ', '5,260,000', '8.9%'],
    ['รวม', '59,000,000', '100%']
  ]);
  setTableStyle(rdBudgetTable);

  body.appendParagraph('');

  // รายละเอียดงบประมาณ Digital
  const digitalBudgetHeader = body.appendParagraph('รายละเอียดงบประมาณเทคโนโลยีดิจิทัล (26 ล้านบาท):');
  setBodyStyle(digitalBudgetHeader);
  digitalBudgetHeader.setBold(true);

  const digitalBudgetTable = body.appendTable([
    ['หมวด', 'งบประมาณ (บาท)', '%'],
    ['Cloud Services & Infrastructure', '5,000,000', '19.2%'],
    ['On-premise Hardware & Servers', '8,000,000', '30.8%'],
    ['Software Licenses & Development Tools', '3,000,000', '11.5%'],
    ['Security & Compliance Systems', '2,000,000', '7.7%'],
    ['System Integration & API Development', '8,000,000', '30.8%'],
    ['รวม', '26,000,000', '100%']
  ]);
  setTableStyle(digitalBudgetTable);

  addPageBreak(body);
}

// ====================================
// หน้า 6/7: การตรวจสอบการดำเนินการ
// ====================================
function createPage6(body) {
  addPageHeader(body, 'หน้า 6/7');

  // 4. การตรวจสอบการดำเนินการ
  const section4 = body.appendParagraph('4. การตรวจสอบการดำเนินการตามโครงการ');
  setSectionHeaderStyle(section4);

  // 4.1 หน่วยงานตรวจสอบ
  const sub41 = body.appendParagraph('4.1 กำหนดหน่วยงานตรวจสอบตามขอบข่ายดังนี้');
  setSubSectionStyle(sub41);

  // หน่วยงาน R&D
  const rdAgency = body.appendParagraph('1) ด้านการวิจัยและพัฒนา');
  setBodyStyle(rdAgency);
  rdAgency.setBold(true);

  const nstda = [
    '▪ สำนักงานพัฒนาวิทยาศาสตร์และเทคโนโลยีแห่งชาติ (สวทช.)',
    '  ที่อยู่: เลขที่ 111 อุทยานวิทยาศาสตร์ประเทศไทย ถนนพหลโยธิน ตำบลคลองหนึ่ง อำเภอคลองหลวง จังหวัดปทุมธานี 12120',
    '  โทร: 0 2564 7000 ต่อ 1334-1338',
    '  อีเมล: ipd-psr@nstda.or.th'
  ];
  nstda.forEach(function(line) {
    const p = body.appendParagraph(line);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // หน่วยงาน Digital
  const digitalAgency = body.appendParagraph('3) ด้านการใช้เทคโนโลยีดิจิทัล');
  setBodyStyle(digitalAgency);
  digitalAgency.setBold(true);

  const depa = [
    '▪ สำนักงานส่งเสริมเศรษฐกิจดิจิทัล (depa)',
    '  ที่อยู่: เลขที่ 234/431 ซอยลาดพร้าว 10 ถนนลาดพร้าว แขวงจอมพล เขตจตุจักร กรุงเทพมหานคร 10900',
    '  โทร: 02 026 2333 ต่อ 1000 หรือ ต่อ 1073',
    '  อีเมล: ddi-ip@depa.or.th'
  ];
  depa.forEach(function(line) {
    const p = body.appendParagraph(line);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // 4.2 การประสานงาน
  const sub42 = body.appendParagraph('4.2 การประสานงานหน่วยงานตรวจสอบ');
  setSubSectionStyle(sub42);

  const coordination = body.appendParagraph('☑ ผู้ขอรับการส่งเสริมต้องประสานและนัดหมายหน่วยงานตรวจสอบในข้อ 4.1 ตามขอบข่ายที่ยื่นขอรับการส่งเสริม เพื่อดำเนินการตรวจสอบให้แล้วเสร็จก่อนครบกำหนดในแต่ละงวดงาน');
  setBodyStyle(coordination);

  addPageBreak(body);
}

// ====================================
// หน้า 7/7: ตัวชี้วัดและผลกระทบเชิงบวก
// ====================================
function createPage7(body) {
  addPageHeader(body, 'หน้า 7/7');

  // 5. ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด
  const section5 = body.appendParagraph('5. ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด');
  setSectionHeaderStyle(section5);

  // งวดที่ 1
  const phase1Header = body.appendParagraph('งวดที่ 1 (ภายใน 6 เดือน นับจากวันที่ออกบัตรส่งเสริม)');
  setBodyStyle(phase1Header);
  phase1Header.setBold(true);

  addFormField(body, 'ประมาณการลงทุน', '42.5 ล้านบาท');
  addFormField(body, 'มูลค่าการสนับสนุน', '21.25 ล้านบาท (50%)');

  const kpi1Header = body.appendParagraph('ตัวชี้วัด:');
  setBodyStyle(kpi1Header);

  KPI_PHASE1.forEach(function(kpi, index) {
    const p = body.appendParagraph('  ' + (index + 1) + '. ' + kpi);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // งวดที่ 2
  const phase2Header = body.appendParagraph('งวดที่ 2 (ภายใน 12 เดือน นับจากวันที่ออกบัตรส่งเสริม)');
  setBodyStyle(phase2Header);
  phase2Header.setBold(true);

  addFormField(body, 'ประมาณการลงทุน', '42.5 ล้านบาท');
  addFormField(body, 'มูลค่าการสนับสนุน', '21.25 ล้านบาท (50%)');

  const kpi2Header = body.appendParagraph('ตัวชี้วัด:');
  setBodyStyle(kpi2Header);

  KPI_PHASE2.forEach(function(kpi, index) {
    const p = body.appendParagraph('  ' + (index + 1) + '. ' + kpi);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // 6. ผลกระทบเชิงบวก
  const section6 = body.appendParagraph('6. ผลกระทบเชิงบวกต่อเศรษฐกิจและสังคมที่ประเทศจะได้รับจากการให้ส่งเสริมโครงการนี้');
  setSectionHeaderStyle(section6);

  // ตารางผลกระทบ
  const impactTable = body.appendTable([
    ['เชิงปริมาณ', 'เชิงคุณภาพ']
  ]);
  setTableStyle(impactTable);

  // ผลประโยชน์ทางตรง - ปริมาณ
  const directQuantHeader = body.appendParagraph('ผลประโยชน์ทางตรง (เชิงปริมาณ):');
  setBodyStyle(directQuantHeader);
  directQuantHeader.setBold(true);

  SECTION6_DATA.directQuantitative.forEach(function(item) {
    const p = body.appendParagraph('- ' + item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // ผลประโยชน์ทางตรง - คุณภาพ
  const directQualHeader = body.appendParagraph('ผลประโยชน์ทางตรง (เชิงคุณภาพ):');
  setBodyStyle(directQualHeader);
  directQualHeader.setBold(true);

  SECTION6_DATA.directQualitative.forEach(function(item) {
    const p = body.appendParagraph('- ' + item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // ผลประโยชน์ทางอ้อม - ปริมาณ
  const indirectQuantHeader = body.appendParagraph('ผลประโยชน์ทางอ้อม (เชิงปริมาณ):');
  setBodyStyle(indirectQuantHeader);
  indirectQuantHeader.setBold(true);

  SECTION6_DATA.indirectQuantitative.forEach(function(item) {
    const p = body.appendParagraph('- ' + item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // ผลประโยชน์ทางอ้อม - คุณภาพ
  const indirectQualHeader = body.appendParagraph('ผลประโยชน์ทางอ้อม (เชิงคุณภาพ):');
  setBodyStyle(indirectQualHeader);
  indirectQualHeader.setBold(true);

  SECTION6_DATA.indirectQualitative.forEach(function(item) {
    const p = body.appendParagraph('- ' + item);
    setBodyStyle(p);
  });

  addPageBreak(body);
}

// ====================================
// เอกสารแนบ A: R&D
// ====================================
function createAttachmentA(body) {
  addPageHeader(body, 'เอกสารแนบ A หน้า 1/4');

  const title = body.appendParagraph('เอกสารแนบ A');
  setTitleStyle(title);

  const subtitle = body.appendParagraph('แบบประกอบการพิจารณาข้อเสนอโครงการ');
  setSubtitleStyle(subtitle);

  const subtitle2 = body.appendParagraph('ด้านการวิจัยและพัฒนา');
  setSubtitleStyle(subtitle2);

  body.appendParagraph('');

  // ข้อมูลโครงการ
  addFormField(body, 'ชื่อโครงการ', PROJECT_DATA.nameEN);
  addFormField(body, 'บริษัท', COMPANY_DATA.nameTH);
  addFormField(body, 'เลขทะเบียน', COMPANY_DATA.registrationNo);
  addFormField(body, 'งบประมาณ R&D', formatNumber(PROJECT_DATA.rdInvestment) + ' บาท');

  body.appendParagraph('');

  // 1. การดำเนินการ
  const rd1 = body.appendParagraph('1. การดำเนินการเพื่อเพิ่มขีดความสามารถในการแข่งขัน');
  setSectionHeaderStyle(rd1);

  const rd1Selection = body.appendParagraph('☑ ด้านการวิจัยและพัฒนา (หน่วยงานตรวจสอบ: สวทช.)');
  setBodyStyle(rd1Selection);

  body.appendParagraph('');

  // 2. ลักษณะการดำเนินการ
  const rd2 = body.appendParagraph('2. ลักษณะการดำเนินการ');
  setSectionHeaderStyle(rd2);

  const rd2Selection = body.appendParagraph('☑ ดำเนินการเอง');
  setBodyStyle(rd2Selection);

  body.appendParagraph('');

  // 3. แผนการวิจัยและพัฒนา
  const rd3 = body.appendParagraph('3. แผนการวิจัยและพัฒนา');
  setSectionHeaderStyle(rd3);

  // 3.3 รูปแบบการวิจัย
  const rd33 = body.appendParagraph('3) รูปแบบการวิจัยและพัฒนา');
  setSubSectionStyle(rd33);

  const rdTypes = [
    '☑ การวิจัยประยุกต์',
    '☑ การพัฒนาเชิงทดลอง',
    '☑ การออกแบบทางวิศวกรรมหรืออิเล็กทรอนิกส์',
    '☑ การพัฒนากระบวนการผลิตใหม่'
  ];
  rdTypes.forEach(function(type) {
    const p = body.appendParagraph(type);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // 4) วัตถุประสงค์
  const rd34 = body.appendParagraph('4) วัตถุประสงค์การวิจัยและพัฒนา');
  setSubSectionStyle(rd34);

  const objectives = [
    '1. พัฒนา Physical AI Agent (Aslan Mascot) ที่มีระบบเซนเซอร์อิเล็กทรอนิกส์ (Microphone, Speaker, LED, Touch sensors) พร้อมระบบ Wireless connectivity (WiFi, Bluetooth) และ Embedded AI OS สำหรับ Real-time inference สามารถสื่อสารภาษาไทยได้เป็นธรรมชาติ กำลังการผลิต 10,000 หน่วย/ปี',
    '2. พัฒนา Thai Natural Language Processing (NLP) Engine โดยสร้างโมเดล Deep Learning สำหรับประมวลผลภาษาไทยในบริบททางการเงิน รวบรวม Thai Scam Corpus Dataset ไม่น้อยกว่า 10,000 ตัวอย่าง ฝึกโมเดลให้มี F1 Score ≥ 85% รองรับ 100,000 คำขอ/วัน',
    '3. พัฒนา Multi-Agent AI Platform ระบบ 5 AI Agents ทำงานร่วมกันแบบ collaborative และ autonomous พร้อม Message Passing Framework ที่มีความแม่นยำ ≥ 99% รองรับ 50,000 ผู้ใช้พร้อมกัน',
    '4. พัฒนา Explainable AI (XAI) Framework ที่สามารถอธิบายการตัดสินใจของ AI เป็นภาษาไทยที่เข้าใจง่าย รองรับความโปร่งใสตามหลักการ Digital Governance รองรับ 100,000 การวิเคราะห์/วัน',
    '5. พัฒนา Mobile Application (iOS/Android) เชื่อมต่อกับ Physical AI Agent พร้อม Dashboard แสดงผล Real-time analysis และ Learning Center รองรับ 100,000 ผู้ใช้งานจริง'
  ];

  objectives.forEach(function(obj) {
    const p = body.appendParagraph(obj);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // 5) แผนการดำเนินงาน
  const rd35 = body.appendParagraph('5) แผนการดำเนินงาน');
  setSubSectionStyle(rd35);

  addFormField(body, 'สั่งซื้อเครื่องจักร', 'เดือน 1 พ.ศ. 2569');
  addFormField(body, 'ติดตั้งเครื่องจักรเสร็จสิ้น', 'เดือน 3 พ.ศ. 2569');

  body.appendParagraph('');

  // 6) ความคาดหวัง
  const rd36 = body.appendParagraph('6) ความคาดหวังและประโยชน์ที่คาดว่าจะได้รับจากการวิจัยและพัฒนา');
  setSubSectionStyle(rd36);

  const expectations = [
    '1. ผลิตภัณฑ์และการให้บริการ: Physical AI Agent ผลิตได้ 10,000 หน่วย/ปี มูลค่า 150-200 ล้านบาท/ปี, Multi-Agent Platform รองรับ 50,000 ผู้ใช้พร้อมกัน 24/7/365, Thai NLP Service รองรับ 100,000 คำขอ/วัน ความแม่นยำ ≥85%',
    '2. การขยายตลาดและผู้ใช้บริการ: กลุ่มเด็ก (Financial Literacy) เป้าหมาย 20,000 ครอบครัว, กลุ่มผู้สูงอายุ (Scam Protection) 30,000 ราย, กลุ่มมืออาชีพ (Pro Traders) 5,000 ราย ปีแรก',
    '3. ผลกระทบทางเศรษฐกิจและสังคม: ลดความเสียหายจากการถูกหลอกลวง (Scam losses) ประมาณ 52,500 ล้านบาท/ปี, สร้างนวัตกรรม AI ครั้งแรกในประเทศไทย, สร้างงานวิจัยคุณภาพสูง: สิทธิบัตร ≥2 ฉบับ, บทความวิชาการ ≥3 ฉบับ',
    '4. ทรัพย์สินทางปัญญาและฐานข้อมูล: Thai Scam Corpus Dataset ≥10,000 ตัวอย่าง (เปิดให้ชุมชนวิจัยใช้), XAI Framework สำหรับ AI ที่โปร่งใสและอธิบายได้',
    '5. การพัฒนาบุคลากร: พัฒนาทีมวิจัย 15 คน (2 PhD, 5 Master\'s, 8 Bachelor\'s), สร้างความเชี่ยวชาญด้าน AI, NLP, Multi-Agent Systems ในไทย'
  ];

  expectations.forEach(function(exp) {
    const p = body.appendParagraph(exp);
    setBodyStyle(p);
  });

  addPageBreak(body);

  // หน้า 2 - บุคลากร
  addPageHeader(body, 'เอกสารแนบ A หน้า 2/4');

  // 7) บุคลากร
  const rd37 = body.appendParagraph('7) แผนการจ้างงานบุคลากรด้านวิจัยและพัฒนาประจำโครงการ (Full-time equivalent)');
  setSubSectionStyle(rd37);

  const personnelTableData = [['ลำดับ', 'ตำแหน่ง', 'วุฒิการศึกษา', 'สัญชาติ', 'จำนวน (คน)', 'เงินเดือนรวม (บาท/ปี)']];

  PERSONNEL.forEach(function(p, index) {
    personnelTableData.push([
      (index + 1).toString(),
      p.position,
      p.education,
      'ไทย ' + p.thaiCount + ' / ต่างชาติ ' + p.foreignCount,
      p.totalCount.toString(),
      p.annualSalary
    ]);
  });

  // รวม
  personnelTableData.push(['รวมทั้งสิ้น', '', '', '', '15', '23,000,000']);

  const personnelTable = body.appendTable(personnelTableData);
  setTableStyle(personnelTable);

  body.appendParagraph('');

  const personnelNote = body.appendParagraph('หมายเหตุ: รวมเงินเดือน R&D Personnel 23M บาท/ปี, ค่าใช้จ่าย Personnel รวม 36.24M บาท (รวม benefits, training, overhead)');
  personnelNote.setFontSize(10);
  personnelNote.setItalic(true);

  addPageBreak(body);

  // หน้า 3 - งบประมาณ
  addPageHeader(body, 'เอกสารแนบ A หน้า 3/4');

  const rd4 = body.appendParagraph('4. ประมาณการเงินลงทุนและ/หรือค่าใช้จ่ายสำหรับโครงการที่ยื่นข้อเสนอ');
  setSectionHeaderStyle(rd4);

  const rdExpenseTable = body.appendTable([
    ['ประเภทเงินลงทุนและ/หรือค่าใช้จ่าย', 'มูลค่า (ล้านบาท)'],
    ['1) ค่าจ้างหรือเงินเดือน (Personnel)', '36.24'],
    ['2) ค่าเครื่องมือหรืออุปกรณ์ (Equipment)', '9.00'],
    ['3) ค่าวัตถุดิบหรือวัสดุ/ข้อมูลและ Dataset', '8.50'],
    ['4) ค่าใช้จ่ายทางตรงอื่น ๆ', '5.26'],
    ['รวมทั้งสิ้น', '59.00']
  ]);
  setTableStyle(rdExpenseTable);

  body.appendParagraph('');

  // รายละเอียดกิจกรรม R&D
  const rdActivitiesHeader = body.appendParagraph('รายละเอียดกิจกรรม R&D (5 กิจกรรม):');
  setBodyStyle(rdActivitiesHeader);
  rdActivitiesHeader.setBold(true);

  const rdActivitiesTable = body.appendTable([
    ['#', 'กิจกรรม R&D', 'งบประมาณ (บาท)', '%', 'ระยะเวลา'],
    ['1', 'Physical AI Agent Development', '12,980,000', '22.0%', '12 เดือน'],
    ['2', 'Explainable AI (XAI) Framework', '11,240,000', '19.1%', '10 เดือน'],
    ['3', 'Thai NLP Model for Finance', '15,780,000', '26.7%', '12 เดือน'],
    ['4', 'Fraud Pattern Recognition', '10,500,000', '17.8%', '10 เดือน'],
    ['5', 'Multi-Agent System Integration', '8,500,000', '14.4%', '12 เดือน'],
    ['', 'รวมทั้งสิ้น', '59,000,000', '100.0%', '']
  ]);
  setTableStyle(rdActivitiesTable);

  addPageBreak(body);

  // หน้า 4 - ตัวชี้วัด
  addPageHeader(body, 'เอกสารแนบ A หน้า 4/4');

  const rd5 = body.appendParagraph('5. ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด');
  setSectionHeaderStyle(rd5);

  const rdKpiTable = body.appendTable([
    ['งวดงาน', 'ตัวชี้วัด', 'ประมาณการลงทุน (ล้านบาท)', 'สัดส่วน (%)', 'มูลค่าสนับสนุน (ล้านบาท)'],
    ['งวดที่ 1\n(ภายใน 6 เดือน)', '1. Working Prototype 3 AI Agents\n2. Thai NLP F1 Score ≥80%\n3. Message Passing ≥99%\n4. ใช้งบ 50%', '29.5', '50', '14.75'],
    ['งวดที่ 2\n(ภายใน 12 เดือน)', '1. Production System 5 AI Agents\n2. F1 Score ≥85%\n3. สิทธิบัตร ≥2 ฉบับ\n4. บทความ ≥3 ฉบับ\n5. Dataset ≥10,000 ตัวอย่าง', '29.5', '50', '14.75'],
    ['รวม', '', '59.0', '', '29.5']
  ]);
  setTableStyle(rdKpiTable);

  // ลายเซ็น
  body.appendParagraph('');
  addSignatureBlock(body, 'กรรมการผู้จัดการ', COMPANY_DATA.contactPerson);

  addPageBreak(body);
}

// ====================================
// เอกสารแนบ C: Digital Technology
// ====================================
function createAttachmentC(body) {
  addPageHeader(body, 'เอกสารแนบ C หน้า 1/4');

  const title = body.appendParagraph('เอกสารแนบ C');
  setTitleStyle(title);

  const subtitle = body.appendParagraph('แบบประกอบการพิจารณาข้อเสนอโครงการ');
  setSubtitleStyle(subtitle);

  const subtitle2 = body.appendParagraph('ด้านการใช้เทคโนโลยีดิจิทัล');
  setSubtitleStyle(subtitle2);

  body.appendParagraph('');

  // ข้อมูลโครงการ
  addFormField(body, 'ชื่อโครงการ', PROJECT_DATA.nameEN);
  addFormField(body, 'บริษัท', COMPANY_DATA.nameTH);
  addFormField(body, 'เลขทะเบียน', COMPANY_DATA.registrationNo);
  addFormField(body, 'งบประมาณเทคโนโลยีดิจิทัล', formatNumber(PROJECT_DATA.digitalInvestment) + ' บาท');

  body.appendParagraph('');

  // 1. การดำเนินการ
  const digital1 = body.appendParagraph('1. การดำเนินการเพื่อเพิ่มขีดความสามารถในการแข่งขัน');
  setSectionHeaderStyle(digital1);

  const digital1Selection = body.appendParagraph('☑ การใช้เทคโนโลยีดิจิทัล เพื่อปรับปรุงประสิทธิภาพด้วยเทคโนโลยีที่ทันสมัย (หน่วยตรวจสอบ: depa)');
  setBodyStyle(digital1Selection);

  body.appendParagraph('');

  // 2. สาระสำคัญ
  const digital2 = body.appendParagraph('2. สาระสำคัญของการนำเทคโนโลยีดิจิทัลมาใช้');
  setSectionHeaderStyle(digital2);

  const digitalTypes = [
    '☑ การนำซอฟต์แวร์ โปรแกรม หรือระบบสารสนเทศ มาใช้ในการเชื่อมโยงภายในองค์กรอย่างเป็นระบบ (Integrated) และเชื่อมโยงภายนอกองค์กร (Connected) โดยต้องมีการเชื่อมโยงข้อมูลอย่างน้อย 3 ฟังก์ชัน',
    '☑ การประยุกต์ใช้ปัญญาประดิษฐ์ (Artificial Intelligence หรือ AI) Machine Learning การนำ Big Data มาใช้หรือการวิเคราะห์ข้อมูล (Data Analytics)'
  ];
  digitalTypes.forEach(function(type) {
    const p = body.appendParagraph(type);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // ตารางรายละเอียดเทคโนโลยี
  const digitalDetailTable = body.appendTable([
    ['ขั้นตอนการทำงาน', 'เทคโนโลยีดิจิทัล', 'หน่วยงานที่เกี่ยวข้อง', 'การเชื่อมโยงข้อมูล', 'ประสิทธิภาพที่เพิ่มขึ้น'],
    ['AI Model Training', 'GPU Cloud (AWS/GCP)', 'ทีม AI Research', 'เชื่อมต่อกับ Data Pipeline', 'เร่งการ Train Model 10x'],
    ['Data Processing', 'Kubernetes Cluster', 'ทีม DevOps', 'รองรับ Auto-scaling', 'Uptime ≥99.5%'],
    ['API Services', 'API Gateway (Kong)', 'ทีม Backend', 'เชื่อมต่อ Banking APIs', 'Response <200ms'],
    ['Security', 'SIEM + WAF', 'ทีม Security', 'Real-time Monitoring', 'ตรวจจับภัยคุกคาม 24/7'],
    ['Mobile App', 'Flutter + Firebase', 'ทีม Frontend', 'เชื่อมต่อ Physical AI Agent', 'รองรับ 100K users']
  ]);
  setTableStyle(digitalDetailTable);

  addPageBreak(body);

  // หน้า 2
  addPageHeader(body, 'เอกสารแนบ C หน้า 2/4');

  // 4. การสนับสนุนอุตสาหกรรม
  const digital4 = body.appendParagraph('4. การสนับสนุนอุตสาหกรรมการพัฒนาเทคโนโลยีดิจิทัลที่พัฒนาโดยผู้ประกอบการในประเทศ');
  setSectionHeaderStyle(digital4);

  const domesticNote = body.appendParagraph('โครงการนี้ใช้เทคโนโลยีดิจิทัลที่พัฒนาในประเทศร่วมด้วย โดยมีสัดส่วนการลงทุนในประเทศ ≥30%');
  setBodyStyle(domesticNote);

  body.appendParagraph('');

  // 5. รายละเอียดเทคโนโลยีดิจิทัล
  const digital5 = body.appendParagraph('5. รายละเอียดเทคโนโลยีดิจิทัลที่จะนำมาใช้');
  setSectionHeaderStyle(digital5);

  const digitalTechTable = body.appendTable([
    ['รายการเทคโนโลยีดิจิทัล', 'ประเทศ', 'จำนวน (ชุด)', 'มูลค่า (ล้านบาท)'],
    ['1. เทคโนโลยีดิจิทัลต่างประเทศ', '', '', ''],
    ['   1.1 Cloud Services (AWS/GCP/Azure)', 'ต่างประเทศ', '1', '3.5'],
    ['   1.2 Software Licenses (Enterprise)', 'ต่างประเทศ', '15', '2.0'],
    ['   1.3 Security Tools (Palo Alto, etc.)', 'ต่างประเทศ', '2', '1.2'],
    ['รวมมูลค่าต่างประเทศ', '', '', '6.7'],
    ['2. เทคโนโลยีดิจิทัลในประเทศ', '', '', ''],
    ['   2.1 On-premise Servers', 'ไทย', '8', '8.0'],
    ['   2.2 Network Equipment', 'ไทย', '1', '2.8'],
    ['   2.3 System Integration Services', 'ไทย', '1', '8.0'],
    ['   2.4 Cloud Hosting (Thai DC)', 'ไทย', '1', '0.5'],
    ['รวมมูลค่าในประเทศ', '', '', '19.3'],
    ['รวมมูลค่าทั้งสิ้น', '', '', '26.0'],
    ['สัดส่วนการลงทุนในประเทศ', '', '', '74.2%']
  ]);
  setTableStyle(digitalTechTable);

  addPageBreak(body);

  // หน้า 3
  addPageHeader(body, 'เอกสารแนบ C หน้า 3/4');

  // 6. แผนการดำเนินงาน
  const digital6 = body.appendParagraph('6. แผนการดำเนินงาน');
  setSectionHeaderStyle(digital6);

  addFormField(body, 'เริ่มสั่งซื้อซอฟต์แวร์/เช่าหรือใช้บริการ Cloud', 'เดือน 1 พ.ศ. 2569');
  addFormField(body, 'พัฒนาและนำซอฟต์แวร์ขึ้นใช้งานจริง', 'เดือน 6 พ.ศ. 2569');

  body.appendParagraph('');

  // 7. ประมาณการเงินลงทุน
  const digital7 = body.appendParagraph('7. ประมาณการเงินลงทุนและ/หรือค่าใช้จ่ายสำหรับโครงการที่ยื่นข้อเสนอ');
  setSectionHeaderStyle(digital7);

  const digitalExpenseTable = body.appendTable([
    ['ประเภทเงินลงทุนและ/หรือค่าใช้จ่าย', 'มูลค่า (ล้านบาท)'],
    ['1) ค่าซอฟต์แวร์ (พัฒนา/ปรับปรุงใหม่ และสำเร็จรูป)', '3.0'],
    ['2) ค่าใช้จ่ายในการเช่า/ใช้บริการ Cloud หรือ Data Center', '5.0'],
    ['3) ค่าฮาร์ดแวร์และอุปกรณ์เครือข่าย', '10.0'],
    ['4) ค่า System Integration และ API Development', '8.0'],
    ['รวมทั้งสิ้น', '26.0']
  ]);
  setTableStyle(digitalExpenseTable);

  body.appendParagraph('');

  // รายละเอียดหมวดหมู่
  const digitalCategoriesHeader = body.appendParagraph('รายละเอียดเทคโนโลยีดิจิทัล (5 หมวดหมู่):');
  setBodyStyle(digitalCategoriesHeader);
  digitalCategoriesHeader.setBold(true);

  const digitalCategoriesTable = body.appendTable([
    ['#', 'หมวดหมู่', 'งบประมาณ (บาท)', '%'],
    ['1', 'Cloud Services & Infrastructure', '5,000,000', '19.2%'],
    ['2', 'On-premise Hardware & Servers', '8,000,000', '30.8%'],
    ['3', 'Software Licenses & Development Tools', '3,000,000', '11.5%'],
    ['4', 'Security & Compliance Systems', '2,000,000', '7.7%'],
    ['5', 'System Integration & API Development', '8,000,000', '30.8%'],
    ['', 'รวมทั้งสิ้น', '26,000,000', '100.0%']
  ]);
  setTableStyle(digitalCategoriesTable);

  addPageBreak(body);

  // หน้า 4 - ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด (ภาพรวมทั้งโครงการ)
  addPageHeader(body, 'เอกสารแนบ C หน้า 4/6');

  const digital8 = body.appendParagraph('8. ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด');
  setSectionHeaderStyle(digital8);

  const digital8Note = body.appendParagraph('(พิจารณากำหนดตามความเหมาะสมของโครงการ - รวม R&D + Digital)');
  digital8Note.setFontSize(10);
  digital8Note.setItalic(true);
  digital8Note.setFontFamily('Sarabun');

  body.appendParagraph('');

  // 8.1 สรุปภาพรวมทั้งโครงการ
  const digital81 = body.appendParagraph('8.1 สรุปภาพรวมการลงทุนและเงินสนับสนุนทั้งโครงการ');
  setSubSectionStyle(digital81);

  const overviewTable = body.appendTable([
    ['หมวดการลงทุน', 'งบประมาณลงทุน (ล้านบาท)', 'อัตราสนับสนุน', 'เงินสนับสนุนสูงสุด (ล้านบาท)', 'หน่วยรับรอง'],
    ['R&D (เอกสารแนบ A)', '59.0', '50%', '29.5', 'สวทช.'],
    ['Digital Technology (เอกสารแนบ C)', '26.0', '50%', '13.0', 'depa'],
    ['รวมทั้งโครงการ', '85.0', '50%', '42.5', '']
  ]);
  setTableStyle(overviewTable);

  body.appendParagraph('');

  // 8.2 ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด
  const digital82 = body.appendParagraph('8.2 ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด');
  setSubSectionStyle(digital82);

  const kpiSummaryTable = body.appendTable([
    ['งวดงาน', 'ประมาณการลงทุน (ล้านบาท)', 'สัดส่วน (%)', 'เงินสนับสนุน (ล้านบาท)'],
    ['งวดที่ 1 (6 เดือน)', '42.5 (R&D: 29.5 + Digital: 13.0)', '50%', '21.25 (R&D: 14.75 + Digital: 6.5)'],
    ['งวดที่ 2 (12 เดือน)', '42.5 (R&D: 29.5 + Digital: 13.0)', '50%', '21.25 (R&D: 14.75 + Digital: 6.5)'],
    ['รวมทั้งสิ้น', '85.0', '100%', '42.5']
  ]);
  setTableStyle(kpiSummaryTable);

  addPageBreak(body);

  // หน้า 5 - รายละเอียดตัวชี้วัดรายงวด
  addPageHeader(body, 'เอกสารแนบ C หน้า 5/6');

  // งวดที่ 1
  const phase1Header = body.appendParagraph('งวดที่ 1: ตัวชี้วัดความคืบหน้าของโครงการ (ภายใน 6 เดือน นับจากวันที่ออกบัตรส่งเสริม)');
  setSubSectionStyle(phase1Header);

  const phase1RD = body.appendParagraph('ส่วน R&D (สวทช.):');
  setBodyStyle(phase1RD);
  phase1RD.setBold(true);

  const phase1RDItems = [
    '- Physical AI Agent: Prototype แรกสำเร็จ (3 AI Agents working)',
    '- Thai NLP: Model Training เบื้องต้น (F1 Score ≥ 80%)',
    '- Multi-Agent: Framework พื้นฐาน (Message Passing ≥ 99%)',
    '- จัดซื้ออุปกรณ์ R&D 50% (GPU Servers, Workstations)',
    '- บุคลากรวิจัย: ฝึกอบรม ≥ 80%'
  ];
  phase1RDItems.forEach(function(item) {
    const p = body.appendParagraph(item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  const phase1Digital = body.appendParagraph('ส่วน Digital (depa):');
  setBodyStyle(phase1Digital);
  phase1Digital.setBold(true);

  const phase1DigitalItems = [
    '- จัดซื้อและติดตั้ง On-premise Hardware 50% (Servers, Network)',
    '- ตั้งค่า Cloud Infrastructure (AWS/GCP) พร้อมใช้งาน',
    '- ติดตั้ง Basic Security Systems (Firewall, IDS/IPS)',
    '- Setup Development Environment',
    '- Deploy Beta version บน Staging'
  ];
  phase1DigitalItems.forEach(function(item) {
    const p = body.appendParagraph(item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  const phase1Conditions = body.appendParagraph('เงื่อนไขการเบิกจ่าย:');
  setBodyStyle(phase1Conditions);
  phase1Conditions.setBold(true);

  const phase1ConditionItems = [
    '1. มีการลงทุนจริงตามขอบข่ายที่ได้รับอนุมัติ ในสัดส่วนไม่น้อยกว่าร้อยละ 50 ของมูลค่าการลงทุน พร้อมหลักฐานการชำระเงิน',
    '2. มีหนังสือรับรองผลการดำเนินการตามตัวชี้วัดจาก สวทช. หรือ depa (ตามขอบข่ายการดำเนินการ)'
  ];
  phase1ConditionItems.forEach(function(item) {
    const p = body.appendParagraph(item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  // งวดที่ 2
  const phase2Header = body.appendParagraph('งวดที่ 2: ตัวชี้วัดผลการดำเนินการตามเป้าหมาย (ภายใน 12 เดือน นับจากวันที่ออกบัตรส่งเสริม)');
  setSubSectionStyle(phase2Header);

  const phase2RD = body.appendParagraph('ส่วน R&D (สวทช.):');
  setBodyStyle(phase2RD);
  phase2RD.setBold(true);

  const phase2RDItems = [
    '- Production System: ครบ 5 AI Agents, Uptime ≥ 99.5%',
    '- Thai NLP: F1 Score ≥ 85%, False Positive ≤ 10%',
    '- Performance: Response Time < 200ms, Throughput ≥ 500 TPS',
    '- ยื่นสิทธิบัตร ≥ 2 ฉบับ',
    '- ตีพิมพ์บทความวิชาการ ≥ 3 ฉบับ',
    '- UAT กับผู้ใช้จริง ≥ 50 คน',
    '- Thai Scam Dataset ≥ 10,000 ตัวอย่าง'
  ];
  phase2RDItems.forEach(function(item) {
    const p = body.appendParagraph(item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  const phase2Digital = body.appendParagraph('ส่วน Digital (depa):');
  setBodyStyle(phase2Digital);
  phase2Digital.setBold(true);

  const phase2DigitalItems = [
    '- จัดซื้อและติดตั้ง Hardware ครบถ้วน 100%',
    '- Production Deployment รองรับ 50,000 Concurrent Users',
    '- System Uptime ≥ 99.5%',
    '- API Response Time < 200ms (p95)',
    '- ผ่าน Security Audit จากหน่วยงานรับรอง',
    '- Go Live บริการจริง'
  ];
  phase2DigitalItems.forEach(function(item) {
    const p = body.appendParagraph(item);
    setBodyStyle(p);
  });

  body.appendParagraph('');

  const phase2Conditions = body.appendParagraph('เงื่อนไขการเบิกจ่าย:');
  setBodyStyle(phase2Conditions);
  phase2Conditions.setBold(true);

  const phase2ConditionItems = [
    '1. มีการลงทุนจริงครบถ้วนตามขอบข่ายที่ได้รับอนุมัติ พร้อมหลักฐานการชำระเงิน',
    '2. มีหนังสือรับรองผลการดำเนินการตามตัวชี้วัดจาก สวทช. หรือ depa (ตามขอบข่ายการดำเนินการ)'
  ];
  phase2ConditionItems.forEach(function(item) {
    const p = body.appendParagraph(item);
    setBodyStyle(p);
  });

  addPageBreak(body);

  // หน้า 6 - สรุปและ Hardware list
  addPageHeader(body, 'เอกสารแนบ C หน้า 6/6');

  // 8.3 สรุปการเบิกจ่ายแยกหมวด
  const digital83 = body.appendParagraph('8.3 สรุปการเบิกจ่ายเงินสนับสนุนรายงวด (แยกตามหมวด)');
  setSubSectionStyle(digital83);

  const summaryByCategory = body.appendTable([
    ['รายการ', 'R&D ลงทุน', 'R&D สนับสนุน', 'Digital ลงทุน', 'Digital สนับสนุน', 'รวมลงทุน', 'รวมสนับสนุน'],
    ['งบประมาณทั้งหมด', '59.0', '29.5', '26.0', '13.0', '85.0', '42.5'],
    ['งวดที่ 1 (50%)', '29.5', '14.75', '13.0', '6.5', '42.5', '21.25'],
    ['งวดที่ 2 (50%)', '29.5', '14.75', '13.0', '6.5', '42.5', '21.25']
  ]);
  setTableStyle(summaryByCategory);

  const unitNote = body.appendParagraph('หน่วย: ล้านบาท');
  unitNote.setFontSize(10);
  unitNote.setItalic(true);
  unitNote.setFontFamily('Sarabun');

  body.appendParagraph('');

  // 8.4 Hardware List
  const digital84 = body.appendParagraph('8.4 รายการ Hardware และอุปกรณ์ที่ได้รับการสนับสนุน');
  setSubSectionStyle(digital84);

  // R&D Hardware
  const rdHwHeader = body.appendParagraph('อุปกรณ์ R&D (เอกสารแนบ A) - รับรองโดย สวทช.');
  setBodyStyle(rdHwHeader);
  rdHwHeader.setBold(true);

  const rdHardwareTable = body.appendTable([
    ['#', 'รายการ', 'มูลค่าลงทุน (บาท)', 'สนับสนุน 50% (บาท)'],
    ['1', 'GPU Server (NVIDIA A100 80GB) x 4 เครื่อง', '6,000,000', '3,000,000'],
    ['2', 'Development Workstations x 15 เครื่อง', '1,500,000', '750,000'],
    ['3', 'Physical AI Prototyping Equipment', '800,000', '400,000'],
    ['4', 'Testing Devices (Smartphones, Tablets, IoT) x 50', '700,000', '350,000'],
    ['รวม Hardware R&D', '', '9,000,000', '4,500,000']
  ]);
  setTableStyle(rdHardwareTable);

  body.appendParagraph('');

  // Digital Hardware
  const digitalHwHeader = body.appendParagraph('อุปกรณ์ Digital (เอกสารแนบ C) - รับรองโดย depa');
  setBodyStyle(digitalHwHeader);
  digitalHwHeader.setBold(true);

  const digitalHardwareTable = body.appendTable([
    ['#', 'รายการ', 'มูลค่าลงทุน (บาท)', 'สนับสนุน 50% (บาท)'],
    ['5', 'Application Server (Dual Xeon, 256GB RAM) x 4', '2,000,000', '1,000,000'],
    ['6', 'Database Server (512GB RAM, NVMe) x 2', '1,600,000', '800,000'],
    ['7', 'Kubernetes Nodes (Master 3 + Worker 10)', '1,950,000', '975,000'],
    ['8', 'Next-Gen Firewall (Palo Alto) x 2', '1,200,000', '600,000'],
    ['9', 'Load Balancer (F5) x 2', '800,000', '400,000'],
    ['10', 'Enterprise Storage (SAN, 100TB)', '1,500,000', '750,000'],
    ['11', 'Network Switches (Core, 10Gb) x 4', '600,000', '300,000'],
    ['12', 'UPS & Power Management x 2', '600,000', '300,000'],
    ['รวม Hardware Digital', '', '10,250,000', '5,125,000']
  ]);
  setTableStyle(digitalHardwareTable);

  body.appendParagraph('');

  // สรุป Hardware ทั้งหมด
  const totalHwHeader = body.appendParagraph('สรุป Hardware ทั้งหมด');
  setBodyStyle(totalHwHeader);
  totalHwHeader.setBold(true);

  const totalHardwareTable = body.appendTable([
    ['หมวด', 'มูลค่าลงทุน (บาท)', 'สนับสนุน 50% (บาท)'],
    ['R&D (สวทช.)', '9,000,000', '4,500,000'],
    ['Digital (depa)', '10,250,000', '5,125,000'],
    ['รวม Hardware ทั้งหมด', '19,250,000', '9,625,000']
  ]);
  setTableStyle(totalHardwareTable);

  body.appendParagraph('');

  const disclaimerNote = body.appendParagraph('* หมายเหตุ: จะได้รับอนุมัติเบิกจ่ายเงินสนับสนุนไม่เกินมูลค่าเงินที่ชำระจริง');
  disclaimerNote.setFontSize(10);
  disclaimerNote.setItalic(true);
  disclaimerNote.setFontFamily('Sarabun');

  // ลายเซ็น
  body.appendParagraph('');
  addSignatureBlock(body, 'กรรมการผู้จัดการ', COMPANY_DATA.contactPerson);
}

// ====================================
// HELPER FUNCTIONS
// ====================================

function addPageHeader(body, pageNum) {
  const header = body.appendParagraph(pageNum);
  header.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  header.setFontSize(10);
  header.setFontFamily('Sarabun');
}

function addPageBreak(body) {
  body.appendPageBreak();
}

function setTitleStyle(paragraph) {
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  paragraph.setFontSize(16);
  paragraph.setBold(true);
  paragraph.setFontFamily('Sarabun');
}

function setSubtitleStyle(paragraph) {
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  paragraph.setFontSize(12);
  paragraph.setFontFamily('Sarabun');
}

function setSectionHeaderStyle(paragraph) {
  paragraph.setFontSize(14);
  paragraph.setBold(true);
  paragraph.setFontFamily('Sarabun');
  paragraph.setSpacingBefore(12);
  paragraph.setSpacingAfter(6);
}

function setSubSectionStyle(paragraph) {
  paragraph.setFontSize(12);
  paragraph.setBold(true);
  paragraph.setFontFamily('Sarabun');
  paragraph.setSpacingBefore(8);
  paragraph.setSpacingAfter(4);
}

function setBodyStyle(paragraph) {
  paragraph.setFontSize(11);
  paragraph.setFontFamily('Sarabun');
  paragraph.setLineSpacing(1.5);
}

function setTableStyle(table) {
  // ตั้งค่า header row
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    const cell = headerRow.getCell(i);
    cell.setBackgroundColor('#E0E0E0');
    const paragraph = cell.getChild(0).asParagraph();
    paragraph.setBold(true);
    paragraph.setFontSize(10);
    paragraph.setFontFamily('Sarabun');
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  // ตั้งค่า body rows
  for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      const paragraph = cell.getChild(0).asParagraph();
      paragraph.setFontSize(10);
      paragraph.setFontFamily('Sarabun');
    }
  }

  // ตั้งค่า border
  table.setBorderWidth(1);
}

function addFormField(body, label, value) {
  const p = body.appendParagraph(label + ': ' + value);
  p.setFontSize(11);
  p.setFontFamily('Sarabun');
}

function addSignatureBlock(body, title, name) {
  body.appendParagraph('');

  const signLine = body.appendParagraph('ลงชื่อ ________________________________');
  signLine.setFontSize(11);
  signLine.setFontFamily('Sarabun');
  signLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const nameP = body.appendParagraph('(' + name + ')');
  nameP.setFontSize(11);
  nameP.setFontFamily('Sarabun');
  nameP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const titleP = body.appendParagraph(title);
  titleP.setFontSize(11);
  titleP.setFontFamily('Sarabun');
  titleP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const companyP = body.appendParagraph(COMPANY_DATA.nameTH);
  companyP.setFontSize(11);
  companyP.setFontFamily('Sarabun');
  companyP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const dateP = body.appendParagraph('วันที่ ________________');
  dateP.setFontSize(11);
  dateP.setFontFamily('Sarabun');
  dateP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function formatNumber(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// ====================================
// UTILITY FUNCTIONS
// ====================================

/**
 * สร้างเอกสาร PDF จาก Google Doc
 */
function exportToPDF() {
  const docUrl = createBOIForm();
  const docId = docUrl.split('/d/')[1].split('/')[0];
  const doc = DocumentApp.openById(docId);
  const blob = doc.getAs('application/pdf');
  blob.setName('BOI_Form_ASLAN_WEALTH.pdf');
  DriveApp.createFile(blob);
  Logger.log('สร้าง PDF สำเร็จ!');
}

/**
 * ทดสอบการสร้างฟอร์ม
 */
function testCreateForm() {
  try {
    const url = createBOIForm();
    Logger.log('SUCCESS: ' + url);
    return url;
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    return null;
  }
}
