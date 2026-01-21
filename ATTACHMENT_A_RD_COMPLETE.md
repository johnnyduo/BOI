# เอกสารแนบ A: รายละเอียดกิจกรรมวิจัยและพัฒนา (R&D)

---

## ข้อมูลโครงการ

| รายการ | ข้อมูล |
|--------|--------|
| **ชื่อโครงการ** | ASLAN AI GUARDIAN PLATFORM |
| **บริษัท** | บริษัท อัสลาน เวลธ์ จำกัด (ASLAN WEALTH COMPANY LIMITED) |
| **เลขทะเบียน** | 0105567220030 |
| **งบประมาณ R&D** | **59,000,000 บาท** |

---

## ส่วนที่ 1: วัตถุประสงค์การวิจัย

### 1.1 วัตถุประสงค์หลัก (Main Objectives)

โครงการวิจัยและพัฒนาระบบปัญญาประดิษฐ์หลายตัวแทน (Multi-Agent AI System) เพื่อการปกป้องทางการเงินสำหรับประชาชนไทย มีวัตถุประสงค์ดังนี้:

#### 1. พัฒนา Physical AI Agent (Aslan Mascot)
- สร้างหุ่นยนต์/ตุ๊กตา AI รูปสิงโต "Aslan" ที่มีระบบเซนเซอร์อิเล็กทรอนิกส์ (Microphone, Speaker, LED, Touch sensors)
- ติดตั้งระบบ Wireless connectivity (WiFi, Bluetooth) สำหรับเชื่อมต่อ IoT
- พัฒนา Embedded AI OS และระบบประมวลผลแบบฝังตัวสำหรับ Real-time inference
- สามารถสื่อสารภาษาไทยได้เป็นธรรมชาติและให้คำแนะนำทางการเงิน
- **เป้าหมาย:** ผลิต 10,000 หน่วย/ปี สำหรับกลุ่มเด็กและผู้สูงอายุ

#### 2. พัฒนา Thai Natural Language Processing (NLP) Engine
- สร้างโมเดล Deep Learning สำหรับประมวลผลภาษาไทยในบริบททางการเงิน
- รวบรวมและสร้าง Thai Scam Corpus Dataset ขนาดไม่น้อยกว่า 10,000 ตัวอย่าง
- ฝึกโมเดล NLP ให้มี F1 Score ≥ 85% ในการตรวจจับ scam/fraud
- สนับสนุนการถาม-ตอบเกี่ยวกับความรู้ทางการเงินเป็นภาษาไทย
- **เป้าหมาย:** รองรับ 100,000 คำขอ/วัน

#### 3. พัฒนา Multi-Agent AI Platform
- สร้างระบบ 5 AI Agents ทำงานร่วมกันแบบ collaborative และ autonomous
- พัฒนา Message Passing Framework ที่มีความแม่นยำ ≥ 99%
- ระบบ Real-time monitoring และ alert สำหรับภัยคุกคามทางการเงิน
- Integration กับ Banking APIs และ Financial data sources
- **เป้าหมาย:** รองรับ 50,000 ผู้ใช้พร้อมกัน (Concurrent Users)

#### 4. พัฒนา Explainable AI (XAI) Framework
- สร้างโมดูล XAI ที่สามารถอธิบายการตัดสินใจของ AI เป็นภาษาไทยที่เข้าใจง่าย
- รองรับความโปร่งใส (Transparency) ตามหลักการ Digital Governance
- ให้ผู้ใช้เข้าใจว่า AI ตัดสินใจอย่างไร (Why AI made this decision)
- **เป้าหมาย:** รองรับ 100,000 การวิเคราะห์/วัน

#### 5. พัฒนา Mobile Application
- สร้างแอปพลิเคชันมือถือ (iOS/Android) เชื่อมต่อกับ Physical AI Agent
- Dashboard แสดงผล Real-time analysis และ alerts
- Learning Center สำหรับเรียนรู้ทางการเงิน
- **เป้าหมาย:** รองรับ 100,000 ผู้ใช้งานจริง (Active Users)

### 1.2 ผลที่คาดว่าจะได้รับ (Expected Outcomes)

- ระบบ Physical AI Agent ที่ทำงานได้จริงและสามารถผลิตเชิงพาณิชย์ได้ 10,000 หน่วย/ปี
- โมเดล Thai NLP ที่มีความแม่นยำ ≥ 85% F1 Score ในการตรวจจับ scam
- Multi-Agent Platform ที่รองรับผู้ใช้พร้อมกัน 50,000 คน พร้อม Uptime ≥ 99.5%
- XAI Framework ที่อธิบายการตัดสินใจได้เป็นภาษาไทย
- Mobile App ที่มี UX/UI ที่เหมาะกับผู้ใช้ไทยทุกวัย
- ยื่นขอสิทธิบัตร ≥ 2 ฉบับ
- ตีพิมพ์บทความวิชาการ ≥ 3 ฉบับ ในวารสารวิชาการไทย
- Thai Scam Corpus Dataset ≥ 10,000 ตัวอย่าง (เปิดให้ชุมชนวิจัยใช้)

---

## ส่วนที่ 2: รายละเอียดกิจกรรม R&D (5 กิจกรรม)

### กิจกรรม R&D #1: Physical AI Agent Development
**งบประมาณ: 12,980,000 บาท**

**รายละเอียด:**
- วิจัยและออกแบบ Hardware สำหรับ Physical AI Agent (Aslan Mascot)
- พัฒนาระบบเซนเซอร์: Microphone array, Speaker system, LED indicators, Touch sensors
- ออกแบบและพัฒนา Embedded AI Processing Unit
- พัฒนา IoT connectivity modules (WiFi, Bluetooth 5.0)
- สร้าง Prototype และทดสอบกับผู้ใช้จริง (UAT ≥ 50 คน)
- ออกแบบ Industrial Design ที่เป็นมิตรกับเด็กและผู้สูงอายุ

**ระยะเวลา:** 12 เดือน

**KPIs:**
- Working Prototype สามารถสื่อสารภาษาไทยได้ภายใน 6 เดือน
- Production-ready design ภายใน 12 เดือน
- ผ่านการทดสอบ UAT ≥ 50 คน, ความพึงพอใจ ≥ 3.5/5

---

### กิจกรรม R&D #2: Explainable AI (XAI) Framework
**งบประมาณ: 11,240,000 บาท**

**รายละเอียด:**
- วิจัยและพัฒนา XAI algorithms สำหรับอธิบายการตัดสินใจของ AI
- สร้าง Visualization tools สำหรับแสดงผลการวิเคราะห์
- พัฒนา Natural Language Generation (NLG) ภาษาไทยสำหรับคำอธิบาย
- Integration กับ Multi-Agent Platform
- ทดสอบความเข้าใจของผู้ใช้ (User comprehension testing)

**ระยะเวลา:** 10 เดือน

**KPIs:**
- XAI Module ที่สามารถอธิบายการตัดสินใจได้ 100%
- ผู้ใช้เข้าใจคำอธิบาย ≥ 80% (จากการทดสอบ)
- รองรับ 100,000 การวิเคราะห์/วัน

---

### กิจกรรม R&D #3: Thai NLP Model for Finance
**งบประมาณ: 15,780,000 บาท**

**รายละเอียด:**
- รวบรวม Thai Financial Corpus และ Scam Dataset (≥ 10,000 ตัวอย่าง)
- Pre-processing และ Data cleaning
- ฝึก Deep Learning models (Transformer-based: BERT, GPT variants for Thai)
- Fine-tuning สำหรับ scam detection และ financial Q&A
- Model optimization สำหรับ inference speed
- Deployment บน Cloud และ Edge devices

**ระยะเวลา:** 12 เดือน

**KPIs:**
- F1 Score ≥ 80% ภายใน 6 เดือน
- F1 Score ≥ 85% ภายใน 12 เดือน
- False Positive Rate ≤ 10%
- Response Time < 200ms
- Dataset คุณภาพ ≥ 75% (จาก human evaluation)

---

### กิจกรรม R&D #4: Fraud Pattern Recognition
**งบประมาณ: 10,500,000 บาท**

**รายละเอียด:**
- วิจัยและสร้าง Pattern Mining algorithms สำหรับตรวจจับรูปแบบการฉ้อโกง
- พัฒนา Anomaly Detection models
- สร้าง Real-time Alert System
- Integration กับ Banking APIs และ Transaction data
- ทดสอบกับข้อมูลจริง (ภายใต้ NDA)

**ระยะเวลา:** 10 เดือน

**KPIs:**
- ตรวจจับ scam patterns ได้ ≥ 85% accuracy
- False Positive ≤ 10%
- Real-time alert < 1 วินาที

---

### กิจกรรม R&D #5: Multi-Agent System Integration
**งบประมาณ: 8,500,000 บาท**

**รายละเอียด:**
- ออกแบบ Multi-Agent Architecture (5 Agents)
- พัฒนา Message Passing Framework
- สร้าง Coordination และ Collaboration mechanisms
- Integration ของทั้ง 5 AI Agents
- Performance optimization (Scalability, Reliability)
- Load testing และ Stress testing

**ระยะเวลา:** 12 เดือน

**KPIs:**
- Message Passing accuracy ≥ 99%
- รองรับ 50,000 Concurrent Users
- System Uptime ≥ 99.5%
- Throughput ≥ 500 TPS (Transactions Per Second)

---

## ส่วนที่ 3: สรุปงบประมาณกิจกรรม R&D

| # | กิจกรรม R&D | งบประมาณ (บาท) | % ของ R&D | ระยะเวลา |
|---|-------------|----------------|-----------|----------|
| 1 | Physical AI Agent Development | 12,980,000 | 22.0% | 12 เดือน |
| 2 | Explainable AI (XAI) Framework | 11,240,000 | 19.1% | 10 เดือน |
| 3 | Thai NLP Model for Finance | 15,780,000 | 26.7% | 12 เดือน |
| 4 | Fraud Pattern Recognition | 10,500,000 | 17.8% | 10 เดือน |
| 5 | Multi-Agent System Integration | 8,500,000 | 14.4% | 12 เดือน |
| **รวมทั้งสิ้น** | | **59,000,000** | **100.0%** | |

---

## ส่วนที่ 4: หมวดหมู่ค่าใช้จ่าย R&D

| # | หมวดหมู่ | งบประมาณ (บาท) | % ของ R&D | อ้างอิง AOR3/2568 |
|---|----------|----------------|-----------|-------------------|
| 1 | **บุคลากรวิจัย (Personnel)** | 36,240,000 | 61.4% | หน้า 5 ข้อ 1 |
|   | - AI Scientists (PhD): 2 คน | | | |
|   | - Senior Engineers (Master's): 5 คน | | | |
|   | - Developers (Bachelor's): 8 คน | | | |
|   | - รวม: 15 คน | | | |
| 2 | **อุปกรณ์และเครื่องมือ (Equipment)** | 9,000,000 | 15.3% | หน้า 5 ข้อ 2 |
|   | - GPU Servers (NVIDIA A100) | | | |
|   | - Development Workstations | | | |
|   | - Physical AI prototyping equipment | | | |
|   | - Testing devices | | | |
| 3 | **ข้อมูลและ Dataset (Data)** | 8,500,000 | 14.4% | หน้า 5 ข้อ 3 |
|   | - Thai Financial Corpus | | | |
|   | - Scam Dataset Collection | | | |
|   | - Data labeling services | | | |
|   | - Data storage and management | | | |
| 4 | **อื่นๆ (Other Expenses)** | 5,260,000 | 8.9% | หน้า 5 ข้อ 4-6 |
|   | - Software licenses for R&D tools | | | |
|   | - Cloud computing for training | | | |
|   | - Testing and validation | | | |
|   | - Documentation and reporting | | | |
| **รวมทั้งสิ้น** | | **59,000,000** | **100.0%** | |

**หมายเหตุ:** ค่าใช้จ่ายทั้งหมดสอดคล้องกับ AOR3/2568 หน้า 5 (รายการค่าใช้จ่ายที่ได้รับการสนับสนุนสำหรับ R&D)

---

## ส่วนที่ 5: เครื่องจักรและอุปกรณ์สำหรับ R&D

| # | รายการ | จำนวน | มูลค่า (บาท) | รวม (บาท) |
|---|--------|-------|--------------|-----------|
| 1 | GPU Server (NVIDIA A100 80GB) สำหรับ Deep Learning Training | 4 เครื่อง | 1,500,000 | 6,000,000 |
| 2 | Development Workstations (High-end) | 15 เครื่อง | 100,000 | 1,500,000 |
| 3 | Physical AI Prototyping Equipment (3D Printers, Electronics Assembly) | 1 ชุด | 800,000 | 800,000 |
| 4 | Testing Devices (Smartphones, Tablets, IoT sensors) | 50 เครื่อง | 14,000 | 700,000 |
| **รวมทั้งสิ้น (ประมาณการ)** | | | | **9,000,000** |

---

## ส่วนที่ 6: Timeline และ Milestones

### งวดที่ 1: เดือน 1-6 (1 มิ.ย. 2569 - 30 พ.ย. 2569)

**เงินสนับสนุน:** 14,750,000 บาท (50% ของ R&D Support 29.5M)

**กิจกรรมหลัก:**
- Physical AI Agent: Prototype แรก (3 AI Agents working)
- Thai NLP: Model Training เบื้องต้น (F1 Score ≥ 80%)
- Multi-Agent: Framework พื้นฐาน (Message Passing ≥ 99%)
- Integration Testing
- บุคลากร: ฝึกอบรม ≥ 80%

---

### งวดที่ 2: เดือน 7-12 (1 ธ.ค. 2569 - 31 พ.ค. 2570)

**เงินสนับสนุน:** 14,750,000 บาท (50% ของ R&D Support 29.5M)

**กิจกรรมหลัก:**
- Production System: ครบ 5 AI Agents, Uptime ≥ 99.5%
- Thai NLP: F1 Score ≥ 85%, False Positive ≤ 10%
- Performance: Response Time < 200ms, Throughput ≥ 500 TPS
- ยื่นสิทธิบัตร ≥ 2 ฉบับ
- ตีพิมพ์วิชาการ ≥ 3 ฉบับ
- UAT กับ ≥ 50 คน
- Thai Scam Dataset ≥ 10,000 ตัวอย่าง

---

## ลงนาม

**ลงชื่อ:** ________________________________

**นายสุรเกียรติ ยาวะโนภาส**
กรรมการผู้จัดการ
บริษัท อัสลาน เวลธ์ จำกัด

**วันที่:** ________________
