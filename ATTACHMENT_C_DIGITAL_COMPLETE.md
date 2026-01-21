# เอกสารแนบ C
## รายละเอียดการใช้เทคโนโลยีดิจิทัล

---

**ชื่อโครงการ:** ASLAN AI GUARDIAN PLATFORM
**บริษัท:** บริษัท อัสลาน เวลธ์ จำกัด (ASLAN WEALTH COMPANY LIMITED)
**เลขทะเบียน:** 0105567220030
**งบประมาณเทคโนโลยีดิจิทัล:** **26,000,000 บาท**

---

## ส่วนที่ 1: วัตถุประสงค์การใช้เทคโนโลยีดิจิทัล

### 1.1 วัตถุประสงค์หลัก

โครงการลงทุนในเทคโนโลยีดิจิทัลเพื่อสนับสนุนการพัฒนาและให้บริการระบบ ASLAN AI GUARDIAN PLATFORM ดังนี้:

**1. สร้าง Cloud Infrastructure ที่ปลอดภัยและ Scalable**
- Deploy บน Cloud Platform (AWS/Google Cloud/Azure) สำหรับรองรับผู้ใช้จำนวนมาก
- Auto-scaling สำหรับรองรับ Peak load (50,000 concurrent users)
- High Availability (HA) และ Disaster Recovery (DR)
- CDN สำหรับเร่งความเร็วในการเข้าถึง

**2. สร้างระบบ On-premise Hardware สำหรับ Critical Operations**
- Production Servers สำหรับ Mission-critical services
- Database Servers สำหรับเก็บข้อมูลผู้ใช้
- Kubernetes Cluster สำหรับ Container orchestration
- Network infrastructure (Firewalls, Load balancers, Switches)

**3. จัดหาซอฟต์แวร์และเครื่องมือสำหรับการพัฒนา**
- Development Tools และ IDEs
- Database Licenses (Enterprise edition)
- Monitoring และ Analytics tools
- Project Management และ Collaboration platforms

**4. สร้างระบบรักษาความปลอดภัย (Security & Compliance)**
- Firewall และ Intrusion Detection/Prevention Systems (IDS/IPS)
- Data Encryption (at rest และ in transit)
- Identity และ Access Management (IAM)
- Security Auditing และ Compliance tools

**5. พัฒนาระบบ Integration และ APIs**
- API Gateway สำหรับเชื่อมต่อกับ Banking Systems
- Message Queue systems (Kafka, RabbitMQ)
- Integration middleware
- DevOps tools สำหรับ CI/CD pipeline

### 1.2 ผลที่คาดว่าจะได้รับ

- Cloud Infrastructure ที่รองรับ 100,000+ requests/day
- System Uptime ≥ 99.5% (SLA)
- Response Time < 200ms สำหรับ API calls
- ผ่าน Security Audit และได้รับใบรับรองจาก depa
- ระบบ Backup และ Disaster Recovery ที่ครบถ้วน
- ระบบ Monitoring แบบ Real-time สำหรับตรวจสอบ Performance

---

## ส่วนที่ 2: รายละเอียดเทคโนโลยีดิจิทัล (5 หมวดหมู่)

### หมวดที่ 1: Cloud Services & Infrastructure (5,000,000 บาท)

**รายละเอียด:**

**Cloud Platform Subscription (AWS/GCP/Azure):**
- Compute Instances (EC2, Compute Engine) สำหรับ Application Servers
- Managed Database Services (RDS, Cloud SQL)
- Object Storage (S3, Cloud Storage) สำหรับ Files และ Backups
- CDN Services สำหรับเร่งความเร็ว

**AI/ML Services:**
- GPU Instances สำหรับ Model inference
- Managed AI/ML platforms
- Data Pipeline services

**Monitoring & Analytics:**
- CloudWatch, Stackdriver
- Application Performance Monitoring (APM)
- Log Analytics

**ระยะเวลา:** 12 เดือน (Subscription)
**เป้าหมาย:** รองรับ 100,000+ requests/day, Auto-scaling

---

### หมวดที่ 2: On-premise Hardware & Servers (8,000,000 บาท)

**รายละเอียด:**

**Production Servers:**
- Application Servers (High-performance) x 4 เครื่อง
- Database Servers (Enterprise-grade) x 2 เครื่อง
- Backup Servers x 2 เครื่อง

**Kubernetes Cluster:**
- Master Nodes x 3 เครื่อง
- Worker Nodes x 10 เครื่อง

**Network Equipment:**
- Enterprise Firewall x 2 ชุด (HA)
- Load Balancer x 2 ชุด
- Core Switch x 4 ตัว
- UPS และ Power Management

**Storage Systems:**
- SAN/NAS Storage (Enterprise) ≥ 100TB
- Backup Storage ≥ 200TB

**เป้าหมาย:** Uptime ≥ 99.5%, High Availability

---

### หมวดที่ 3: Software Licenses & Development Tools (3,000,000 บาท)

**รายละเอียด:**

**Database Licenses:**
- PostgreSQL Enterprise (Support)
- MongoDB Enterprise
- Redis Enterprise

**Development Tools:**
- JetBrains Suite (IntelliJ, PyCharm, etc.) x 15 licenses
- Visual Studio Enterprise x 15 licenses
- GitHub Enterprise (for code repository)

**Monitoring & Analytics:**
- Datadog/New Relic (APM)
- Splunk/ELK Stack (Log management)
- Grafana/Prometheus

**Project Management:**
- Jira Software (Project tracking)
- Confluence (Documentation)
- Slack Enterprise (Communication)

**ระยะเวลา:** 12 เดือน (Annual licenses)

---

### หมวดที่ 4: Security & Compliance Systems (2,000,000 บาท)

**รายละเอียด:**

**Security Hardware:**
- Next-Gen Firewall (Palo Alto/Fortinet)
- IDS/IPS Systems
- Web Application Firewall (WAF)

**Security Software:**
- Endpoint Protection (Antivirus, Anti-malware)
- SIEM (Security Information and Event Management)
- Vulnerability Scanner
- Penetration Testing Tools

**Data Protection:**
- Encryption Software (TLS, AES-256)
- Data Loss Prevention (DLP)
- Backup and Recovery solutions

**Compliance Tools:**
- Audit logging systems
- Compliance monitoring tools
- Security certification preparation

**เป้าหมาย:** ผ่าน Security Audit, รับใบรับรองจาก depa

---

### หมวดที่ 5: System Integration & API Development (8,000,000 บาท)

**รายละเอียด:**

**API Gateway & Management:**
- Kong/Apigee API Gateway
- API Management platform
- Rate limiting และ Throttling

**Message Queue & Streaming:**
- Apache Kafka Cluster
- RabbitMQ for messaging
- Event streaming infrastructure

**Integration Middleware:**
- ESB (Enterprise Service Bus)
- ETL Tools สำหรับ Data transformation
- Banking API Integration modules

**DevOps & CI/CD:**
- Jenkins/GitLab CI pipelines
- Docker Registry
- Kubernetes (K8s) management tools
- Terraform for Infrastructure as Code

**Professional Services:**
- System Integration consulting
- Technical support และ Maintenance
- Training สำหรับทีมพัฒนา

**เป้าหมาย:** API Response Time < 200ms, Throughput ≥ 500 TPS

---

## ส่วนที่ 3: สรุปงบประมาณเทคโนโลยีดิจิทัล

| # | หมวดหมู่ | งบประมาณ (บาท) | % ของ Digital | อ้างอิง AOR3 |
|---|----------|----------------|---------------|--------------|
| 1 | Cloud Services & Infrastructure | 5,000,000 | 19.2% | หน้า 6 ข้อ 1 |
| 2 | On-premise Hardware & Servers | 8,000,000 | 30.8% | หน้า 6 ข้อ 2 |
| 3 | Software Licenses & Development Tools | 3,000,000 | 11.5% | หน้า 6 ข้อ 3 |
| 4 | Security & Compliance Systems | 2,000,000 | 7.7% | หน้า 6 ข้อ 4 |
| 5 | System Integration & API Development | 8,000,000 | 30.8% | หน้า 6 ข้อ 5 |
| **รวม** | **รวมทั้งสิ้น** | **26,000,000** | **100.0%** | |

**หมายเหตุ:** ค่าใช้จ่ายทั้งหมดสอดคล้องกับ AOR3/2568 หน้า 6 (รายการค่าใช้จ่ายที่ได้รับการสนับสนุนสำหรับเทคโนโลยีดิจิทัล)

---

## ส่วนที่ 4: รายการ Hardware และ Infrastructure หลัก

| # | รายการ | จำนวน | มูลค่า/หน่วย (บาท) | รวม (บาท) |
|---|--------|-------|-------------------|-----------|
| 1 | Application Server (High-performance, Dual Xeon, 256GB RAM) | 4 | 500,000 | 2,000,000 |
| 2 | Database Server (Enterprise-grade, 512GB RAM, NVMe Storage) | 2 | 800,000 | 1,600,000 |
| 3 | Kubernetes Nodes (Master: 3, Worker: 10) | 13 | 150,000 | 1,950,000 |
| 4 | Next-Gen Firewall (Palo Alto PA-3260 or equivalent) | 2 | 600,000 | 1,200,000 |
| 5 | Load Balancer (F5 or equivalent) | 2 | 400,000 | 800,000 |
| 6 | Enterprise Storage (SAN, 100TB usable) | 1 | 1,500,000 | 1,500,000 |
| 7 | Network Switches (Core, 10Gb) | 4 | 150,000 | 600,000 |
| 8 | UPS & Power Management (Redundant) | 2 | 300,000 | 600,000 |
| **รวม** | **รวมทั้งสิ้น (ประมาณการหลัก)** | | | **10,250,000** |

*หมายเหตุ: รายการข้างต้นเป็นส่วนหลัก ยังมี Hardware เสริมและอุปกรณ์เครือข่ายเพิ่มเติม*

---

## ส่วนที่ 5: สถาปัตยกรรมระบบ (System Architecture)

### 5.1 Infrastructure Layers

**1. Presentation Layer (Frontend)**
- Mobile Apps (iOS/Android) - Flutter/React Native
- Physical AI Agent Interface - Embedded Linux
- Web Dashboard - React.js
- CDN - CloudFront/Cloud CDN

**2. API Gateway Layer**
- Kong/Apigee API Gateway
- Rate Limiting & Throttling
- Authentication & Authorization (OAuth 2.0, JWT)
- SSL/TLS Termination

**3. Application Layer (Backend Services)**
- 5 AI Agents (Microservices)
- Thai NLP Service
- XAI Service
- Fraud Detection Service
- User Management Service
- Notification Service

**4. Data Layer**
- PostgreSQL (Primary Database)
- MongoDB (Document Store)
- Redis (Cache & Session)
- Elasticsearch (Search & Analytics)
- Kafka (Event Streaming)

**5. Infrastructure Layer**
- Kubernetes Cluster (Container Orchestration)
- Cloud Services (AWS/GCP/Azure)
- On-premise Servers (Hybrid)
- Monitoring (Prometheus, Grafana, Datadog)
- Logging (ELK Stack)

### 5.2 Security Layers

- **Network Security:** Firewall, IDS/IPS, WAF
- **Application Security:** Input validation, XSS/CSRF protection
- **Data Security:** Encryption at rest (AES-256), in transit (TLS 1.3)
- **Access Control:** RBAC, MFA
- **Audit & Compliance:** SIEM, Log management

---

## ส่วนที่ 6: เป้าหมาย Performance และ Scalability

| Metric | เป้าหมาย | วิธีการวัด |
|--------|---------|-----------|
| System Uptime | ≥ 99.5% | Monitoring tools (24/7) |
| API Response Time | < 200ms (p95) | APM (Datadog/New Relic) |
| Concurrent Users | 50,000 users | Load testing (JMeter/Gatling) |
| Throughput | ≥ 500 TPS | Performance monitoring |
| Database Queries | < 50ms (avg) | Database profiling |
| Security Scans | Weekly | Vulnerability scanner |
| Backup Frequency | Daily (Full), Hourly (Incremental) | Backup monitoring |
| Recovery Time Objective (RTO) | < 4 hours | DR testing (quarterly) |

---

## ส่วนที่ 7: Timeline การติดตั้งและ Deployment

### งวดที่ 1: เดือน 1-6 (1 มิ.ย. 2569 - 30 พ.ย. 2569)

**เงินสนับสนุน:** 6,500,000 บาท (50% ของ Digital Support 13M)

**กิจกรรม:**
- จัดซื้อและติดตั้ง On-premise Hardware (50%)
- ตั้งค่า Cloud Infrastructure
- ติดตั้ง Basic Security systems
- Setup Development environment
- Deploy Beta version

---

### งวดที่ 2: เดือน 7-12 (1 ธ.ค. 2569 - 31 พ.ค. 2570)

**เงินสนับสนุน:** 6,500,000 บาท (50% ของ Digital Support 13M)

**กิจกรรม:**
- จัดซื้อและติดตั้ง Hardware ที่เหลือ
- Complete Production deployment
- Full Security implementation
- Integration testing
- Performance tuning
- ผ่าน Security Audit และได้รับใบรับรองจาก depa
- Go Live บริการจริง

---

## ส่วนที่ 8: ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด

*(พิจารณากำหนดตามความเหมาะสมของโครงการ - รวม R&D + Digital)*

---

### 8.1 สรุปภาพรวมการลงทุนและเงินสนับสนุนทั้งโครงการ

| หมวดการลงทุน | งบประมาณลงทุน (ล้านบาท) | อัตราสนับสนุน | เงินสนับสนุนสูงสุด (ล้านบาท) | หน่วยรับรอง |
|-------------|------------------------|--------------|---------------------------|------------|
| **R&D (เอกสารแนบ A)** | 59.0 | 50% | 29.5 | สวทช. |
| **Digital Technology (เอกสารแนบ C)** | 26.0 | 50% | 13.0 | depa |
| **รวมทั้งโครงการ** | **85.0** | **50%** | **42.5** | |

---

### 8.2 ตัวชี้วัดและมูลค่าเงินสนับสนุนรายงวด

| งวดงาน | ตัวชี้วัด | ประมาณการลงทุน (ล้านบาท) | สัดส่วน (%) | เงินสนับสนุน (ล้านบาท)* |
|--------|----------|-------------------------|------------|------------------------|
| **งวดที่ 1** (6 เดือน) | ดูรายละเอียดด้านล่าง | 42.5 (R&D: 29.5 + Digital: 13.0) | 50% | 21.25 (R&D: 14.75 + Digital: 6.5) |
| **งวดที่ 2** (12 เดือน) | ดูรายละเอียดด้านล่าง | 42.5 (R&D: 29.5 + Digital: 13.0) | 50% | 21.25 (R&D: 14.75 + Digital: 6.5) |
| **รวมทั้งสิ้น** | | **85.0** | **100%** | **42.5** |

---

### งวดที่ 1: ตัวชี้วัดความคืบหน้าของโครงการ (ภายใน 6 เดือน)

**1. ตัวชี้วัดความคืบหน้าของโครงการ:**
*(ดูตัวอย่างจากคู่มือการกรอกขอเสนอโครงการ ข้อ 1)*

**ส่วน R&D (สวทช.):**
- Physical AI Agent: Prototype แรกสำเร็จ (3 AI Agents working)
- Thai NLP: Model Training เบื้องต้น (F1 Score ≥ 80%)
- Multi-Agent: Framework พื้นฐาน (Message Passing ≥ 99%)
- จัดซื้ออุปกรณ์ R&D 50% (GPU Servers, Workstations)
- บุคลากรวิจัย: ฝึกอบรม ≥ 80%

**ส่วน Digital (depa):**
- จัดซื้อและติดตั้ง On-premise Hardware 50% (Servers, Network)
- ตั้งค่า Cloud Infrastructure (AWS/GCP) พร้อมใช้งาน
- ติดตั้ง Basic Security Systems (Firewall, IDS/IPS)
- Setup Development Environment
- Deploy Beta version บน Staging

**2. มีการลงทุนจริงตามขอบข่ายที่ได้รับอนุมัติ ในสัดส่วนไม่น้อยกว่าร้อยละ 50** ของมูลค่าการลงทุน พร้อมหลักฐานการชำระเงิน เช่น ใบเสร็จรับเงิน หรือหลักฐานการโอนเงิน เป็นต้น

**3. มีหนังสือรับรองผลการดำเนินการตามตัวชี้วัดจาก สวทช. หรือ depa** (ตามขอบข่ายการดำเนินการ)

---

### งวดที่ 2: ตัวชี้วัดผลการดำเนินการตามเป้าหมาย (ภายใน 12 เดือน)

**1. ตัวชี้วัดผลการดำเนินการตามเป้าหมาย:**

**ส่วน R&D (สวทช.):**
- Production System: ครบ 5 AI Agents, Uptime ≥ 99.5%
- Thai NLP: F1 Score ≥ 85%, False Positive ≤ 10%
- Performance: Response Time < 200ms, Throughput ≥ 500 TPS
- ยื่นสิทธิบัตร ≥ 2 ฉบับ
- ตีพิมพ์บทความวิชาการ ≥ 3 ฉบับ
- UAT กับผู้ใช้จริง ≥ 50 คน
- Thai Scam Dataset ≥ 10,000 ตัวอย่าง

**ส่วน Digital (depa):**
- จัดซื้อและติดตั้ง Hardware ครบถ้วน 100%
- Production Deployment รองรับ 50,000 Concurrent Users
- System Uptime ≥ 99.5%
- API Response Time < 200ms (p95)
- ผ่าน Security Audit จากหน่วยงานรับรอง
- Go Live บริการจริง

**2. มีการลงทุนจริงครบถ้วน** ตามขอบข่ายที่ได้รับอนุมัติ พร้อมหลักฐานการชำระเงิน เช่น ใบเสร็จรับเงิน หรือหลักฐานการโอนเงิน เป็นต้น

**3. มีหนังสือรับรองผลการดำเนินการตามตัวชี้วัดจาก สวทช. หรือ depa** (ตามขอบข่ายการดำเนินการ)

---

### 8.3 สรุปการเบิกจ่ายเงินสนับสนุนรายงวด (แยกตามหมวด)

| รายการ | R&D ลงทุน | R&D สนับสนุน | Digital ลงทุน | Digital สนับสนุน | รวมลงทุน | รวมสนับสนุน |
|--------|----------|-------------|--------------|-----------------|---------|------------|
| งบประมาณทั้งหมด | 59.0 | 29.5 | 26.0 | 13.0 | **85.0** | **42.5** |
| งวดที่ 1 (50%) | 29.5 | 14.75 | 13.0 | 6.5 | **42.5** | **21.25** |
| งวดที่ 2 (50%) | 29.5 | 14.75 | 13.0 | 6.5 | **42.5** | **21.25** |

*หน่วย: ล้านบาท*

---

### 8.4 รายการ Hardware และอุปกรณ์ที่ได้รับการสนับสนุน

#### อุปกรณ์ R&D (เอกสารแนบ A) - รับรองโดย สวทช.

| # | รายการ | มูลค่าลงทุน (บาท) | สนับสนุน 50% (บาท) |
|---|--------|------------------|-------------------|
| 1 | GPU Server (NVIDIA A100 80GB) x 4 เครื่อง | 6,000,000 | 3,000,000 |
| 2 | Development Workstations x 15 เครื่อง | 1,500,000 | 750,000 |
| 3 | Physical AI Prototyping Equipment | 800,000 | 400,000 |
| 4 | Testing Devices (Smartphones, Tablets, IoT) x 50 | 700,000 | 350,000 |
| **รวม Hardware R&D** | | **9,000,000** | **4,500,000** |

#### อุปกรณ์ Digital (เอกสารแนบ C) - รับรองโดย depa

| # | รายการ | มูลค่าลงทุน (บาท) | สนับสนุน 50% (บาท) |
|---|--------|------------------|-------------------|
| 5 | Application Server (Dual Xeon, 256GB RAM) x 4 | 2,000,000 | 1,000,000 |
| 6 | Database Server (512GB RAM, NVMe) x 2 | 1,600,000 | 800,000 |
| 7 | Kubernetes Nodes (Master 3 + Worker 10) | 1,950,000 | 975,000 |
| 8 | Next-Gen Firewall (Palo Alto) x 2 | 1,200,000 | 600,000 |
| 9 | Load Balancer (F5) x 2 | 800,000 | 400,000 |
| 10 | Enterprise Storage (SAN, 100TB) | 1,500,000 | 750,000 |
| 11 | Network Switches (Core, 10Gb) x 4 | 600,000 | 300,000 |
| 12 | UPS & Power Management x 2 | 600,000 | 300,000 |
| **รวม Hardware Digital** | | **10,250,000** | **5,125,000** |

#### สรุป Hardware ทั้งหมด

| หมวด | มูลค่าลงทุน | สนับสนุน 50% |
|------|------------|-------------|
| R&D (สวทช.) | 9,000,000 | 4,500,000 |
| Digital (depa) | 10,250,000 | 5,125,000 |
| **รวม Hardware ทั้งหมด** | **19,250,000** | **9,625,000** |

**\* หมายเหตุ:** จะได้รับอนุมัติเบิกจ่ายเงินสนับสนุนไม่เกินมูลค่าเงินที่ชำระจริง

---

## ลงนาม

**ลงชื่อ:** ________________________________

**นายสุรเกียรติ ยาวะโนภาส**
กรรมการผู้จัดการ
บริษัท อัสลาน เวลธ์ จำกัด

**วันที่:** ________________

---

## ข้อมูลสำหรับกรอกในแบบฟอร์ม

### ติ๊กถูก ✓ ในฟอร์ม:
- ☑ **การใช้เทคโนโลยีดิจิทัล เพื่อปรับปรุงประสิทธิภาพด้วยเทคโนโลยีที่ทันสมัย** (หน่วยตรวจสอบ: depa)
- ☑ **การประยุกต์ใช้ปัญญาประดิษฐ์ (AI) Machine Learning การนำ Big Data มาใช้หรือการวิเคราะห์ข้อมูล**

### ตารางข้อ 2 (สาระสำคัญของการนำเทคโนโลยีดิจิทัลมาใช้):

| ขั้นตอนการทำงาน | ชื่อเทคโนโลยีดิจิทัล | หน่วยงานที่เกี่ยวข้อง | อธิบายการเชื่อมโยงข้อมูล | ประสิทธิภาพที่เพิ่มขึ้น |
|----------------|---------------------|---------------------|------------------------|----------------------|
| 1. AI Model Training | GPU Cloud Infrastructure (AWS/GCP) | ทีม AI Research, Cloud Provider | เชื่อมต่อ Training Pipeline กับ Data Lake, Model Registry | Training time ลด 60%, Cost ลด 40% |
| 2. NLP Processing | Thai NLP Engine + LLM APIs | ทีม NLP, OpenAI/Anthropic | API Integration กับ Multi-Agent Platform | Accuracy 85%, Response <200ms |
| 3. Multi-Agent Orchestration | Kubernetes + Message Queue | ทีม Platform, DevOps | Service Mesh เชื่อมต่อ 5 AI Agents | Uptime 99.5%, Throughput 500 TPS |
| 4. Data Analytics | Big Data Pipeline (Spark/Kafka) | ทีม Data, Analytics | Real-time streaming จาก User interactions | Process 1M events/day |
| 5. Security & Monitoring | SIEM + WAF + Monitoring Stack | ทีม Security, IT | Centralized logging, Alert system | Threat detection <5 min |

---

*เอกสารนี้เป็นส่วนหนึ่งของการขอรับการส่งเสริมการลงทุนจาก BOI*
*จัดทำโดย: บริษัท อัสลาน เวลธ์ จำกัด*
*วันที่: มกราคม 2569*
