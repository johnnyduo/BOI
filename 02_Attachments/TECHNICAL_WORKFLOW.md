# ASLAN AI GUARDIAN PLATFORM - Technical Workflow

## BOI Submission & Project Implementation Workflow

---

## 1. BOI Document Submission Flow

```mermaid
flowchart TD
    subgraph PREPARATION["Phase 1: Document Preparation"]
        A1[Company Documents] --> A2[Project Proposal]
        A2 --> A3[Attachment A: R&D]
        A2 --> A4[Attachment C: Digital]
        A2 --> A5[Attachment D: Products]
        A3 --> A6[NDA Agreement]
        A4 --> A6
        A5 --> A6
    end

    subgraph SUBMISSION["Phase 2: BOI Submission"]
        B1[BOI Online Portal] --> B2{Document Review}
        B2 -->|Complete| B3[Application Accepted]
        B2 -->|Incomplete| B4[Return for Revision]
        B4 --> A2
    end

    subgraph EVALUATION["Phase 3: BOI Evaluation"]
        C1[Technical Review] --> C2[NSTDA: R&D Verification]
        C1 --> C3[depa: Digital Verification]
        C2 --> C4{Approval Decision}
        C3 --> C4
        C4 -->|Approved| C5[BOI Certificate Issued]
        C4 -->|Rejected| C6[Appeal/Resubmit]
    end

    A6 --> B1
    B3 --> C1
    C5 --> D1[Project Kickoff]

    style PREPARATION fill:#e1f5fe
    style SUBMISSION fill:#fff3e0
    style EVALUATION fill:#e8f5e9
```

---

## 2. Required Documents Checklist

```mermaid
flowchart LR
    subgraph CORE["Core Documents"]
        D1["1. Project Proposal<br/>ข้อเสนอโครงการ"]
        D2["2. Attachments<br/>เอกสารแนบ"]
    end

    subgraph ATTACHMENTS["Attachment Details"]
        A1["A: R&D Activities<br/>กิจกรรมวิจัยและพัฒนา<br/>59M THB"]
        A2["C: Digital Technology<br/>เทคโนโลยีดิจิทัล<br/>26M THB"]
        A3["Process Flow Diagram<br/>Before/After Digital"]
    end

    subgraph LEGAL["Legal Documents"]
        L1["3. NDA Agreement<br/>บันทึกข้อตกลงฯ<br/>2 copies"]
        L2["4. Power of Attorney<br/>หนังสือมอบอำนาจ<br/>if applicable"]
        L3["5. ID Copies<br/>สำเนาบัตรประชาชน"]
    end

    D2 --> A1
    D2 --> A2
    A2 --> A3

    style CORE fill:#bbdefb
    style ATTACHMENTS fill:#c8e6c9
    style LEGAL fill:#ffecb3
```

---

## 3. Project Implementation Timeline (12 Months)

```mermaid
gantt
    title ASLAN AI GUARDIAN - 12 Month Project Timeline
    dateFormat  YYYY-MM
    section Phase 1 (Month 1-6)
    R&D Activity 1: Physical AI Agent Development    :a1, 2026-06, 6M
    R&D Activity 2: XAI Framework Development        :a2, 2026-06, 5M
    R&D Activity 3: Thai NLP Model Training          :a3, 2026-07, 5M
    Milestone 1: Working Prototype (3 Agents)        :milestone, m1, 2026-11, 1d

    section Phase 2 (Month 7-12)
    R&D Activity 4: Fraud Pattern Recognition        :a4, 2026-11, 4M
    R&D Activity 5: Multi-Agent Integration          :a5, 2026-12, 4M
    Digital Infrastructure Deployment                :d1, 2027-01, 4M
    UAT & Testing                                    :t1, 2027-03, 2M
    Milestone 2: Production System (5 Agents)        :milestone, m2, 2027-05, 1d
```

---

## 4. Multi-Agent AI Architecture

```mermaid
flowchart TB
    subgraph INPUT["Input Layer"]
        I1[Voice Input<br/>Thai Speech]
        I2[Text Input<br/>Thai Text]
        I3[Transaction Data<br/>Financial Records]
        I4[Behavioral Data<br/>User Patterns]
    end

    subgraph AGENTS["5 AI Agents Layer"]
        AG1["Agent 1: NLP & Voice<br/>Thai Language Processing<br/>Speech Recognition"]
        AG2["Agent 2: Quantitative Risk<br/>Risk Scoring<br/>Financial Analysis"]
        AG3["Agent 3: Behavioral<br/>Pattern Detection<br/>Anomaly Detection"]
        AG4["Agent 4: Fraud Pattern<br/>Scam Detection<br/>Threat Intelligence"]
        AG5["Agent 5: Meta-Decision<br/>Consensus Building<br/>Explainable Output"]
    end

    subgraph OUTPUT["Output Layer"]
        O1[Physical AI Agent<br/>Aslan Mascot]
        O2[Mobile App<br/>Dashboard & Alerts]
        O3[API Services<br/>Enterprise Integration]
    end

    I1 --> AG1
    I2 --> AG1
    I3 --> AG2
    I4 --> AG3

    AG1 --> AG5
    AG2 --> AG5
    AG3 --> AG5
    AG4 --> AG5

    AG5 --> O1
    AG5 --> O2
    AG5 --> O3

    style INPUT fill:#e3f2fd
    style AGENTS fill:#f3e5f5
    style OUTPUT fill:#e8f5e9
```

---

## 5. Budget Allocation Flow

```mermaid
pie showData
    title Total Investment: 85M THB
    "R&D Personnel (36.24M)" : 36.24
    "R&D Equipment (9M)" : 9
    "R&D Dataset (8.5M)" : 8.5
    "R&D Other (5.26M)" : 5.26
    "Digital Cloud (5M)" : 5
    "Digital Hardware (8M)" : 8
    "Digital Software (3M)" : 3
    "Digital Security (2M)" : 2
    "Digital Integration (8M)" : 8
```

---

## 6. R&D Activities Breakdown

```mermaid
flowchart LR
    subgraph RD["R&D Activities: 59M THB"]
        R1["R&D #1<br/>Physical AI Agent<br/>12.98M THB"]
        R2["R&D #2<br/>XAI Framework<br/>11.24M THB"]
        R3["R&D #3<br/>Thai NLP Model<br/>15.78M THB"]
        R4["R&D #4<br/>Fraud Pattern<br/>10.5M THB"]
        R5["R&D #5<br/>Multi-Agent Integration<br/>8.5M THB"]
    end

    R1 --> P1[Physical AI<br/>Prototype]
    R2 --> P2[Explainable AI<br/>Module]
    R3 --> P3[Thai NLP<br/>Engine]
    R4 --> P4[Fraud Detection<br/>System]
    R5 --> P5[Integrated<br/>Platform]

    P1 --> FINAL[Production System<br/>5 AI Agents]
    P2 --> FINAL
    P3 --> FINAL
    P4 --> FINAL
    P5 --> FINAL

    style RD fill:#e8eaf6
    style FINAL fill:#c8e6c9
```

---

## 7. Digital Technology Implementation

```mermaid
flowchart TB
    subgraph BEFORE["Before Digital Implementation"]
        B1[Manual Scam Detection<br/>Human Review Only]
        B2[No Real-time Alerts<br/>Delayed Response]
        B3[Limited Thai NLP<br/>Foreign AI Only]
        B4[No Financial AI<br/>Basic Tools]
    end

    subgraph DIGITAL["Digital Technology: 26M THB"]
        D1["Cloud Infrastructure<br/>AWS/Azure GPU<br/>5M THB"]
        D2["On-premise Hardware<br/>GPU Servers<br/>8M THB"]
        D3["Software & AI Tools<br/>TensorFlow, PyTorch<br/>3M THB"]
        D4["Security Systems<br/>Encryption, Monitoring<br/>2M THB"]
        D5["Integration & API<br/>Mobile App, Platform<br/>8M THB"]
    end

    subgraph AFTER["After Digital Implementation"]
        A1[AI-Powered Detection<br/>Real-time Analysis]
        A2[Instant Alerts<br/><200ms Response]
        A3[Thai NLP Native<br/>Local AI Solution]
        A4[Multi-Agent AI<br/>Advanced Platform]
    end

    B1 -.->|Transform| D1
    B2 -.->|Transform| D2
    B3 -.->|Transform| D3
    B4 -.->|Transform| D5

    D1 --> A1
    D2 --> A2
    D3 --> A3
    D4 --> A1
    D5 --> A4

    style BEFORE fill:#ffcdd2
    style DIGITAL fill:#bbdefb
    style AFTER fill:#c8e6c9
```

---

## 8. KPI Verification Flow

```mermaid
flowchart TD
    subgraph PHASE1["Phase 1 KPIs (Month 1-6)"]
        K1["Working Prototype<br/>3 AI Agents"]
        K2["Thai NLP Model<br/>F1 Score ≥80%"]
        K3["Multi-Agent Comm<br/>Message Pass ≥99%"]
        K4["Integration Test<br/>Complete"]
        K5["Documentation<br/>≥80% Complete"]
        K6["Budget Utilization<br/>50% ±15%"]
        K7["Personnel Training<br/>15 people ≥80%"]
    end

    subgraph VERIFY1["Verification Bodies"]
        V1["NSTDA<br/>R&D Verification"]
        V2["depa<br/>Digital Verification"]
    end

    subgraph PHASE2["Phase 2 KPIs (Month 7-12)"]
        K8["Production System<br/>5 AI Agents"]
        K9["Performance<br/>Accuracy ≥85%<br/>Response <200ms<br/>Uptime ≥99.5%"]
        K10["Patents<br/>≥2 Filed"]
        K11["Publications<br/>≥3 Papers"]
        K12["Thai Scam Corpus<br/>≥10,000 Samples"]
        K13["UAT<br/>≥50 Users<br/>Satisfaction ≥3.5/5"]
        K14["Certification<br/>depa Certified"]
    end

    K1 --> V1
    K2 --> V1
    K3 --> V1
    K4 --> V2
    K5 --> V1
    K6 --> V1
    K7 --> V1

    V1 --> |Approve| M1["Milestone 1<br/>21.25M THB Released"]
    V2 --> |Approve| M1

    K8 --> V1
    K9 --> V2
    K10 --> V1
    K11 --> V1
    K12 --> V1
    K13 --> V2
    K14 --> V2

    V1 --> |Approve| M2["Milestone 2<br/>21.25M THB Released"]
    V2 --> |Approve| M2

    style PHASE1 fill:#e3f2fd
    style PHASE2 fill:#e8f5e9
    style VERIFY1 fill:#fff3e0
```

---

## 9. Physical AI Agent System Architecture

```mermaid
flowchart TB
    subgraph HARDWARE["Physical AI Hardware"]
        H1["Microphone Array<br/>Thai Voice Input"]
        H2["Speaker System<br/>Thai Voice Output"]
        H3["LED Display<br/>Visual Feedback"]
        H4["IoT Sensors<br/>Environment Detection"]
        H5["Embedded AI Chip<br/>Edge Processing"]
    end

    subgraph CONNECTIVITY["Connectivity Layer"]
        C1["WiFi Module<br/>Cloud Connection"]
        C2["Bluetooth<br/>Mobile App Sync"]
        C3["Embedded OS<br/>Real-time Processing"]
    end

    subgraph CLOUD["Cloud AI Platform"]
        CL1["Multi-Agent<br/>AI Platform"]
        CL2["Thai NLP<br/>Engine"]
        CL3["Fraud Detection<br/>System"]
    end

    subgraph MOBILE["Mobile Application"]
        M1["Dashboard"]
        M2["Alerts"]
        M3["Learning Center"]
    end

    H1 --> C3
    H4 --> C3
    C3 --> H5
    H5 --> H2
    H5 --> H3

    C3 <--> C1
    C3 <--> C2

    C1 <--> CL1
    CL1 <--> CL2
    CL1 <--> CL3

    C2 <--> M1
    M1 --> M2
    M1 --> M3

    style HARDWARE fill:#ffecb3
    style CONNECTIVITY fill:#b3e5fc
    style CLOUD fill:#d1c4e9
    style MOBILE fill:#c8e6c9
```

---

## 10. BOI Compliance Matrix

```mermaid
flowchart LR
    subgraph CRITERIA["BOI Category 8.2.12 Criteria"]
        C1["Criterion 1<br/>Electronic Sensors<br/>for Data Detection"]
        C2["Criterion 2<br/>Wireless<br/>Connectivity"]
        C3["Criterion 3<br/>Embedded<br/>Processing System"]
    end

    subgraph PRODUCT["Physical AI Agent Features"]
        P1["Microphone<br/>Speaker<br/>LED Sensors<br/>IoT Sensors"]
        P2["WiFi Module<br/>Bluetooth<br/>Connectivity"]
        P3["Embedded AI Chip<br/>Real-time OS<br/>Edge Processing"]
    end

    subgraph STATUS["Compliance Status"]
        S1["PASS"]
        S2["PASS"]
        S3["PASS"]
    end

    C1 --> P1 --> S1
    C2 --> P2 --> S2
    C3 --> P3 --> S3

    S1 --> FINAL["100% Compliant<br/>Category 8.2.12<br/>Smart Electronics"]
    S2 --> FINAL
    S3 --> FINAL

    style CRITERIA fill:#e3f2fd
    style PRODUCT fill:#fff3e0
    style STATUS fill:#c8e6c9
    style FINAL fill:#a5d6a7
```

---

## 11. Fund Disbursement Flow

```mermaid
sequenceDiagram
    participant BOI as BOI Office
    participant NSTDA as NSTDA
    participant DEPA as depa
    participant COMPANY as Aslan Wealth

    Note over BOI,COMPANY: Total Support: 42.5M THB

    BOI->>COMPANY: Project Approval

    rect rgb(225, 245, 254)
        Note over BOI,COMPANY: Phase 1 (Month 1-6)
        COMPANY->>NSTDA: Submit R&D Progress Report
        COMPANY->>DEPA: Submit Digital Progress Report
        NSTDA->>BOI: R&D Verification (50%)
        DEPA->>BOI: Digital Verification (50%)
        BOI->>COMPANY: Disbursement 1: 21.25M THB
    end

    rect rgb(232, 245, 233)
        Note over BOI,COMPANY: Phase 2 (Month 7-12)
        COMPANY->>NSTDA: Submit Final R&D Report
        COMPANY->>DEPA: Submit Final Digital Report
        NSTDA->>BOI: R&D Verification (100%)
        DEPA->>BOI: Digital Verification (100%)
        BOI->>COMPANY: Disbursement 2: 21.25M THB
    end

    Note over BOI,COMPANY: Project Complete
```

---

## 12. Data Flow Architecture

```mermaid
flowchart TB
    subgraph SOURCES["Data Sources"]
        S1["Thai Scam Reports<br/>Bank of Thailand"]
        S2["Financial News<br/>Thai Media"]
        S3["User Interactions<br/>App & Physical AI"]
        S4["Transaction Patterns<br/>Anonymous Data"]
    end

    subgraph PROCESSING["Data Processing"]
        P1["Data Collection<br/>ETL Pipeline"]
        P2["Thai NLP<br/>Text Processing"]
        P3["Feature Engineering<br/>Pattern Extraction"]
        P4["Model Training<br/>GPU Cluster"]
    end

    subgraph STORAGE["Data Storage"]
        ST1["Thai Scam Corpus<br/>≥10,000 Samples"]
        ST2["Model Repository<br/>5 AI Models"]
        ST3["Analytics DB<br/>Real-time Metrics"]
    end

    subgraph OUTPUT["Service Output"]
        O1["Real-time API<br/><200ms Response"]
        O2["Batch Analytics<br/>Daily Reports"]
        O3["Model Updates<br/>Continuous Learning"]
    end

    S1 --> P1
    S2 --> P1
    S3 --> P1
    S4 --> P1

    P1 --> P2
    P2 --> P3
    P3 --> P4

    P4 --> ST1
    P4 --> ST2
    P3 --> ST3

    ST1 --> O3
    ST2 --> O1
    ST3 --> O2

    style SOURCES fill:#ffecb3
    style PROCESSING fill:#b3e5fc
    style STORAGE fill:#d1c4e9
    style OUTPUT fill:#c8e6c9
```

---

## Quick Reference

| Document | Location | Purpose |
|----------|----------|---------|
| Project Proposal | `01_Submission/BOI_COMPLETE_CONSOLIDATED.docx` | Main BOI form data |
| Attachment A | `02_Attachments/attachment_a_rd_COMPLETE.html` | R&D activities detail |
| Attachment C | `02_Attachments/attachment_c_digital_COMPLETE.html` | Digital technology detail |
| Attachment D | `02_Attachments/attachment_d_product_details.html` | Product specifications |
| NDA | `05_Government_Documents/บันทึกข้อตกลงฯ.pdf` | Non-disclosure agreement |

---

## Investment Summary

| Category | Amount | BOI Support |
|----------|--------|-------------|
| R&D Activities | 59,000,000 THB | 29,500,000 THB |
| Digital Technology | 26,000,000 THB | 13,000,000 THB |
| **Total** | **85,000,000 THB** | **42,500,000 THB** |

---

*Document Version: 2.0*
*Last Updated: January 2025*
