# คู่มือติดตั้ง LINE Tag
## Aslan Wealth Channel | Technical Setup Guide
**Website**: https://aslanwealth.com | **LINE OA**: @aslanwealth

---

## 1. LINE Tag คืออะไร?

LINE Tag (หรือ LINE Pixel) คือโค้ดติดตามที่ติดตั้งบนเว็บไซต์เพื่อ:
- ติดตามพฤติกรรมผู้เยี่ยมชม
- วัดผล Conversions
- สร้าง Custom Audiences สำหรับ Retargeting
- วิเคราะห์ประสิทธิภาพโฆษณา

---

## 2. ขั้นตอนการสร้าง LINE Tag

### Step 1: เข้าสู่ LINE Ads Manager
```
1. ไปที่ https://admanager.line.biz
2. Login ด้วย Business ID
3. เลือก Ad Account ที่ต้องการ
```

### Step 2: สร้าง LINE Tag
```
1. คลิก "Tools" (เครื่องมือ) ที่เมนูด้านบน
2. เลือก "Tracking (LINE Tags)"
3. คลิก "Create LINE Tag"
4. ตั้งชื่อ Tag (เช่น "Aslan Wealth - Main")
5. คลิก "Create"
6. ระบบจะแสดง Tag ID และ Base Code
```

### Step 3: Copy Base Code
```javascript
// ตัวอย่าง Base Code ที่ได้รับ
<script>
(function(g,d,o){
  g._ltq=g._ltq||[];g._lt=g._lt||function(){g._ltq.push(arguments)};
  var h=d.getElementsByTagName("head")[0];
  var j=d.createElement("script");j.async=1;
  j.src="https://tagjs.line-scdn.net/tag.js?t="+new Date().getTime();
  h.appendChild(j);
})(window, document);
_lt('init', {
  customerLinkId: 'YOUR_TAG_ID'  // แทนที่ด้วย Tag ID ของคุณ
});
_lt('send', 'pv', ['YOUR_TAG_ID']);
</script>
```

---

## 3. การติดตั้ง Base Code

### วิธีที่ 1: ติดตั้งโดยตรงใน HTML
```html
<!DOCTYPE html>
<html>
<head>
    <title>Aslan Wealth Channel</title>

    <!-- LINE Tag Base Code -->
    <script>
    (function(g,d,o){
      g._ltq=g._ltq||[];g._lt=g._lt||function(){g._ltq.push(arguments)};
      var h=d.getElementsByTagName("head")[0];
      var j=d.createElement("script");j.async=1;
      j.src="https://tagjs.line-scdn.net/tag.js?t="+new Date().getTime();
      h.appendChild(j);
    })(window, document);
    _lt('init', {customerLinkId: 'YOUR_TAG_ID'});
    _lt('send', 'pv', ['YOUR_TAG_ID']);
    </script>
    <!-- End LINE Tag -->

</head>
<body>
    <!-- เนื้อหาเว็บไซต์ -->
</body>
</html>
```

### วิธีที่ 2: ติดตั้งผ่าน Google Tag Manager (GTM)

#### Step 1: สร้าง Tag ใน GTM
```
1. เข้า Google Tag Manager
2. คลิก "Tags" → "New"
3. ตั้งชื่อ: "LINE Tag - Base Code"
4. เลือก Tag Type: "Custom HTML"
5. วาง Base Code ในช่อง HTML
```

#### Step 2: ตั้งค่า Trigger
```
1. คลิก "Triggering"
2. เลือก "All Pages"
3. บันทึก Tag
4. Publish Container
```

### วิธีที่ 3: ติดตั้งบน WordPress

#### ใช้ Plugin (แนะนำ)
```
1. ติดตั้ง Plugin "Insert Headers and Footers"
2. ไปที่ Settings → Insert Headers and Footers
3. วาง Base Code ในช่อง "Scripts in Header"
4. บันทึก
```

#### แก้ไข header.php โดยตรง
```php
<!-- ใน wp-content/themes/[your-theme]/header.php -->
<!-- ก่อน </head> -->

<?php
// LINE Tag Base Code
?>
<script>
(function(g,d,o){
  g._ltq=g._ltq||[];g._lt=g._lt||function(){g._ltq.push(arguments)};
  var h=d.getElementsByTagName("head")[0];
  var j=d.createElement("script");j.async=1;
  j.src="https://tagjs.line-scdn.net/tag.js?t="+new Date().getTime();
  h.appendChild(j);
})(window, document);
_lt('init', {customerLinkId: 'YOUR_TAG_ID'});
_lt('send', 'pv', ['YOUR_TAG_ID']);
</script>

</head>
```

---

## 4. การตั้งค่า Conversion Events

### Standard Events ที่แนะนำ

#### 1. Lead Generation (กรอกฟอร์ม)
```javascript
// เรียกใช้เมื่อผู้ใช้กรอกฟอร์มสำเร็จ
_lt('send', 'cv', {
  type: 'Lead'
}, ['YOUR_TAG_ID']);
```

#### 2. Registration (สมัครสมาชิก)
```javascript
// เรียกใช้เมื่อสมัครสมาชิกสำเร็จ
_lt('send', 'cv', {
  type: 'Registration'
}, ['YOUR_TAG_ID']);
```

#### 3. Purchase (ซื้อสินค้า/คอร์ส)
```javascript
// เรียกใช้เมื่อชำระเงินสำเร็จ
_lt('send', 'cv', {
  type: 'Purchase',
  value: 3000,  // มูลค่าการซื้อ (บาท)
  currency: 'THB'
}, ['YOUR_TAG_ID']);
```

#### 4. Add to Cart
```javascript
// เรียกใช้เมื่อเพิ่มสินค้าลงตะกร้า
_lt('send', 'cv', {
  type: 'AddToCart',
  value: 3000,
  currency: 'THB'
}, ['YOUR_TAG_ID']);
```

---

## 5. ตัวอย่างการติดตั้ง Events

### ตัวอย่าง 1: ฟอร์มลงทะเบียน
```html
<form id="registration-form" onsubmit="trackRegistration()">
    <input type="text" name="name" placeholder="ชื่อ">
    <input type="email" name="email" placeholder="อีเมล">
    <input type="tel" name="phone" placeholder="เบอร์โทร">
    <button type="submit">สมัครรับข่าวสาร</button>
</form>

<script>
function trackRegistration() {
    // Track Lead Event
    _lt('send', 'cv', {
        type: 'Lead'
    }, ['YOUR_TAG_ID']);
}
</script>
```

### ตัวอย่าง 2: ปุ่มซื้อคอร์ส
```html
<button onclick="trackPurchase(3000)">
    ซื้อคอร์ส - 3,000 บาท
</button>

<script>
function trackPurchase(amount) {
    _lt('send', 'cv', {
        type: 'Purchase',
        value: amount,
        currency: 'THB'
    }, ['YOUR_TAG_ID']);
}
</script>
```

### ตัวอย่าง 3: ดาวน์โหลด E-Book
```html
<a href="/ebook.pdf" onclick="trackDownload()">
    ดาวน์โหลด E-Book ฟรี
</a>

<script>
function trackDownload() {
    _lt('send', 'cv', {
        type: 'Lead'
    }, ['YOUR_TAG_ID']);
}
</script>
```

---

## 6. ตั้งค่า Custom Conversion

### Step 1: สร้าง Custom Conversion ใน LINE Ads Manager
```
1. ไปที่ Tools → Tracking (LINE Tags)
2. เลือก Tab "Conversions"
3. คลิก "Create Conversion"
4. ตั้งชื่อ (เช่น "Course Purchase")
5. เลือก Event Type (Purchase, Lead, etc.)
6. ตั้งค่า Attribution Window (แนะนำ 30 วัน)
7. บันทึก
```

### Step 2: ใช้ Conversion ID ในโค้ด
```javascript
// Conversion ที่กำหนดเอง
_lt('send', 'cv', {
  type: 'Conversion',
  id: 'YOUR_CONVERSION_ID'  // ID จาก LINE Ads Manager
}, ['YOUR_TAG_ID']);
```

---

## 7. ตรวจสอบการติดตั้ง

### วิธีที่ 1: ตรวจสอบใน LINE Ads Manager
```
1. ไปที่ Tools → Tracking (LINE Tags)
2. ดูสถานะ Tag:
   - "Active" = ติดตั้งถูกต้อง
   - "Inactive" = ยังไม่ได้รับข้อมูล (รอ 24-48 ชม.)
   - "Error" = มีปัญหาในการติดตั้ง
```

### วิธีที่ 2: ตรวจสอบผ่าน Browser Console
```
1. เปิดเว็บไซต์ที่ติดตั้ง LINE Tag
2. กด F12 เพื่อเปิด Developer Tools
3. ไปที่ Tab "Console"
4. พิมพ์: _ltq
5. ถ้าเห็น Array แสดงว่าติดตั้งถูกต้อง
```

### วิธีที่ 3: ตรวจสอบ Network Requests
```
1. เปิด Developer Tools → Network
2. Filter: "line-scdn"
3. Refresh หน้าเว็บ
4. ควรเห็น Request ไปยัง tag.js
```

---

## 8. การสร้าง Audiences จาก LINE Tag

### Website Visitors Audience
```
1. ไปที่ Tools → Audiences
2. คลิก "Create Audience"
3. เลือก "Website Traffic"
4. ตั้งชื่อ Audience
5. เลือกเงื่อนไข:
   - All website visitors
   - Visitors of specific pages
   - Visitors who spent X time
6. เลือก Retention Period (เช่น 30 วัน)
7. บันทึก
```

### Conversion Audience
```
1. สร้าง Audience จากคนที่ทำ Conversion แล้ว
2. ใช้เพื่อ:
   - Exclude ออกจากแคมเปญ (ไม่ยิงซ้ำ)
   - สร้าง Lookalike Audience
   - Upsell แคมเปญ
```

---

## 9. Cross-Domain Tracking

### กรณีมีหลายโดเมน
```javascript
// ติดตั้งบนทุกโดเมนที่ต้องการ Track
<script>
(function(g,d,o){
  g._ltq=g._ltq||[];g._lt=g._lt||function(){g._ltq.push(arguments)};
  var h=d.getElementsByTagName("head")[0];
  var j=d.createElement("script");j.async=1;
  j.src="https://tagjs.line-scdn.net/tag.js?t="+new Date().getTime();
  h.appendChild(j);
})(window, document);

// ใช้ Tag ID เดียวกันทุกโดเมน
_lt('init', {
  customerLinkId: 'YOUR_TAG_ID',
  crossDomain: true  // เปิด Cross-Domain Tracking
});
_lt('send', 'pv', ['YOUR_TAG_ID']);
</script>
```

---

## 10. Troubleshooting

### ปัญหาที่พบบ่อย

#### 1. Tag แสดง "Inactive"
```
สาเหตุ:
- เพิ่งติดตั้ง (รอ 24-48 ชม.)
- โค้ดติดตั้งผิดตำแหน่ง
- Tag ID ไม่ถูกต้อง

วิธีแก้:
✅ ตรวจสอบว่าโค้ดอยู่ใน <head>
✅ ตรวจสอบ Tag ID ตรงกับ Ads Manager
✅ ล้าง Cache แล้วทดสอบใหม่
```

#### 2. Conversion ไม่ถูก Track
```
สาเหตุ:
- Event Code เรียกใช้ไม่ถูกต้อง
- ผู้ใช้บล็อก JavaScript
- Page Load ช้าเกินไป

วิธีแก้:
✅ ตรวจสอบว่า Event ถูกเรียกหลังจาก Base Code Load
✅ ใช้ addEventListener แทน onclick inline
✅ ทดสอบใน Console
```

#### 3. Audience Size = 0
```
สาเหตุ:
- Traffic น้อยเกินไป
- Retention Period สั้นเกินไป
- เงื่อนไข Audience แคบเกินไป

วิธีแก้:
✅ เพิ่ม Retention Period เป็น 90-180 วัน
✅ ขยายเงื่อนไข (เช่น All Visitors)
✅ รอให้มี Traffic มากขึ้น
```

---

## 11. Best Practices

### Do's ✅
- ติดตั้ง Base Code ก่อน Event Code เสมอ
- ใช้ Tag ID เดียวกันทุกหน้า
- ตั้งค่า Conversion ที่สำคัญให้ครบ
- ตรวจสอบการทำงานหลังติดตั้ง
- อัพเดท Tag เมื่อมี Version ใหม่

### Don'ts ❌
- อย่าติดตั้ง Tag ซ้ำหลายครั้งในหน้าเดียว
- อย่าใส่ Event Code ก่อน Base Code
- อย่าลืม Publish เมื่อใช้ GTM
- อย่าเปลี่ยน Tag ID บ่อย

---

## 12. Quick Reference

### Base Code Template
```javascript
<script>
(function(g,d,o){
  g._ltq=g._ltq||[];g._lt=g._lt||function(){g._ltq.push(arguments)};
  var h=d.getElementsByTagName("head")[0];
  var j=d.createElement("script");j.async=1;
  j.src="https://tagjs.line-scdn.net/tag.js?t="+new Date().getTime();
  h.appendChild(j);
})(window, document);
_lt('init', {customerLinkId: 'YOUR_TAG_ID'});
_lt('send', 'pv', ['YOUR_TAG_ID']);
</script>
```

### Event Code Templates
```javascript
// Page View (อัตโนมัติจาก Base Code)
_lt('send', 'pv', ['YOUR_TAG_ID']);

// Lead
_lt('send', 'cv', {type: 'Lead'}, ['YOUR_TAG_ID']);

// Registration
_lt('send', 'cv', {type: 'Registration'}, ['YOUR_TAG_ID']);

// Purchase
_lt('send', 'cv', {
  type: 'Purchase',
  value: 3000,
  currency: 'THB'
}, ['YOUR_TAG_ID']);

// Add to Cart
_lt('send', 'cv', {type: 'AddToCart'}, ['YOUR_TAG_ID']);

// Custom Conversion
_lt('send', 'cv', {
  type: 'Conversion',
  id: 'YOUR_CONVERSION_ID'
}, ['YOUR_TAG_ID']);
```

---

*เอกสารนี้เป็นส่วนหนึ่งของ LINE Ads Campaign สำหรับ Aslan Wealth Channel*
