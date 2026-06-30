# دليل إعداد Apps Script Backend

## 1. إنشاء Google Sheet
1. افتح Google Sheets وانشئ ملفًا جديدًا.
2. انسخ الـ **ID** من رابط الملف (الجزء الطويل بين `/d/` و `/edit`).
3. ضعه في ملف `Code.gs` في المتغير `CONFIG.SHEET_ID`.

## 2. إنشاء Google Apps Script
1. من الشيت: **Extensions → Apps Script**.
2. احذف `Code.gs` الافتراضي.
3. ألصق كود `Code.gs` المرفق.
4. عدّل:
   - `SHEET_ID`
   - `ADMIN_EMAIL` (بريد الموجه)
   - `AUTO_APPROVE_REGISTRATION` (true/false)

## 3. الجداول المطلوبة
سيتم إنشاؤها تلقائيًا عند أول طلب، لكن يمكنك إنشاؤها يدويًا لتجنب الأخطاء:
- `Profiles`
- `Interests`
- `AdminChatRequests`
- `AdminChatMessages`
- `Messages`
- `RegistrationRequests`

## 4. نشر السكربت
1. **Deploy → New deployment**.
2. اختر **Web app**.
3. **Execute as: Me**.
4. **Who has access: Anyone**.
5. انسخ رابط الـ Web App URL.

## 5. ربط HTML بالـ Backend
في ملف `index-netflix (1).html`، عدّل:
```javascript
const CONFIG = {
  API_URL: 'https://script.google.com/macros/s/XXXXXXXX/exec', // ← ضع رابطك هنا
  GOOGLE_CLIENT_ID: '...',
  IMGBB_KEY: '...'
};
```

## 6. إنشاء حساب الموجه (Admin)
أول ما يسجل دخوله بريد الموجه، يتم اعتباره admin تلقائيًا (بناءً على `ADMIN_EMAIL`).

## ملاحظة أمان مهمة
في نهاية ملف HTML يوجد كود مشبوه يخص Cloudflare. إذا لم تضعه أنت، قم بإزالته قبل النشر.
