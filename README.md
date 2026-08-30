<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Al Zajel Express Trading Company</title>
  <!-- مكتبة تصدير ملفات Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    * {
      box-sizing: border-box;
      font-family: system-ui, -apple-system, sans-serif;
    }

    body {
      background-color: #f4f6f9;
      margin: 0;
      padding: 15px;
      display: flex;
      justify-content: center;
    }

    .card {
      width: 100%;
      max-width: 400px;
      background: #ffffff;
      border-radius: 16px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
      overflow: hidden;
      padding: 20px;
    }

    /* الهيدر الرئيسي */
    .header {
      background-color: #0056b3;
      color: white;
      text-align: center;
      padding: 20px 10px;
      border-radius: 12px;
      margin-bottom: 20px;
    }

    .header h2 {
      margin: 0 0 8px 0;
      font-size: 1.2rem;
    }

    .header p {
      margin: 0;
      font-size: 0.95rem;
      opacity: 0.9;
    }

    /* زر الكاميرا */
    .btn-camera {
      width: 100%;
      background-color: #0056b3;
      color: white;
      border: none;
      padding: 14px;
      border-radius: 8px;
      font-size: 1rem;
      font-weight: bold;
      cursor: pointer;
    }

    .divider {
      border: 0;
      border-top: 1px solid #e2e8f0;
      margin: 20px 0;
    }

    /* قسم اختيار الملف - منع التداخل */
    .file-container {
      text-align: center;
    }

    .file-container label {
      display: block;
      font-weight: bold;
      margin-bottom: 12px;
      color: #333;
    }

    .file-input-wrapper {
      width: 100%;
      border: 1px solid #ccc;
      border-radius: 8px;
      padding: 8px;
      background-color: #fafafa;
    }

    input[type="file"] {
      width: 100%;
      direction: rtl;
      font-size: 0.85rem;
    }

    /* عداد الشحنات */
    .counter-box {
      background-color: #f0f4f8;
      padding: 14px;
      border-radius: 8px;
      text-align: center;
      font-size: 1.1rem;
      font-weight: bold;
      color: #1a202c;
      margin-bottom: 15px;
    }

    /* زر تحميل Excel */
    .btn-excel {
      width: 100%;
      background-color: #ffc107;
      color: #000;
      border: none;
      padding: 14px;
      border-radius: 8px;
      font-size: 1rem;
      font-weight: bold;
      cursor: pointer;
    }
  </style>
</head>
<body>

  <div class="card">
    <!-- الهيدر -->
    <div class="header">
      <h2>Al Zajel Express Trading Company</h2>
      <p>نظام مسح الشحنات</p>
    </div>

    <!-- تشغيل الكاميرا -->
    <button class="btn-camera" id="start-camera">
      📷 تشغيل الكاميرا لمسح الشحنة
    </button>

    <hr class="divider">

    <!-- رفع الملف دون تداخل -->
    <div class="file-container">
      <label>أو اختر صورة للباركود من الجهاز:</label>
      <div class="file-input-wrapper">
        <input type="file" id="barcode-file" accept="image/*">
      </div>
    </div>

    <hr class="divider">

    <!-- العداد -->
    <div class="counter-box">
      عدد الشحنات: <span id="count">0</span>
    </div>

    <!-- زر التحميل -->
    <button class="btn-excel" id="export-excel">📥 تحميل ملف Excel</button>
  </div>

  <script>
    // قائمة حفظ الباركودات
    let scannedBarcodes = [];

    // تحديث رقم العداد
    function updateCounter() {
      document.getElementById('count').innerText = scannedBarcodes.length;
    }

    // تصدير ملف Excel بنفس ترتيب الصورة تمامًا (عمود AWBNUMBER واسم Sheet SaleRent Details)
    function exportToExcel() {
      if (scannedBarcodes.length === 0) {
        alert("لا توجد شحنات ممسوحة للتصدير!");
        return;
      }

      const excelData = scannedBarcodes.map(code => ({
        "AWBNUMBER": code
      }));

      const worksheet = XLSX.utils.json_to_sheet(excelData);
      const workbook = XLSX.utils.book_new();
      
      // تسمية ورقة العمل بالضبط كما بالصورة
      XLSX.utils.book_append_sheet(workbook, worksheet, "SaleRent Details");

      XLSX.writeFile(workbook, "Shipments_List.xlsx");
    }

    document.getElementById('export-excel').addEventListener('click', exportToExcel);
  </script>

</body>
</html>
