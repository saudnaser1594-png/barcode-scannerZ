<!DOCTYPE html>
<html lang="ar" dir="rtl">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ماسح باركود الشحنات - زاجل</title>

    <!-- مكتبة مسح الباركود -->
    <script src="https://unpkg.com/html5-qrcode" type="text/javascript"></script>

    <!-- مكتبة تصدير Excel -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <style>
        * {
            box-sizing: border-box;
        }

        body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            text-align: center;
            background: #f4f6f9;
            margin: 0;
            padding: 20px 10px;
            color: #2c3e50;
        }

        .card {
            max-width: 480px;
            margin: 0 auto;
            background: #ffffff;
            border-radius: 20px;
            box-shadow: 0 8px 30px rgba(0, 0, 0, 0.08);
            padding: 24px 18px;
        }

        /* الشعار والتصاميم العلوية */
        .logo-container {
            margin-bottom: 12px;
        }

        .logo {
            width: 120px;
            height: 120px;
            border-radius: 18px;
            object-fit: contain;
        }

        h1 {
            font-size: 22px;
            font-weight: 800;
            margin: 10px 0 20px 0;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }

        /* الأزرار الرئيسية */
        .btn-camera {
            background: #007bff;
            color: white;
            border: none;
            width: 100%;
            padding: 14px;
            border-radius: 12px;
            font-size: 17px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(0, 123, 255, 0.25);
            transition: background 0.2s;
        }

        .btn-camera:active {
            background: #0056b3;
        }

        #reader {
            width: 100%;
            margin-top: 15px;
            border-radius: 12px;
            overflow: hidden;
            display: none;
        }

        hr {
            border: 0;
            border-top: 1px solid #e9ecef;
            margin: 22px 0;
        }

        .section-title {
            font-size: 16px;
            font-weight: bold;
            margin-bottom: 12px;
        }

        /* اختيار صورة باركود */
        .file-input-wrapper {
            margin: 10px 0;
        }

        input[type="file"] {
            width: 100%;
            padding: 10px;
            border: 1px solid #ced4da;
            border-radius: 10px;
            background: #f8f9fa;
            font-size: 14px;
        }

        .hint-text {
            font-size: 12px;
            color: #6c757d;
            margin-top: 8px;
            line-height: 1.5;
        }

        /* قسم التحكم بالقائمة */
        .list-title {
            font-size: 18px;
            font-weight: bold;
            margin: 15px 0;
        }

        .action-buttons {
            display: flex;
            gap: 10px;
            margin-bottom: 15px;
        }

        .btn-delete-last {
            flex: 1;
            background: #f8f9fa;
            color: #333;
            border: 1px solid #ced4da;
            padding: 10px;
            border-radius: 10px;
            font-size: 15px;
            font-weight: bold;
            cursor: pointer;
        }

        .btn-delete-all {
            flex: 1;
            background: #fff;
            color: #dc3545;
            border: 1px solid #dc3545;
            padding: 10px;
            border-radius: 10px;
            font-size: 15px;
            font-weight: bold;
            cursor: pointer;
        }

        /* جدول الشحنات */
        .table-container {
            max-height: 250px;
            overflow-y: auto;
            border: 1px solid #dee2e6;
            border-radius: 10px;
            margin-bottom: 18px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }

        th {
            background: #f8f9fa;
            color: #212529;
            padding: 12px 8px;
            border-bottom: 2px solid #dee2e6;
            position: sticky;
            top: 0;
        }

        th.col-num {
            width: 20%;
            border-left: 1px solid #dee2e6;
        }

        td {
            padding: 10px 8px;
            border-bottom: 1px solid #e9ecef;
            word-break: break-all;
        }

        td.col-num {
            font-weight: bold;
            border-left: 1px solid #e9ecef;
            background: #fafafa;
        }

        td.col-awb {
            direction: ltr;
            font-family: monospace;
            font-weight: 600;
            color: #005baa;
        }

        /* زر تحميل Excel */
        .btn-excel {
            background: #28a745;
            color: white;
            border: none;
            width: 100%;
            padding: 14px;
            border-radius: 12px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(40, 167, 69, 0.25);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }

        .btn-excel:active {
            background: #218838;
        }
    </style>
</head>

<body>

    <div class="card">

        <!-- الشعار -->
        <div class="logo-container">
            <img src="zajil_logo.png" alt="زاجل Zajil" class="logo">
        </div>

        <!-- العنوان -->
        <h1>📦 ماسح باركود الشحنات</h1>

        <!-- زر تشغيل الكاميرا -->
        <button class="btn-camera" onclick="startCamera()">📷 تشغيل الكاميرا</button>

        <div id="reader"></div>

        <hr>

        <!-- اختيار صورة -->
        <div class="section-title">أو اختر صورة باركود من الجهاز:</div>

        <div class="file-input-wrapper">
            <input type="file" id="qr-input-file" accept="image/*">
        </div>

        <div class="hint-text">
            يدعم QR والباركود الخطي 1D مثل Code 128 و Code 39 و EAN وغيرها حسب دعم الجهاز والمتصفح.
        </div>

        <hr>

        <!-- قائمة الشحنات -->
        <div class="list-title">
            قائمة الشحنات (<span id="count">0</span>)
        </div>

        <!-- أزرار الحذف -->
        <div class="action-buttons">
            <button class="btn-delete-last" onclick="deleteLastShipment()">حذف آخر شحنة</button>
            <button class="btn-delete-all" onclick="clearShipments()">حذف الكل</button>
        </div>

        <!-- الجدول -->
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th class="col-num">#</th>
                        <th>رقم الشحنة</th>
                    </tr>
                </thead>
                <tbody id="table-body">
                    <!-- تضاف الشحنات هنا تلقائياً -->
                </tbody>
            </table>
        </div>

        <!-- زر تحميل الإكسل -->
        <button class="btn-excel" onclick="exportToExcel()">
            📫 تحميل ملف Excel
        </button>

    </div>

    <script>
        let shipments = [];
        const html5QrCode = new Html5Qrcode("reader");

        // إضافة الشحنة للسيستم
        function onScanSuccess(decodedText) {
            decodedText = decodedText.trim();
            if (!decodedText) return;

            if (!shipments.includes(decodedText)) {
                shipments.push(decodedText);
                updateTable();
            } else {
                alert("هذه الشحنة مضافة مسبقاً:\n" + decodedText);
            }
        }

        // تحديث عرض الجدول والعداد
        function updateTable() {
            const tbody = document.getElementById("table-body");
            tbody.innerHTML = "";

            shipments.forEach((num, index) => {
                const row = document.createElement("tr");
                row.innerHTML = `
                    <td class="col-num">${index + 1}</td>
                    <td class="col-awb">${num}</td>
                `;
                tbody.appendChild(row);
            });

            document.getElementById("count").innerText = shipments.length;
        }

        // تشغيل كاميرا الجوال
        function startCamera() {
            const reader = document.getElementById("reader");
            reader.style.display = "block";

            html5QrCode.start(
                { facingMode: "environment" },
                { fps: 10, qrbox: { width: 280, height: 140 } },
                (decodedText) => {
                    onScanSuccess(decodedText);
                    html5QrCode.stop().then(() => {
                        reader.style.display = "none";
                    }).catch(() => {});
                },
                (errorMessage) => {}
            ).catch(err => {
                alert("يرجى السماح بصلاحية الكاميرا للمتصفح");
                reader.style.display = "none";
            });
        }

        // المسح من صورة من الجهاز
        document.getElementById("qr-input-file").addEventListener("change", e => {
            if (e.target.files.length === 0) return;
            const imageFile = e.target.files[0];

            html5QrCode.scanFile(imageFile, true)
                .then(decodedText => {
                    onScanSuccess(decodedText);
                })
                .catch(err => {
                    alert("تعذر قراءة الباركود من الصورة المختارة");
                });
        });

        // حذف آخر شحنة
        function deleteLastShipment() {
            if (shipments.length > 0) {
                shipments.pop();
                updateTable();
            }
        }

        // حذف جميع الشحنات
        function clearShipments() {
            if (shipments.length === 0) return;
            if (confirm("هل أنت متأكد من حذف قائمة الشحنات بالكامل؟")) {
                shipments = [];
                updateTable();
            }
        }

        // ======================================================
        // تصدير Excel (العمود A يحتوي على أرقام الشحنات مباشرة)
        // ======================================================
        function exportToExcel() {
            if (shipments.length === 0) {
                alert("لا توجد شحنات مسجلة للتصدير!");
                return;
            }

            // الهيدر AWBNUMBER وتحته مباشرة أرقام الشحنات داخل العمود A
            const data = shipments.map(num => ({
                "AWBNUMBER": num
            }));

            const worksheet = XLSX.utils.json_to_sheet(data);
            const workbook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(workbook, worksheet, "AWBNUMBER");

            // توسيع العمود A لتظهر الأرقام كاملة
            worksheet["!cols"] = [{ wch: 35 }];

            XLSX.writeFile(workbook, "Al_Zajel_Express_Shipments.xlsx");
        }
    </script>

</body>
</html>
