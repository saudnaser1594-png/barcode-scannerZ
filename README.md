<!DOCTYPE html>
<html lang="ar" dir="rtl">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Al Zajel Express Trading Company</title>

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
            display: flex;
            justify-content: center;
        }

        .logo {
            width: 110px;
            height: 110px;
            border-radius: 18px;
            object-fit: cover;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }

        h1 {
            font-size: 19px;
            font-weight: 800;
            margin: 10px 0 20px 0;
            color: #005baa;
        }

        /* زر الكاميرا */
        .btn-camera {
            background: #007bff;
            color: white;
            border: none;
            width: 100%;
            padding: 14px;
            border-radius: 12px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(0, 123, 255, 0.25);
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
            margin: 20px 0;
        }

        .section-title {
            font-size: 15px;
            font-weight: bold;
            margin-bottom: 12px;
        }

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

        .count-box {
            background: #f1f3f5;
            border-radius: 10px;
            padding: 12px;
            font-size: 16px;
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
            font-size: 14px;
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
            font-size: 14px;
            font-weight: bold;
            cursor: pointer;
        }

        .table-container {
            max-height: 220px;
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

        td {
            padding: 12px 10px;
            border-bottom: 1px solid #e9ecef;
            word-break: break-all;
        }

        tr:last-child td {
            border-bottom: none;
        }

        td.col-num {
            width: 20%;
            font-weight: bold;
            border-left: 1px solid #e9ecef;
            background: #fafafa;
            text-align: center;
        }

        td.col-awb {
            direction: ltr;
            font-family: monospace;
            font-weight: 600;
            color: #005baa;
            text-align: center;
        }

        .btn-excel {
            background: #ffc107;
            color: #000;
            border: none;
            width: 100%;
            padding: 14px;
            border-radius: 12px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(255, 193, 7, 0.3);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }
    </style>
</head>

<body>

    <div class="card">

        <!-- الشعار المدمج مباشرة بالملف -->
        <div class="logo-container">
            <img src="zajil_logo.png" alt="زاجل Zajil" class="logo">
        </div>

        <h1>Al Zajel Express Trading Company</h1>

        <!-- زر الكاميرا -->
        <button class="btn-camera" onclick="startCamera()">📷 تشغيل الكاميرا لمسح الشحنة</button>

        <div id="reader"></div>

        <hr>

        <div class="section-title">أو اختر صورة للباركود من الجهاز:</div>
        <div class="file-input-wrapper">
            <input type="file" id="qr-input-file" accept="image/*">
        </div>

        <hr>

        <div class="count-box">
            عدد الشحنات: <span id="count">0</span>
        </div>

        <div class="action-buttons">
            <button class="btn-delete-last" onclick="deleteLastShipment()">حذف آخر شحنة</button>
            <button class="btn-delete-all" onclick="clearShipments()">حذف الكل</button>
        </div>

        <!-- قائمة الشحنات الممسوحة -->
        <div class="table-container">
            <table>
                <tbody id="table-body"></tbody>
            </table>
        </div>

        <!-- زر تحميل الإكسل -->
        <button class="btn-excel" onclick="exportToExcel()">
            📥 تحميل ملف Excel
        </button>

    </div>

    <script>
        let shipments = [];
        const html5QrCode = new Html5Qrcode("reader");

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

        function deleteLastShipment() {
            if (shipments.length > 0) {
                shipments.pop();
                updateTable();
            }
        }

        function clearShipments() {
            if (shipments.length === 0) return;
            if (confirm("هل أنت متأكد من حذف قائمة الشحنات بالكامل؟")) {
                shipments = [];
                updateTable();
            }
        }

        function exportToExcel() {
            if (shipments.length === 0) {
                alert("لا توجد شحنات مسجلة للتصدير!");
                return;
            }

            const data = shipments.map(num => ({
                "AWBNUMBER": num
            }));

            const worksheet = XLSX.utils.json_to_sheet(data);
            const workbook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(workbook, worksheet, "AWBNUMBER");
            worksheet["!cols"] = [{ wch: 35 }];

            XLSX.writeFile(workbook, "Al_Zajel_Express_Shipments.xlsx");
        }
    </script>

</body>
</html>
