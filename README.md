# barcode-scannerZ
An app for scanning shipment barcodes and exporting the data to ExcelZ.
<!DOCTYPE html>
<html lang="ar" dir="rtl">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <title>Al Zajel Express Trading Company</title>

    <!-- مكتبة مسح الباركود -->
    <script src="https://unpkg.com/html5-qrcode" type="text/javascript"></script>

    <!-- مكتبة Excel -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <style>
        body {
            font-family: system-ui, sans-serif;
            text-align: center;
            background: #f5f7fa;
            margin: 0;
            padding: 20px;
        }

        .header {
            background: #005baa;
            color: white;
            padding: 20px;
            border-radius: 15px 15px 0 0;
        }

        .header h1 {
            margin: 5px 0;
            font-size: 24px;
        }

        .header p {
            margin: 5px 0;
            font-size: 14px;
        }

        .logo {
            width: 100px;
            height: 100px;
            object-fit: contain;
            background: white;
            border-radius: 50%;
            padding: 5px;
        }

        .card {
            max-width: 550px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 5px 20px rgba(0,0,0,0.12);
            overflow: hidden;
        }

        .content {
            padding: 20px;
        }

        .category {
            background: #ffd400;
            color: #003b73;
            padding: 12px;
            border-radius: 10px;
            font-weight: bold;
            margin-bottom: 15px;
            font-size: 20px;
        }

        button,
        input[type="file"] {
            margin: 8px 0;
            padding: 12px 15px;
            border-radius: 8px;
            border: 1px solid #ccc;
            font-size: 16px;
            cursor: pointer;
        }

        .btn-main {
            background: #005baa;
            color: white;
            border: none;
            width: 100%;
        }

        .btn-main:hover {
            background: #004785;
        }

        .btn-excel {
            background: #ffd400;
            color: #003b73;
            border: none;
            width: 100%;
            font-weight: bold;
        }

        .btn-clear {
            background: #dc3545;
            color: white;
            border: none;
            width: 100%;
        }

        #reader {
            width: 100%;
            margin-top: 15px;
            display: none;
        }

        table {
            width: 100%;
            margin-top: 20px;
            border-collapse: collapse;
        }

        th,
        td {
            border: 1px solid #ddd;
            padding: 10px;
            text-align: center;
        }

        th {
            background: #005baa;
            color: white;
        }

        td:first-child {
            font-weight: bold;
        }

        .awb {
            direction: ltr;
            font-weight: bold;
            color: #005baa;
        }

        hr {
            border: 0;
            border-top: 1px solid #ddd;
            margin: 20px 0;
        }

        .counter {
            background: #f1f5f9;
            padding: 10px;
            border-radius: 8px;
            font-weight: bold;
        }
    </style>
</head>

<body>

<div class="card">

    <!-- رأس الموقع -->
    <div class="header">

        <!-- ضع صورة الشركة هنا -->
        <img
            src="logo.png"
            class="logo"
            alt="Al Zajel Express Trading Company">

        <h1>Al Zajel Express Trading Company</h1>
        <p>نظام مسح الشحنات</p>

    </div>

    <div class="content">

        <!-- الفئة A -->
        <div class="category">
            A — AWBNUMBER
        </div>

        <!-- تشغيل الكاميرا -->
        <button class="btn-main" onclick="startCamera()">
            📷 تشغيل الكاميرا لمسح الشحنة
        </button>

        <div id="reader"></div>

        <hr>

        <!-- رفع صورة -->
        <label for="qr-input-file">
            <b>أو اختر صورة للباركود من الجهاز:</b>
        </label>

        <br>

        <input
            type="file"
            id="qr-input-file"
            accept="image/*">

        <hr>

        <!-- عدد الشحنات -->
        <div class="counter">
            عدد الشحنات:
            <span id="count">0</span>
        </div>

        <!-- جدول الشحنات -->
        <table id="shipments-table">

            <thead>
                <tr>
                    <th>#</th>
                    <th>AWBNUMBER</th>
                </tr>
            </thead>

            <tbody id="table-body"></tbody>

        </table>

        <br>

        <!-- تصدير Excel -->
        <button
            class="btn-excel"
            onclick="exportToExcel()">
            📥 تحميل ملف Excel
        </button>

        <!-- حذف القائمة -->
        <button
            class="btn-clear"
            onclick="clearShipments()">
            🗑️ مسح جميع الشحنات
        </button>

    </div>

</div>


<script>

    // قائمة الشحنات
    let shipments = [];

    // إنشاء الماسح
    const html5QrCode = new Html5Qrcode("reader");


    // ==========================================
    // إضافة الشحنة
    // ==========================================

    function onScanSuccess(decodedText) {

        decodedText = decodedText.trim();

        if (!decodedText) {
            return;
        }

        // منع التكرار
        if (!shipments.includes(decodedText)) {

            shipments.push(decodedText);

            updateTable();

            alert(
                "تم مسح الشحنة بنجاح\n\nAWBNUMBER:\n" +
                decodedText
            );

        } else {

            alert(
                "هذا الرقم تم مسحه سابقاً!\n\n" +
                decodedText
            );

        }
    }


    // ==========================================
    // تحديث الجدول
    // ==========================================

    function updateTable() {

        const tbody =
            document.getElementById("table-body");

        tbody.innerHTML = "";

        shipments.forEach((num, index) => {

            const row = document.createElement("tr");

            row.innerHTML = `
                <td>${index + 1}</td>
                <td class="awb">${num}</td>
            `;

            tbody.appendChild(row);

        });

        document.getElementById("count").innerText =
            shipments.length;
    }


    // ==========================================
    // تشغيل الكاميرا
    // ==========================================

    function startCamera() {

        const reader =
            document.getElementById("reader");

        reader.style.display = "block";

        html5QrCode.start(

            {
                facingMode: "environment"
            },

            {
                fps: 10,

                qrbox: {
                    width: 300,
                    height: 150
                }
            },

            (decodedText) => {

                onScanSuccess(decodedText);

                html5QrCode.stop()
                    .then(() => {

                        reader.style.display = "none";

                    })
                    .catch(() => {});

            },

            (errorMessage) => {
                // تجاهل أخطاء البحث
            }

        ).catch(err => {

            alert(
                "تعذر تشغيل الكاميرا.\n" +
                "تأكد من السماح للموقع باستخدام الكاميرا."
            );

            reader.style.display = "none";
        });
    }


    // ==========================================
    // قراءة الباركود من صورة
    // ==========================================

    document
        .getElementById("qr-input-file")
        .addEventListener("change", e => {

            if (e.target.files.length === 0) {
                return;
            }

            const imageFile =
                e.target.files[0];

            html5QrCode
                .scanFile(imageFile, true)

                .then(decodedText => {

                    onScanSuccess(decodedText);

                })

                .catch(err => {

                    alert(
                        "تعذر قراءة AWBNUMBER من الصورة.\n\n" +
                        "تأكد من وضوح الباركود."
                    );

                });

        });


    // ==========================================
    // تصدير Excel
    // ==========================================

    function exportToExcel() {

        if (shipments.length === 0) {

            alert(
                "لا توجد شحنات لتصديرها!"
            );

            return;
        }


        /*
         * ملف Excel يحتوي على:
         *
         * الفئة
         * AWBNUMBER
         */

        const data = shipments.map(
            (num, idx) => ({

                "A": "A",

                "AWBNUMBER": num

            })
        );


        // إنشاء ورقة Excel
        const worksheet =
            XLSX.utils.json_to_sheet(data);


        // إنشاء ملف Excel
        const workbook =
            XLSX.utils.book_new();


        XLSX.utils.book_append_sheet(
            workbook,
            worksheet,
            "A - AWBNUMBER"
        );


        // ضبط عرض الأعمدة
        worksheet["!cols"] = [
            { wch: 10 },
            { wch: 30 }
        ];


        // تحميل الملف
        XLSX.writeFile(
            workbook,
            "Al_Zajel_Express_Shipments.xlsx"
        );

    }


    // ==========================================
    // حذف جميع الشحنات
    // ==========================================

    function clearShipments() {

        if (shipments.length === 0) {
            return;
        }

        const confirmDelete =
            confirm(
                "هل أنت متأكد من حذف جميع الشحنات؟"
            );

        if (confirmDelete) {

            shipments = [];

            updateTable();

        }

    }

</script>

</body>
</html>
