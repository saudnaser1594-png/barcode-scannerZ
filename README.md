function exportToExcel() {

    if (shipments.length === 0) {

        showStatus(
            "لا توجد شحنات لتصديرها",
            "error"
        );

        return;
    }

    // إنشاء البيانات في عمود A فقط
    const data = [
        ["AWBNUMBER"]
    ];

    shipments.forEach(function(num) {
        data.push([num]);
    });

    // إنشاء ورقة Excel
    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // جعل العمود A بعرض مناسب
    worksheet["!cols"] = [
        { wch: 30 }
    ];

    // تثبيت الأرقام كنص للحفاظ على الأصفار والأرقام الطويلة
    for (let row = 2; row <= data.length; row++) {

        const cell = worksheet["A" + row];

        if (cell) {
            cell.t = "s";
            cell.v = String(cell.v);
        }
    }

    // إنشاء ملف Excel
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "الشحنات"
    );

    // تحميل الملف
    XLSX.writeFile(
        workbook,
        "AWBNUMBER.xlsx"
    );
}
