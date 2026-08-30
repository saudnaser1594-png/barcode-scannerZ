function exportToExcel() {

    if (shipments.length === 0) {
        alert("لا توجد شحنات لتصديرها!");
        return;
    }

    // كل البيانات في العمود A فقط
    const data = [["AWBNUMBER"]];

    // إضافة أرقام الشحنات تحت AWBNUMBER
    shipments.forEach(function(number) {
        data.push([String(number)]);
    });

    // إنشاء ورقة Excel
    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // عرض العمود A
    worksheet["!cols"] = [
        { wch: 35 }
    ];

    // إجبار أرقام الشحنات على أن تكون نصاً
    for (let i = 2; i <= data.length; i++) {
        const cell = worksheet["A" + i];

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

    // اسم الملف
    XLSX.writeFile(
        workbook,
        "AWBNUMBER.xlsx"
    );
}
