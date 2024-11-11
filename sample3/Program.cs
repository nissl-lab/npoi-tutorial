using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using (var workbook = new XSSFWorkbook())
{ 
    var sheet = workbook.CreateSheet("Sheet1");
    var dataformat = workbook.CreateDataFormat();

    //A1 shows 2 digits number
    var cell1 = sheet.CreateRow(0).CreateCell(0);
    cell1.SetCellValue(1.234);
    SetCellFormat(workbook, cell1, dataformat.GetFormat("0.00"));

    //A2 shows scientific number
    var cell2 = sheet.CreateRow(1).CreateCell(0);
    cell2.SetCellValue(3.141592654);
    SetCellFormat(workbook, cell2, dataformat.GetFormat("0.00E+00"));

    //A3 shows percentage
    var cell3 = sheet.CreateRow(2).CreateCell(0);
    cell3.SetCellValue(0.9933);
    SetCellFormat(workbook, cell3, dataformat.GetFormat("0.00%"));

    //A4 shows money
    var cell4 = sheet.CreateRow(3).CreateCell(0);
    cell4.SetCellValue(2000);
    SetCellFormat(workbook, cell4, dataformat.GetFormat("$#,##0"));

    //A5 shows phone number
    var cell5 = sheet.CreateRow(4).CreateCell(0);
    cell5.SetCellValue(02168883222);
    SetCellFormat(workbook, cell5, dataformat.GetFormat("000-00000000"));

    //A6 shows date and time
    var cell6 = sheet.CreateRow(5).CreateCell(0);
    cell6.SetCellFormula("DateValue(\"2005-11-11\")+TIMEVALUE(\"11:11:11\")");
    SetCellFormat(workbook, cell6, dataformat.GetFormat("m/d/yy h:mm"));

    using (var file = File.Create("sample2.xlsx"))
    { 
        workbook.Write(file);
    }
}

static void SetCellFormat(IWorkbook workbook, ICell cell, short formatId)
{ 
    var cellStyle = workbook.CreateCellStyle();
    cellStyle.DataFormat = formatId;
    cell.CellStyle=cellStyle; 
}