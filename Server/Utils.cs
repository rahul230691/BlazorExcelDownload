using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorDownload.Server
{
    public class Utils
    {
        public static ExcelPackage createExcelPackage(IWebHostEnvironment _hostingEnvironment)
        {
            var package = new ExcelPackage();
            package.Workbook.Properties.Title = "Test Report";
            package.Workbook.Properties.Author = "rahul2306";
            package.Workbook.Properties.Subject = "Test Report";
            package.Workbook.Properties.Keywords = "Testing";


            var worksheet = package.Workbook.Worksheets.Add("Employee");

            //First add the headers
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Gender";
            worksheet.Cells[1, 4].Value = "Salary (in $)";

            //Add values

            var numberformat = "#,##0";
            var dataCellStyleName = "TableNumber";
            var numStyle = package.Workbook.Styles.CreateNamedStyle(dataCellStyleName);
            numStyle.Style.Numberformat.Format = numberformat;

            worksheet.Cells[2, 1].Value = 1;
            worksheet.Cells[2, 2].Value = "Rahul";
            worksheet.Cells[2, 3].Value = "M";
            worksheet.Cells[2, 4].Value = 50000;
            worksheet.Cells[2, 4].Style.Numberformat.Format = numberformat;

            worksheet.Cells[3, 1].Value = 2;
            worksheet.Cells[3, 2].Value = "Duy";
            worksheet.Cells[3, 3].Value = "M";
            worksheet.Cells[3, 4].Value = 50000;
            worksheet.Cells[3, 4].Style.Numberformat.Format = numberformat;

            worksheet.Cells[4, 1].Value = 3;
            worksheet.Cells[4, 2].Value = "Steve";
            worksheet.Cells[4, 3].Value = "M";
            worksheet.Cells[4, 4].Value = 45000;
            worksheet.Cells[4, 4].Style.Numberformat.Format = numberformat;

            // Add to table / Add summary row
            var tbl = worksheet.Tables.Add(new ExcelAddressBase(fromRow: 1, fromCol: 1, toRow: 4, toColumn: 4), "Data");
            tbl.ShowHeader = true;
            tbl.TableStyle = TableStyles.Dark9;
            tbl.ShowTotal = true;
            tbl.Columns[3].DataCellStyleName = dataCellStyleName;
            tbl.Columns[3].TotalsRowFunction = RowFunctions.Sum;
            worksheet.Cells[5, 4].Style.Numberformat.Format = numberformat;

            // AutoFitColumns
            worksheet.Cells[1, 1, 4, 4].AutoFitColumns();

            return package;
        }
    }
}
