using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Data.SqlClient;
using System.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Linq;
using OfficeOpenXml.Table;
using System.Collections.Generic;

namespace EPPlusExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var package = new ExcelPackage())
            {
         
                var worksheet = package.Workbook.Worksheets.Add("Weekly Usage");
                worksheet.Cells["A1"].LoadFromCollection(CustomerEntity.GetCustomers(), true, TableStyles.Medium6);
                ApplyStyleToSheet(ref worksheet);

                package.SaveAs(new FileInfo(@"C:\test\excel.xlsx"));
            }
        }

        static void ApplyStyleToSheet(ref ExcelWorksheet worksheet)
        {
                   
            //Format the header 
            using (ExcelRange range = worksheet.Cells["A1:B1"])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(90, 135, 210));
                range.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
            {
                worksheet.Column(i).AutoFit();
            }
        }
    }
}
