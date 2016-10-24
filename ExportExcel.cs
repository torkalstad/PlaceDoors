using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
/* To work eith EPPlus library */
using OfficeOpenXml;
using System.Diagnostics;
using System.Drawing;
using OfficeOpenXml.Style;

namespace SOM.RevitTools.PlaceDoors
{
    class ExportExcel
    {
        public string ExportToExcel(List<ObjDoors> LinkedModelDoors, string WorkSheetName)
        {
            // Set the file name and get the output directory
            var fileName = DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + "-Revit-Doors" + ".xlsx";
            var outputDir = @"C:\";

            // Create the file using the FileInfo object
            FileInfo file = new FileInfo(outputDir + fileName);

            // Create the package and make sure you wrap it in a using statement
            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Door Information");

                // --------- Data and styling goes here -------------- //
                // Add some formatting to the worksheet
                worksheet.TabColor = Color.Blue;
                worksheet.DefaultRowHeight = 12;
                worksheet.Row(1).Height = 20;
                worksheet.Row(2).Height = 18;


                // Start adding the header at first row.
                worksheet.Cells[1, 1].Value = "Door Name";
                worksheet.Cells[1, 2].Value = "doorId";
                worksheet.Cells[1, 3].Value = "doorWidth";
                worksheet.Cells[1, 4].Value = "doorHeight";
                worksheet.Cells[1, 5].Value = "doorSillHeight";
                worksheet.Cells[1, 6].Value = "X";
                worksheet.Cells[1, 7].Value = "Y";
                worksheet.Cells[1, 8].Value = "Z";
                worksheet.Cells[1, 9].Value = "level";

                // Format the first row of the heade, but only columns 1-10;
                using (var range = worksheet.Cells[1, 1, 1, 10])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Size = 12;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    range.Style.Font.Color.SetColor(Color.Black);
                }

                // Keep track of the row that we're on, but start with two to skip the header
                int rowNumber = 2;
                foreach (var exportObj in LinkedModelDoors)
                {
                    worksheet.Cells[rowNumber, 1].Value = exportObj.doorName;
                    worksheet.Cells[rowNumber, 2].Value = exportObj.doorId;
                    worksheet.Cells[rowNumber, 3].Value = exportObj.doorWidth;
                    worksheet.Cells[rowNumber, 4].Value = exportObj.doorHeight;
                    worksheet.Cells[rowNumber, 5].Value = exportObj.doorSillHeight;
                    worksheet.Cells[rowNumber, 6].Value = exportObj.X;
                    worksheet.Cells[rowNumber, 7].Value = exportObj.Y;
                    worksheet.Cells[rowNumber, 8].Value = exportObj.Z;
                    worksheet.Cells[rowNumber, 9].Value = exportObj.Z;
                    rowNumber++;
                }

                // Fit the columns according to its content
                worksheet.Column(1).Width = 20;
                worksheet.Column(2).Width = 15;
                worksheet.Column(3).Width = 10;
                worksheet.Column(4).Width = 10;
                worksheet.Column(5).Width = 10;
                worksheet.Column(6).Width = 10;
                worksheet.Column(7).Width = 10;
                worksheet.Column(8).Width = 10;
                worksheet.Column(9).Width = 10;

                // Set some document properties
                package.Workbook.Properties.Title = "Revit Doors";
                package.Workbook.Properties.Author = "danny.bentley";
                package.Workbook.Properties.Company = "Parachense";

                // save our new workbook and we are done!
                try
                {
                    package.Save();
                    package.Stream.Close();
                    package.Dispose();

                }

                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + "Exception raised - " + ex.Message);
                }

                //-------- Now leaving the using statement
            } // Outside the using statement

            return file.ToString();
        }
    }
}
