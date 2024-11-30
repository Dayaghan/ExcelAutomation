﻿using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ExcelHandlingDotnetpractice
{
    public class Leaver_Master
    {
        public static void LeaverMaster(string filePath, string destinationFolder)
        {
            try
            {
                string ascendcodes = "C:/Automation/McAafee software/Automation_Ascent_Codes/Ascent Codes.xlsx";
                string outputFilePath = "C:/Users/Dayaghanl/Desktop/output.xlsx";
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets["Leavers Data"];// Assuming the data is in the first worksheet
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    using (var outputPackage = new ExcelPackage())
                    {
                        int employeenumber = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                        int dateofLeaving = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payroll End Date");
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Leaver_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "Date Of Leaving (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 3].Value = "Payroll Code";
                        outputWorksheet.Cells[1, 4].Value = "InactiveID";
                        outputWorksheet.Cells[1, 5].Value = "Date Of Resign (YYYY-MM-DD)";
                        for (int row = 2; row <= lastRow; row++)
                        {
                            var HRID = inputWorkSheet.Cells[row, employeenumber].GetValue<string>();
                            var DateOfLeaving = inputWorkSheet.Cells[row, dateofLeaving].GetValue<string>();
                            DateOfLeaving = DateOfLeaving.Replace(" ", "");
                            if (DateOfLeaving.Length == 10)
                            {
                                outputWorksheet.Cells[row, 2].Value = DateOfLeaving;
                            }
                            outputWorksheet.Cells[row, 1].Value = HRID;
                            outputWorksheet.Cells[row, 3].Value ="9";
                            outputWorksheet.Cells[row, 4].Value = "3";
                            outputWorksheet.Cells[row, 5].Value = DateOfLeaving;
                        }
                        string newFileName = Path.Combine(destinationFolder, "Leaver_Master " + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                        Console.WriteLine("Leaver_Master Excel file created successfully!");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}