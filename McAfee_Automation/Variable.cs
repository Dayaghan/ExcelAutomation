using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHandlingDotnetpractice
{
    public class Variable
    {
        public static void Variable_Pay_Inputs_Data(string filePath, string destinationFolder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                string ascendcodes = "C:/Automation/McAafee software/Automation_Ascent_Codes/Ascent Codes.xlsx";
                string outputFilePath = Path.Combine(destinationFolder, "Variable_Pay_Summary.xlsx");
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets["Variable Pay Inputs Data"];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    // Get column numbers for relevant headers
                    int hridCol = Program.getColumnNumber(filePath, inputWorkSheet.Name, "HR ID");
                    int payElementCol = Program.getColumnNumber(filePath, inputWorkSheet.Name, "Pay Element Short Code");
                    int amountCol = Program.getColumnNumber(filePath, inputWorkSheet.Name, "Amount");

                    // Data structures to store unique pay elements and employee data
                    var employeeData = new Dictionary<string, Dictionary<string, double>>();
                    var payElementCodes = new HashSet<string>();

                    // Read data from input sheet
                    for (int row = 2; row <= lastRow; row++)
                    {
                        string hrid = inputWorkSheet.Cells[row, hridCol].GetValue<string>();
                        string payElement = inputWorkSheet.Cells[row, payElementCol].GetValue<string>();
                        string amountText = inputWorkSheet.Cells[row, amountCol].GetValue<string>();
                        double amount = double.TryParse(amountText, out var parsedAmount) ? parsedAmount : 0;
                        // Add pay element to the set
                        payElementCodes.Add(payElement);
                        // Add or update employee data
                        if (!employeeData.ContainsKey(hrid))
                        {
                            employeeData[hrid] = new Dictionary<string, double>();
                        }
                        if (!employeeData[hrid].ContainsKey(payElement))
                        {
                            employeeData[hrid][payElement] = 0;
                        }
                        employeeData[hrid][payElement] += amount;
                    }
                    // Write the output file
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Summary");
                        // Write headers
                        outputWorksheet.Cells[1, 1].Value = "HR ID";
                        int colIndex = 2;
                        var payElementList = new List<string>(payElementCodes);
                        foreach (var payElement in payElementList)
                        {
                            outputWorksheet.Cells[1, colIndex].Value = payElement;
                            colIndex++;
                        }
                        // Write employee data
                        int rowIndex = 2;
                        foreach (var kvp in employeeData)
                        {
                            string hrid = kvp.Key;
                            outputWorksheet.Cells[rowIndex, 1].Value = hrid;

                            for (int i = 0; i < payElementList.Count; i++)
                            {
                                string payElement = payElementList[i];
                                double amount = kvp.Value.ContainsKey(payElement) ? kvp.Value[payElement] : 0;
                                outputWorksheet.Cells[rowIndex, i + 2].Value = amount;
                            }

                            rowIndex++;
                        }
                        for (int column = 1; column <= payElementCodes.Count + 1; column++)
                        {
                            // Get the header value of the current column
                            string temp = outputWorksheet.Cells[1, column].GetValue<string>(); // Correctly reference the column header
                            temp = Program.ShrinkString(temp);
                            // Check for specific keywords
                            bool containsEncashment = temp.Contains("encashment", StringComparison.OrdinalIgnoreCase);
                            bool containsHoliday = temp.Contains("holiday", StringComparison.OrdinalIgnoreCase);
                            bool containsOvertime = temp.Contains("overtime", StringComparison.OrdinalIgnoreCase);
                            bool containsShift = temp.Contains("shift", StringComparison.OrdinalIgnoreCase);
                            if (containsEncashment || containsHoliday || containsOvertime || containsShift)
                            {
                                // Define the range for the entire column
                                var columnRange = outputWorksheet.Cells[1, column, lastRow, column];
                                // Apply fill color
                                columnRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                columnRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                            }
                        }
                        // Save output file
                        string newFileName = Path.Combine(destinationFolder, "Variable_" + Path.GetFileName(filePath));
                        // outputPackage.SaveAs(new FileInfo(outputFilePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                    }
                }
                Console.WriteLine($"Variable Excel file created successfully at {outputFilePath}!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
