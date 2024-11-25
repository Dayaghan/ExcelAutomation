using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelHandlingDotnetpractice
{
    public class Benefeciaries_Data
    {

        public static void Beneficiaries_Data(string filePath, string destinationFolder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                string ascendcodes = "C:/Automation/Twilio_Twilio Technology/Automation_Ascent_Codes/Ascent Codes.xlsx";
                string outputFilePath = "C:/Users/Dayaghanl/Desktop/output.xlsx";
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Program.getSheetNumber(filePath, "Beneficiaries Data");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                    int lastRow = inputWorkSheet.Dimension.End.Row;

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Benificieries Data");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "Primary NameAsPerBank";
                        outputWorksheet.Cells[1, 3].Value = "Primary Bank A / c No";
                        outputWorksheet.Cells[1, 4].Value = "Primary IFSC";
                        outputWorksheet.Cells[1, 5].Value = " Primary Bank Code";
                        int hrid = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                        int benefeciaryname = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Beneficiary Name");
                        int acno = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "account number");
                        int ifsc = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "sort code");
                        for (int row = 2; row <= lastRow; row++)
                        {
                            //string whiteColour = "16777215";
                            //string whiteColorHex = "FFFFFF";  // HEX representation of white color
                            var cell = inputWorkSheet.Cells[row, hrid]; // Example cell to check color
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            // Check if cell has a background color and compare it
                            if (!string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Theme != null)
                            {
                                continue;
                            }
                            var HRID = inputWorkSheet.Cells[row, hrid].GetValue<string>();
                            var BENEFICIARYNAME = inputWorkSheet.Cells[row, benefeciaryname].GetValue<string>();
                            var PrimaryBankAcNO = inputWorkSheet.Cells[row, acno].GetValue<string>();
                            var IFSC = inputWorkSheet.Cells[row, ifsc].GetValue<string>();
                            outputWorksheet.Cells[row, 1].Value = HRID;
                            Regex validCharsRegex = new Regex("[^a-zA-Z ]");
                            BENEFICIARYNAME = validCharsRegex.Replace(BENEFICIARYNAME, "");
                            outputWorksheet.Cells[row, 2].Value = BENEFICIARYNAME;
                            if ((PrimaryBankAcNO.All(char.IsDigit)))
                            {
                                outputWorksheet.Cells[row, 3].Value = PrimaryBankAcNO;
                            }
                            IFSC = IFSC.Replace(" ", "");
                            if (IFSC.Length == 11)
                            {
                                outputWorksheet.Cells[row, 4].Value = IFSC;
                            }
                        }
                        int endRow = outputWorksheet.Dimension.End.Row;
                        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                        {
                            var Ascendsheet = package2.Workbook.Worksheets["Primary Bank"];
                            int bankcode = Program.getColumnNumber(ascendcodes, Ascendsheet.ToString(), "bank code");
                            int bankname = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "bank name");
                            int bankname2 = Program.getColumnNumber(ascendcodes, Ascendsheet.ToString(), "bank name");
                            int lastRow2 = Ascendsheet.Dimension.End.Row;
                            for (int row = 2; row <= endRow; row++)
                            {
                                string Bankname = inputWorkSheet.Cells[row, bankname].GetValue<string>();
                                Bankname = Bankname.ToLower();
                                Bankname = Bankname.Trim();
                                int row2 = 2;
                                for (; row2 <= lastRow2; row2++)
                                {
                                    string temp = Ascendsheet.Cells[row2, bankname2].GetValue<string>();
                                    temp = temp.ToLower();
                                    temp = temp.Trim();
                                    if (temp.Equals(Bankname))
                                    {
                                        outputWorksheet.Cells[row, 5].Value = Ascendsheet.Cells[row2, bankcode].GetValue<string>();
                                        break;
                                    }
                                }
                            }
                        }
                        string newFileName = Path.Combine(destinationFolder, "Benificieries Data_" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                    }
                }
                Console.WriteLine("Beneficiaries Data Excel file created successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
