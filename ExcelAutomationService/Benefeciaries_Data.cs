﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelAutomationService
{
    public class Benefeciaries_Data
    {
        public static void Beneficiaries_Data(string ascendcodes,string filePath, string destinationFolder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Service1.getSheetNumber(filePath, "Beneficiaries Data");
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
                        outputWorksheet.Cells[1, 6].Value = "Bank Name";
                        int hrid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                        int bn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Bank Name");
                        int benefeciaryname = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Beneficiary Name");
                        int acno = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "account number");
                        int ifsc = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "sort code");
                        int row3 = 2;
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
                            outputWorksheet.Cells[row3, 1].Value = HRID;
                            Regex validCharsRegex = new Regex("[^a-zA-Z ]");
                            BENEFICIARYNAME = validCharsRegex.Replace(BENEFICIARYNAME, "");
                            outputWorksheet.Cells[row3, 2].Value = BENEFICIARYNAME;
                            if ((PrimaryBankAcNO.All(char.IsDigit)))
                            {
                                outputWorksheet.Cells[row3, 3].Value = PrimaryBankAcNO;
                            }
                            IFSC = IFSC.Replace(" ", "");
                            if (IFSC.Length == 11)
                            {
                                outputWorksheet.Cells[row3, 4].Value = IFSC;
                            }
                            var bankname = inputWorkSheet.Cells[row, bn].GetValue<string>();
                            outputWorksheet.Cells[row3, 6].Value = bankname;
                            row3++;
                        }
                        row3 = 2;
                        int endRow = outputWorksheet.Dimension.End.Row;
                        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                        {
                            //int n = Service1.getSheetNumber(ascendcodes, "%bank%");
                            var Ascendsheet = package2.Workbook.Worksheets["Banks Detailed"];
                            int bankcode = Service1.getColumnNumber(ascendcodes, Ascendsheet.ToString(), "Code");
                            int Bankname = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Bank Name");
                            int bankname2 = Service1.getColumnNumber(ascendcodes, Ascendsheet.ToString(), "Name of Bank");
                            int lastRow2 = Ascendsheet.Dimension.End.Row;
                            for (int row = 2; row <= endRow; row++)
                            {
                                string bankname = outputWorksheet.Cells[row, 6].GetValue<string>();
                                bankname = Service1.ShrinkString(bankname);
                                bankname = bankname.Replace("ltd", "");
                                bankname = bankname.Replace("limited", "");
                                bankname = bankname.Replace("pvt", "");
                                bankname = bankname.Replace(".", "");
                                bool containsBank = bankname.Contains("bank");
                                if (!containsBank)
                                {
                                    bankname = bankname + "bank";
                                }
                                int row2 = 2;
                                for (; row2 <= lastRow2; row2++)
                                {
                                    string temp = Ascendsheet.Cells[row2, bankname2].GetValue<string>();
                                    temp = Service1.ShrinkString(temp);
                                    if (temp.Equals(bankname))
                                    {
                                        outputWorksheet.Cells[row, 5].Value = Ascendsheet.Cells[row2, bankcode].GetValue<string>();

                                    }
                                }
                            }
                            outputWorksheet.DeleteColumn(6);
                        }
                        string newFileName = Path.Combine(destinationFolder, "Benificieries Data_" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                    }
                }
                Service1.Log("Beneficiaries Data Excel file created successfully!");
            }
            catch (Exception ex)
            {
                Service1.Log($"An error occurred: {ex.Message}");
            }
        }
    }
}