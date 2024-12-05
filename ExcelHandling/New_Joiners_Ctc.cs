using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHandlingDotnetpractice
{
    public class New_Joiners_Ctc
    {
        public static void CTC_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets["Payments and Deductions"];
                    var joinerandChangesSheet = package.Workbook.Worksheets["Joiner and Changes "];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int lastRow2 = joinerandChangesSheet.Dimension.End.Row;
                    int employee_Number = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int payelementdescription = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Description");
                    int town = Program.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "PT Location");
                    int hrid = Program.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Hr id");
                    // int payelementdescription = Program.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int annualctc = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Amount");
                    int payfreq = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY - MM - DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        outputWorksheet.Cells[1, 4].Value = "001 Basic Salary";
                        outputWorksheet.Cells[1, 5].Value = "Location";
                        outputWorksheet.Cells[1, 6].Value = "003 HRA";
                        outputWorksheet.Cells[1, 7].Value = "004 LTA";
                        outputWorksheet.Cells[1, 8].Value = "006 Cash Allowance";
                        outputWorksheet.Cells[1, 9].Value = "014 Employer NPS";
                        outputWorksheet.Cells[1, 10].Value = "031 Stipend";
                        outputWorksheet.Cells[1, 11].Value = "012 Connect";
                        outputWorksheet.Cells[1, 12].Value = "Employer PF";
                        HashSet<string> HRID = new HashSet<string>();
                        int row2 = 2;
                        for (int row = 2; row <= lastRow2; row++)
                        {
                            var cell = joinerandChangesSheet.Cells[row, hrid];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase))
                            {
                                HRID.Add(inputWorkSheet.Cells[row, employee_Number].GetValue<string>());
                                outputWorksheet.Cells[row2, 1].Value = joinerandChangesSheet.Cells[row, hrid].GetValue<string>();
                                outputWorksheet.Cells[row2, 5].Value = joinerandChangesSheet.Cells[row, town].GetValue<string>();
                                row2++;
                            }
                        }
                        row2 = 2;
                        //foreach (string t in HRID)
                        //{
                        //    outputWorksheet.Cells[row2, 1].Value = t;
                        //    for (int row = 2; row <= lastRow; row++)
                        //    {
                        //        if (t.Equals(joinerandChangesSheet.Cells[row,hrid].GetValue<string>())) {
                        //            outputWorksheet.Cells[row2, 5].Value = joinerandChangesSheet.Cells[row, town].GetValue<string>();
                        //            break;
                        //        }
                        //    }
                        //    row2++;
                        //}
                        //row2 = 2;
                        //foreach (string t in HRID)
                        //{
                        //    for (int row = 2; row <= lastRow; row++)
                        //    {
                        //        if (inputWorkSheet.Cells[row, employee_Number].GetValue<string>().Equals(t))
                        //        {
                        //            outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].GetValue<string>();
                        //            if (inputWorkSheet.Cells[row, payfreq].GetValue<string>().Equals("Annual"))
                        //            {
                        //                outputWorksheet.Cells[row2, 3].Value = inputWorkSheet.Cells[row, annualctc].GetValue<string>();
                        //                double monthly40 = (inputWorkSheet.Cells[row, annualctc].GetValue<double>()) / 30.0;
                        //                monthly40 = Math.Round(monthly40);
                        //                outputWorksheet.Cells[row2, 4].Value = monthly40;
                        //                outputWorksheet.Cells[row2, 9].Value = 0;
                        //                outputWorksheet.Cells[row2, 10].Value = 0;
                        //                outputWorksheet.Cells[row2, 12].Value = Math.Round((monthly40) * 12 / 100);
                        //                outputWorksheet.Cells[row2, 6].Value = Math.Round(outputWorksheet.Cells[row2, 4].GetValue<double>() * 0.4);
                        //                outputWorksheet.Cells[row2, 7].Value = 0;
                        //            }
                        //        }
                        //    }
                        //    row2++;
                        //}
                        for (int row = 2; row <= outputWorksheet.Dimension.End.Row; row++)
                        {
                            for (row2 = 2; row2 <= lastRow; row2++)
                            {
                                var cell = outputWorksheet.Cells[row, 1];
                                // Get the background color of the cell
                                var bgColor = cell.Style.Fill.BackgroundColor;
                                if (outputWorksheet.Cells[row, 1].GetValue<string>().Equals(inputWorkSheet.Cells[row2, employee_Number].GetValue<string>()))
                                {
                                    outputWorksheet.Cells[row, 2].Value = inputWorkSheet.Cells[row2, witheffectfrom].GetValue<string>();
                                    if (inputWorkSheet.Cells[row2, payelementdescription].GetValue<string>().Contains("connect", StringComparison.OrdinalIgnoreCase))
                                    {
                                        outputWorksheet.Cells[row, 11].Value = inputWorkSheet.Cells[row2, annualctc].GetValue<double>();
                                    }
                                    if (inputWorkSheet.Cells[row2, payfreq].GetValue<string>().Contains("annual", StringComparison.OrdinalIgnoreCase))
                                    {
                                        //annual = inputWorkSheet.Cells[row, annualctc].GetValue<double>();
                                        outputWorksheet.Cells[row, 3].Value = inputWorkSheet.Cells[row2, annualctc].GetValue<double>();
                                        //double monthly40 = (inputWorkSheet.Cells[row2, annualctc].GetValue<double>()) / 30.0;
                                        //monthly40 = Math.Round(monthly40);
                                        //outputWorksheet.Cells[row, 4].Value = monthly40;
                                        //outputWorksheet.Cells[row, 9].Value = 0;
                                        //outputWorksheet.Cells[row, 10].Value = 0;
                                        //outputWorksheet.Cells[row, 12].Value = Math.Round((monthly40) * 12 / 100);
                                        //outputWorksheet.Cells[row, 6].Value = Math.Round(outputWorksheet.Cells[row, 4].GetValue<double>() * 0.4);
                                        //outputWorksheet.Cells[row, 7].Value = 0;
                                    }
                                }
                            }
                        }
                        //for (int row = 2; row <= lastRow; row++)
                        //{
                        //    if ((outputWorksheet.Cells[row, 5].GetValue<string>()) != null)
                        //    {
                        //        string city = outputWorksheet.Cells[row, 5].GetValue<string>();
                        //        city = Program.ShrinkString(city);
                        //        bool delhi = city.Contains("delhi", StringComparison.OrdinalIgnoreCase);
                        //        bool mumbai = city.Contains("mumbai", StringComparison.OrdinalIgnoreCase);
                        //        bool maharashtra = city.Contains("maharashtra", StringComparison.OrdinalIgnoreCase);
                        //        bool tamilnadu = city.Contains("tamilnadu", StringComparison.OrdinalIgnoreCase);
                        //        bool newdelhi = city.Contains("newdelhi", StringComparison.OrdinalIgnoreCase);
                        //        bool chennai = city.Contains("chennai", StringComparison.OrdinalIgnoreCase);
                        //        bool kolkata = city.Contains("kolkata", StringComparison.OrdinalIgnoreCase);
                        //        bool calcutta = city.Contains("calcutta", StringComparison.OrdinalIgnoreCase);
                        //        if (delhi || mumbai || newdelhi || chennai || kolkata || maharashtra || tamilnadu)
                        //        {
                        //            outputWorksheet.Cells[row, 6].Value = outputWorksheet.Cells[row, 4].GetValue<double>() / 2;
                        //        }
                        //        outputWorksheet.Cells[row, 6].Value = Math.Round(outputWorksheet.Cells[row, 6].GetValue<double>());
                        //        outputWorksheet.Cells[row, 8].Value = Math.Round((outputWorksheet.Cells[row, 3].GetValue<double>()) / 12 - outputWorksheet.Cells[row, 4].GetValue<double>() - outputWorksheet.Cells[row, 6].GetValue<double>() - outputWorksheet.Cells[row, 7].GetValue<double>() - outputWorksheet.Cells[row, 9].GetValue<double>());
                        //    }
                        //}
                        //row2 = 2;
                        //foreach (string t in HRID)
                        //{
                        //    outputWorksheet.Cells[row2, 8].Value = Math.Round((outputWorksheet.Cells[row2, 3].GetValue<double>()) / 12 - outputWorksheet.Cells[row2, 4].GetValue<double>() - outputWorksheet.Cells[row2, 6].GetValue<double>() - outputWorksheet.Cells[row2, 7].GetValue<double>() - outputWorksheet.Cells[row2, 9].GetValue<double>());
                        //    for (int row = 2; row <= lastRow; row++)
                        //    {
                        //        if (inputWorkSheet.Cells[row, employee_Number].GetValue<string>().Equals(t))
                        //        {
                        //            string temp = inputWorkSheet.Cells[row, payelementdescription].GetValue<string>();
                        //            temp = temp.ToLower();
                        //            temp = temp.Replace(" ", "");
                        //            if (temp.Contains("connect"))
                        //            {
                        //                outputWorksheet.Cells[row2, 11].Value = inputWorkSheet.Cells[row, annualctc].GetValue<string>();
                        //            }
                        //        }
                        //    }
                        //    row2++;
                        //}
                        outputWorksheet.Column(3).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(4).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(6).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(7).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(8).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(9).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(10).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(11).Style.Numberformat.Format = "0.00";
                        //outputWorksheet.Column(12).Style.Numberformat.Format = "0.00";
                        string newFileName = Path.Combine(destinationFolder, "NEW_Joiners_Ctc_Master_" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                    }
                    Console.WriteLine("CTC Excel file created successfully!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
