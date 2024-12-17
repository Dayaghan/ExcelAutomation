using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomationService
{
    public class Existing_Changes_Master
    {
        public static void Existing_changes_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int row, row2;
                    int IP = Service1.getSheetNumber(filePath, "Joiner and Changes ");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];
                    var BenefeciariesDataSheet = package.Workbook.Worksheets["Beneficiaries Data"];
                    //var OrgAssignmentsDataSheet = package.Workbook.Worksheets["Org Assignments"];
                    //int OrgAssignmentsDataSheetLastRow = OrgAssignmentsDataSheet.Dimension.End.Row;
                    int employeenumber = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int fn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "firstname");
                    int mn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "middlename");
                    int ln = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "surname");
                    int gender = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Gender");
                    int erelation = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "relation");
                    int dateofleaving = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "payroll end date");
                    int add1 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 01");
                    int add2 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 02");
                    int add3 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 03");
                    int town = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "town");
                    int pincode = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "ZIP / Postal Code");
                    int marriedornot = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "marital status");
                    int ifsccode = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "sort code");
                    int acno = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "account number");
                    int dob = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "date of birth");
                    int payrollstartdate = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payroll Start Date");
                    int jobtitle = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "job title");
                    int pancard = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Permanent Account Number (PAN)");
                    int emailid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Email Address");
                    int pension = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Pension Scheme");
                    int nationality = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Nationality");
                    int Aadhar = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Aadhaar Card Number");
                    int uan = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Universal Account Number (UAN)");
                    int ptlocation = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), " PT Location");
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int lastColumn = inputWorkSheet.Dimension.End.Column;
                    int BenefeciarieslastRow = BenefeciariesDataSheet.Dimension.End.Row;
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Existing_Changes_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "Gender -M/F/T";
                        outputWorksheet.Cells[1, 3].Value = "First Name";
                        outputWorksheet.Cells[1, 4].Value = "Middle Name";
                        outputWorksheet.Cells[1, 5].Value = "lastName";
                        outputWorksheet.Cells[1, 6].Value = "Fathers/Husband Name";
                        outputWorksheet.Cells[1, 7].Value = "EmpRelation";
                        outputWorksheet.Cells[1, 8].Value = "Display Name";
                        outputWorksheet.Cells[1, 9].Value = "Marital Status (B/S/M/W)";
                        outputWorksheet.Cells[1, 10].Value = "Spouse Name";
                        outputWorksheet.Cells[1, 11].Value = "No. of Children";
                        outputWorksheet.Cells[1, 12].Value = "Date Of Leaving (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 13].Value = "Reason Of Leaving (S/R/C/D/P)";
                        outputWorksheet.Cells[1, 14].Value = "Present Address 1";
                        outputWorksheet.Cells[1, 15].Value = "Present Address 2";
                        outputWorksheet.Cells[1, 16].Value = "Present Address 3";
                        outputWorksheet.Cells[1, 17].Value = "Present City";
                        outputWorksheet.Cells[1, 18].Value = "Present State Code";
                        outputWorksheet.Cells[1, 19].Value = "Present PinCode";
                        outputWorksheet.Cells[1, 20].Value = "Present Phone";
                        outputWorksheet.Cells[1, 21].Value = "Permanent Address 1";
                        outputWorksheet.Cells[1, 22].Value = "Permanent Address 2";
                        outputWorksheet.Cells[1, 23].Value = "Permanent Address 3";
                        outputWorksheet.Cells[1, 24].Value = "Permanent City";
                        outputWorksheet.Cells[1, 25].Value = "Permanent State Code";
                        outputWorksheet.Cells[1, 26].Value = "Permanent PinCode";
                        outputWorksheet.Cells[1, 27].Value = "Permanent Phone";
                        outputWorksheet.Cells[1, 28].Value = "Primary Bank Code";
                        outputWorksheet.Cells[1, 29].Value = "Primary IFSC";
                        outputWorksheet.Cells[1, 30].Value = "Primary Bank A/c No";
                        outputWorksheet.Cells[1, 31].Value = "Secondary Bank Code";
                        outputWorksheet.Cells[1, 32].Value = "Secondary IFSC";
                        outputWorksheet.Cells[1, 33].Value = "Secondary Bank A/c No";
                        outputWorksheet.Cells[1, 34].Value = "Payroll Code";
                        outputWorksheet.Cells[1, 35].Value = "Date Of Joining (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 36].Value = "Training Start Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 37].Value = "Probation Start Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 38].Value = "Date of Confirmation (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 39].Value = "Date of Retirement (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 40].Value = "Date Of Birth (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 41].Value = "Marriage Anniversary Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 42].Value = "Super Annuation Number";
                        outputWorksheet.Cells[1, 43].Value = "SA wef Dt. (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 44].Value = "Super Annuation Percent";
                        outputWorksheet.Cells[1, 45].Value = "Super Annuation Max Limit";
                        outputWorksheet.Cells[1, 46].Value = "Gratuity Number";
                        outputWorksheet.Cells[1, 47].Value = "Gratuity wef Dt. (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 48].Value = "Gratuity %";
                        outputWorksheet.Cells[1, 49].Value = "FPS Number";
                        outputWorksheet.Cells[1, 50].Value = "Category Code";
                        outputWorksheet.Cells[1, 51].Value = "Status Code";
                        outputWorksheet.Cells[1, 52].Value = "Grade Code";
                        outputWorksheet.Cells[1, 53].Value = "Designation";
                        outputWorksheet.Cells[1, 54].Value = "Cost Centre Code";
                        outputWorksheet.Cells[1, 55].Value = "Business Area Code";
                        outputWorksheet.Cells[1, 56].Value = "Location Code";
                        outputWorksheet.Cells[1, 57].Value = "Leave Approver";
                        outputWorksheet.Cells[1, 58].Value = "Occupation Code";
                        outputWorksheet.Cells[1, 59].Value = "Qualification";
                        outputWorksheet.Cells[1, 60].Value = "Permanent A/c No.";
                        outputWorksheet.Cells[1, 61].Value = "P.F. Registration Code";
                        outputWorksheet.Cells[1, 62].Value = "P.F. A/c No.(10 Digits)";
                        outputWorksheet.Cells[1, 63].Value = "PF wef Dt. (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 64].Value = "E.S.I. No.";
                        outputWorksheet.Cells[1, 65].Value = "ESIC Clinic";
                        outputWorksheet.Cells[1, 66].Value = "Blood Group";
                        outputWorksheet.Cells[1, 67].Value = "Emergency Phone No.";
                        outputWorksheet.Cells[1, 68].Value = "Emergency Contact Person";
                        outputWorksheet.Cells[1, 69].Value = "Email ID";
                        outputWorksheet.Cells[1, 70].Value = "Reports To Emp";
                        outputWorksheet.Cells[1, 71].Value = "Passport No";
                        outputWorksheet.Cells[1, 72].Value = "Passport Validity (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 73].Value = "Mobile No";
                        outputWorksheet.Cells[1, 74].Value = "Web User Name";
                        outputWorksheet.Cells[1, 75].Value = "Web User Password";
                        outputWorksheet.Cells[1, 76].Value = "User Profiles";
                        outputWorksheet.Cells[1, 77].Value = "Voluntary PF %";
                        outputWorksheet.Cells[1, 78].Value = "Date Of Resign (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 79].Value = "Photo File Path";
                        outputWorksheet.Cells[1, 80].Value = "MICR";
                        outputWorksheet.Cells[1, 81].Value = "MICR2";
                        outputWorksheet.Cells[1, 82].Value = "IsLocalAuthentication";
                        outputWorksheet.Cells[1, 83].Value = "Userdefined 1 Code";
                        if (filePath.Contains("mcafee") || filePath.Contains("musarubra"))
                        {
                            outputWorksheet.Cells[1, 83].Value = "Pay Scale";
                        }
                        outputWorksheet.Cells[1, 84].Value = "Userdefined 2 Code";
                        outputWorksheet.Cells[1, 85].Value = "Note";
                        outputWorksheet.Cells[1, 86].Value = "Userdefined 4 Code";
                        outputWorksheet.Cells[1, 87].Value = "Userdefined 5 Code";
                        outputWorksheet.Cells[1, 88].Value = "Userdefined 6";
                        outputWorksheet.Cells[1, 89].Value = "Userdefined 7 Code";
                        outputWorksheet.Cells[1, 90].Value = "Userdefined 8 Code";
                        outputWorksheet.Cells[1, 91].Value = "Userdefined 9";
                        outputWorksheet.Cells[1, 92].Value = "Userdefined 10 Code";
                        outputWorksheet.Cells[1, 93].Value = "Userdefined 11 Code";
                        outputWorksheet.Cells[1, 94].Value = "Aadhaar Card No";
                        outputWorksheet.Cells[1, 95].Value = "Primary Name As Per Bank";
                        outputWorksheet.Cells[1, 96].Value = "Secondary Name As Per Bank";
                        outputWorksheet.Cells[1, 97].Value = "InactiveID";
                        outputWorksheet.Cells[1, 98].Value = "Inactive Notes";
                        outputWorksheet.Cells[1, 99].Value = "Process Last Month In FNF";
                        outputWorksheet.Cells[1, 100].Value = "UAN";
                        outputWorksheet.Cells[1, 101].Value = "Pension Scheme";
                        outputWorksheet.Cells[1, 102].Value = "Personal Email ID";
                        outputWorksheet.Cells[1, 103].Value = "Group Joining Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 104].Value = "PRAN";
                        outputWorksheet.Cells[1, 105].Value = "Nationality";
                        outputWorksheet.Cells[1, 106].Value = "Religion";
                        outputWorksheet.Cells[1, 107].Value = "Personal Mobile No";
                        outputWorksheet.Cells[1, 108].Value = "Company Superannuation No";
                        outputWorksheet.Cells[1, 109].Value = "Notice Period Confirmed Days/Months";
                        outputWorksheet.Cells[1, 110].Value = "Notice Period Confirmed Type";
                        outputWorksheet.Cells[1, 111].Value = "Notice Period Probation Days/Months";
                        outputWorksheet.Cells[1, 112].Value = "Notice Period Probation Type";
                        outputWorksheet.Cells[1, 113].Value = "Labour Indentification No";
                        outputWorksheet.Cells[1, 114].Value = "Division Code";
                        outputWorksheet.Cells[1, 115].Value = "Training End Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 116].Value = "Probation End Date (YYYY-MM-DD)";
                        int row7 = 2;
                        for (row = 2; row <= lastRow; row++)
                        {

                            var cell = inputWorkSheet.Cells[row, employeenumber];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF"))
                            {

                                var HRID = inputWorkSheet.Cells[row, employeenumber].GetValue<string>();
                                outputWorksheet.Cells[row7, 1].Value = HRID;

                                var Firstname = inputWorkSheet.Cells[row, fn].GetValue<string>();
                                var LastName = inputWorkSheet.Cells[row, ln].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, fn];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 3].Value = Firstname;
                                    outputWorksheet.Cells[row7, 8].Value = Firstname+" "+LastName;
                                }
                                cell = inputWorkSheet.Cells[row, mn];
                                var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {

                                    outputWorksheet.Cells[row7, 4].Value = MiddleName;
                                }
                                
                                cell = inputWorkSheet.Cells[row, ln];
                                //var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 5].Value = LastName;
                                    outputWorksheet.Cells[row7, 8].Value = Firstname + " " + LastName;
                                }
                                var Gender = inputWorkSheet.Cells[row, gender].GetValue<string>();
                                Gender = Gender.ToLower();
                                Gender = Gender.Replace(" ", "");
                                switch (Gender)
                                {
                                    case "male":
                                        Gender = "M";
                                        break;
                                    case "female":
                                        Gender = "F";
                                        break;
                                    case "transgender":
                                        Gender = "T";
                                        break;
                                }
                                var columnRange = outputWorksheet.Cells[1, 55, lastRow, 55];
                                columnRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                columnRange.Style.Fill.BackgroundColor.SetColor(Color.Red);
                                //outputWorksheet.Cells[row7, 2].Value = Gender;
                                var MaritalStatus = inputWorkSheet.Cells[row, marriedornot].GetValue<string>();
                                MaritalStatus = MaritalStatus.ToUpper();
                                MaritalStatus = MaritalStatus.Replace(" ", "");
                                cell = inputWorkSheet.Cells[row, marriedornot];
                                //var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    switch (MaritalStatus)
                                    {
                                        case "BACHELOR":
                                            MaritalStatus = "B";
                                            break;
                                        case "B":
                                            MaritalStatus = "B";
                                            break;
                                        case "BACHLOR":
                                            MaritalStatus = "B";
                                            break;
                                        case "MARRIED":
                                            MaritalStatus = "M";
                                            break;
                                        case "M":
                                            MaritalStatus = "M";
                                            break;
                                        case "WIDOW":
                                            MaritalStatus = "W";
                                            break;
                                        case "W":
                                            MaritalStatus = "W";
                                            break;
                                        case "WIDOWED":
                                            MaritalStatus = "W";
                                            break;
                                        default:
                                            MaritalStatus = "B";
                                            break;
                                    }
                                    outputWorksheet.Cells[row7, 9].Value = MaritalStatus;
                                }
                                var UAN = inputWorkSheet.Cells[row, uan].GetValue<string>();
				                cell = inputWorkSheet.Cells[row, uan];
				                bgColor = cell.Style.Fill.BackgroundColor;
				                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
				                {
   				                    outputWorksheet.Cells[row7, 100].Value = UAN;
				                }
                                var Address = inputWorkSheet.Cells[row, add1].GetValue<string>();
                                // outputWorksheet.Cells[row7, 14].Value = Address;
                                Address = inputWorkSheet.Cells[row, add2].GetValue<string>();
                                // outputWorksheet.Cells[row7, 15].Value = Address;
                                Address = inputWorkSheet.Cells[row, add3].GetValue<string>();
                                //outputWorksheet.Cells[row7, 16].Value = Address;
                                Address = inputWorkSheet.Cells[row, town].GetValue<string>();
                                //outputWorksheet.Cells[row7, 17].Value = Address;
                                //Address = inputWorkSheet.Cells[row, pincode].GetValue<string>();
                                // outputWorksheet.Cells[row7, 19].Value = Address;
                                var email = inputWorkSheet.Cells[row, emailid].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, emailid];
                                //var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 69].Value = email;
                                }
                                var date = inputWorkSheet.Cells[row, dob].GetValue<string>();
                                date = date.Replace(" ", "");
                                if ((date.Length == 10) && (date[4] == '-'))
                                {
                                    //outputWorksheet.Cells[row7, 40].Value = date;
                                }
                                date = inputWorkSheet.Cells[row, payrollstartdate].GetValue<string>();
                                date = date.Replace(" ", "");
                                if ((date.Length == 10) && (date[4] == '-'))
                                {
                                    //outputWorksheet.Cells[row7, 35].Value = date;
                                    //outputWorksheet.Cells[row7, 63].Value = date;
                                    //outputWorksheet.Cells[row7, 103].Value = date;
                                }
                                var pan = inputWorkSheet.Cells[row, pancard].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, pancard];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    pan = pan.Replace(" ", "");
                                    if ((pan.Length == 10) && (pan[3] == 'P'))
                                    {
                                        outputWorksheet.Cells[row7, 60].Value = pan;
                                    }
                                    if (pan.Length == 0)
                                    {
                                        outputWorksheet.Cells[row7, 60].Value = "PANNOTAVBLE";
                                    }
                                }
                                var adhaar = (inputWorkSheet.Cells[row, Aadhar].GetValue<string>()).Replace(" ", "");
                                cell = inputWorkSheet.Cells[row, Aadhar];
                                //var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    if ((adhaar.Length == 12) && (adhaar.All(char.IsDigit)))
                                    {
                                        outputWorksheet.Cells[row7, 94].Value = adhaar;
                                    }
                                }
                               // var UAN = inputWorkSheet.Cells[row, uan].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, jobtitle];
                                var JobTitle = inputWorkSheet.Cells[row, jobtitle].GetValue<string>();
                                //var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 53].Value = JobTitle;
                                }
                                var PTLocation = inputWorkSheet.Cells[row, ptlocation].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, ptlocation];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 56].Value = PTLocation;
                                }
                                if (filePath.Contains("mcafee") || filePath.Contains("musarubra"))
                                {
                                    int EmployeeGrade = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Grade");
                                    var grade = inputWorkSheet.Cells[row, EmployeeGrade].GetValue<string>();
                                    cell = inputWorkSheet.Cells[row, EmployeeGrade];
                                    // Get the background color of the cell
                                    bgColor = cell.Style.Fill.BackgroundColor;
                                    if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF")) { 
                                        if (grade.Contains("grade"))
                                        {
                                        grade = grade.Substring(grade.Length - 2);
                                        }
                                    outputWorksheet.Cells[row7, 83].Value = grade;
                                    }
                                }
                                var Nationality = inputWorkSheet.Cells[row, nationality].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, nationality];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    // outputWorksheet.Cells[row7, 105].Value = Nationality;
                                }
                                //outputWorksheet.Cells[row7, 105].Value = Nationality;
                                var Pension = inputWorkSheet.Cells[row, pension].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, pension];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    Pension = Pension.ToLower();
                                    Pension = Pension.Replace(" ", "");
                                    switch (Pension)
                                    {
                                        case "yes":
                                            Pension = "1";
                                            break;
                                        case "no":
                                            Pension = "0";
                                            break;
                                        case "0":
                                            Pension = "0";
                                            break;
                                        case "1":
                                            Pension = "1";
                                            break;
                                        default:
                                            Pension = "0";
                                            break;
                                    }
                                    outputWorksheet.Cells[row7, 101].Value = Pension;
                                }
                                row7++;
                            }
                        }
                        lastColumn = outputWorksheet.Dimension.End.Column;
                        lastRow = outputWorksheet.Dimension.End.Row;
                        for (int i = 1; i <= lastRow; i++)
                        {
                            int j;
                            for (j = 1; j <= lastColumn; j++)
                            {
                                string cellValue = outputWorksheet.Cells[i, j].GetValue<string>();
                                if ((cellValue == null) || (cellValue == "") || (cellValue == " "))
                                {
                                    outputWorksheet.Cells[i, j].Value = "!!!";
                                }
                            }
                        }
                        for (int j = 2; j <= lastColumn; j++)
                        {
                            var cell = inputWorkSheet.Cells[1, j];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF"))
                            {
                                string sheetname = outputWorksheet.Cells[1, j].GetValue<string>();
                                var outputWorksheets = outputPackage.Workbook.Worksheets.Add(sheetname);
                                outputWorksheets.Cells[1, 1].Value = "HR ID";
                                outputWorksheets.Cells[1, 2].Value = sheetname;
                                int m = 2;
                                int k = 2;
                                for (int l = 2; l <= lastRow; l++)
                                {
                                    if (outputWorksheet.Cells[l, j].GetValue<string>() != "!!!")
                                    {
                                        outputWorksheets.Cells[m, k - 1].Value = outputWorksheet.Cells[l, 1].GetValue<string>();
                                        outputWorksheets.Cells[m, k].Value = outputWorksheet.Cells[l, j].GetValue<string>();
                                        m++;
                                    }
                                }
                            }
                        }
                        //Logic To delete the Empty Sheets.
                        int worksheetCount = outputPackage.Workbook.Worksheets.Count;
                        for (int i = worksheetCount - 1; i >= 1; i--)
                        {
                            var testworksheet = outputPackage.Workbook.Worksheets[i];
                            string temp = testworksheet.Cells[2, 1].GetCellValue<string>();
                            if ((temp == null) || (temp == ""))
                            {
                                outputPackage.Workbook.Worksheets.Delete(i);
                            }
                        }
                        string newFileName = Path.Combine(destinationFolder, "Existing_Changes_Master" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.Workbook.Worksheets.Delete(0);
                        if (outputPackage.Workbook.Worksheets.Count != 0)
                        {
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.Log("Existing_Changes_Master Excel file created successfully!");
                        }
                        else
                        {
                            Service1.Log("No existing employee changes");
                        }
                        
                    }
                }
            }
            catch (Exception ex)
            {
                Service1.ErrorCount++;
                Service1.Log($"An error occurred: {ex.Message}");
            }
        }
    }
}