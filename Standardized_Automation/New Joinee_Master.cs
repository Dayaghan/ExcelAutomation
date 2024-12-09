using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Formats.Tar;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PayrollAutomationService
{
    public class New_Joinee_Master
    {
        public static void NewJoinee_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int row, row2;
                    int IP = Program.getSheetNumber(filePath, "Joiner and Changes ");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];
                    var BenefeciariesDataSheet = package.Workbook.Worksheets["Beneficiaries Data"];
                    //var OrgAssignmentsDataSheet = package.Workbook.Worksheets["Org Assignments"];
                    //int OrgAssignmentsDataSheetLastRow = OrgAssignmentsDataSheet.Dimension.End.Row;
                    int employeenumber = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int Aadhar = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Aadhaar Card Number");
                    int uan = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Universal Account Number (UAN)");
                    int PreferredName = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Preferred Name");
                    int fn = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "firstname");
                    int mn = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "middlename");
                    int ln = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "surname");
                    int gender = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Gender");
                    int erelation = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "relation");
                    int dateofleaving = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "payroll end date");
                    int add1 = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 01");
                    int add2 = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 02");
                    int add3 = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 03");
                    int town = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "town");
                    int pincode = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "ZIP / Postal Code");
                    int marriedornot = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "marital status");
                    int ifsccode = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "sort code");
                    int acno = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "account number");
                    int dob = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "date of birth");
                    int payrollstartdate = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payroll Start Date");
                    int jobtitle = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "job title");
                    int pancard = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Permanent Account Number (PAN)");
                    int emailid = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Email Address");
                    int pension = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Pension Scheme");
                    int bfhrid = Program.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Hr id");
                    int primarynameasperbank = Program.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Beneficiary Name");
                    int bfbankname = Program.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Bank Name");
                    int bfifsc = Program.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Sort Code");
                    int bfacno = Program.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Account Number");
                    int nationality = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Nationality");
                    int ptlocation = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), " PT Location");
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int BenefeciarieslastRow = BenefeciariesDataSheet.Dimension.End.Row;
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("New Joinee_Master");
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
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase))
                            {
                                var HRID = inputWorkSheet.Cells[row, 2].GetValue<string>();
                                outputWorksheet.Cells[row7, 1].Value = HRID;
                                var Firstname = inputWorkSheet.Cells[row, fn].GetValue<string>();
                                var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                var LastName = inputWorkSheet.Cells[row, ln].GetValue<string>();
                                // var FatherOrHusband = inputWorkSheet.Cells[row, 7].GetValue<string>();
                                //var Relation = inputWorkSheet.Cells[row, 7].GetValue<string>();
                                outputWorksheet.Cells[row7, 3].Value = Firstname;
                                outputWorksheet.Cells[row7, 4].Value = MiddleName;
                                if (LastName == "")
                                {
                                    outputWorksheet.Cells[row7, 5].Value = ".";
                                }
                                else
                                {
                                    outputWorksheet.Cells[row7, 5].Value = LastName;
                                }
                                outputWorksheet.Cells[row7, 8].Value = Firstname + " " + LastName;
                                var Gender = inputWorkSheet.Cells[row, gender].GetValue<string>();
                                Gender = Program.ShrinkString(Gender);
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
                                outputWorksheet.Cells[row7, 2].Value = Gender;
                                var MaritalStatus = inputWorkSheet.Cells[row, marriedornot].GetValue<string>();
                                MaritalStatus = MaritalStatus.ToUpper();
                                MaritalStatus = MaritalStatus.Replace(" ", "");
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
                                var Address = inputWorkSheet.Cells[row, add1].GetValue<string>();
                                outputWorksheet.Cells[row7, 14].Value = Address;
                                Address = inputWorkSheet.Cells[row, add2].GetValue<string>();
                                outputWorksheet.Cells[row7, 15].Value = Address;
                                Address = inputWorkSheet.Cells[row, add3].GetValue<string>();
                                outputWorksheet.Cells[row7, 16].Value = Address;
                                Address = inputWorkSheet.Cells[row, town].GetValue<string>();
                                outputWorksheet.Cells[row7, 17].Value = Address;
                                Address = inputWorkSheet.Cells[row, pincode].GetValue<string>();
                                outputWorksheet.Cells[row7, 19].Value = Address;
                                var email = inputWorkSheet.Cells[row, emailid].GetValue<string>();
                                outputWorksheet.Cells[row7, 69].Value = email;
                                outputWorksheet.Cells[row7, 74].Value = HRID;
                                outputWorksheet.Cells[row7, 31].Value = "00000";
                                var date = inputWorkSheet.Cells[row, 9].GetValue<string>();
                                date = date.Replace(" ", "");
                                if ((date.Length == 10) && (date[4] == '-'))
                                {
                                    outputWorksheet.Cells[row7, 40].Value = date;
                                }
                                date = inputWorkSheet.Cells[row, 14].GetValue<string>();
                                date = date.Replace(" ", "");
                                if ((date.Length == 10) && (date[4] == '-'))
                                {
                                    outputWorksheet.Cells[row7, 35].Value = date;
                                    outputWorksheet.Cells[row7, 63].Value = date;
                                    outputWorksheet.Cells[row7, 103].Value = date;
                                }
                                var pan = inputWorkSheet.Cells[row, pancard].GetValue<string>();
                                pan = pan.Replace(" ", "");
                                if ((pan.Length == 10) && (pan[3] == 'P'))
                                {
                                    outputWorksheet.Cells[row7, 60].Value = pan;
                                }
                                if (pan.Length == 0)
                                {
                                    outputWorksheet.Cells[row7, 60].Value = "PANNOTAVBLE";
                                }
                                var adhaar = (inputWorkSheet.Cells[row, Aadhar].GetValue<string>()).Replace(" ", "");

                                if ((adhaar.Length == 12) && (adhaar.All(char.IsDigit)))
                                {
                                    outputWorksheet.Cells[row7, 94].Value = adhaar;
                                }
                                var UAN = inputWorkSheet.Cells[row, uan].GetValue<string>();
                                outputWorksheet.Cells[row7, 100].Value = UAN;
                                var JobTitle = inputWorkSheet.Cells[row, 13].GetValue<string>();
                                outputWorksheet.Cells[row7, 53].Value = JobTitle;
                                var PTLocation = inputWorkSheet.Cells[row, ptlocation].GetValue<string>();
                                outputWorksheet.Cells[row7, 56].Value = PTLocation;
                                var Nationality = inputWorkSheet.Cells[row, nationality].GetValue<string>();
                                outputWorksheet.Cells[row7, 105].Value = Nationality;
                                var Pension = inputWorkSheet.Cells[row, pension].GetValue<string>();
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
                                //for (row2 = 2; row2 <= OrgAssignmentsDataSheetLastRow; row2++)
                                //{
                                //    var id = OrgAssignmentsDataSheet.Cells[row2, 1].GetValue<string>();
                                //    if (id.Equals(HRID))
                                //    {
                                //        //outputWorksheet.Cells[row7, 54].Value = OrgAssignmentsDataSheet.Cells[row2, 7].GetValue<string>(); ;

                                //    }
                                //}
                                //int n = Program.getSheetNumber(filePath, "Cost Center code");
                                //var inputcostcentersheet = package.Workbook.Worksheets[n];
                                //int inputcostcenterLastRow = inputcostcentersheet.Dimension.End.Row;
                                //var locationcode = "!!!";
                                //for (int i = 2; i <= inputcostcenterLastRow; i++)
                                //{
                                //    var id = inputcostcentersheet.Cells[i, 1].GetValue<string>();
                                //    if (HRID.Equals(id))
                                //    {
                                //        locationcode = inputcostcentersheet.Cells[i, 2].GetValue<string>();
                                //        outputWorksheet.Cells[row7, 54].Value = locationcode;
                                //    }
                                //}
                                using (var package4 = new ExcelPackage(new FileInfo(ascendcodes)))
                                {
                                    var AscendStatusCode = package4.Workbook.Worksheets["Status"];
                                    outputWorksheet.Cells[row7, 51].Value = AscendStatusCode.Cells[2, 1].GetValue<string>(); ;
                                    var AscendGradeCode = package4.Workbook.Worksheets["Grades"];
                                    outputWorksheet.Cells[row7, 52].Value = AscendGradeCode.Cells[2, 1].GetValue<string>(); ;
                                    var AscendBusinessAreaCode = package4.Workbook.Worksheets["Business Area"];
                                    outputWorksheet.Cells[row7, 55].Value = AscendBusinessAreaCode.Cells[2, 1].GetValue<string>(); ;
                                    var AscendCategoryCode = package4.Workbook.Worksheets["Categories"];
                                    outputWorksheet.Cells[row7, 50].Value = AscendCategoryCode.Cells[2, 1].GetValue<string>(); ;
                                    var AscendOccupationCode = package4.Workbook.Worksheets["Occupations"];
                                    outputWorksheet.Cells[row7, 58].Value = AscendOccupationCode.Cells[2, 1].GetValue<string>();
                                    //var AscendPayrollCode = package4.Workbook.Worksheets["Payroll Code"];
                                    //outputWorksheet.Cells[row7, 34].Value = AscendPayrollCode.Cells[2, 1].GetValue<string>();
                                    var AscendPFRegistrationCode = package4.Workbook.Worksheets["P.F. Registration Code"];
                                    outputWorksheet.Cells[row7, 61].Value = AscendPFRegistrationCode.Cells[2, 2].GetValue<string>();
                                    //var AscendLocation = package4.Workbook.Worksheets["Location"];
                                    //string location = inputWorkSheet.Cells[row, ptlocation].GetValue<string>();
                                    //location = Program.ShrinkString(location);
                                    //for (row2 = 2; row2 <= BenefeciarieslastRow; row2++)
                                    //{
                                    //    string loc = AscendLocation.Cells[row2, 2].GetValue<string>();
                                    //    loc = Program.ShrinkString(loc);
                                    //    if (loc.Equals(location))
                                    //    {
                                    //        outputWorksheet.Cells[row7, 56].Value = AscendLocation.Cells[row2, 1].GetValue<string>();
                                    //    }
                                    //}
                                    //outputWorksheet.Cells[row7, 61].Value = AscendPFRegistrationCode.Cells[2, 2].GetValue<string>();
                                    //var AscendCostCenterSheet = package4.Workbook.Worksheets["Cost Centre"];
                                    //int AscendCostCenterLastRow = AscendCostCenterSheet.Dimension.End.Row;
                                    //for (int i = 2; i <= AscendCostCenterLastRow; i++)
                                    //{
                                    //    if (Program.ShrinkString(locationcode).Equals(Program.ShrinkString(AscendCostCenterSheet.Cells[i, 2].GetValue<string>())))
                                    //    {
                                    //        outputWorksheet.Cells[row7, 54].Value = AscendCostCenterSheet.Cells[i, 1].GetValue<string>();
                                    //        break;
                                    //    }
                                    //}
                                }
                                outputWorksheet.Cells[row7, 101].Value = Pension;
                                for (row2 = 2; row2 <= BenefeciarieslastRow; row2++)
                                {
                                    var id = BenefeciariesDataSheet.Cells[row2, bfhrid].GetValue<string>();
                                    if (HRID.Equals(id))
                                    {
                                        Regex validCharsRegex = new Regex("[^a-zA-Z ]");
                                        outputWorksheet.Cells[row7, 95].Value = validCharsRegex.Replace(BenefeciariesDataSheet.Cells[row2, primarynameasperbank].GetValue<string>(), "");
                                        var ifsc = BenefeciariesDataSheet.Cells[row2, bfifsc].GetValue<string>();
                                        ifsc = ifsc.Replace(" ", "");
                                        if (ifsc.Length == 11) { outputWorksheet.Cells[row7, 29].Value = ifsc; }
                                        outputWorksheet.Cells[row7, 30].Value = BenefeciariesDataSheet.Cells[row2, bfacno].GetValue<string>();
                                        var bankname = BenefeciariesDataSheet.Cells[row2, bfbankname].GetValue<string>();
                                        bankname = Program.ShrinkString(bankname);
                                        bankname = bankname.Replace("ltd", "");
                                        bankname = bankname.Replace("limited", "");
                                        bankname = bankname.Replace("pvt", "");
                                        bankname = bankname.Replace(".", "");
                                        bool containsBank = bankname.Contains("bank", StringComparison.OrdinalIgnoreCase);
                                        if (!containsBank)
                                        {
                                            bankname = bankname + "bank";
                                        }
                                        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                                        {
                                            var Ascendsheet = package2.Workbook.Worksheets["Banks Detailed"];
                                            int AscendLastRow = Ascendsheet.Dimension.End.Row;
                                            for (int row5 = 2; row5 <= AscendLastRow; row5++)
                                            {
                                                var bank = Ascendsheet.Cells[row5, 2].GetValue<string>();
                                                bank = Program.ShrinkString(bank);
                                                if (bank.Equals(bankname))
                                                {
                                                    outputWorksheet.Cells[row7, 28].Value = Ascendsheet.Cells[row5, 1].GetValue<string>();
                                                }
                                            }
                                        }
                                    }
                                }
                                row7++;
                            }
                        }
                        string newFileName = Path.Combine(destinationFolder, "New Joinee_Master" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                        Console.WriteLine("New_joinee_Master Excel file created successfully!");
                        File.Delete(filePath);
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