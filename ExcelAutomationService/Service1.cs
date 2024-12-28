using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Timer = System.Timers.Timer;

namespace ExcelAutomationService
{
    public partial class Service1 : ServiceBase
    {
        public static string errors = @"E:/PAYROLL_SERVER/Automation/Errors";
        public static string archived = @"E:/PAYROLL_SERVER/Automation/Archived";
        public static int ErrorCount=0;
        Timer timer = new Timer();
        string sourceFolder = @"E:\PAYROLL_SERVER\Automation\Input";     // Folder to watch for Excel files
        public static string destination = @"E:/PAYROLL_SERVER/Automation/output";
        string destinationFolder = @"E:/PAYROLL_SERVER/Automation/output";
        string ascendcodes = "E:/PAYROLL_SERVER/Automation/Twilio_Twilio Technology/Automation_Ascent_Codes/Ascent Codes.xlsx";
       
        public Service1()
        {
            InitializeComponent();
        }
        //Method to get position of column
        public static int getColumnNumber(string filepath, string worksheetname, string columnname)
        {
            try
            {
                columnname = columnname.ToLower();
                columnname = columnname.Replace(" ", "");
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets[worksheetname];
                    int col = 1;
                    int totalColumns = inputWorkSheet.Dimension.End.Column;
                    for (col = 1; col <= totalColumns; col++)
                    {
                        string temp = inputWorkSheet.Cells[1, col].Text.ToLower();
                        temp = temp.Replace(" ", "");
                        if (columnname.Equals(temp))
                        {
                            return col; // Return the column number if the header matches
                        }
                    }
                    col = -1;
                    if (col == -1)
                    {
                        Log(columnname + " column was not found in " + worksheetname + " of " + filepath + " file.");
                        ErrorCount++;
                    }
                    return col;
                }
            }
            catch (Exception e)
            {
                PathLog(columnname+" column was not found in"+worksheetname+" of "+filepath+" file.");
                throw;
            }
        }
        //Method to get position of Sheet
        public static int getSheetNumber(string filepath, string worksheetname)
        {
            try
            {
                worksheetname = ShrinkString(worksheetname);
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    int worksheetCount = package.Workbook.Worksheets.Count;
                    int i = 0;
                    for (i = worksheetCount - 1; i >= 0; i--)
                    {
                        string temp = package.Workbook.Worksheets[i].Name;
                        temp = ShrinkString(temp);
                        if (temp.Equals(worksheetname))
                        {
                            return i;
                        }
                    }
                    i = 0;
                    if (i==0) {
                        PathLog(worksheetname + " sheet was not found in " + filepath);
                        ErrorCount++;
                    }
                    return i;
                }
            }
            catch (Exception e)
            {
                PathLog(e.Message);
                throw;
            }
        }
        //Method to Validate Aadhaar
        public static string ValidateAadhar(string sheetname, string hrid, string adhaar)
        {
            adhaar = adhaar.Replace(" ", "");
            if (adhaar.Length == 0)
            {
                PathLog(hrid + "  aadhar number not given in " + sheetname + " sheet.");
                return "";
            }
            if ((adhaar.Length == 12) && (adhaar.All(char.IsDigit)&&(adhaar.Length != 0)))
                return adhaar;
            else
            {
                PathLog(hrid+"  adhaar number "+ adhaar + " is not valid in " + sheetname + " sheet.");
                return "";
            }
        }
        //Method to Validate PAN
        public static string ValidatePAN(string sheetname, string hrid, string pan)
        {
            pan = pan.Replace(" ", "");
            if ((pan.Length == 10) && (pan[3] == 'P'))
                return pan;
            if (pan.Length == 0)
            {
                PathLog( hrid + " PAN is empty in " + sheetname + " sheet.");
                return "PANNOTAVBLE";
            }
            else
            {
                PathLog(hrid + "  pan number "+pan+" is not valid in " + sheetname + " sheet.");
                return "";
            }
        }
        public static string ValidateDate(string date)
        {
            return date;
        }
        //Method to Validate IFSC
        public static string ValidateIFSC(string sheetname, string hrid, string ifsc)
        {
            ifsc = ifsc.Replace(" ", "");
            if (ifsc.Length == 0)
            {
                PathLog(hrid+" IFSC code is not given in " + sheetname + " sheet.");
                return "";
            }
            if (ifsc.Length == 11)
                return ifsc;
            else
            {
                PathLog(hrid+" IFSC "+ ifsc +" is not valid in " + sheetname + " sheet.");
                return "";
            }
        }
        public static string ShrinkString(string input)
        {
            if (input != null)
            {
                input = input.ToLower();
                input = input.Replace(" ", "");
                return input;
            }
            return "";
        }
        protected override void OnStart(string[] args)
        {
            timer.Interval = 1000;
            timer.Enabled = true;
            if (!Directory.Exists(sourceFolder) || !Directory.Exists(destinationFolder))
            {
                Console.WriteLine("Source or destination folder does not exist. Please check paths.");
                return;
            }
            Log("Service started");
            Log("Watching for Excel files in " + sourceFolder);
            FileSystemWatcher watcher = new FileSystemWatcher(sourceFolder, "*.xlsx")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
            };
            watcher.Created += async (sender, e) => await ProcessFile(ascendcodes, e.FullPath, destinationFolder);
            watcher.EnableRaisingEvents = true;
            
            //Log("Press Enter to exit...");
            Console.ReadLine();
        }
        public static async Task ProcessFile(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                DateTime now = DateTime.Now;

                // Format the month and year as "Month_Year"
                string formattedDate = $"{now:dd_MMMM_yyyy}";
                string foldername = Path.GetFileName(filePath);
                foldername = foldername.Replace(".xlsx", "");
                string filename = Path.GetFileName(filePath.ToLower());
                string[] directories = Directory.GetDirectories(destinationFolder);

                // Extract only the folder names
                //method to find right folder
                string[] folderNames = Array.ConvertAll(directories, dir => Path.GetFileName(dir.ToLower()));
                foreach (string folderName in folderNames)
                {
                    if (!folderName.Contains(' '))
                    {
                        if (filename.ToLower().Contains(folderName.ToLower()))
                        {
                            Console.WriteLine(folderName);
                            destinationFolder = destinationFolder + "/" + folderName;
                            string[] referencefile=Directory.GetFiles((destinationFolder), "*.xlsx");
                            ascendcodes = destinationFolder + "/" + Path.GetFileName(referencefile[0]);
                            destinationFolder = destinationFolder + "/" + folderName + " " + formattedDate;
                            destination = destinationFolder;
                            Console.WriteLine(ascendcodes);
                            break;
                        }
                    }
                    //in case of spaces in folder name
                    else
                    {
                        string[] parts = folderName.Split(' ');
                        int count = parts.Length;
                        int temp = 0;
                        foreach (string part in parts)
                        {
                            if (filename.ToLower().Contains(part.ToLower()))
                            {
                                temp++;
                            }
                        }
                        if (temp == count)
                        {
                            destinationFolder = destinationFolder + "/" + folderName;
                            string[] referencefile = Directory.GetFiles((destinationFolder), "*.xlsx");
                            ascendcodes = destinationFolder + "/" + Path.GetFileName(referencefile[0]);
                            destinationFolder = destinationFolder + "/" + folderName + " " + formattedDate;
                            destination = destinationFolder;
                            Console.WriteLine(ascendcodes);
                            break;
                        }
                    }
                }
                if (!Directory.Exists(foldername))
                {
                    Directory.CreateDirectory(destinationFolder);
                }
                // Ensure file is fully available by checking in a loop until it's accessible
                for (int retries = 0; retries < 5; retries++)
                {
                    try
                    {
                        using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                        {
                            stream.Close();
                            break; // If accessible, break the loop
                        }
                    }
                    catch (IOException)
                    {
                        await Task.Delay(500); // Wait and retry if file is still being written
                    }
                }
                // Call the relevant methods to process the file
                await Task.Run(() => Benefeciaries_Data.Beneficiaries_Data(ascendcodes, filePath, destinationFolder));
                await Task.Run(() => Leaver_Master.LeaverMaster(ascendcodes, filePath, destinationFolder));
                await Task.Run(() => Joiner_Leaver_Master.JoinerLeaverMaster(ascendcodes, filePath, destinationFolder));
                await Task.Run(() => Variable.Variable_Pay_Inputs_Data(ascendcodes, filePath, destinationFolder));
                //await Task.Run(() => New_Joiners_Ctc.CTC_Master(ascendcodes, filePath, destinationFolder));
                await Task.Run(() => Existing_Changes_Master.Existing_changes_Master(ascendcodes, filePath, destinationFolder));
                await Task.Run(() => New_Joinee_Master.NewJoinee_Master(ascendcodes, filePath, destinationFolder));
               
               //action after processing
                if (!Directory.Exists(archived))
                {
                    Directory.CreateDirectory(archived);
                }
                if (!Directory.Exists(errors))
                {
                    Directory.CreateDirectory(errors);
                }
                if (File.Exists(filePath))
                {
                    if (ErrorCount == 0)
                    {
                        if (File.Exists(archived+"/"+Path.GetFileName(filePath))) { 
                            File.Delete(filePath); 
                        }
                        else { 
                        File.Move(filePath, Path.Combine(archived, Path.GetFileName(filePath)));
                        }
                    }
                    else
                    {
                        if (File.Exists(errors + "/" + Path.GetFileName(filePath))) {
                            File.Delete(filePath); 
                        }
                        else
                        {
                            File.Move(filePath, Path.Combine(errors, Path.GetFileName(filePath)));
                        }
                        ErrorCount = 0;
                    }
                }
                Log($"Processed file: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
               Log($"Error processing file {Path.GetFileName(filePath)}: {ex.Message}");
               File.Move(filePath, Path.Combine(errors, Path.GetFileName(filePath)));
            }
        }
        protected override void OnStop()
        {
            Log("Service stopped.");
        }
        public static void Log(string message)
        {
            try
            {
                DateTime today = DateTime.Today;
                string _logFilePath = @"E:\PAYROLL_SERVER\Automation\ServiceLogs\" + today.ToString("dd/MMMM/yyyy")+"_PayrollAutomationService.log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                //Log(ex.Message);
                // Fail silently if logging fails to avoid crashing the service
            }
        }
        public static void PathLog(string message)
        {
            try
            {
                DateTime today = DateTime.Today;
                string _logFilePath = destination +"/"+ today.ToString("dd/MMMM/yyyy") + "_PayrollAutomationService.log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                //Log(ex.Message);
                // Fail silently if logging fails to avoid crashing the service
            }
        }
    }
}
