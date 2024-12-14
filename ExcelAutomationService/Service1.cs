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
using OfficeOpenXml;
using Timer = System.Timers.Timer;

namespace ExcelAutomationService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        string sourceFolder = @"C:\Automation";     // Folder to watch for Excel files
        string destinationFolder = @"C:/Automation/output";
        string ascendcodes = "C:/Automation/Twilio_Twilio Technology/Automation_Ascent_Codes/Ascent Codes.xlsx";
        
        public Service1()
        {
            InitializeComponent();
        }
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
                    return col;
                }
            }
            catch (Exception e)
            {
                Log(e.Message);
                throw;
            }
        }
        public static int getSheetNumber(string filepath, string worksheetname)
        {
            try
            {
                worksheetname = ShrinkString(worksheetname);
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    int worksheetCount = package.Workbook.Worksheets.Count;
                    for (int i = worksheetCount - 1; i >= 0; i--)
                    {
                        string temp = package.Workbook.Worksheets[i].Name;
                        temp = ShrinkString(temp);
                        if (temp.Equals(worksheetname))
                        {
                            return i;
                        }
                    }
                    return -1;
                }
            }
            catch (Exception e)
            {
                Log(e.Message);
                throw;
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
            FileSystemWatcher watcher = new FileSystemWatcher(sourceFolder, "*.xlsx")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
            };
            watcher.Created += async (sender, e) => await ProcessFile(ascendcodes, e.FullPath, destinationFolder);
            watcher.EnableRaisingEvents = true;
            Log("Service started");
            Log("Watching for Excel files in " + sourceFolder);
            //Log("Press Enter to exit...");
            Console.ReadLine();
        }
        private static async Task ProcessFile(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                string foldername = Path.GetFileName(filePath);

                if (foldername.IndexOf('_') != -1)
                {
                    ascendcodes = destinationFolder + "/" + foldername.Substring(0, foldername.IndexOf('_')) + "/" + "Automation_Ascent_Codes/Ascent Codes.xlsx";
                    destinationFolder = destinationFolder + "/" + foldername.Substring(0, foldername.IndexOf('_')) + "/" + foldername.Substring(foldername.IndexOf('_') + 1);
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
                Console.WriteLine($"Processed file: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file {Path.GetFileName(filePath)}: {ex.Message}");
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
                string _logFilePath = @"C:\ServiceLogs\AutomationService.log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Log(ex.Message );
                // Fail silently if logging fails to avoid crashing the service
            }
        }
    }
}
