using ExcelHandlingDotnetpractice;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;

public class Program
{
    public static int getColumnNumber(string filepath, string worksheetname, string columnname)
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
    public static int getSheetNumber(string filepath, string worksheetname)
    {
        worksheetname = ShrinkString(worksheetname);
        using (var package = new ExcelPackage(new FileInfo(filepath)))
        {
            int worksheetCount = package.Workbook.Worksheets.Count;
            for (int i = worksheetCount - 1; i >= 0; i--)
            {
                string temp= package.Workbook.Worksheets[i].Name;
                temp=ShrinkString(temp);
                if (temp.Equals(worksheetname)) {
                    return i;
                }
            }
            return -1;
        }
    }
    public static string ShrinkString(string input)
    {
        input=input.ToLower();
        input=input.Replace(" ","");
        return input;
    }
    public static async Task Main(string[] args)
    {
        string sourceFolder = @"C:\Automation\Twilio_Twilio Technology";     // Folder to watch for Excel files
        string destinationFolder = @"C:\Automation_Results";
        if (!Directory.Exists(sourceFolder) || !Directory.Exists(destinationFolder))
        {
            Console.WriteLine("Source or destination folder does not exist. Please check paths.");
            return;
        }
        FileSystemWatcher watcher = new FileSystemWatcher(sourceFolder, "*.xlsx")
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
        };
        watcher.Created += async (sender, e) => await ProcessFile(e.FullPath, destinationFolder);
        watcher.EnableRaisingEvents = true;
        Console.WriteLine("Watching for Excel files in " + sourceFolder);
        Console.WriteLine("Press Enter to exit...");
        Console.ReadLine();
    }
    public static string isColoured(int i,int j,string filepath, string worksheetname) {
        return filepath; 
    }
    private static async Task ProcessFile(string filePath, string destinationFolder)
    {
        try
        {
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
            await Task.Run(() => Benefeciaries_Data.Beneficiaries_Data(filePath, destinationFolder));
            await Task.Run(() => Leaver_Master.LeaverMaster(filePath, destinationFolder));
            await Task.Run(() => Joiner_Leaver_Master.JoinerLeaverMaster(filePath, destinationFolder));
            await Task.Run(() => Variable.Variable_Pay_Inputs_Data(filePath, destinationFolder));
            await Task.Run(() => New_Joiners_Ctc.CTC_Master(filePath, destinationFolder));
            await Task.Run(() => Existing_Changes_Master.Existing_changes_Master(filePath, destinationFolder));
            await Task.Run(() => New_Joinee_Master.NewJoinee_Master(filePath, destinationFolder));

            Console.WriteLine($"Processed file: {Path.GetFileName(filePath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing file {Path.GetFileName(filePath)}: {ex.Message}");
        }
    }
}
