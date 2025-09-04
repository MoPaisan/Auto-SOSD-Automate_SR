using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace SOSD_AutomateSR
{
    class Program
    {
        static void Main(string[] args)
        {
            WebManager webManager = new WebManager();
            ExcelProcessor excelProcessor = new ExcelProcessor();

            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("Starting SOSD Automate SR");
            Console.WriteLine("--------------------------------------------------------------------------------------");

            string PathCurrentDir = Directory.GetCurrentDirectory();
            string PathProfileFF = PathCurrentDir + "\\Profiles\\SOSD_AutomateSR";
            
            // Read Config File
            string[] configStr = webManager.ReadTextFileAsArray(PathCurrentDir + "\\SOSD_AutomateSR.config");

            string UserMyOffice = webManager.GetConfig(configStr, "ResuMyoffice");
            string passMyOffice = webManager.GetConfig(configStr, "SsapWordMyoffice");
            string URLMyOffice = webManager.GetConfig(configStr, "URLMyofficeWeb");

            string UserIM = webManager.GetConfig(configStr, "ResuIM");
            string passIM = webManager.GetConfig(configStr, "SsapWordIM");
            string URLIM = webManager.GetConfig(configStr, "URLIM");
            string titleReport = webManager.GetConfig(configStr, "titleReport");
            string StatusCase = webManager.GetConfig(configStr, "StatusCase");
            string DayRange = webManager.GetConfig(configStr, "DayRange");

            string PathDownload = webManager.GetConfig(configStr, "PathDownload");
            string PathReportSOSD = webManager.GetConfig(configStr, "PathReportSOSD");
            string PathOutputReport = webManager.GetConfig(configStr, "PathOutputReport");
            string PathTemplateCloseCase = webManager.GetConfig(configStr, "PathTemplateCloseCase");
            string PathTemplateCheckMO = webManager.GetConfig(configStr, "PathTemplateCheckMO");
            string PathMappingFile = webManager.GetConfig(configStr, "PathMappingFile");
            string PathReportDaily = webManager.GetConfig(configStr, "PathReportDaily");
            string PathReportMonthly = webManager.GetConfig(configStr, "PathReportMonthly");

            // สร้าง Path โฟลเดอร์ปลายทางตามโครงสร้าง ปี\เดือน\วัน
            DateTime now = DateTime.Now;
            string yearFolder = now.ToString("yyyy");
            string monthFolder = now.ToString("MM");
            string dateFolder = now.ToString("dd");
            string timeFolder = now.ToString("HHmm");

            // รวม Path ปลายทางใหม่: baseTargetPath\Year\Month\Day\Time
            string TargetPathReport = Path.Combine(PathReportSOSD, yearFolder, monthFolder, dateFolder, timeFolder);
            if (!Directory.Exists(TargetPathReport))
            {
                Directory.CreateDirectory(TargetPathReport);
                Console.WriteLine($"Created target directory: {TargetPathReport}");
            }

            string TargetPathOutputReport = Path.Combine(PathOutputReport, yearFolder, monthFolder, dateFolder, timeFolder);
            if (!Directory.Exists(TargetPathOutputReport))
            {
                Directory.CreateDirectory(TargetPathOutputReport);
                Console.WriteLine($"Created target directory: {TargetPathOutputReport}");
            }

            //string TargetPathReport = @"D:\SOSD_AutomateSR\ReportSOSDPromotion\2025\09\03\1558";
            //string TargetPathOutputReport = @"D:\SOSD_AutomateSR\ReportCloseCase\2025\09\03\1558";

            // กำหนดชื่อไฟล์ Output ใหม่ตามรูปแบบ Import_CloseCaseIM_YYYMMDD_hhmmss.xlsx
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputImportFileName = $"Import_CloseCaseIM_{timestamp}.xlsx";
            string outputImportFilePath = Path.Combine(TargetPathOutputReport, outputImportFileName);

            // กำหนดชื่อไฟล์ Output สำหรับ Batch Template
            string outputBatchFileName = $"PSP_SR_{timestamp}.xlsx";
            string outputBatchFilePath = Path.Combine(TargetPathOutputReport, outputBatchFileName);

            ExcelProcessor processor = new ExcelProcessor();
            IWebDriver driverIM = null;
            IWebDriver driverMO = null;


            try
            {
                //Download Report SOSD Promotion
                //เริ่มต้น WebDriver สำหรับ IM
                driverIM = webManager.WebIM(URLIM, UserIM, passIM, PathProfileFF);
                webManager.SelectMyReportIM(driverIM, titleReport);
                webManager.ExportPendingReview(driverIM, PathDownload, TargetPathReport, StatusCase, DayRange, titleReport);

                //รวมไฟล์รายงาน
                Console.WriteLine("--------------------------------------------------------------------------------------");
                string outputCombinedFileName = "CombinedSOSDPromotionReport_All.xlsx";
                string masterExcelFilePath = Path.Combine(TargetPathReport, outputCombinedFileName);

                string combinedFilePath = processor.CombineExcelReports(TargetPathReport, titleReport, outputCombinedFileName, "สรุปข้อมูลโปรโมชั่น", true);

                if (combinedFilePath != null)
                {
                    Console.WriteLine($"\nรายงานรวมถูกบันทึกที่: {combinedFilePath}");
                }
                else
                {
                    Console.WriteLine("\nไม่สามารถสร้างรายงานรวมได้");
                    return;
                }

                // กำหนดโฟลเดอร์ดาวน์โหลดเริ่มต้นสำหรับเบราว์เซอร์
                string browserDownloadFolder = Path.Combine(Path.GetTempPath(), "SeleniumBrowserDownloads");
                if (!Directory.Exists(browserDownloadFolder))
                {
                    Directory.CreateDirectory(browserDownloadFolder);
                }

                ChromeOptions options = new ChromeOptions();
                options.AddUserProfilePreference("download.default_directory", browserDownloadFolder);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("plugins.always_open_pdf_externally", true);

                // ประมวลผลข้อมูล Excel และสร้างไฟล์สำหรับนำเข้า Close Case และ Batch Order
                Console.WriteLine("--------------------------------------------------------------------------------------");
                Console.WriteLine("Starting Excel data processing and web interaction...");
                string generatedFilePath = processor.ProcessAndExportFilteredData(
                    masterExcelFilePath,
                    PathTemplateCloseCase,
                    TargetPathOutputReport,
                    outputImportFilePath,
                    outputBatchFilePath,
                    driverIM,
                    PathMappingFile,
                    PathTemplateCheckMO
                );
                webManager.LogoutIM(driverIM);
                driverIM.Quit();
                Thread.Sleep(1000);


                //หาชื่อไฟล์ที่มีคำว่า "PSP_SR_"
                string searchDirectory = TargetPathOutputReport;

                // หาไฟล์ PSP_SR_ ที่ไม่ลงท้ายด้วย _err หรือ _ok
                string[] allPSPFiles = Directory.GetFiles(searchDirectory, "PSP_SR_*.xlsx");
                string foundFileNameSR = null;

                foreach (string file in allPSPFiles)
                {
                    string fileName = Path.GetFileName(file);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);

                    // ตรวจสอบว่าไฟล์ไม่ลงท้ายด้วย _err หรือ _ok
                    if (!fileNameWithoutExt.EndsWith("_err") && !fileNameWithoutExt.EndsWith("_ok"))
                    {
                        foundFileNameSR = fileName;
                        break;
                    }
                }

                string fileNameSR = foundFileNameSR != null ? Path.GetFileNameWithoutExtension(foundFileNameSR) : null;
                //Console.WriteLine($"Found Summary Report File: {foundFileNameSR}");
                //Console.WriteLine($"Found Summary Report File: {fileNameSR}");
                string foundFileNameCloseCase = webManager.FindFileNameWithPrefix(searchDirectory, "Import_CloseCaseIM_");
                string fileNameClsoeCase = Path.GetFileNameWithoutExtension(foundFileNameCloseCase);

                if (!string.IsNullOrEmpty(foundFileNameSR))
                {
                    string PathMarketing = Path.Combine(TargetPathOutputReport, foundFileNameSR);
                    string PathCloseCase = Path.Combine(TargetPathOutputReport, foundFileNameCloseCase);

                    // สร้าง path สำหรับไฟล์ต้นทางที่มี suffix _ok และ _err
                    string sourceFileOkPath = Path.Combine(TargetPathOutputReport, fileNameSR + "_ok.xlsx");
                    string sourceFileErrPath = Path.Combine(TargetPathOutputReport, fileNameSR + "_err.xlsx");

                    //อัพโหลดไฟล์ที่ Batch Order
                    Console.WriteLine("--------------------------------------------------------------------------------------");
                    driverMO = webManager.WebMyOffice(URLMyOffice, UserMyOffice, passMyOffice, PathProfileFF);
                    webManager.SelsectMenuOrderMyofiice(driverMO);
                    webManager.SelectBatchTopicMyofiice(driverMO, PathMarketing);

                    //Download ไฟล์จาก batch mornitoring
                    webManager.SelectMonitoringAndProcessFile(driverMO, TargetPathOutputReport, fileNameSR);
                    webManager.LogoutMyofiice(driverMO);
                    driverMO.Quit();

                    bool fileOk = processor.CreateImportCloseCaseIM(
                       sourceFileOkPath,                               // ไฟล์ต้นทาง (มี Case ID ในคอลัมน์ AI, AJ, AK)
                       PathCloseCase,                           // ไฟล์ปลายทาง (template)
                       "Resolved-Completed",                            // Case Status
                       "ดำเนินการเรียบร้อย รบกวนตรวจสอบความถูกต้องอีกครั้งค่ะ",       // Comment
                       null,                                           // Source sheet (ใช้ sheet แรก)
                       null,                                           // Target sheet (ใช้ sheet แรก)  
                       2                                               // เริ่มจากแถวที่ 2 (ข้าม header)
                    );

                    if (fileOk)
                    {
                        Console.WriteLine("Import_CloseCaseIM Complete");
                    }

                    bool fileErr = processor.CreateImportCloseCaseIM(
                       sourceFileErrPath,                              // ไฟล์ต้นทาง (มี Case ID ในคอลัมน์ AI, AJ, AK)
                       PathCloseCase,                           // ไฟล์ปลายทาง (template)
                       "Resolved-Rejected",                            // Case Status
                       "ไม่สามารถดำเนินการได้ ยอดเงินไม่เพียงพอ",                 // Comment
                       null,                                           // Source sheet (ใช้ sheet แรก)
                       null,                                           // Target sheet (ใช้ sheet แรก)  
                       2                                               // เริ่มจากแถวที่ 2 (ข้าม header)
                    );

                    if (fileErr)
                    {
                        Console.WriteLine("Import_CloseCaseIM Reject");
                    }

                    //Import Close Case ใน IM
                    driverIM = webManager.WebIM(URLIM, UserIM, passIM, PathProfileFF);
                    webManager.SelectManagerToolsIM(driverIM);
                    webManager.ImportCloseCase(driverIM, PathCloseCase);
                    webManager.LogoutIM(driverIM);

                    //สร้างรายงานสรุป
                    Console.WriteLine("--------------------------------------------------------------------------------------");
                    Console.WriteLine("Process summary report");

                    string monthlyReportPath = Path.Combine(PathReportMonthly, yearFolder, monthFolder);
                    if (!Directory.Exists(monthlyReportPath))
                    {
                        Directory.CreateDirectory(monthlyReportPath);
                        Console.WriteLine($"Created target directory: {monthlyReportPath}");
                    }

                    string dailyReportPath = Path.Combine(PathReportDaily, yearFolder, monthFolder, dateFolder);
                    if (!Directory.Exists(dailyReportPath))
                    {
                        Directory.CreateDirectory(dailyReportPath);
                        Console.WriteLine($"Created target directory: {dailyReportPath}");
                    }

                    string naOfferingCodeReportPath = Path.Combine(TargetPathOutputReport, "ManualCheck_CaseOfferingCode.xlsx");
                    string summaryReportFilePath = processor.SummarizeProcessedData(
                        masterExcelFilePath,
                        PathCloseCase,
                        PathMarketing,
                        dailyReportPath,
                        naOfferingCodeReportPath
                    );

                    if (summaryReportFilePath != null)
                    {
                        Console.WriteLine($"Summary report saved to: {summaryReportFilePath}");
                    }
                    else
                    {
                        Console.WriteLine("Failed to generate summary report.");
                    }

                    // สรุปเดือนปัจจุบัน
                    string monthlyReport = excelProcessor.CreateMonthlySummary(
                        dailyReportPath,
                        monthlyReportPath
                    );

                    if (monthlyReport != null)
                    {
                        Console.WriteLine($"Summary report saved to: {monthlyReport}");
                    }
                    else
                    {
                        Console.WriteLine("Failed to generate summary report.");
                    }
                }
                else
                {
                    Console.WriteLine("ไม่พบไฟล์ Summary report");
                }
            }
            catch
            {
                // ปิด WebDriver เสมอเมื่อเสร็จสิ้นการทำงาน
                if (driverIM != null)
                {
                    try
                    {
                        driverIM.Quit();
                        driverIM.Dispose();
                        Console.WriteLine("IM WebDriver closed and disposed.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error closing IM WebDriver: {ex.Message}");
                    }
                }

                if (driverMO != null)
                {
                    try
                    {
                        driverMO.Quit();
                        driverMO.Dispose();
                        Console.WriteLine("MyOffice WebDriver closed and disposed.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error closing MyOffice WebDriver: {ex.Message}");
                    }
                }
            }

            Console.WriteLine("\nProcess completed. Press any key to exit...");
            Console.ReadKey();
        }
    }
}
