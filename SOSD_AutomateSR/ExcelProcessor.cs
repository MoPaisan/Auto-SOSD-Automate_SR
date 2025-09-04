using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Data; // สำหรับ DataTable
using System.Drawing;
using System.IO.Packaging;

namespace SOSD_AutomateSR
{
    internal class ExcelProcessor
    {
        WebManager webManager = new WebManager();
        // คลาสภายในสำหรับเก็บข้อมูลแต่ละแถวจาก Excel Master File
        private class MasterExcelRow
        {
            public string CaseId { get; set; }
            public string Topic { get; set; }
            public string Reason { get; set; }
            public string Doc { get; set; }
            public string CreatedBy { get; set; }
            public string Mobile { get; set; }
            public string ResolutionComment { get; set; }
            public int RowIndex { get; set; } // สำหรับ Debugging หรืออ้างอิงแถวต้นฉบับ
        }

        // คลาสภายในสำหรับเก็บข้อมูลที่จะเขียนลงไฟล์ Batch
        private class BatchEntry
        {
            public string OfferingCode { get; set; }
            public string UserName { get; set; }
            public string MobileNo { get; set; }
            public string ServiceRequest1 { get; set; }
            public string Category1 { get; set; }
            public string SubCategory1 { get; set; }
        }


        private readonly int defaultShortWait = 1000;
        private readonly int defaultLongWait = 3000;
        private readonly int initialLoadWait = 6000;


        /// <summary>
        /// รวมข้อมูลจากไฟล์ Excel หลายไฟล์ที่มีชื่อตามรูปแบบที่กำหนด
        /// เข้าไปใน Worksheet เดียวกันของไฟล์ Excel ใหม่
        /// </summary>
        /// <param name="directoryPath">พาธของโฟลเดอร์ที่มีไฟล์ Excel ที่ต้องการรวม</param>
        /// <param name="titleReport">ชื่อไฟล์หลักที่ต้องการรวม (เช่น "ReportSOSDPromotion_ServiceRequest")</param>
        /// <param name="outputFileName">ชื่อไฟล์ Excel ที่จะใช้บันทึกข้อมูลที่รวมแล้ว (ค่าเริ่มต้นคือ "CombinedReport.xlsx")</param>
        /// <param name="mainSheetName">ชื่อของ Worksheet หลักที่จะใช้รวมข้อมูล (ค่าเริ่มต้นคือ "Combined Data")</param>
        /// <param name="hasHeader">ระบุว่าไฟล์ Excel ต้นฉบับมีแถว Header หรือไม่ (ค่าเริ่มต้นคือ true)</param>
        /// <returns>พาธของไฟล์ Excel ที่รวมข้อมูลแล้ว หรือ null หากไม่พบไฟล์ที่จะรวม</returns>
        public string CombineExcelReports(string directoryPath, string titleReport, string outputFileName = "CombinedReport.xlsx", string mainSheetName = "Combined Data", bool hasHeader = true)
        {
            // กำหนด LicenseContext สำหรับ EPPlus (สำคัญมาก)
            // ถ้าคุณใช้ Commercial License ให้ใช้ LicenseContext.Commercial
            // ถ้าคุณใช้ Non-Commercial License (สำหรับ Free tier) ให้ใช้ LicenseContext.NonCommercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // รูปแบบของชื่อไฟล์ที่คุณต้องการรวม (ใช้ Regex เพื่อจับคู่ชื่อไฟล์ที่ไม่สนวันที่)
            // ตัวอย่างชื่อไฟล์: ReportSOSDPromotion_ServiceRequest_2025-07-08_15-43-29.xlsx
            // Regex: ^{titleReport}_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.xlsx$
            // อธิบาย Regex:
            // ^                 - เริ่มต้นของสตริง
            // {titleReport}_    - ชื่อไฟล์ส่วนที่ส่งมาจาก parameter + underscore
            // \d{4}             - ตัวเลข 4 หลัก (สำหรับปี)
            // -                 - เครื่องหมายขีด
            // \d{2}             - ตัวเลข 2 หลัก (สำหรับเดือน, วัน, ชั่วโมง, นาที, วินาที)
            // \.xlsx$           - นามสกุลไฟล์ .xlsx ที่ต้องอยู่ท้ายสุดของสตริง
            string fileNamePattern = $@"^{Regex.Escape(titleReport)}_\d{{4}}-\d{{2}}-\d{{2}}_\d{{2}}-\d{{2}}-\d{{2}}\.xlsx$";
            Regex regex = new Regex(fileNamePattern, RegexOptions.IgnoreCase); // ไม่สนใจตัวพิมพ์เล็กพิมพ์ใหญ่

            try
            {
                Console.WriteLine($"Checking directory existence for: '{directoryPath}'");
                if (!Directory.Exists(directoryPath))
                {
                    Console.WriteLine($"Error: Directory not found at {directoryPath}");
                    return null;
                }

                // ค้นหาไฟล์ Excel ทั้งหมดในโฟลเดอร์ที่ตรงตามรูปแบบชื่อ
                var excelFilesToCombine = Directory.GetFiles(directoryPath, "*.xlsx")
                                                   .Where(file => regex.IsMatch(Path.GetFileName(file)))
                                                   .OrderBy(File.GetCreationTime) // เรียงตามเวลาสร้างเพื่อความสอดคล้อง
                                                   .ToList();

                if (!excelFilesToCombine.Any())
                {
                    Console.WriteLine($"No Excel files matching pattern '{titleReport}_YYYY-MM-DD_HH-mm-ss.xlsx' found in {directoryPath}.");
                    return null;
                }

                Console.WriteLine($"Found {excelFilesToCombine.Count} Excel files to combine.");

                // กำหนดพาธสำหรับไฟล์ Output
                string outputPath = Path.Combine(directoryPath, outputFileName);

                // ถ้าไฟล์ Output มีอยู่แล้ว ให้ลบทิ้งก่อน
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                    Console.WriteLine($"Deleted existing output file: {outputPath}");
                }

                // สร้าง ExcelPackage ใหม่สำหรับไฟล์ Combined
                using (var combinedPackage = new ExcelPackage())
                {
                    // สร้าง Worksheet หลักที่จะใช้รวมข้อมูลทั้งหมด
                    var combinedWorksheet = combinedPackage.Workbook.Worksheets.Add(mainSheetName);
                    int currentRow = 1; // แถวปัจจุบันที่จะเขียนข้อมูลใน Worksheet รวม

                    // สำหรับแต่ละไฟล์ Excel ที่พบ
                    foreach (var filePath in excelFilesToCombine)
                    {
                        Console.WriteLine($"Processing file: {Path.GetFileName(filePath)}");
                        using (var sourcePackage = new ExcelPackage(new FileInfo(filePath)))
                        {
                            var sourceWorksheet = sourcePackage.Workbook.Worksheets.FirstOrDefault();
                            if (sourceWorksheet == null || sourceWorksheet.Dimension == null)
                            {
                                Console.WriteLine($"Warning: No data or worksheet found in {Path.GetFileName(filePath)}. Skipping.");
                                continue;
                            }

                            // กำหนดช่วงของข้อมูลที่จะคัดลอก
                            int startRow = sourceWorksheet.Dimension.Start.Row;
                            int endRow = sourceWorksheet.Dimension.End.Row;
                            int startCol = sourceWorksheet.Dimension.Start.Column;
                            int endCol = sourceWorksheet.Dimension.End.Column;

                            // ถ้าเป็นไฟล์แรก และมี Header ให้คัดลอก Header ด้วย
                            if (combinedPackage.Workbook.Worksheets.Count == 1 && currentRow == 1 && hasHeader)
                            {
                                // คัดลอก Header จากไฟล์แรก
                                for (int col = startCol; col <= endCol; col++)
                                {
                                    combinedWorksheet.Cells[currentRow, col].Value = sourceWorksheet.Cells[startRow, col].Value;
                                }
                                currentRow++; // เลื่อนไปแถวถัดไปสำหรับข้อมูล
                                startRow++; // เริ่มคัดลอกข้อมูลจากแถวถัดจาก Header ในไฟล์ต้นฉบับ
                            }
                            else if (hasHeader)
                            {
                                // ถ้าไม่ใช่ไฟล์แรก และมี Header ให้ข้ามแถว Header ของไฟล์ต้นฉบับ
                                startRow++;
                            }

                            // คัดลอกข้อมูล (ไม่รวม Header หาก hasHeader เป็น true และไม่ใช่ไฟล์แรก)
                            for (int row = startRow; row <= endRow; row++)
                            {
                                for (int col = startCol; col <= endCol; col++)
                                {
                                    combinedWorksheet.Cells[currentRow, col].Value = sourceWorksheet.Cells[row, col].Value;
                                }
                                currentRow++;
                            }
                            Console.WriteLine($"Copied data from '{Path.GetFileName(filePath)}'. Total rows in combined sheet: {currentRow - 1}");
                        }
                    }

                    // ปรับขนาดคอลัมน์อัตโนมัติ (Optional: เพื่อให้อ่านง่ายขึ้น)
                    combinedWorksheet.Cells[combinedWorksheet.Dimension.Address].AutoFitColumns();

                    // บันทึกไฟล์ Combined
                    combinedPackage.SaveAs(new FileInfo(outputPath));
                    Console.WriteLine($"Successfully combined files to: {outputPath}");
                    return outputPath;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while combining Excel files: {ex.Message}");
                // Optional: Console.WriteLine(ex.StackTrace); // สำหรับ Debugging เพิ่มเติม
                return null;

            }

        }

        /// <summary>
        /// Filter Excel To Temp File
        /// </summary>
        /// <param name="masterFilePath">พาธเต็มของไฟล์ Excel ข้อมูลหลัก</param>
        /// <param name="templateFilePath">พาธเต็มของไฟล์ Excel Template (Import_CloseCaseIM)</param>
        /// <param name="outputDirectory">โฟลเดอร์สำหรับบันทึกไฟล์ Excel ที่สร้างขึ้นใหม่</param>
        /// <param name="outputImportFilePath"> กำหนดชื่อไฟล์ close case </param>
        /// <param name="outputBatchFilePath"> กำหนดชื่อไฟล์ PSP_SR </param>
        /// <param name="driver">IWebDriver instance สำหรับการโต้ตอบกับเว็บ (จำเป็นสำหรับเงื่อนไขใหม่)</param>
        /// <param name="mappingFilePath">พาธเต็มของไฟล์ Excel Mapping (สำหรับ Feature Code -> Offering Code)</param>
        /// <param name="templateBatchFilePath">พาธเต็มของไฟล์ Excel Template สำหรับ Batch</param>
        /// <returns>พาธเต็มของไฟล์ Excel ที่สร้างขึ้นใหม่ หากสำเร็จ; Null หากเกิดข้อผิดพลาด</returns>

        public string ProcessAndExportFilteredData(string masterFilePath, string templateFilePath, string outputDirectory, string outputImportFilePath, string outputBatchFilePath, IWebDriver driver, string mappingFilePath, string templateBatchFilePath)
        {
            // ตั้งค่า LicenseContext ที่นี่อีกครั้ง เพื่อให้แน่ใจว่าถูกตั้งค่าเมื่อเมธอดถูกเรียก
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // รายการของเหตุผล (Reason) ที่ต้องการกรอง
            List<string> reasonsToFilter = new List<string>
        {
            "กลุ่มที่ 2 กรณีได้รับต่อโปรโมชั่นเดิม (A to A)",
            "กลุ่มที่ 3 กรณีได้รับต่อโปรโมชั่นเดิม (A to B)",
            "ส่งขอโปรโมชั่นกลุ่ม A To A",
            "ส่งขอโปรโมชั่นกลุ่ม A To B"
        };

            // ชื่อคอลัมน์ที่จะอ่านจากไฟล์ Master
            const string masterCaseIdColumnName = "Case ID";
            const string masterTopicColumnName = "Topic";
            const string masterReasonColumnName = "Reason";
            const string masterDocColumnName = "Doc#";
            const string masterCreatedByColumnName = "Created By";
            const string masterMobileColumnName = "Mobile";
            const string masterCategoryColumnName = "Category";
            const string masterResolutionCommentColumnName = "Resolution Comment";

            // ชื่อคอลัมน์ที่จะเขียนลงไฟล์ Import_CloseCaseIM Template
            const string templateCaseIdColumnName = "Case ID";
            const string templateTopicColumnName = "Topic";
            const string templateCaseStatusColumnName = "Case Status";
            const string templateCommentColumnName = "Comment";

            // ชื่อคอลัมน์ที่จะอ่านจากไฟล์ Mapping
            const string mappingFeatureCodeColumnName = "Feature Code";
            const string mappingOfferingCodeColumnName = "Offering Code";

            // ชื่อคอลัมน์ที่จะเขียนลงไฟล์ Batch Template
            const string batchOfferingCodeColName = "OfferingCode1";
            const string batchUserNameColName = "UserName";
            const string batchMobileNoColName = "mobileNo";
            const string batchActionStatus1Name = "ActionStatus1";
            const string batchBypassProductRule1Name = "BypassProductRule1";
            const string batchBypassSMSCompleteFlag1Name = "bypassSMSCompleteFlag1";
            const string batchServiceRequestColName = "ServiceRequest1";
            const string batchCategoryColName = "Category1";
            const string batchSubCategoryColName = "SubCategory1";

            // กำหนดค่าคงที่สำหรับการกรองข้อมูล
            const string fixedActionStatus1Value = "Add";
            const string fixedBypassProductRule1Value = "Y";
            const string fixedBypassSMSCompleteFlag1Value = "N";

            // ค่าคงที่ที่จะใส่ในไฟล์ Import_CloseCaseIM Template
            const string caseStatusRejected = "Resolved-Rejected";
            const string commentForRejected = "ไม่สามารถดำเนินการได้ รบกวนขออนุมัติจาก MKT Owner เพิ่มเติมค่ะ";

            const string caseStatusForCompleted = "Resolved-Completed";
            const string commentForCompleted = "ดำเนินการเรียบร้อย รบกวนตรวจสอบความถูกต้องอีกครั้งค่ะ";

        // List สำหรับเก็บข้อมูลที่จะเขียนลงไฟล์ Batch
        List<BatchEntry> batchEntries = new List<BatchEntry>();
        
        // ✅ เพิ่ม List สำหรับเก็บเคสที่ OfferingCode = N/A (ให้ user เช็คเอง)
        List<MasterExcelRow> naOfferingCases = new List<MasterExcelRow>();            try
            {
                // ตรวจสอบพาธ
                if (!File.Exists(masterFilePath))
                {
                    Console.WriteLine($"Error: Master Excel file not found at '{masterFilePath}'.");
                    return null;
                }
                if (!File.Exists(templateFilePath))
                {
                    Console.WriteLine($"Error: Import_CloseCaseIM Template file not found at '{templateFilePath}'.");
                    return null;
                }
                if (!File.Exists(mappingFilePath))
                {
                    Console.WriteLine($"Error: Mapping file not found at '{mappingFilePath}'.");
                    return null;
                }
                if (!File.Exists(templateBatchFilePath))
                {
                    Console.WriteLine($"Error: Batch Template file not found at '{templateBatchFilePath}'.");
                    return null;
                }
                if (!Directory.Exists(outputDirectory))
                {
                    Console.WriteLine($"Output directory '{outputDirectory}' does not exist. Creating it...");
                    Directory.CreateDirectory(outputDirectory);
                }

                // โหลดข้อมูล Mapping (Feature Code -> Offering Code)
                Dictionary<string, string> mappingData = LoadMappingData(mappingFilePath, mappingFeatureCodeColumnName, mappingOfferingCodeColumnName);
                if (mappingData == null)
                {
                    Console.WriteLine("Failed to load mapping data. Aborting process.");
                    return null;
                }
                Console.WriteLine($"Loaded {mappingData.Count} entries from mapping file.");


                // ✅ คัดลอกไฟล์ Template ไปยังตำแหน่ง Output เพื่อเป็นไฟล์ใหม่ที่เราจะเขียนข้อมูลลงไป
                File.Copy(templateFilePath, outputImportFilePath, true); // true = overwrite if exists
                Console.WriteLine($"Copied Import_CloseCaseIM template file to '{outputImportFilePath}'.");

                // ✅ โหลดไฟล์ Excel ข้อมูลหลักและอ่านข้อมูลทั้งหมดเข้าสู่หน่วยความจำ
                List<MasterExcelRow> allMasterRows = new List<MasterExcelRow>();
                using (var masterPackage = new ExcelPackage(new FileInfo(masterFilePath)))
                {
                    var masterWorksheet = masterPackage.Workbook.Worksheets.FirstOrDefault();
                    if (masterWorksheet == null || masterWorksheet.Dimension == null)
                    {
                        Console.WriteLine($"Error: No data or worksheet found in master file '{masterFilePath}'.");
                        return null;
                    }

                    // ค้นหา Column Index ของคอลัมน์ที่ต้องการใน Master File
                    int masterCaseIdCol = -1;
                    int masterTopicCol = -1;
                    int masterReasonCol = -1;
                    int masterDocCol = -1;
                    int masterCreatedByCol = -1;
                    int masterMobileCol = -1;
                    int masterCategoryCol = -1;
                    int masterResolutionCommentCol = -1;

                    for (int col = 1; col <= masterWorksheet.Dimension.End.Column; col++)
                    {
                        string header = masterWorksheet.Cells[1, col].Text.Trim(); // อ่าน Header แถวแรก
                        if (header.Equals(masterCaseIdColumnName, StringComparison.OrdinalIgnoreCase))
                            masterCaseIdCol = col;
                        else if (header.Equals(masterTopicColumnName, StringComparison.OrdinalIgnoreCase))
                            masterTopicCol = col;
                        else if (header.Equals(masterReasonColumnName, StringComparison.OrdinalIgnoreCase))
                            masterReasonCol = col;
                        else if (header.Equals(masterDocColumnName, StringComparison.OrdinalIgnoreCase))
                            masterDocCol = col;
                        else if (header.Equals(masterCreatedByColumnName, StringComparison.OrdinalIgnoreCase))
                            masterCreatedByCol = col;
                        else if (header.Equals(masterMobileColumnName, StringComparison.OrdinalIgnoreCase))
                            masterMobileCol = col;
                        else if (header.Equals(masterCategoryColumnName, StringComparison.OrdinalIgnoreCase))
                            masterCategoryCol = col;
                        else if (header.Equals(masterResolutionCommentColumnName, StringComparison.OrdinalIgnoreCase))
                            masterResolutionCommentCol = col;
                    }

                    if (masterCaseIdCol == -1 || masterTopicCol == -1 || masterReasonCol == -1 || masterDocCol == -1 || masterCreatedByCol == -1 || masterMobileCol == -1 || masterCategoryCol == -1)
                    {
                        Console.WriteLine("Error: Missing one or more required columns (Case ID, Topic, Reason, Doc#, Created By, Mobile, Category) in master Excel file.");
                        return null;
                    }

                    // อ่านข้อมูลทั้งหมดจาก Master File เข้า List
                    // *** เพิ่มการตรวจสอบเพื่อข้ามแถวที่ Case ID ว่างเปล่า ***
                    Console.WriteLine($"\n--- DEBUG: Reading master file data ---");
                    int totalRowsRead = 0;
                    int skippedEmptyCaseId = 0;
                    int skippedCategoryMismatch = 0;
                    int validRowsAdded = 0;
                    
                    for (int row = masterWorksheet.Dimension.Start.Row + 1; row <= masterWorksheet.Dimension.End.Row; row++)
                    {
                        totalRowsRead++;
                        string caseId = masterWorksheet.Cells[row, masterCaseIdCol].Text.Trim();
                        string category = masterWorksheet.Cells[row, masterCategoryCol].Text.Trim();
                        string reason = masterWorksheet.Cells[row, masterReasonCol].Text.Trim();

                        // DEBUG: แสดงข้อมูลทุกแถวที่อ่าน
                        Console.WriteLine($"DEBUG Row {row}: CaseID='{caseId}', Category='{category}', Reason='{reason}'");

                        // ถ้า Case ID ว่างเปล่าหรือเป็นช่องว่าง ให้ข้ามแถวนั้นไป
                        if (string.IsNullOrWhiteSpace(caseId))
                        {
                            skippedEmptyCaseId++;
                            Console.WriteLine($"❌ SKIP Row {row}: Case ID is empty or whitespace.");
                            continue; // ข้ามไปแถวถัดไป
                        }

                        // *** เพิ่มการตรวจสอบ Category - เอาเฉพาะ "Promotion Prepaid Main On top Data ปัญหา" ***
                        bool categoryMatches = category.Equals("Promotion Prepaid Main On top Data ปัญหา", StringComparison.OrdinalIgnoreCase);
                        Console.WriteLine($"DEBUG Row {row}: Category Match = {categoryMatches} ('{category}' vs 'Promotion Prepaid Main On top Data ปัญหา')");
                        
                        if (!categoryMatches)
                        {
                            skippedCategoryMismatch++;
                            Console.WriteLine($"❌ SKIP Row {row}: Category '{category}' does not match required category.");
                            continue; // ข้ามไปแถวถัดไป
                        }

                        // อ่าน Resolution Comment (อาจจะ null ถ้าไม่มีคอลัมน์นี้)
                        string resolutionComment = "";
                        if (masterResolutionCommentCol != -1)
                        {
                            resolutionComment = masterWorksheet.Cells[row, masterResolutionCommentCol].Text.Trim();
                        }

                        // DEBUG: ตรวจสอบ Resolution Comment ที่มีคำ PHX หรือ Connection Timeout
                        if (!string.IsNullOrEmpty(resolutionComment) && 
                            (resolutionComment.IndexOf("PHX", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             resolutionComment.IndexOf("Connection Timeout", StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            Console.WriteLine($"🔍 DEBUG: Row {row}, Case ID '{caseId}' has interesting Resolution Comment: '{resolutionComment}'");
                        }

                        allMasterRows.Add(new MasterExcelRow
                        {
                            CaseId = caseId,
                            Topic = masterWorksheet.Cells[row, masterTopicCol].Text,
                            Reason = reason,
                            Doc = masterWorksheet.Cells[row, masterDocCol].Text.Trim(),
                            CreatedBy = masterWorksheet.Cells[row, masterCreatedByCol].Text.Trim(),
                            Mobile = masterWorksheet.Cells[row, masterMobileCol].Text.Trim(),
                            ResolutionComment = resolutionComment,
                            RowIndex = row
                        });
                        
                        validRowsAdded++;
                        Console.WriteLine($"✅ ADDED Row {row}: Case ID '{caseId}' successfully added to processing list.");
                        Console.WriteLine("--------------------------------------------------------------------------------------");
                    }
                    
                    Console.WriteLine($"\n📊 SUMMARY - Master File Reading:");
                    Console.WriteLine($"   Total rows read: {totalRowsRead}");
                    Console.WriteLine($"   Skipped (empty Case ID): {skippedEmptyCaseId}");
                    Console.WriteLine($"   Skipped (category mismatch): {skippedCategoryMismatch}");
                    Console.WriteLine($"   Valid rows added: {validRowsAdded}");
                    Console.WriteLine($"   Final count in allMasterRows: {allMasterRows.Count}");
                }

                // ✅ แบ่งข้อมูลออกเป็นสามกลุ่มตามลำดับความสำคัญ
                // 1. กลุ่มที่มี PHX 20002 หรือ Connection Timeout ใน Resolution Comment (สำคัญที่สุด)
                Console.WriteLine("--------------------------------------------------------------------------------------");
                Console.WriteLine("\n--- DEBUG: Checking all rows for PHX/Connection Timeout patterns ---");
                var phxCompletedRows = new List<MasterExcelRow>();
                
                foreach (var row in allMasterRows)
                {
                    bool hasPHX = !string.IsNullOrEmpty(row.ResolutionComment) && 
                                  row.ResolutionComment.IndexOf("20002", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool hasConnectionTimeout = !string.IsNullOrEmpty(row.ResolutionComment) && 
                                               row.ResolutionComment.IndexOf("Connection Timeout", StringComparison.OrdinalIgnoreCase) >= 0;
                    
                    Console.WriteLine($"\n 🔍 DEBUG: Case ID '{row.CaseId}' - PHX={hasPHX}, ConnectionTimeout={hasConnectionTimeout}, ResComment='{row.ResolutionComment}'");
                    
                    if (hasPHX || hasConnectionTimeout)
                    {
                        string keyword = hasPHX ? "PHX 20002" : "Connection Timeout";
                        Console.WriteLine($"✅ FOUND {keyword} in Case ID '{row.CaseId}' (Row {row.RowIndex}). Adding to PHX group.");
                        phxCompletedRows.Add(row);
                    }
                }

                // 2. กลุ่มที่ตรงกับ reasonsToFilter (แต่ไม่ใช่ PHX หรือ Connection Timeout cases)
                Console.WriteLine("\n--- DEBUG: Checking rows for matching reasons ---");
                var matchingReasonRows = new List<MasterExcelRow>();
                foreach (var r in allMasterRows)
                {
                    bool isPHXOrTimeout = (!string.IsNullOrEmpty(r.ResolutionComment) && 
                                          (r.ResolutionComment.IndexOf("20002", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                           r.ResolutionComment.IndexOf("Connection Timeout", StringComparison.OrdinalIgnoreCase) >= 0));
                    bool matchesReason = reasonsToFilter.Contains(r.Reason, StringComparer.OrdinalIgnoreCase);
                    
                    Console.WriteLine($"🔍 DEBUG: Case ID '{r.CaseId}' - Reason='{r.Reason}', MatchesFilter={matchesReason}, IsPHXOrTimeout={isPHXOrTimeout}");
                    
                    if (matchesReason && !isPHXOrTimeout)
                    {
                        Console.WriteLine($"✅ ADDED to Matching Reason group: Case ID '{r.CaseId}' with reason '{r.Reason}'");
                        matchingReasonRows.Add(r);
                    }
                }

                // 3. กลุ่มที่เหลือ (ต้องเช็คเว็บ) - ไม่ใช่ PHX/Connection Timeout และไม่ตรงกับ reasonsToFilter
                Console.WriteLine("\n--- DEBUG: Checking rows that need web verification ---");
                var otherReasonRows = new List<MasterExcelRow>();
                foreach (var r in allMasterRows)
                {
                    bool isPHXOrTimeout = (!string.IsNullOrEmpty(r.ResolutionComment) && 
                                          (r.ResolutionComment.IndexOf("20002", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                           r.ResolutionComment.IndexOf("Connection Timeout", StringComparison.OrdinalIgnoreCase) >= 0));
                    bool matchesReason = reasonsToFilter.Contains(r.Reason, StringComparer.OrdinalIgnoreCase);
                    
                    if (!matchesReason && !isPHXOrTimeout)
                    {
                        Console.WriteLine($"✅ ADDED to Web Check group: Case ID '{r.CaseId}' with reason '{r.Reason}'");
                        otherReasonRows.Add(r);
                    }
                }

                Console.WriteLine($"\n📊 GROUPING SUMMARY:");
                Console.WriteLine($"   🔴 PHX/Connection Timeout group: {phxCompletedRows.Count} cases (will be marked as Completed)");
                Console.WriteLine($"   🟡 Matching reason group: {matchingReasonRows.Count} cases (will be marked as Rejected)");
                Console.WriteLine($"   🔵 Web check group: {otherReasonRows.Count} cases (need web verification)");
                Console.WriteLine($"   📋 Total processed: {phxCompletedRows.Count + matchingReasonRows.Count + otherReasonRows.Count} of {allMasterRows.Count}");

                // Process file close case IM
                // ✅ โหลดไฟล์ Excel Template ที่เพิ่งคัดลอกมาเพื่อเขียนข้อมูล
                using (var outputPackage = new ExcelPackage(new FileInfo(outputImportFilePath)))
                {
                    var outputWorksheet = outputPackage.Workbook.Worksheets.FirstOrDefault();
                    if (outputWorksheet == null)
                    {
                        Console.WriteLine($"Error: No worksheet found in output template file '{outputImportFilePath}'.");
                        return null;
                    }

                    // ค้นหา Column Index ของคอลัมน์ที่ต้องการใน Output File (Template)
                    int templateCaseIdCol = -1;
                    int templateTopicCol = -1;
                    int templateCaseStatusCol = -1;
                    int templateCommentCol = -1;

                    for (int col = 1; col <= outputWorksheet.Dimension.End.Column; col++)
                    {
                        string header = outputWorksheet.Cells[1, col].Text.Trim(); // อ่าน Header แถวแรก
                        if (header.Equals(templateCaseIdColumnName, StringComparison.OrdinalIgnoreCase))
                            templateCaseIdCol = col;
                        else if (header.Equals(templateTopicColumnName, StringComparison.OrdinalIgnoreCase))
                            templateTopicCol = col;
                        else if (header.Equals(templateCaseStatusColumnName, StringComparison.OrdinalIgnoreCase))
                            templateCaseStatusCol = col;
                        else if (header.Equals(templateCommentColumnName, StringComparison.OrdinalIgnoreCase))
                            templateCommentCol = col;
                    }

                    if (templateCaseIdCol == -1 || templateTopicCol == -1 || templateCaseStatusCol == -1 || templateCommentCol == -1)
                    {
                        Console.WriteLine("Error: Missing one or more required columns (Case ID, Topic, Case Status, Comment) in output template file.");
                        return null;
                    }

                    // เริ่มเขียนข้อมูลจากแถวถัดจาก Header ในไฟล์ Output
                    int outputImportRow = 2;

                    // ✅ ประมวลผลกลุ่ม PHX cases ก่อน (มีความสำคัญสูงสุด)
                    Console.WriteLine("\n--- Processing PHX 20002 and Connection Timeout cases ---");
                    foreach (var rowData in phxCompletedRows)
                    {
                        // ตรวจสอบว่าเป็น PHX หรือ Connection Timeout
                        string detectedKeyword = "";
                        if (!string.IsNullOrEmpty(rowData.ResolutionComment))
                        {
                            if (rowData.ResolutionComment.IndexOf("20002", StringComparison.OrdinalIgnoreCase) >= 0)
                                detectedKeyword = "PHX 20002";
                            else if (rowData.ResolutionComment.IndexOf("Connection Timeout", StringComparison.OrdinalIgnoreCase) >= 0)
                                detectedKeyword = "Connection Timeout";
                        }
                        
                        Console.WriteLine($"Case ID '{rowData.CaseId}' (Resolution Comment contains {detectedKeyword}) - Completed.");
                        
                        // เขียนข้อมูลลงใน Output File และตั้งค่าสี Font เป็นสีแดง
                        outputWorksheet.Cells[outputImportRow, templateCaseIdCol].Value = rowData.CaseId;
                        outputWorksheet.Cells[outputImportRow, templateCaseIdCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputWorksheet.Cells[outputImportRow, templateTopicCol].Value = rowData.Topic;
                        outputWorksheet.Cells[outputImportRow, templateTopicCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputWorksheet.Cells[outputImportRow, templateCaseStatusCol].Value = caseStatusForCompleted;
                        outputWorksheet.Cells[outputImportRow, templateCaseStatusCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputWorksheet.Cells[outputImportRow, templateCommentCol].Value = commentForCompleted;
                        outputWorksheet.Cells[outputImportRow, templateCommentCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputImportRow++;
                    }

                    // ✅ ประมวลผลกลุ่มที่ตรงตามเงื่อนไข Reason (แต่ไม่ใช่ PHX cases)
                    Console.WriteLine("\n--- Processing rows matching filter reasons ---");
                    foreach (var rowData in matchingReasonRows)
                    {
                        Console.WriteLine($"Case ID '{rowData.CaseId}' (Reason: '{rowData.Reason}') - Reject.");
                        
                        // เขียนข้อมูลลงใน Output File และตั้งค่าสี Font เป็นสีแดง
                        outputWorksheet.Cells[outputImportRow, templateCaseIdCol].Value = rowData.CaseId;
                        outputWorksheet.Cells[outputImportRow, templateCaseIdCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputWorksheet.Cells[outputImportRow, templateTopicCol].Value = rowData.Topic;
                        outputWorksheet.Cells[outputImportRow, templateTopicCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputWorksheet.Cells[outputImportRow, templateCaseStatusCol].Value = caseStatusRejected;
                        outputWorksheet.Cells[outputImportRow, templateCaseStatusCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputWorksheet.Cells[outputImportRow, templateCommentCol].Value = commentForRejected;
                        outputWorksheet.Cells[outputImportRow, templateCommentCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        outputImportRow++;
                    }

                    // ✅ ประมวลผลกลุ่มที่ไม่ตรงตามเงื่อนไข Reason (ต้องเข้าเว็บ)
                    Console.WriteLine("\n--- Processing rows not matching filter reasons (Web check required) ---");
                    WebDriverWait webWait = new WebDriverWait(driver, TimeSpan.FromSeconds(45)); // เพิ่มเวลารอจาก 30 เป็น 45 วินาที
                    int webCheckCount = 0;
                    int webCheckSuccess = 0;
                    int webCheckFailed = 0;

                    foreach (var rowData in otherReasonRows)
                    {
                        webCheckCount++;
                        Console.WriteLine($"\n🌐 WEB CHECK {webCheckCount}/{otherReasonRows.Count}: Starting check for Case ID '{rowData.CaseId}' (Reason: '{rowData.Reason}')");
                        
                        try
                        {
                            // ✅ Reset context ก่อนเริ่มทุกเคส
                            driver.SwitchTo().DefaultContent();
                            // ✅ รอจน iframe โหลดเสร็จ
                            Console.WriteLine($"⏳ Step 1: Waiting for iframe PegaGadget0Ifr...");
                            webWait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.XPath("//iframe[@id='PegaGadget0Ifr']")));
                            Console.WriteLine($"✅ Step 1: Successfully switched to iframe PegaGadget0Ifr");

                            // ✅ หา element My Work หลังเข้า iframe แล้ว
                            Console.WriteLine($"⏳ Step 2: Looking for My Work element...");
                            IWebElement myReportsElement = webWait.Until(ExpectedConditions.ElementIsVisible(
                                By.XPath("//li[@title='My Work']")));
                            myReportsElement.Click();
                            Thread.Sleep(defaultLongWait);
                            Console.WriteLine($"✅ Step 2: Successfully clicked My Work");

                            // ✅ หาเมนู Search All Case ภายใน iframe เดียวกัน
                            Console.WriteLine($"⏳ Step 3: Looking for Search All Case menu...");
                            IWebElement menuSearch = webWait.Until(ExpectedConditions.ElementIsVisible(
                                By.XPath("//div[@aria-label='Search All Case']")));
                            menuSearch.Click();
                            Thread.Sleep(defaultLongWait);
                            Console.WriteLine($"✅ Step 3: Successfully clicked Search All Case");

                            // ✅ ระบุ Locator ที่ถูกต้องสำหรับช่องใส่ Case ID บนเว็บ
                            Console.WriteLine($"⏳ Step 4: Looking for Case ID input field...");
                            IWebElement caseIdInputField = webWait.Until(ExpectedConditions.ElementIsVisible(By.Id("37ecd95d")));
                            caseIdInputField.Clear();
                            caseIdInputField.SendKeys(rowData.CaseId + Keys.Enter);
                            Thread.Sleep(defaultLongWait * 2); // เพิ่มเวลารอให้ผลลัพธ์โหลด
                            Console.WriteLine($"✅ Step 4: Successfully entered Case ID '{rowData.CaseId}' and pressed Enter");

                            // ✅ คลิกเข้าไปดูรายละเอียดของ Case ID
                            Console.WriteLine($"⏳ Step 5: Looking for Case ID link to click...");
                            IWebElement caseIdLinkElement = webWait.Until(ExpectedConditions.ElementToBeClickable(
                                By.XPath($"//a[contains(text(), '{rowData.CaseId}')]")));
                            caseIdLinkElement.Click();
                            Thread.Sleep(defaultLongWait * 2); // เพิ่มเวลารอให้หน้ารายละเอียดโหลด
                            Console.WriteLine($"✅ Step 5: Successfully clicked Case ID link");

                            // ✅ กลับสู่ context หลัก (optional ถ้า iframe เดียว)
                            Console.WriteLine($"⏳ Step 6: Switching back to default content...");
                            driver.SwitchTo().DefaultContent();

                            // ✅ รอจน iframe โหลดเสร็จ
                            Console.WriteLine($"⏳ Step 7: Waiting for iframe PegaGadget1Ifr...");
                            webWait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.XPath("//iframe[@id='PegaGadget1Ifr']")));
                            Console.WriteLine($"✅ Step 7: Successfully switched to iframe PegaGadget1Ifr");
                            
                            // ✅ เลื่อนหน้าจอไปยัง Element ของ table Case History
                            // และรอให้ Element นั้นปรากฏ (ปรับปรุงให้เร็วขึ้น)
                            Console.WriteLine($"⏳ Step 8: Looking for Case History section...");
                            IWebElement targetDivElement = webWait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//h2[text()='Case History']")));
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", targetDivElement);
                            Thread.Sleep(1000); // ลดเวลารอจาก 2000 เป็น 1000 ms
                            Console.WriteLine($"✅ Step 8: Successfully found and scrolled to Case History");

                            // ✅ หา Element ที่มีข้อความ "Marketing" (ปรับปรุงให้เร็วขึ้น)
                            Console.WriteLine($"⏳ Step 9: Looking for Marketing approval status...");
                            string marketingElementXPath = "//div[@data-layout-id='202402161557340851']//tr[@pl_index='1'][.//td[@headers='a2']/div/span[contains(text(), 'Marketing')] and .//td[@headers='a4']/div/span[contains(text(), 'Approved')] ]";

                            string PositionText = "Unknown"; // กำหนดค่าเริ่มต้น
                            bool marketingFound = false;

                            try
                            {
                                // ใช้ WebDriverWait แบบสั้นเฉพาะสำหรับ Marketing check (10 วินาที)
                                WebDriverWait shortWait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                                IWebElement marketingTextElement = shortWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(marketingElementXPath)));
                                PositionText = marketingTextElement.Text.Trim();
                                Console.WriteLine($"🔍 Step 9: Found marketing element with text: '{PositionText}'");

                                // ใช้ IndexOf เพื่อตรวจสอบว่ามีคำว่า "Marketing" หรือไม่ (ไม่สนใจตัวพิมพ์เล็ก-ใหญ่)
                                if (PositionText.IndexOf("Marketing", StringComparison.OrdinalIgnoreCase) != -1)
                                {
                                    marketingFound = true;
                                    Console.WriteLine($"✅ Step 9: Marketing approval FOUND - Case will go to batch processing");
                                }
                                else
                                {
                                    marketingFound = false;
                                    Console.WriteLine($"❌ Step 9: Marketing approval NOT FOUND - Case will be rejected");
                                }
                            }
                            catch (WebDriverTimeoutException)
                            {
                                // ถ้าไม่เจอ Marketing element ภายใน 10 วินาที ให้ถือว่าไม่มี approval
                                Console.WriteLine($"⏰ Step 9: Marketing element not found within 10 seconds - treating as no approval");
                                marketingFound = false;
                            }

                            // ✅ ตรวจสอบเงื่อนไข: ถ้าข้อความสถานะไม่มีคำว่า "Marketing" ให้บันทึกลง Excel
                            // (หมายถึงถ้า marketingFound เป็น false หรือหาเจอแต่ข้อความไม่มีคำว่า Marketing)
                            if (!marketingFound) // ถ้าหา Element ไม่เจอ หรือหาเจอแต่ไม่มีคำว่า Marketing (ตาม Logic ที่แก้ไขด้านบน)
                            {
                                webCheckSuccess++; // นับจำนวนการเช็คที่สำเร็จ (ถึงแม้จะเป็น reject)
                                Console.WriteLine($"Case ID '{rowData.CaseId}' (Reason: '{rowData.Reason}') - Reject (no marketing approval).");

                                outputWorksheet.Cells[outputImportRow, templateCaseIdCol].Value = rowData.CaseId;
                                outputWorksheet.Cells[outputImportRow, templateCaseIdCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                                outputWorksheet.Cells[outputImportRow, templateTopicCol].Value = rowData.Topic;
                                outputWorksheet.Cells[outputImportRow, templateTopicCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                                outputWorksheet.Cells[outputImportRow, templateCaseStatusCol].Value = caseStatusRejected;
                                outputWorksheet.Cells[outputImportRow, templateCaseStatusCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                                outputWorksheet.Cells[outputImportRow, templateCommentCol].Value = commentForRejected;
                                outputWorksheet.Cells[outputImportRow, templateCommentCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                                outputImportRow++;

                            }
                        
                            else
                            {
                                webCheckSuccess++; // นับจำนวนการเช็คที่สำเร็จ
                                Console.WriteLine($"Case ID '{rowData.CaseId}' contains Marketing approval. Preparing data for Batch Excel.");

                                // ✅ ดึงค่า Offering Code จากไฟล์ Mapping
                                string offeringCode = "N/A"; // ค่าเริ่มต้นหากหาไม่เจอ
                                if (!string.IsNullOrWhiteSpace(rowData.Doc) && mappingData.ContainsKey(rowData.Doc))
                                {
                                    offeringCode = mappingData[rowData.Doc];
                                    //Console.WriteLine($"Found Offering Code '{offeringCode}' for Doc# '{rowData.Doc}'.");
                                }
                                else
                                {
                                    Console.WriteLine($"Warning: Doc# '{rowData.Doc}' not found in mapping file or is empty. Using 'N/A'.");
                                }

                                // ✅ ตรวจสอบ OfferingCode = N/A แล้วแยกการจัดการ
                                if (offeringCode.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"🔍 Case ID '{rowData.CaseId}' has OfferingCode = N/A - Adding to manual check list.");
                                    naOfferingCases.Add(rowData);
                                    
                                    // ❌ ไม่ใส่ในไฟล์ Import_CloseCaseIM (ให้ user จัดการเอง)
                                    // ❌ ไม่ใส่ใน batchEntries เพราะไม่มี offering code ที่ถูกต้อง
                                }
                                else
                                {
                                    // ✅ เฉพาะเคสที่มี OfferingCode ที่ถูกต้องเท่านั้นที่ใส่ลง Batch
                                    Console.WriteLine($"✅ Case ID '{rowData.CaseId}' has valid OfferingCode '{offeringCode}' - Adding to batch processing.");
                                    
                                    // ✅ เตรียมข้อมูลสำหรับไฟล์ PSP_SR
                                    batchEntries.Add(new BatchEntry
                                    {
                                        OfferingCode = offeringCode,
                                        UserName = rowData.CreatedBy,
                                        MobileNo = rowData.Mobile,
                                        ServiceRequest1 = rowData.CaseId,
                                        Category1 = rowData.CaseId,
                                        SubCategory1 = rowData.CaseId 
                                    });
                                }
                            }


                            // ✅ หาปุ่ม Close และคลิก (ปรับปรุงให้เร็วขึ้น)
                            IWebElement closeMenu = webWait.Until(ExpectedConditions.ElementToBeClickable(
                                By.XPath("//div[@data-node-id='CPMReviewDriverWrapper']//button[@title='Close']")));
                            closeMenu.Click();
                            Thread.Sleep(defaultShortWait); // ลดเวลารอจาก defaultLongWait เป็น defaultShortWait
                            // ✅ กลับสู่ context หลัก (optional ถ้า iframe เดียว)
                            driver.SwitchTo().DefaultContent();

                        }
                        catch (Exception webEx)
                        {
                            webCheckFailed++;
                            Console.WriteLine($"❌ WEB CHECK FAILED for Case ID '{rowData.CaseId}': {webEx.Message}");
                            Console.WriteLine($"⚠️ This case will be skipped from processing");
                            
                            // ในกรณีที่เกิด Exception ก่อนที่จะหาปุ่ม Close เจอ
                            // เราต้องพยายามหาปุ่ม Close อีกครั้งเพื่อไม่ให้ค้างอยู่ในหน้าเดิม
                            try
                            {
                                // ลองสลับไป DefaultContent ก่อน (ถ้าไม่แน่ใจว่าตอนนี้อยู่ IFrame ไหน)
                                driver.SwitchTo().DefaultContent();
                                Console.WriteLine("⏳ Attempting to switch to Default Content to find Close button after error.");

                                // ลองหาปุ่ม Close อีกครั้ง
                                IWebElement closeButtonOnError = webWait.Until(ExpectedConditions.ElementToBeClickable(
                                    By.XPath("//div[@data-node-id='CPMReviewDriverWrapper']//button[@title='Close']")));
                                closeButtonOnError.Click();
                                Console.WriteLine("✅ Successfully clicked 'Close' button after error.");
                                Thread.Sleep(2000); // รอให้หน้าจอปิด
                            }
                            catch (Exception innerEx)
                            {
                                Console.WriteLine($"❌ Critical: Could not close page for Case ID '{rowData.CaseId}' after error: {innerEx.Message}");
                                // หากยังปิดไม่ได้ อาจจะต้องพิจารณา driver.Navigate().GoToUrl(webAppUrl); เพื่อกลับไปหน้าเริ่มต้น
                                // หรือ driver.Quit() เพื่อจบการทำงานทั้งหมด
                            }
                        }
                    }

                    Console.WriteLine($"\n📊 FINAL WEB CHECK SUMMARY:");
                    Console.WriteLine($"   Total web checks attempted: {webCheckCount}");
                    Console.WriteLine($"   Successful web checks: {webCheckSuccess}");
                    Console.WriteLine($"   Failed web checks: {webCheckFailed}");
                    Console.WriteLine($"   Cases added to batch processing: {batchEntries.Count}");
                    Console.WriteLine($"   🔍 Cases with N/A OfferingCode (require manual check): {naOfferingCases.Count}");

                    // ✅ ปรับขนาดคอลัมน์อัตโนมัติและบันทึกไฟล์ Import_CloseCaseIM
                    outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                    outputPackage.Save();

                    Console.WriteLine($"\n✅ Successfully processed and exported Import_CloseCaseIM to '{outputImportFilePath}'.");
                    Console.WriteLine($"📊 FINAL SUMMARY - Import_CloseCaseIM File:");
                    Console.WriteLine($"   Total rows in file: {outputImportRow - 2}");
                    Console.WriteLine($"   - PHX/Connection Timeout (Completed): {phxCompletedRows.Count}");
                    Console.WriteLine($"   - Matching Reason (Rejected): {matchingReasonRows.Count}");
                    Console.WriteLine($"   - Web Check (Rejected): {webCheckSuccess - batchEntries.Count}"); // แก้ไขให้แสดงจำนวนที่ถูกต้อง
                    Console.WriteLine($"   ⚠️ Cases with N/A OfferingCode (excluded): {naOfferingCases.Count}");
                } // End using (var outputPackage ... )

                // ✅ สร้างไฟล์รายงานแยกสำหรับเคส N/A OfferingCode
                if (naOfferingCases.Any())
                {
                    string naReportPath = Path.Combine(outputDirectory, "ManualCheck_CaseOfferingCode.xlsx");
                    CreateNAOfferingCodeReport(naOfferingCases, naReportPath);
                    Console.WriteLine($"\n📋 Created manual check report for N/A OfferingCode cases: '{naReportPath}'");
                    Console.WriteLine($"   ⚠️ Please manually review {naOfferingCases.Count} cases that require Doc# mapping verification.");
                }

                // ✅ เขียนข้อมูลลงไฟล์ PSP_SR
                if (batchEntries.Any())
                {
                    Console.WriteLine($"\n--- Writing {batchEntries.Count} entries to PSP_SR Excel file ---");

                    // *** เพิ่มการ Debug: ตรวจสอบพาธของไฟล์ PSP_SR Template และ Output ***
                    //Console.WriteLine($"Batch Template Path: '{templateBatchFilePath}'");
                    //Console.WriteLine($"Output Batch File Path: '{outputBatchFilePath}'");

                    // คัดลอก Template Batch
                    File.Copy(templateBatchFilePath, outputBatchFilePath, true);
                    Console.WriteLine($"Copied batch template file to '{outputBatchFilePath}'.");

                    using (var batchPackage = new ExcelPackage(new FileInfo(outputBatchFilePath)))
                    {
                        var batchWorksheet = batchPackage.Workbook.Worksheets["InputFile"];
                        //var batchWorksheet = batchPackage.Workbook.Worksheets.FirstOrDefault();
                        Console.WriteLine($"Batch worksheet name: '{batchWorksheet.Name}'");
                        if (batchWorksheet == null)
                        {
                            Console.WriteLine($"Error: No worksheet found in batch template file '{outputBatchFilePath}'.");
                            return outputImportFilePath;
                        }

                        // ค้นหา Column Index ของคอลัมน์ที่ต้องการใน Batch File                  
                        int batchUserNameCol = -1;
                        int batchMobileNoCol = -1;
                        int batchOfferingCodeCol = -1;
                        int batchActionStatus1Col = -1;
                        int batchBypassProductRule1Col = -1;
                        int batchBypassSMSCompleteFlag1Col = -1;
                        int batchServiceRequestCol = -1;
                        int batchCategoryCol = -1;
                        int batchSubCategoryCol = -1;
                        // *** Debug: Loop เพื่ออ่าน Header และแสดงค่าที่อ่านได้ ***
                        // อ่านจากแถวที่ 1 (สมมติว่า Header อยู่ที่แถว 1)
                        for (int col = 1; col <= batchWorksheet.Dimension.End.Column; col++)
                        {
                            var cell = batchWorksheet.Cells[1, col]; // อ้างอิงเซลล์
                            string headerText = cell.Text.Trim(); // ค่าที่แสดงผล
                            object headerValue = cell.Value; // ค่าดิบในเซลล์ (อาจเป็น string, int, double, null)

                            // *** MODIFICATION 2: Add a fallback check if headerText is numeric or empty ***
                            // ถ้า headerText ยังคงเป็น '1', '2', ฯลฯ หรือว่างเปล่า ให้ลองแปลง headerValue เป็น string
                            string effectiveHeaderText = headerText;
                            if (string.IsNullOrWhiteSpace(effectiveHeaderText) || (headerValue is double && (double)headerValue == col))
                            {
                                effectiveHeaderText = headerValue?.ToString()?.Trim() ?? "";
                                Console.WriteLine($"  -> Fallback: Using Header Value as text: '{effectiveHeaderText}'");
                            }

                            //Console.WriteLine($"Checking column {col}: Cell Address='{cell.Address}', Header Text='{headerText}', Header Value='{headerValue ?? "NULL"}' (Type: {headerValue?.GetType().Name ?? "null"})");
                            if (headerText.Equals(batchOfferingCodeColName, StringComparison.OrdinalIgnoreCase))
                                batchOfferingCodeCol = col;
                            else if (headerText.Equals(batchUserNameColName, StringComparison.OrdinalIgnoreCase))
                                batchUserNameCol = col;
                            else if (headerText.Equals(batchMobileNoColName, StringComparison.OrdinalIgnoreCase))
                                batchMobileNoCol = col;
                            else if (headerText.Equals(batchActionStatus1Name, StringComparison.OrdinalIgnoreCase))
                                batchActionStatus1Col = col;
                            else if (headerText.Equals(batchBypassProductRule1Name, StringComparison.OrdinalIgnoreCase))
                                batchBypassProductRule1Col = col;
                            else if (headerText.Equals(batchBypassSMSCompleteFlag1Name, StringComparison.OrdinalIgnoreCase))
                                batchBypassSMSCompleteFlag1Col = col;
                            else if (headerText.Equals(batchServiceRequestColName, StringComparison.OrdinalIgnoreCase))
                                batchServiceRequestCol = col;
                            else if (headerText.Equals(batchCategoryColName, StringComparison.OrdinalIgnoreCase))
                                batchCategoryCol = col;
                            else if (headerText.Equals(batchSubCategoryColName, StringComparison.OrdinalIgnoreCase))
                                batchSubCategoryCol = col;
                        }

                        if (batchOfferingCodeCol == -1 || batchUserNameCol == -1 || batchMobileNoCol == -1 ||
                            batchActionStatus1Col == -1 || batchBypassProductRule1Col == -1 || batchBypassSMSCompleteFlag1Col == -1 ||
                            batchServiceRequestCol == -1 || batchCategoryCol == -1 || batchSubCategoryCol == -1)
                        {
                            Console.WriteLine("Error: Missing one or more required columns in batch template file.");

                            return outputImportFilePath;
                        }

                        int currentBatchRow = 2; // เริ่มเขียนจากแถวที่ 2 (หลัง Header)
                        foreach (var entry in batchEntries)
                        {
                            batchWorksheet.Cells[currentBatchRow, batchOfferingCodeCol].Value = entry.OfferingCode;
                            batchWorksheet.Cells[currentBatchRow, batchUserNameCol].Value = entry.UserName;
                            batchWorksheet.Cells[currentBatchRow, batchMobileNoCol].Value = entry.MobileNo;
                            batchWorksheet.Cells[currentBatchRow, batchActionStatus1Col].Value = fixedActionStatus1Value;
                            batchWorksheet.Cells[currentBatchRow, batchBypassProductRule1Col].Value = fixedBypassProductRule1Value;
                            batchWorksheet.Cells[currentBatchRow, batchBypassSMSCompleteFlag1Col].Value = fixedBypassSMSCompleteFlag1Value;
                            batchWorksheet.Cells[currentBatchRow, batchServiceRequestCol].Value = entry.ServiceRequest1;
                            batchWorksheet.Cells[currentBatchRow, batchCategoryCol].Value = entry.Category1;
                            batchWorksheet.Cells[currentBatchRow, batchSubCategoryCol].Value = entry.SubCategory1;
                            currentBatchRow++;
                        }

                        batchWorksheet.Cells[batchWorksheet.Dimension.Address].AutoFitColumns();
                        batchPackage.Save();
                        Console.WriteLine($"Successfully created Batch Excel file: '{outputBatchFilePath}'.");
                    }
                }
                else
                {
                    Console.WriteLine("\nNo entries to write to Batch Excel file.");
                }

                return outputImportFilePath; // คืนค่าพาธของไฟล์ Import_CloseCaseIM
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during Excel processing: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return null;
            }
        }

        /// <summary>
        /// Helper method สำหรับโหลดข้อมูล Mapping จากไฟล์ Excel
        /// </summary>
        /// <param name="filePath">พาธของไฟล์ Mapping</param>
        /// <param name="featureCodeColName">ชื่อคอลัมน์ Feature Code ในไฟล์ Mapping</param>
        /// <param name="offeringCodeColName">ชื่อคอลัมน์ Offering Code ในไฟล์ Mapping</param>
        /// <returns>Dictionary ที่มี Feature Code เป็น Key และ Offering Code เป็น Value</returns>
        private Dictionary<string, string> LoadMappingData(string filePath, string featureCodeColName, string offeringCodeColName)
        {
            Dictionary<string, string> mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null || worksheet.Dimension == null)
                    {
                        Console.WriteLine($"Warning: No data or worksheet found in mapping file '{filePath}'.");
                        return null;
                    }

                    int featureCodeCol = -1;
                    int offeringCodeCol = -1;

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string header = worksheet.Cells[1, col].Text.Trim();
                        if (header.Equals(featureCodeColName, StringComparison.OrdinalIgnoreCase))
                            featureCodeCol = col;
                        else if (header.Equals(offeringCodeColName, StringComparison.OrdinalIgnoreCase))
                            offeringCodeCol = col;
                    }

                    if (featureCodeCol == -1 || offeringCodeCol == -1)
                    {
                        Console.WriteLine($"Error: Missing '{featureCodeColName}' or '{offeringCodeColName}' column in mapping file '{filePath}'.");
                        return null;
                    }

                    for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string featureCode = worksheet.Cells[row, featureCodeCol].Text.Trim();
                        string offeringCode = worksheet.Cells[row, offeringCodeCol].Text.Trim();

                        if (!string.IsNullOrWhiteSpace(featureCode))
                        {
                            if (!mapping.ContainsKey(featureCode))
                            {
                                mapping.Add(featureCode, offeringCode);
                            }
                            else
                            {
                                Console.WriteLine($"Warning: Duplicate Feature Code '{featureCode}' found in mapping file. Using first occurrence.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading mapping data from '{filePath}': {ex.Message}");
                return null;
            }
            return mapping;
        }


        /// <summary>
        /// หาหมายเลขแถวสุดท้ายที่มีข้อมูลจริงๆ ใน worksheet
        /// </summary>
        /// <param name="worksheet">Worksheet ที่ต้องการตรวจสอบ</param>
        /// <returns>หมายเลขแถวสุดท้ายที่มีข้อมูล หากไม่มีจะคืนค่าเป็น 0</returns>
        private int GetLastUsedRow(ExcelWorksheet worksheet)
        {
            // ถ้า worksheet ไม่มี Dimension (ว่างเปล่า) หรือ null ให้คืนค่า 0
            if (worksheet == null || worksheet.Dimension == null)
            {
                return 0;
            }

            // เริ่มวนลูปจากแถวสุดท้ายที่ EPPlus คิดว่ามีข้อมูลขึ้นมาข้างบน
            for (int row = worksheet.Dimension.End.Row; row >= 1; row--)
            {
                // เช็คว่าแถวนั้นมีเซลล์ใดที่มีค่า (ไม่ใช่ช่องว่าง) หรือไม่
                // ToArray() ใช้เพื่อแปลง Range เป็น Array เพื่อให้วนลูปได้
                var rowCells = worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column].ToArray();

                // ใช้ LINQ เพื่อตรวจสอบว่ามีเซลล์ใดในแถวที่มีค่าไม่ใช่ null หรือ string.Empty
                if (rowCells.Any(cell => cell.Value != null && !string.IsNullOrWhiteSpace(cell.Text)))
                {
                    // ถ้าเจอแถวแรกที่มีข้อมูล ให้คืนค่าหมายเลขแถวนั้น
                    return row;
                }
            }

            // ถ้าวนลูปจนหมดแล้วไม่เจอข้อมูลเลย ให้คืนค่า 0
            return 0;
        }

        /// <summary>
        /// สรุปจำนวนเคสจากไฟล์ Excel ที่ถูกประมวลผล และบันทึกลงในไฟล์ Excel ใหม่พร้อมการจัดรูปแบบ
        /// </summary>
        /// <param name="masterFilePath">พาธเต็มของไฟล์ Excel ข้อมูลหลักต้นฉบับ</param>
        /// <param name="importCloseCaseIMFilePath">พาธเต็มของไฟล์ Excel Import_CloseCaseIM ที่ถูกสร้าง</param>
        /// <param name="batchFilePath">พาธเต็มของไฟล์ Excel Batch ที่ถูกสร้าง (ถ้ามี)</param>
        /// <param name="naOfferingCodeFilePath">พาธเต็มของไฟล์ Manual Check N/A OfferingCode (ถ้ามี)</param>
        /// <param name="outputDirectory">โฟลเดอร์สำหรับบันทึกไฟล์ Excel สรุปผลลัพธ์</param>
        /// <returns>พาธเต็มของไฟล์ Excel สรุปผลลัพธ์ที่สร้างขึ้น หากสำเร็จ; Null หากเกิดข้อผิดพลาด</returns>
        public string SummarizeProcessedData(string masterFilePath, string importCloseCaseIMFilePath, string batchFilePath, string outputDirectory, string naOfferingCodeFilePath = null)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            int totalPromotionPrepaidCases = 0;
            int completedCases = 0;
            int rejectedCases = 0;
            int naOfferingCodeCasesFromReport = 0; // จากไฟล์ ManualCheck_NAOfferingCode.xlsx
            List<string> naOfferingCodeCases = new List<string>(); // จากไฟล์ Batch (legacy support)

            // 1. นับจำนวนเคสทั้งหมดจากไฟล์ Master ที่ Category = "Promotion Prepaid*"
            try
            {
                if (File.Exists(masterFilePath))
                {
                    using (var package = new ExcelPackage(new FileInfo(masterFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        if (worksheet != null && worksheet.Dimension != null)
                        {
                            // ค้นหาคอลัมน์ Category
                            int categoryCol = -1;
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                string header = worksheet.Cells[1, col].Text.Trim();
                                if (header.Equals("Category", StringComparison.OrdinalIgnoreCase))
                                {
                                    categoryCol = col;
                                    break;
                                }
                            }

                            if (categoryCol != -1)
                            {
                                // นับแถวที่ Category เริ่มต้นด้วย "Promotion Prepaid"
                                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    string category = worksheet.Cells[row, categoryCol].Text.Trim();
                                    if (!string.IsNullOrEmpty(category) && 
                                        category.IndexOf("Promotion Prepaid", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        totalPromotionPrepaidCases++;
                                    }
                                }
                                Console.WriteLine($"พบเคส Promotion Prepaid ทั้งหมด: {totalPromotionPrepaidCases} เคส");
                            }
                            else
                            {
                                Console.WriteLine("ไม่พบคอลัมน์ Category ในไฟล์ Master");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Master file not found at '{masterFilePath}'. Cannot count total cases.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error counting total cases from master file: {ex.Message}");
            }

            // 2. นับจำนวนเคสที่ Complete และ Reject จากไฟล์ Import_CloseCaseIM
            try
            {
                if (File.Exists(importCloseCaseIMFilePath))
                {
                    using (var package = new ExcelPackage(new FileInfo(importCloseCaseIMFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        if (worksheet != null && worksheet.Dimension != null)
                        {
                            // ค้นหาคอลัมน์ Case Status
                            int caseStatusCol = -1;
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                string header = worksheet.Cells[1, col].Text.Trim();
                                if (header.Equals("Case Status", StringComparison.OrdinalIgnoreCase))
                                {
                                    caseStatusCol = col;
                                    break;
                                }
                            }

                            if (caseStatusCol != -1)
                            {
                                // นับตาม Case Status
                                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    string caseStatus = worksheet.Cells[row, caseStatusCol].Text.Trim();
                                    if (!string.IsNullOrEmpty(caseStatus))
                                    {
                                        if (caseStatus.IndexOf("Resolved-Completed", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            completedCases++;
                                        }
                                        else if (caseStatus.IndexOf("Resolved-Rejected", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            rejectedCases++;
                                        }
                                    }
                                }
                                Console.WriteLine($"พบเคส Completed: {completedCases} เคส, Rejected: {rejectedCases} เคส");
                            }
                            else
                            {
                                Console.WriteLine("ไม่พบคอลัมน์ Case Status ในไฟล์ Import_CloseCaseIM");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Import_CloseCaseIM file not found at '{importCloseCaseIMFilePath}'. Cannot count cases by status.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error counting cases from Import_CloseCaseIM file: {ex.Message}");
            }

            // 3. นับจำนวนเคสที่ Marketing Approve และเคสที่มี OfferingCode = N/A
            try
            {
                if (File.Exists(batchFilePath))
                {
                    using (var package = new ExcelPackage(new FileInfo(batchFilePath)))
                    {
                        // เข้าถึง worksheet ที่ชื่อ "InputFile" โดยตรง
                        var worksheet = package.Workbook.Worksheets["InputFile"];
                        if (worksheet != null && worksheet.Dimension != null)
                        {
                            
                            // ค้นหาคอลัมน์ OfferingCode1 และ ServiceRequest1
                            int offeringCodeCol = -1;
                            int serviceRequestCol = -1;
                            
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                string header = worksheet.Cells[1, col].Text.Trim();
                                if (header.Equals("OfferingCode1", StringComparison.OrdinalIgnoreCase))
                                {
                                    offeringCodeCol = col;
                                }
                                else if (header.Equals("ServiceRequest1", StringComparison.OrdinalIgnoreCase))
                                {
                                    serviceRequestCol = col;
                                }
                            }

                            if (offeringCodeCol != -1 && serviceRequestCol != -1)
                            {
                                
                                // นับและรวบรวมข้อมูล OfferingCode = N/A เท่านั้น
                                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    string offeringCode = worksheet.Cells[row, offeringCodeCol].Text.Trim();
                                    string serviceRequest = worksheet.Cells[row, serviceRequestCol].Text.Trim();
                                    
                                    if (!string.IsNullOrEmpty(serviceRequest))
                                    {
                                        // ตรวจสอบ OfferingCode = N/A
                                        if (offeringCode.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                                        {
                                            naOfferingCodeCases.Add(serviceRequest);
                                        }
                                    }
                                }
                                Console.WriteLine($"พบเคส OfferingCode = N/A: {naOfferingCodeCases.Count} เคส");
                            }
                            else
                            {
                                Console.WriteLine("ไม่พบคอลัมน์ OfferingCode1 หรือ ServiceRequest1 ในไฟล์ Batch");
                            }
                        }
                        else
                        {
                            Console.WriteLine("ไม่พบ worksheet ชื่อ 'InputFile' ในไฟล์ Batch หรือ worksheet ว่างเปล่า");
                            
                            // แสดงรายชื่อ worksheet ทั้งหมดเพื่อ debug
                            Console.WriteLine("รายชื่อ worksheet ในไฟล์:");
                            foreach (var ws in package.Workbook.Worksheets)
                            {
                                Console.WriteLine($"  - {ws.Name} (Hidden: {ws.Hidden})");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Batch file not found at '{batchFilePath}'. Cannot count marketing approved cases.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error counting marketing approved cases from batch file: {ex.Message}");
            }

            // ✅ 4. นับจำนวนเคสจากไฟล์ ManualCheck_NAOfferingCode.xlsx 
            try
            {
                if (!string.IsNullOrEmpty(naOfferingCodeFilePath) && File.Exists(naOfferingCodeFilePath))
                {
                    using (var package = new ExcelPackage(new FileInfo(naOfferingCodeFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        if (worksheet != null && worksheet.Dimension != null)
                        {
                            // ✅ นับเฉพาะแถวที่มี Case ID (ไม่นับ Summary section)
                            naOfferingCodeCasesFromReport = 0;
                            
                            // หาคอลัมน์ Case ID (คอลัมน์แรก)
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                string caseId = worksheet.Cells[row, 1].Text.Trim();
                                
                                // ถ้าเจอคำว่า "SUMMARY:" หรือแถวว่าง ให้หยุดนับ
                                if (string.IsNullOrEmpty(caseId) || 
                                    caseId.Equals("SUMMARY:", StringComparison.OrdinalIgnoreCase) ||
                                    caseId.StartsWith("Total cases", StringComparison.OrdinalIgnoreCase) ||
                                    caseId.StartsWith("Action:", StringComparison.OrdinalIgnoreCase) ||
                                    caseId.StartsWith("After updating", StringComparison.OrdinalIgnoreCase))
                                {
                                    break; // หยุดนับเมื่อเจอ Summary section
                                }
                                
                                naOfferingCodeCasesFromReport++;
                            }
                            
                            Console.WriteLine($"พบเคส N/A OfferingCode จากไฟล์รายงาน: {naOfferingCodeCasesFromReport} เคส");
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(naOfferingCodeFilePath))
                {
                    Console.WriteLine($"Warning: N/A OfferingCode report file not found at '{naOfferingCodeFilePath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error counting N/A OfferingCode cases from report file: {ex.Message}");
            }

            // 5. สร้างไฟล์ Excel สรุปผลลัพธ์ใหม่
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string summaryFileName = $"Summary_Report_{timestamp}.xlsx";
            string summaryFilePath = Path.Combine(outputDirectory, summaryFileName);

            try
            {
                using (var summaryPackage = new ExcelPackage())
                {
                    var summaryWorksheet = summaryPackage.Workbook.Worksheets.Add("Summary");

                    // Header
                    summaryWorksheet.Cells[1, 1].Value = "Category";
                    summaryWorksheet.Cells[1, 2].Value = "Count";

                    using (var range = summaryWorksheet.Cells[1, 1, 1, 2])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    int currentRow = 2;
                    
                    // ข้อมูลสรุป
                    summaryWorksheet.Cells[currentRow, 1].Value = "Total Promotion Prepaid Cases in Master File";
                    summaryWorksheet.Cells[currentRow, 2].Value = totalPromotionPrepaidCases;
                    currentRow++;

                    summaryWorksheet.Cells[currentRow, 1].Value = "Cases Completed (Import_CloseCaseIM)";
                    summaryWorksheet.Cells[currentRow, 2].Value = completedCases;
                    currentRow++;

                    summaryWorksheet.Cells[currentRow, 1].Value = "Cases Rejected (Import_CloseCaseIM)";
                    summaryWorksheet.Cells[currentRow, 2].Value = rejectedCases;
                    currentRow++;

                    // ✅ ปรับปรุง: ใช้ข้อมูลจากไฟล์รายงาน N/A OfferingCode แทน
                    int totalNACases = naOfferingCodeCasesFromReport > 0 ? naOfferingCodeCasesFromReport : naOfferingCodeCases.Count;
                    summaryWorksheet.Cells[currentRow, 1].Value = "Cases with N/A OfferingCode (Need Manual Check)";
                    summaryWorksheet.Cells[currentRow, 2].Value = totalNACases;
                    currentRow++;

                    // เพิ่มข้อมูลสรุปการตรวจสอบ
                    summaryWorksheet.Cells[currentRow, 1].Value = "Total Cases Processed";
                    summaryWorksheet.Cells[currentRow, 2].Value = completedCases + rejectedCases + totalNACases;
                    summaryWorksheet.Cells[currentRow, 2].Style.Font.Bold = true;
                    summaryWorksheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    summaryWorksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                    currentRow++;

                    // สีของข้อมูล
                    using (var range = summaryWorksheet.Cells[2, 1, currentRow - 1, 2])
                    {
                        range.Style.Font.Color.SetColor(Color.Green);
                    }

                    // เพิ่ม Worksheet สำหรับแสดงรายการเคสที่มี N/A OfferingCode
                    List<string> finalNACases = naOfferingCodeCases; // legacy จากไฟล์ batch
                    
                    // ✅ ถ้ามีไฟล์ ManualCheck_NAOfferingCode.xlsx ให้อ่านข้อมูลจากไฟล์นั้นแทน
                    if (!string.IsNullOrEmpty(naOfferingCodeFilePath) && File.Exists(naOfferingCodeFilePath))
                    {
                        finalNACases = new List<string>();
                        try
                        {
                            using (var naPackage = new ExcelPackage(new FileInfo(naOfferingCodeFilePath)))
                            {
                                var naWorksheet = naPackage.Workbook.Worksheets.FirstOrDefault();
                                if (naWorksheet != null && naWorksheet.Dimension != null)
                                {
                                    // ✅ อ่าน Case ID จากคอลัมน์แรก (ข้าม header และ Summary section)
                                    for (int row = 2; row <= naWorksheet.Dimension.End.Row; row++)
                                    {
                                        string caseId = naWorksheet.Cells[row, 1].Text.Trim();
                                        
                                        // ถ้าเจอคำว่า "SUMMARY:" หรือแถวว่าง ให้หยุดอ่าน
                                        if (string.IsNullOrEmpty(caseId) || 
                                            caseId.Equals("SUMMARY:", StringComparison.OrdinalIgnoreCase) ||
                                            caseId.StartsWith("Total cases", StringComparison.OrdinalIgnoreCase) ||
                                            caseId.StartsWith("Action:", StringComparison.OrdinalIgnoreCase) ||
                                            caseId.StartsWith("After updating", StringComparison.OrdinalIgnoreCase))
                                        {
                                            break; // หยุดอ่านเมื่อเจอ Summary section
                                        }
                                        
                                        finalNACases.Add(caseId);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not read N/A OfferingCode report file: {ex.Message}");
                            finalNACases = naOfferingCodeCases; // fallback ให้ใช้ข้อมูลเดิม
                        }
                    }
                    
                    if (finalNACases.Count > 0)
                    {
                        var naWorksheet = summaryPackage.Workbook.Worksheets.Add("N/A OfferingCode Cases");
                        
                        // Header
                        naWorksheet.Cells[1, 1].Value = "Case ID";
                        naWorksheet.Cells[1, 2].Value = "Status";
                        naWorksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        naWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                        naWorksheet.Cells[1, 1].Style.Font.Bold = true;
                        naWorksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        naWorksheet.Cells[1, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        naWorksheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                        naWorksheet.Cells[1, 2].Style.Font.Bold = true;
                        naWorksheet.Cells[1, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        // รายการเคส
                        int naRow = 2;
                        foreach (string caseId in finalNACases)
                        {
                            naWorksheet.Cells[naRow, 1].Value = caseId;
                            naWorksheet.Cells[naRow, 1].Style.Font.Color.SetColor(Color.Red);
                            naWorksheet.Cells[naRow, 2].Value = "Requires Doc# mapping verification";
                            naWorksheet.Cells[naRow, 2].Style.Font.Color.SetColor(Color.DarkOrange);
                            naRow++;
                        }
                        
                        naWorksheet.Cells[naWorksheet.Dimension.Address].AutoFitColumns();
                    }

                    summaryWorksheet.Cells[summaryWorksheet.Dimension.Address].AutoFitColumns();

                    summaryPackage.SaveAs(new FileInfo(summaryFilePath));
                    
                    if (finalNACases.Count > 0)
                    {
                        Console.WriteLine($"พบเคสที่ต้องตรวจสอบ OfferingCode เพิ่มเติม: {finalNACases.Count} เคส");
                        Console.WriteLine($"📋 รายการ Case ID ที่ต้องตรวจสอบ Doc# mapping:");
                        foreach (string caseId in finalNACases.Take(10)) // แสดงแค่ 10 เคสแรก
                        {
                            Console.WriteLine($"   - {caseId}");
                        }
                        if (finalNACases.Count > 10)
                        {
                            Console.WriteLine($"   ... และอีก {finalNACases.Count - 10} เคส");
                        }
                    }
                    
                    return summaryFilePath;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating summary Excel file: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return null;
            }
        }


        /// <summary>
        /// สร้างไฟล์ Import_CloseCaseIM จากไฟล์ที่มีข้อมูล Case ID ในคอลัมน์ AI, AJ, AK
        /// </summary>
        /// <param name="sourceFilePath">พาธของไฟล์ต้นทาง (ไฟล์ที่มี Case ID)</param>
        /// <param name="targetFilePath">พาธของไฟล์ปลายทาง (template Import_CloseCaseIM)</param>
        /// <param name="caseStatus">ค่าที่ต้องการใส่ในคอลัมน์ Case Status</param>
        /// <param name="comment">ค่าที่ต้องการใส่ในคอลัมน์ Comment</param>
        /// <param name="sourceSheetName">ชื่อ sheet ต้นทาง (null = sheet แรก)</param>
        /// <param name="targetSheetName">ชื่อ sheet ปลายทาง (null = sheet แรก)</param>
        /// <param name="startFromRow">เริ่มคัดลอกจากแถวที่ (1-based, default=2 เพื่อข้าม header)</param>
        /// <returns>true ถ้าสำเร็จ, false ถ้าไม่สำเร็จ</returns>
        public bool CreateImportCloseCaseIM(
            string sourceFilePath,
            string targetFilePath,
            string caseStatus,
            string comment,
            string sourceSheetName = null,
            string targetSheetName = null,
            int startFromRow = 2)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // ตรวจสอบไฟล์
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"ไม่พบไฟล์ต้นทาง: {sourceFilePath}");
                    return false;
                }

                if (!File.Exists(targetFilePath))
                {
                    Console.WriteLine($"ไม่พบไฟล์ปลายทาง: {targetFilePath}");
                    return false;
                }

                using (var sourcePackage = new ExcelPackage(new FileInfo(sourceFilePath)))
                using (var targetPackage = new ExcelPackage(new FileInfo(targetFilePath)))
                {
                    // เลือก worksheet ต้นทาง
                    var sourceWorksheet = string.IsNullOrEmpty(sourceSheetName)
                        ? sourcePackage.Workbook.Worksheets.FirstOrDefault()
                        : sourcePackage.Workbook.Worksheets[sourceSheetName];

                    // เลือก worksheet ปลายทาง
                    var targetWorksheet = string.IsNullOrEmpty(targetSheetName)
                        ? targetPackage.Workbook.Worksheets.FirstOrDefault()
                        : targetPackage.Workbook.Worksheets[targetSheetName];

                    if (sourceWorksheet == null || targetWorksheet == null)
                    {
                        Console.WriteLine("ไม่พบ worksheet ที่ระบุ");
                        return false;
                    }

                    // กำหนด column index สำหรับไฟล์ต้นทาง (AI=35, AJ=36, AK=37)
                    int[] caseIdColumns = { 35, 36, 37 };

                    // ค้นหา column index ของไฟล์ปลายทาง
                    int targetCaseIdCol = -1;
                    int targetTopicCol = -1;
                    int targetCaseStatusCol = -1;
                    int targetCommentCol = -1;

                    // หา header ในไฟล์ปลายทาง
                    for (int col = 1; col <= targetWorksheet.Dimension?.Columns; col++)
                    {
                        var cellValue = targetWorksheet.Cells[1, col].Value?.ToString()?.Trim();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        if (cellValue.Equals("Case ID", StringComparison.OrdinalIgnoreCase))
                            targetCaseIdCol = col;
                        else if (cellValue.Equals("Topic", StringComparison.OrdinalIgnoreCase))
                            targetTopicCol = col;
                        else if (cellValue.Equals("Case Status", StringComparison.OrdinalIgnoreCase))
                            targetCaseStatusCol = col;
                        else if (cellValue.Equals("Comment", StringComparison.OrdinalIgnoreCase))
                            targetCommentCol = col;
                    }

                    // ตรวจสอบว่าพบคอลัมน์ที่จำเป็นทั้งหมดหรือไม่
                    if (targetCaseIdCol == -1 || targetTopicCol == -1 || targetCaseStatusCol == -1 || targetCommentCol == -1)
                    {
                        Console.WriteLine("ไม่พบคอลัมน์ที่จำเป็นในไฟล์ปลายทาง (Case ID, Topic, Case Status, Comment)");
                        return false;
                    }

                    Console.WriteLine($"พบคอลัมน์: Case ID={targetCaseIdCol}, Topic={targetTopicCol}, Case Status={targetCaseStatusCol}, Comment={targetCommentCol}");

                    // หาแถวว่างแรกในไฟล์ปลายทาง (ไฟล์ปลายทางไม่มี header)
                    int targetStartRow = 1; // เริ่มต้นที่แถวที่ 1
                    if (targetWorksheet.Dimension != null)
                    {
                        // หาแถวสุดท้ายที่มีข้อมูลจริง ๆ แล้วเขียนต่อจากแถวถัดไป
                        int lastUsedRow = GetLastUsedRow(targetWorksheet);
                        targetStartRow = lastUsedRow + 1;
                        Console.WriteLine($"พบข้อมูลเก่าในไฟล์ปลายทางจนถึงแถวที่ {lastUsedRow}, จะเขียนต่อจากแถวที่ {targetStartRow}");
                    }
                    else
                    {
                        Console.WriteLine("ไฟล์ปลายทางว่างเปล่า จะเขียนข้อมูลใหม่จากแถวที่ 1");
                    }

                    // คัดลอกข้อมูล
                    int copiedRows = 0;
                    int sourceMaxRow = sourceWorksheet.Dimension?.Rows ?? 0;

                    // ปรับ startFromRow สำหรับไฟล์ต้นทางที่ไม่มี header
                    int actualStartRow = 1; // เริ่มจากแถวที่ 1 เสมอ (ไฟล์ต้นทางไม่มี header)
                    
                    Console.WriteLine($"เริ่มประมวลผลข้อมูลจากแถวที่ {actualStartRow} ถึงแถวที่ {sourceMaxRow} (ไฟล์ต้นทางไม่มี header)");

                    for (int sourceRow = actualStartRow; sourceRow <= sourceMaxRow; sourceRow++)
                    {
                        // หา Case ID จากคอลัมน์ AI, AJ, AK (เลือกคอลัมน์แรกที่มีข้อมูล)
                        string caseId = null;
                        foreach (int colIndex in caseIdColumns)
                        {
                            var cellValue = sourceWorksheet.Cells[sourceRow, colIndex].Value?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                caseId = cellValue;
                                break; // ใช้ค่าแรกที่เจอ
                            }
                        }

                        // ข้ามแถวที่ไม่มี Case ID
                        if (string.IsNullOrEmpty(caseId))
                        {
                            continue;
                        }

                        // เขียนข้อมูลลงไฟล์ปลายทาง
                        targetWorksheet.Cells[targetStartRow, targetCaseIdCol].Value = caseId;
                        targetWorksheet.Cells[targetStartRow, targetCaseIdCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        targetWorksheet.Cells[targetStartRow, targetTopicCol].Value = "Service Request";
                        targetWorksheet.Cells[targetStartRow, targetTopicCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        targetWorksheet.Cells[targetStartRow, targetCaseStatusCol].Value = caseStatus;
                        targetWorksheet.Cells[targetStartRow, targetCaseStatusCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        targetWorksheet.Cells[targetStartRow, targetCommentCol].Value = comment;
                        targetWorksheet.Cells[targetStartRow, targetCommentCol].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                        //Console.WriteLine($"ประมวลผล Case ID: {caseId} (แถวที่ {sourceRow} -> {targetStartRow})");

                        targetStartRow++;
                        copiedRows++;
                    }

                    // ปรับขนาดคอลัมน์อัตโนมัติ
                    targetWorksheet.Cells[targetWorksheet.Dimension.Address].AutoFitColumns();

                    // บันทึกไฟล์
                    targetPackage.Save();

                    Console.WriteLine($"สร้างไฟล์ Import_CloseCaseIM สำเร็จ: {copiedRows} แถว");
                    Console.WriteLine($"บันทึกที่: {targetFilePath}");
                    
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการสร้างไฟล์ Import_CloseCaseIM: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return false;
            }
        }

        /// <summary>
        /// สร้างไฟล์สรุปรายเดือนจากไฟล์ Summary Reports ทั้งหมดในโฟลเดอร์
        /// </summary>
        /// <param name="summaryReportsDirectory">โฟลเดอร์ที่มีไฟล์ Summary Reports</param>
        /// <param name="outputDirectory">โฟลเดอร์สำหรับบันทึกไฟล์สรุปรายเดือน</param>
        /// <param name="month">เดือนที่ต้องการสรุป (1-12) หรือ 0 สำหรับเดือนปัจจุบัน</param>
        /// <param name="year">ปีที่ต้องการสรุป หรือ 0 สำหรับปีปัจจุบัน</param>
        /// <returns>พาธเต็มของไฟล์สรุปรายเดือนที่สร้างขึ้น หากสำเร็จ; Null หากเกิดข้อผิดพลาด</returns>
        public string CreateMonthlySummary(string summaryReportsDirectory, string outputDirectory, int month = 0, int year = 0)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // กำหนดเดือนและปีที่ต้องการสรุป
            DateTime targetDate = DateTime.Now;
            if (year > 0) targetDate = new DateTime(year, month > 0 ? month : targetDate.Month, 1);
            else if (month > 0) targetDate = new DateTime(targetDate.Year, month, 1);

            string monthName = targetDate.ToString("MMMM yyyy", new System.Globalization.CultureInfo("th-TH"));
            
            try
            {
                // ตรวจสอบโฟลเดอร์
                if (!Directory.Exists(summaryReportsDirectory))
                {
                    Console.WriteLine($"ไม่พบโฟลเดอร์ Summary Reports: {summaryReportsDirectory}");
                    return null;
                }

                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                    Console.WriteLine($"สร้างโฟลเดอร์ Output: {outputDirectory}");
                }

                // ค้นหาไฟล์ Summary_Report ในโฟลเดอร์ที่มีวันที่ตรงกับเดือน/ปีที่ต้องการ
                string searchPattern = $"Summary_Report_{targetDate:yyyyMM}*.xlsx";
                var summaryFiles = Directory.GetFiles(summaryReportsDirectory, searchPattern, SearchOption.AllDirectories)
                                           .OrderBy(f => f)
                                           .ToList();

                if (!summaryFiles.Any())
                {
                    Console.WriteLine($"ไม่พบไฟล์ Summary Report สำหรับ {monthName} ในโฟลเดอร์ {summaryReportsDirectory}");
                    return null;
                }

                Console.WriteLine($"พบไฟล์ Summary Report {summaryFiles.Count} ไฟล์ สำหรับ {monthName}");

                // ตัวแปรสำหรับเก็บข้อมูลรวม
                int totalPromotionPrepaidCases = 0;
                int totalCompletedCases = 0;
                int totalRejectedCases = 0;
                int totalNAOfferingCodeCases = 0;
                
                List<DailyReport> dailyReports = new List<DailyReport>();

                // อ่านข้อมูลจากแต่ละไฟล์ Summary
                foreach (string filePath in summaryFiles)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);
                        
                        // แยกวันที่จากชื่อไฟล์ (Summary_Report_yyyyMMdd_HHmmss)
                        var parts = fileName.Split('_');
                        if (parts.Length >= 3)
                        {
                            string dateStr = parts[2]; // yyyyMMdd
                            if (DateTime.TryParseExact(dateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                            {
                                if (fileDate.Year != targetDate.Year || fileDate.Month != targetDate.Month)
                                    continue; // ข้ามไฟล์ที่ไม่ใช่เดือน/ปีที่ต้องการ
                            }
                        }

                        using (var package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            var summarySheet = package.Workbook.Worksheets["Summary"];
                            if (summarySheet != null && summarySheet.Dimension != null)
                            {
                                var dailyReport = new DailyReport
                                {
                                    FileName = fileName,
                                    FilePath = filePath,
                                    Date = Path.GetFileName(filePath)
                                };

                                // อ่านข้อมูลจากคอลัมน์ B (Count)
                                for (int row = 2; row <= summarySheet.Dimension.End.Row; row++)
                                {
                                    string category = summarySheet.Cells[row, 1].Text?.Trim() ?? "";
                                    string countText = summarySheet.Cells[row, 2].Text?.Trim() ?? "0";
                                    
                                    if (int.TryParse(countText, out int count))
                                    {
                                        if (category.Contains("Total Promotion Prepaid Cases"))
                                        {
                                            dailyReport.PromotionPrepaidCases = count;
                                            totalPromotionPrepaidCases += count;
                                        }
                                        else if (category.Contains("Cases Completed"))
                                        {
                                            dailyReport.CompletedCases = count;
                                            totalCompletedCases += count;
                                        }
                                        else if (category.Contains("Cases Rejected"))
                                        {
                                            dailyReport.RejectedCases = count;
                                            totalRejectedCases += count;
                                        }
                                        else if (category.Contains("Cases with N/A OfferingCode"))
                                        {
                                            dailyReport.NAOfferingCodeCases = count;
                                            totalNAOfferingCodeCases += count;
                                        }
                                    }
                                }

                                dailyReports.Add(dailyReport);
                                Console.WriteLine($"อ่านข้อมูลจากไฟล์: {fileName}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"ข้อผิดพลาดในการอ่านไฟล์ {filePath}: {ex.Message}");
                    }
                }

                if (!dailyReports.Any())
                {
                    Console.WriteLine($"ไม่มีข้อมูลที่ถูกต้องสำหรับ {monthName}");
                    return null;
                }

                // สร้างไฟล์สรุปรายเดือน
                string monthlyFileName = $"Monthly_Summary_{targetDate:yyyyMM}.xlsx";
                string monthlyFilePath = Path.Combine(outputDirectory, monthlyFileName);

                // ตรวจสอบว่ามีไฟล์เดิมอยู่หรือไม่
                bool isUpdate = File.Exists(monthlyFilePath);
                if (isUpdate)
                {
                    Console.WriteLine($"พบไฟล์สรุปรายเดือนเดิม จะทำการอัพเดต: {monthlyFileName}");
                }
                else
                {
                    Console.WriteLine($"สร้างไฟล์สรุปรายเดือนใหม่: {monthlyFileName}");
                }

                using (var package = isUpdate ? new ExcelPackage(new FileInfo(monthlyFilePath)) : new ExcelPackage())
                {
                    // ถ้ามีไฟล์เดิมอยู่ ลบ sheets เดิมเพื่อสร้างใหม่ด้วยข้อมูลล่าสุด
                    if (isUpdate)
                    {
                        var existingSummarySheet = package.Workbook.Worksheets["Monthly Summary"];
                        var existingDailySheet = package.Workbook.Worksheets["Daily Details"];
                        
                        if (existingSummarySheet != null)
                            package.Workbook.Worksheets.Delete(existingSummarySheet);
                        if (existingDailySheet != null)
                            package.Workbook.Worksheets.Delete(existingDailySheet);
                    }
                    
                    // Sheet 1: สรุปรวม
                    var summarySheet = package.Workbook.Worksheets.Add("Monthly Summary");
                    
                    // Header
                    summarySheet.Cells[1, 1].Value = $"สรุปรายเดือน {monthName} (อัพเดต: {DateTime.Now:dd/MM/yyyy HH:mm})";
                    summarySheet.Cells[1, 1, 1, 3].Merge = true;
                    using (var range = summarySheet.Cells[1, 1, 1, 3])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.Navy);
                        range.Style.Font.Color.SetColor(Color.White);
                        range.Style.Font.Bold = true;
                        range.Style.Font.Size = 16;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    // ข้อมูลสรุป
                    int currentRow = 3;
                    summarySheet.Cells[currentRow, 1].Value = "หมวดหมู่";
                    summarySheet.Cells[currentRow, 2].Value = "จำนวนรวม";
                    summarySheet.Cells[currentRow, 3].Value = "จำนวนรายงาน";
                    
                    using (var range = summarySheet.Cells[currentRow, 1, currentRow, 3])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    currentRow++;

                    // รายละเอียดสรุป
                    summarySheet.Cells[currentRow, 1].Value = "Total Promotion Prepaid Cases";
                    summarySheet.Cells[currentRow, 2].Value = totalPromotionPrepaidCases;
                    summarySheet.Cells[currentRow, 3].Value = dailyReports.Count;
                    currentRow++;

                    summarySheet.Cells[currentRow, 1].Value = "Cases Completed";
                    summarySheet.Cells[currentRow, 2].Value = totalCompletedCases;
                    summarySheet.Cells[currentRow, 3].Value = dailyReports.Count;
                    currentRow++;

                    summarySheet.Cells[currentRow, 1].Value = "Cases Rejected";
                    summarySheet.Cells[currentRow, 2].Value = totalRejectedCases;
                    summarySheet.Cells[currentRow, 3].Value = dailyReports.Count;
                    currentRow++;

                    summarySheet.Cells[currentRow, 1].Value = "Cases with N/A OfferingCode";
                    summarySheet.Cells[currentRow, 2].Value = totalNAOfferingCodeCases;
                    summarySheet.Cells[currentRow, 3].Value = dailyReports.Count;
                    currentRow++;

                    // Sheet 2: รายละเอียดรายวัน
                    var dailySheet = package.Workbook.Worksheets.Add("Daily Details");
                    
                    // Header สำหรับ Daily Details
                    dailySheet.Cells[1, 1].Value = "วันที่ (ไฟล์)";
                    dailySheet.Cells[1, 2].Value = "Promotion Prepaid";
                    dailySheet.Cells[1, 3].Value = "Completed";
                    dailySheet.Cells[1, 4].Value = "Rejected";
                    dailySheet.Cells[1, 6].Value = "N/A OfferingCode";
                    dailySheet.Cells[1, 7].Value = "ชื่อไฟล์";

                    using (var range = dailySheet.Cells[1, 1, 1, 7])
                    {
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    // ข้อมูลรายวัน
                    int dailyRow = 2;
                    foreach (var report in dailyReports.OrderBy(r => r.Date))
                    {
                        dailySheet.Cells[dailyRow, 1].Value = report.Date;
                        dailySheet.Cells[dailyRow, 2].Value = report.PromotionPrepaidCases;
                        dailySheet.Cells[dailyRow, 3].Value = report.CompletedCases;
                        dailySheet.Cells[dailyRow, 4].Value = report.RejectedCases;
                        dailySheet.Cells[dailyRow, 6].Value = report.NAOfferingCodeCases;
                        dailySheet.Cells[dailyRow, 7].Value = report.FileName;
                        dailyRow++;
                    }

                    // AutoFit columns
                    summarySheet.Cells[summarySheet.Dimension.Address].AutoFitColumns();
                    dailySheet.Cells[dailySheet.Dimension.Address].AutoFitColumns();

                    package.SaveAs(new FileInfo(monthlyFilePath));
                }

                Console.WriteLine($"\n{(isUpdate ? "อัพเดต" : "สร้าง")}ไฟล์สรุปรายเดือนสำเร็จ: {monthlyFilePath}");
                Console.WriteLine($"ข้อมูลจาก {dailyReports.Count} รายงาน ในเดือน {monthName}");
                Console.WriteLine($"รวม Promotion Prepaid Cases: {totalPromotionPrepaidCases} เคส");
                Console.WriteLine($"รวม Completed Cases: {totalCompletedCases} เคส");
                Console.WriteLine($"รวม Rejected Cases: {totalRejectedCases} เคส");
                Console.WriteLine($"รวม N/A OfferingCode Cases: {totalNAOfferingCodeCases} เคส");

                return monthlyFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ข้อผิดพลาดในการสร้างไฟล์สรุปรายเดือน: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return null;
            }
        }

        /// <summary>
        /// สร้างไฟล์รายงานสำหรับเคสที่มี OfferingCode = N/A เพื่อให้ user ตรวจสอบเอง
        /// </summary>
        /// <param name="naOfferingCases">รายการเคสที่มี OfferingCode = N/A</param>
        /// <param name="outputPath">พาธของไฟล์รายงานที่จะสร้าง</param>
        private void CreateNAOfferingCodeReport(List<MasterExcelRow> naOfferingCases, string outputPath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Manual Check Required");
                    
                    // สร้าง Header
                    worksheet.Cells[1, 1].Value = "Case ID";
                    worksheet.Cells[1, 2].Value = "Topic";
                    worksheet.Cells[1, 3].Value = "Reason";
                    worksheet.Cells[1, 4].Value = "Doc#";
                    worksheet.Cells[1, 5].Value = "Created By";
                    worksheet.Cells[1, 6].Value = "Mobile";
                    worksheet.Cells[1, 7].Value = "Status";
                    worksheet.Cells[1, 8].Value = "Action Required";
                    
                    // จัดรูปแบบ Header
                    using (var range = worksheet.Cells[1, 1, 1, 8])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                    }
                    
                    // เขียนข้อมูล
                    int row = 2;
                    foreach (var caseData in naOfferingCases)
                    {
                        worksheet.Cells[row, 1].Value = caseData.CaseId;
                        worksheet.Cells[row, 2].Value = caseData.Topic;
                        worksheet.Cells[row, 3].Value = caseData.Reason;
                        worksheet.Cells[row, 4].Value = caseData.Doc;
                        worksheet.Cells[row, 5].Value = caseData.CreatedBy;
                        worksheet.Cells[row, 6].Value = caseData.Mobile;
                        worksheet.Cells[row, 7].Value = "Marketing Approved - Missing OfferingCode";
                        worksheet.Cells[row, 8].Value = "Please verify Doc# and update mapping file";
                        
                        // ตั้งสีข้อความเป็นสีส้มเพื่อแสดงว่าต้องตรวจสอบ
                        using (var rowRange = worksheet.Cells[row, 1, row, 8])
                        {
                            rowRange.Style.Font.Color.SetColor(System.Drawing.Color.DarkOrange);
                        }
                        
                        row++;
                    }
                    
                    // เพิ่มข้อมูลสรุปที่ด้านล่าง
                    row += 2;
                    worksheet.Cells[row, 1].Value = "SUMMARY:";
                    worksheet.Cells[row, 1].Style.Font.Bold = true;
                    
                    row++;
                    worksheet.Cells[row, 1].Value = $"Total cases requiring manual check: {naOfferingCases.Count}";
                    worksheet.Cells[row, 1].Style.Font.Bold = true;
                    worksheet.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                    
                    row++;
                    worksheet.Cells[row, 1].Value = "Action: Please check Doc# in mapping file and update accordingly";
                    
                    row++;
                    worksheet.Cells[row, 1].Value = "After updating mapping file, re-run the process to include these cases";
                    
                    // ปรับขนาดคอลัมน์อัตโนมัติ
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    
                    // บันทึกไฟล์
                    package.SaveAs(new FileInfo(outputPath));
                }
                
                Console.WriteLine($"✅ Successfully created N/A OfferingCode report: '{outputPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error creating N/A OfferingCode report: {ex.Message}");
            }
        }

        /// <summary>
        /// คลาสสำหรับเก็บข้อมูลรายงานรายวัน
        /// </summary>
        private class DailyReport
        {
            public string FileName { get; set; }
            public string FilePath { get; set; }
            public string Date { get; set; }
            public int PromotionPrepaidCases { get; set; }
            public int CompletedCases { get; set; }
            public int RejectedCases { get; set; }
            public int NAOfferingCodeCases { get; set; }
        }

    }
}
