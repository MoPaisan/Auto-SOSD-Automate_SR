using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using AutoIt;
using System.IO;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Globalization;
using OfficeOpenXml;



namespace SOSD_AutomateSR
{
    internal class WebManager
    {
        private readonly int defaultShortWait = 1000;
        private readonly int defaultLongWait = 3000;
        private readonly int initialLoadWait = 6000;


        public IWebDriver WebMyOffice(string URLMyOffice, string UserMyOffice, string passMyOffice, string PathProfile)
        {
            IWebDriver driverMO = null;

            try
            {
                // ตรวจสอบว่า URL ไม่ว่างเปล่า
                if (string.IsNullOrEmpty(URLMyOffice))
                {
                    throw new ArgumentException("URLMyOffice cannot be null or empty.");
                }

                // สร้าง ChromeDriverService
                var driverService = ChromeDriverService.CreateDefaultService();
                driverService.SuppressInitialDiagnosticInformation = true; // ปิดข้อความเริ่มต้น
                driverService.HideCommandPromptWindow = true; // ซ่อนหน้าต่าง console ของ ChromeDriver

                // สร้าง ChromeOptions และตั้งค่าต่าง ๆ
                ChromeOptions options = new ChromeOptions();

                // ตั้งค่า Profile
                options.AddArgument($"--user-data-dir={PathProfile}");
                // ตั้งค่าที่เกี่ยวกับ certificate
                options.AddArgument("--ignore-certificate-errors");
                options.AddArgument("--allow-insecure-localhost");

                // สำหรับบางองค์กรอาจต้องเปิดใช้ certificate จาก Windows Root Store
                //options.AddArgument("--allow-running-insecure-content");

                // สร้าง ChromeDriver ด้วย Service + Options
                driverMO = new ChromeDriver(driverService, options);

                // Navigate to URL
                driverMO.Navigate().GoToUrl(URLMyOffice);
                Thread.Sleep(initialLoadWait);

                // Handle login
                //HandleLoginMyoffice(driverMO, UserMyOffice, passMyOffice);
                HandleLoginWithEmailAuth(driverMO, UserMyOffice, passMyOffice);
            }
            catch (Exception ex)
            {
                HandleLoginException(ex, UserMyOffice, passMyOffice, driverMO);
            }

            return driverMO;
        }

        public IWebDriver WebIM(string URLIM, string UserIM, string passIM, string PathProfile)
        {
            IWebDriver driverMO = null;

            try
            {
                // ตรวจสอบว่า URL ไม่ว่างเปล่า
                if (string.IsNullOrEmpty(URLIM))
                {
                    throw new ArgumentException("URLMyOffice cannot be null or empty.");
                }

                // สร้าง ChromeDriverService
                var driverService = ChromeDriverService.CreateDefaultService();
                driverService.SuppressInitialDiagnosticInformation = true; // ปิดข้อความเริ่มต้น
                driverService.HideCommandPromptWindow = true; // ซ่อนหน้าต่าง console ของ ChromeDriver

                // สร้าง ChromeOptions และตั้งค่าต่าง ๆ
                ChromeOptions options = new ChromeOptions();

                // ตั้งค่า Profile
                options.AddArgument($"--user-data-dir={PathProfile}");
                // ตั้งค่าที่เกี่ยวกับ certificate
                options.AddArgument("--ignore-certificate-errors");
                options.AddArgument("--allow-insecure-localhost");

                // สำหรับบางองค์กรอาจต้องเปิดใช้ certificate จาก Windows Root Store
                options.AddArgument("--allow-running-insecure-content");

                // สร้าง ChromeDriver ด้วย Service + Options
                driverMO = new ChromeDriver(driverService, options);

                // Navigate to URL
                driverMO.Navigate().GoToUrl(URLIM);
                Thread.Sleep(initialLoadWait);

                // Handle login
                HandleLoginIM(driverMO, UserIM, passIM);
            }
            catch (Exception ex)
            {
                HandleLoginException(ex, UserIM, passIM, driverMO);
            }

            return driverMO;
        }

        private void HandleLoginIM(IWebDriver driver, string username, string password)
        {
            try
            {
                IList<IWebElement> FindTextboxuserName = driver.FindElements(By.Id("txtUserID"));

                if (FindTextboxuserName.Any())
                {
                    // Web form login
                    IWebElement TextboxUsername = driver.FindElement(By.XPath("//input[@id='txtUserID']"));
                    IWebElement TextboxPassword = driver.FindElement(By.XPath("//input[@id='txtPassword']"));
                    IWebElement BtnLogon = driver.FindElement(By.XPath("//button[@id='sub']"));

                    Thread.Sleep(defaultShortWait);
                    TextboxUsername.SendKeys(username);
                    Thread.Sleep(defaultShortWait);
                    TextboxPassword.SendKeys(password);
                    Thread.Sleep(defaultShortWait);
                    BtnLogon.Click();
                    Thread.Sleep(defaultLongWait);
                }
                else
                {
                    // AutoIt login
                    HandleAutoItLogin(username, password);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during login: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }
        private void HandleLoginMyoffice(IWebDriver driver, string username, string password)
        {
            try
            {
                IList<IWebElement> FindTextboxuserName = driver.FindElements(By.Id("username"));

                if (FindTextboxuserName.Any())
                {
                    // Web form login
                    IWebElement TextboxUsername = driver.FindElement(By.XPath("//input[@id='username']"));
                    IWebElement TextboxPassword = driver.FindElement(By.XPath("//input[@id='password']"));
                    IWebElement BtnLogon = driver.FindElement(By.XPath("//button[@type=\"submit\" and text()='Log in']"));

                    Thread.Sleep(defaultShortWait);
                    TextboxUsername.SendKeys(username);
                    Thread.Sleep(defaultShortWait);
                    TextboxPassword.SendKeys(password);
                    Thread.Sleep(defaultShortWait);
                    BtnLogon.Click();
                    Thread.Sleep(defaultLongWait);
                }
                else
                {
                    // AutoIt login
                    HandleAutoItLogin(username, password);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during login: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }
        private void HandleAutoItLogin(string username, string password)
        {
            try
            {
                AutoItX.Send(username);
                Thread.Sleep(defaultShortWait);
                AutoItX.Send("{TAB}");
                Thread.Sleep(defaultShortWait);
                AutoItX.Send(password.Replace("#", "{#}").Replace("!", "{!}"));
                Thread.Sleep(defaultShortWait);
                AutoItX.Send("{ENTER}");
                Thread.Sleep(defaultLongWait);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during AutoIt login: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }
        private void HandleLoginException(Exception ex, string username, string password, IWebDriver driver)
        {
            try
            {
                if (ex.Message.Contains("This site is asking you to sign in"))
                {
                    HandleAutoItLogin(username, password);
                }
                else
                {
                    Console.WriteLine($"Error: {ex.Message}");
                    if (driver != null)
                    {
                        driver.Quit();
                    }
                    Environment.Exit(0);
                }
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Error handling login exception: {innerEx.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }

        public string[] ReadTextFileAsArray(string Path)
        {
            string[] LineArray = { "" };
            try
            {
                LineArray = File.ReadAllLines(Path);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Environment.Exit(0);
            }
            return LineArray;
        }

        /// <summary>
        /// รับค่าการตั้งค่าจากไฟล์ config โดยแยกค่าด้วย => และตัด whitespace
        /// </summary>
        public string GetConfig(string[] ValueConfig, string ValueRequire)
        {
            string ConfigVal = "";
            try
            {
                foreach (string line in ValueConfig)
                {
                    // แยกค่าด้วย => และตัด whitespace
                    string[] parts = line.Split(new[] { "=>" }, StringSplitOptions.None);
                    if (parts.Length == 2)
                    {
                        string key = parts[0].Trim();
                        if (key == ValueRequire)
                        {
                            ConfigVal = parts[1].Trim();
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading config: {ex.Message}");
                Environment.Exit(0);
            }
            return ConfigVal;
        }


        /// <summary>
        /// เลือกเมนู "Order" ในระบบ Myofiice
        /// </summary>
        public void SelsectMenuOrderMyofiice(IWebDriver driver)
        {
            try
            {

                IWebElement menuOrder = driver.FindElement(By.XPath($"//a[@class='text-lg font-medium font-dbHelvethaica'][normalize-space()='Order']"));
                menuOrder.Click();
                Thread.Sleep(defaultShortWait);

                IWebElement menuNasMyOffice = driver.FindElement(By.XPath("//button/a[normalize-space()='Batch Order']"));
                menuNasMyOffice.Click();
                Thread.Sleep(defaultShortWait);

                IWebElement menuNasMyOfficeSecond = driver.FindElement(By.XPath("//li[@class='hover:bg-[#7a7a7a0e] pl-5 rounded-lg , bg-[#7a7a7a0e] text-text-hover']//a[@class='text-lg font-medium font-dbHelvethaica truncate pr-2'][normalize-space()='Batch Order']"));
                menuNasMyOfficeSecond.Click();
                Thread.Sleep(defaultShortWait);

                IWebElement menuNasMyOfficeThird = driver.FindElement(By.XPath("//li[@class='hover:bg-[#7a7a7a0e] pl-5 rounded-lg , bg-[#7a7a7a0e] text-text-hover']//a[@class='truncate py-1 text-[18px] font-medium font-dbHelvethaica cursor-pointer text-text-secondary-color hover:text-text-hover w-full'][normalize-space()='Batch Order']"));
                menuNasMyOfficeThird.Click();
                Thread.Sleep(initialLoadWait);


            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Selsect Menu Orders: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }

        /// <summary>
        /// เลือกเมนู "Batch Monitoring" และประมวลผลไฟล์ในระบบ Myofiice (รวม 2 methods เดิม)
        /// ปรับปรุงเพิ่มการจัดการ error และ retry mechanism
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        /// <param name="outputPath">โฟลเดอร์สำหรับเก็บไฟล์ผลลัพธ์</param>
        /// <param name="foundFileName">ชื่อไฟล์ที่ต้องการค้นหา</param>
        public void SelectMonitoringAndProcessFile(IWebDriver driver, string outputPath, string foundFileName)
        {
            string originalWindowHandle = driver.CurrentWindowHandle;
            int maxRetries = 3;
            int currentRetry = 0;

            try
            {
                // กำหนดเวลาที่ต้องการรอในหน่วยมิลลิวินาที (10 นาที = 600,000 ms)
                Console.WriteLine("Process batch order กำลังประมวลผล กรุณารอประมาณ 10 นาที (~ 600 วินาที)");
                Console.WriteLine("หากต้องการข้ามการรออัตโนมัติ ให้กดปุ่มใดๆ ที่คีย์บอร์ด...");
                int waitTimeInMilliseconds = 600000;
                int checkInterval = 100; // ตรวจสอบทุกๆ 100 มิลลิวินาที

                for (int i = 0; i < waitTimeInMilliseconds; i += checkInterval)
                {
                    if (Console.KeyAvailable)
                    {
                        Console.ReadKey(true); // อ่านการกดปุ่ม (true = ไม่แสดงปุ่มที่กด)
                        Console.WriteLine("\nตรวจพบการกดปุ่ม! ข้ามการรออัตโนมัติแล้ว");
                        break; // ออกจาก loop
                    }
                    Thread.Sleep(checkInterval);
                }

                while (currentRetry < maxRetries)
                {
                    try
                    {
                        currentRetry++;
                        Console.WriteLine($"🔄 การลองครั้งที่ {currentRetry}/{maxRetries}: เริ่มเลือกเมนู Batch Monitoring และประมวลผลไฟล์...");

                        // ตรวจสอบสถานะหน้าเว็บก่อนเริ่ม
                        CheckPageState(driver);

                        // ขั้นตอนที่ 1: เลือกเมนู Batch Monitoring (จาก SelsectMenuMonitoringMyofiice)
                        Console.WriteLine("📋 กำลังเลือกเมนู Batch Order...");
                        WebDriverWait shortWait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        
                        // เพิ่มการรอให้หน้าโหลดเสร็จก่อน
                        Thread.Sleep(defaultShortWait);
                        
                        IWebElement menuNasMyOffice = shortWait.Until(ExpectedConditions.ElementToBeClickable(
                            By.XPath("//button/a[normalize-space()='Batch Order']")));
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuNasMyOffice);
                        Thread.Sleep(defaultShortWait);

                        Console.WriteLine("📋 กำลังเลือกเมนู Batch Monitoring (ระดับ 2)...");
                        IWebElement menuNasMyOfficeSecond = shortWait.Until(ExpectedConditions.ElementToBeClickable(
                            By.XPath("//li[@class='hover:bg-[#7a7a7a0e] pl-5 rounded-lg , bg-[#7a7a7a0e] text-text-hover']//a[@class='text-lg font-medium font-dbHelvethaica truncate pr-2'][normalize-space()='Batch Monitoring']")));
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuNasMyOfficeSecond);
                        Thread.Sleep(defaultShortWait);

                        Console.WriteLine("📋 กำลังเลือกเมนู Batch Monitoring (ระดับ 3)...");
                        IWebElement menuNasMyOfficeThird2 = shortWait.Until(ExpectedConditions.ElementToBeClickable(
                            By.XPath("//li[@class='hover:bg-[#7a7a7a0e] pl-5 rounded-lg , bg-[#7a7a7a0e] text-text-hover']//a[@class='truncate py-1 text-[18px] font-medium font-dbHelvethaica cursor-pointer text-text-secondary-color hover:text-text-hover w-full'][normalize-space()='Batch Monitoring']")));
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuNasMyOfficeThird2);
                        Thread.Sleep(initialLoadWait);

                        Console.WriteLine("✅ เลือกเมนู Batch Monitoring สำเร็จ");

                        // ขั้นตอนที่ 2: จัดการหน้าต่างใหม่และประมวลผลไฟล์ (จาก SelectMonitoringTopicMyofiice)
                        if (HandleNewWindowAndProcess(driver, outputPath, foundFileName, originalWindowHandle))
                        {
                            Console.WriteLine("✅ ประมวลผลไฟล์เสร็จสิ้น");
                            return; // สำเร็จแล้ว ออกจาก method
                        }
                        else
                        {
                            throw new Exception("Failed to process file in new window");
                        }
                    }
                    catch (Exception retryEx)
                    {
                        Console.WriteLine($"❌ การลองครั้งที่ {currentRetry} ล้มเหลว: {retryEx.Message}");
                        
                        // ทำความสะอาดและเตรียมพร้อมสำหรับการลองใหม่
                        CleanupAndReset(driver, originalWindowHandle);
                        
                        if (currentRetry < maxRetries)
                        {
                            Console.WriteLine($"⏳ รอ 5 วินาทีก่อนลองใหม่...");
                            Thread.Sleep(5000);
                        }
                        else
                        {
                            throw new Exception($"ล้มเหลวหลังจากลอง {maxRetries} ครั้ง: {retryEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ เกิดข้อผิดพลาดในการเลือกเมนูและประมวลผลไฟล์: {ex.Message}");
                
                // ทำความสะอาดก่อน re-throw
                CleanupAndReset(driver, originalWindowHandle);
                throw;
            }
        }

        /// <summary>
        /// ตรวจสอบสถานะหน้าเว็บก่อนเริ่มการทำงาน
        /// </summary>
        private void CheckPageState(IWebDriver driver)
        {
            try
            {
                // ตรวจสอบว่าหน้าเว็บยังใช้งานได้อยู่หรือไม่
                string currentUrl = driver.Url;
                string pageTitle = driver.Title;
                
                Console.WriteLine($"🔍 ตรวจสอบสถานะหน้าเว็บ: URL={currentUrl}, Title={pageTitle}");
                
                // ตรวจสอบว่าเป็นหน้า error หรือไม่
                if (currentUrl.Contains("error") || pageTitle.ToLower().Contains("error") || 
                    pageTitle.ToLower().Contains("bad request") || currentUrl.Contains("me.sh"))
                {
                    Console.WriteLine("⚠️ ตรวจพบหน้า error - ไม่ refresh เพื่อป้องกัน Bad Request");
                    // driver.Navigate().Refresh(); // ❌ ลบออกเพราะทำให้เกิด Bad Request
                    // Thread.Sleep(defaultLongWait);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ ไม่สามารถตรวจสอบสถานะหน้าเว็บได้: {ex.Message}");
            }
        }

        /// <summary>
        /// จัดการหน้าต่างใหม่และประมวลผลไฟล์
        /// </summary>
        private bool HandleNewWindowAndProcess(IWebDriver driver, string outputPath, string foundFileName, string originalWindowHandle)
        {
            try
            {
                Console.WriteLine("🔄 กำลังรอหน้าต่างใหม่...");
                
                // รอให้หน้าต่างใหม่เปิดขึ้นมา (รอจนกว่าจะมี 2 หน้าต่าง)
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15)); // เพิ่มเวลารอ
                wait.Until(d => d.WindowHandles.Count > 1);

                // วนลูปเพื่อสลับไปที่หน้าต่างใหม่
                foreach (string windowHandle in driver.WindowHandles)
                {
                    if (windowHandle != originalWindowHandle)
                    {
                        // ถ้า Handle ไม่ใช่ของหน้าต่างเดิม ให้สลับไปที่หน้าต่างนั้น
                        driver.SwitchTo().Window(windowHandle);
                        Console.WriteLine("✅ สลับไปที่หน้าต่างใหม่เรียบร้อยแล้ว");
                        break;
                    }
                }

                // ตรวจสอบว่าหน้าต่างใหม่โหลดสำเร็จหรือไม่
                Thread.Sleep(defaultLongWait);
                string newWindowUrl = driver.Url;
                string newWindowTitle = driver.Title;
                
                Console.WriteLine($"🔍 หน้าต่างใหม่: URL={newWindowUrl}, Title={newWindowTitle}");
                
                // ตรวจสอบว่าเป็นหน้า error หรือไม่
                if (newWindowUrl.Contains("error") || newWindowTitle.ToLower().Contains("error") || 
                    newWindowTitle.ToLower().Contains("bad request") || newWindowUrl.Contains("me.sh"))
                {
                    Console.WriteLine("❌ หน้าต่างใหม่เป็นหน้า error - จะปิดและลองใหม่");
                    driver.Close(); // ปิดหน้าต่างที่มี error
                    driver.SwitchTo().Window(originalWindowHandle);
                    return false;
                }

                // หลังจากสลับหน้าต่างแล้ว หา Element ในหน้าต่างใหม่
                IWebElement selectTopic = wait.Until(d => d.FindElement(By.XPath("//select[@id='selectModalMonitoringBatchTopicId']")));

                // ตรวจสอบว่า element สามารถคลิกได้หรือไม่
                if (selectTopic.Displayed && selectTopic.Enabled)
                {
                    // ใช้ JavaScript เพื่อคลิก element หากไม่สามารถคลิกได้โดยตรง
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("arguments[0].click();", selectTopic);
                    Thread.Sleep(defaultShortWait);

                    // รอจนกว่า option "ChangeOfferingSP" จะปรากฏใน dropdown
                    IWebElement menuTopic = wait.Until(d => d.FindElement(By.XPath("//option[@value='CHANGEPRO_SP']")));
                    menuTopic.Click();
                    Thread.Sleep(defaultShortWait);
                    Console.WriteLine("✅ เลือก Topic: ChangeOfferingSP");
                }
                else
                {
                    Console.WriteLine("❌ Element is not visible or not enabled.");
                    return false;
                }

                if (!string.IsNullOrEmpty(foundFileName))
                {
                    // กรอกชื่อไฟล์ใน input
                    IWebElement filePathInput = wait.Until(d => d.FindElement(By.XPath("//input[@id='inputMonitoringBatchProjectId']")));
                    filePathInput.Clear();
                    filePathInput.SendKeys(foundFileName);
                    AutoItX.Send("{ENTER}");
                    Thread.Sleep(initialLoadWait);
                    Console.WriteLine($"✅ กรอกชื่อไฟล์: {foundFileName}");
                }
                else
                {
                    Console.WriteLine("❌ ไม่พบไฟล์ที่ต้องการ ไม่สามารถดำเนินการต่อได้");
                    return false;
                }

                // คลิกปุ่ม "Query"
                Console.WriteLine("🔍 กำลังค้นหาข้อมูล...");
                IWebElement uploadButton = wait.Until(d => d.FindElement(By.XPath("//button[@id='btnMonitoringBatchConfirmId']")));
                uploadButton.Click();
                Console.WriteLine("✅ คลิกปุ่ม Query สำเร็จ");
                Thread.Sleep(defaultLongWait);

                // รอให้ตารางผลลัพธ์แสดงขึ้นและดาวน์โหลดไฟล์
                Console.WriteLine("📥 กำลังดาวน์โหลดไฟล์ผลลัพธ์...");
                DownloadAndConvertResultFiles(driver, wait, outputPath);

                // สลับกลับไปยัง context หลัก (หน้าเว็บหลัก)
                driver.SwitchTo().Window(originalWindowHandle);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ เกิดข้อผิดพลาดในการจัดการหน้าต่างใหม่: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ทำความสะอาดและรีเซ็ตสถานะ WebDriver
        /// </summary>
        private void CleanupAndReset(IWebDriver driver, string originalWindowHandle)
        {
            try
            {
                Console.WriteLine("🧹 กำลังทำความสะอาดและรีเซ็ต...");
                
                // ปิดหน้าต่างเพิ่มเติมทั้งหมด (ยกเว้นหน้าต่างหลัก)
                var allWindows = driver.WindowHandles;
                foreach (string windowHandle in allWindows)
                {
                    if (windowHandle != originalWindowHandle)
                    {
                        try
                        {
                            driver.SwitchTo().Window(windowHandle);
                            driver.Close();
                            Console.WriteLine("✅ ปิดหน้าต่างเพิ่มเติม");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠️ ไม่สามารถปิดหน้าต่างได้: {ex.Message}");
                        }
                    }
                }

                // กลับไปยังหน้าต่างหลัก
                driver.SwitchTo().Window(originalWindowHandle);
                
                // ❌ ไม่รีเฟรชหน้าเพราะจะทำให้เกิด Bad Request
                // driver.Navigate().Refresh(); 
                // Thread.Sleep(defaultLongWait);
                
                Console.WriteLine("✅ ทำความสะอาดเสร็จสิ้น");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ เกิดปัญหาในการทำความสะอาด: {ex.Message}");
            }
        }

        /// <summary>
        /// เลือกหัวข้อ "ChangeOfferingSP" ในระบบ Myofiice
        /// </summary>
        public void SelectBatchTopicMyofiice(IWebDriver driver, string PathMarketing)
        {
            try
            {
                // รอจนกว่า iframe จะปรากฏบนหน้าเว็บ
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); // รอสูงสุด 10 วินาที

                // สลับ context เข้าไปยัง iframe โดยใช้ XPath
                IWebElement iframeElement = wait.Until(d => d.FindElement(By.XPath("//iframe[@class='h-screen w-full overflow-hidden']")));
                driver.SwitchTo().Frame(iframeElement);

                // รอจนกว่า element "topicId" จะปรากฏใน iframe
                IWebElement selectTopic = wait.Until(d => d.FindElement(By.XPath("//div[@id='mat-select-value-1']")));

                // ตรวจสอบว่า element สามารถคลิกได้หรือไม่
                if (selectTopic.Displayed && selectTopic.Enabled)
                {
                    // ใช้ JavaScript เพื่อคลิก element หากไม่สามารถคลิกได้โดยตรง
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("arguments[0].click();", selectTopic);
                    Thread.Sleep(defaultShortWait);

                    // รอจนกว่า option "ChangeOfferingSP" จะปรากฏใน dropdown
                    IWebElement menuTopic = wait.Until(d => d.FindElement(By.XPath("//span[text()='ChangeOfferingSP']")));
                    menuTopic.Click();
                    Thread.Sleep(defaultShortWait);

                    // คลิก element สำหรับอัพโหลดไฟล์
                    IWebElement fileUploadElement = wait.Until(d => d.FindElement(By.XPath("//input[@id='mat-input-0']")));
                    fileUploadElement.Click();
                    Thread.Sleep(defaultShortWait);

                    // จัดการ File Picker Dialog ด้วย AutoIt
                    HandleFilePickerDialog(PathMarketing);
                    
                }
                else
                {
                    Console.WriteLine("Element is not visible or not enabled.");
                }

                Thread.Sleep(initialLoadWait);

                // สลับกลับไปยัง context หลัก (หน้าเว็บหลัก)
                driver.SwitchTo().DefaultContent();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Select Topic: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }


        /// <summary>
        /// เลือกหัวข้อ "ChangeOfferingSP" ในระบบ Myofiice
        /// </summary>
        public void SelectMonitoringTopicMyofiice(IWebDriver driver, string outputPath, string foundFileName)
        {
            try
            {
                // เก็บ ID ของหน้าต่างปัจจุบัน (หน้าต่างหลัก)
                string originalWindowHandle = driver.CurrentWindowHandle;

                // รอให้หน้าต่างใหม่เปิดขึ้นมา (รอจนกว่าจะมี 2 หน้าต่าง)
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                wait.Until(d => d.WindowHandles.Count > 1);

                // วนลูปเพื่อสลับไปที่หน้าต่างใหม่
                foreach (string windowHandle in driver.WindowHandles)
                {
                    if (windowHandle != originalWindowHandle)
                    {
                        // ถ้า Handle ไม่ใช่ของหน้าต่างเดิม ให้สลับไปที่หน้าต่างนั้น
                        driver.SwitchTo().Window(windowHandle);
                        Console.WriteLine("สลับไปที่หน้าต่างใหม่เรียบร้อยแล้ว");
                        break;
                    }
                }

                // หลังจากสลับหน้าต่างแล้ว ก็สามารถหา Element ในหน้าต่างใหม่ได้
                // ในกรณีนี้โค้ดจะสามารถหา element 'selectModalMonitoringBatchTopicId' ได้แล้ว
                IWebElement selectTopic = wait.Until(d => d.FindElement(By.XPath("//select[@id='selectModalMonitoringBatchTopicId']")));


                // ตรวจสอบว่า element สามารถคลิกได้หรือไม่
                if (selectTopic.Displayed && selectTopic.Enabled)
                {
                    // ใช้ JavaScript เพื่อคลิก element หากไม่สามารถคลิกได้โดยตรง
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("arguments[0].click();", selectTopic);
                    Thread.Sleep(defaultShortWait);

                    // รอจนกว่า option "ChangeOfferingSP" จะปรากฏใน dropdown
                    IWebElement menuTopic = wait.Until(d => d.FindElement(By.XPath("//option[@value='CHANGEPRO_SP']")));
                    menuTopic.Click();
                    Thread.Sleep(defaultShortWait);

                }
                else
                {
                    Console.WriteLine("Element is not visible or not enabled.");
                }

                if (!string.IsNullOrEmpty(foundFileName))
                {
                    // กรอกชื่อไฟล์ใน input
                    IWebElement filePathInput = wait.Until(d => d.FindElement(By.XPath("//input[@id='inputMonitoringBatchProjectId']")));
                    filePathInput.Clear();
                    filePathInput.SendKeys(foundFileName);
                    AutoItX.Send("{ENTER}");
                    Thread.Sleep(initialLoadWait);
                    Console.WriteLine($"กรอกชื่อไฟล์: {foundFileName}");
                }
                else
                {
                    Console.WriteLine("ไม่พบไฟล์ที่ต้องการ ไม่สามารถดำเนินการต่อได้");
                    driver.SwitchTo().DefaultContent();
                    return;
                }

                // คลิกปุ่ม "Query"
                IWebElement uploadButton = wait.Until(d => d.FindElement(By.XPath("//button[@id='btnMonitoringBatchConfirmId']")));
                uploadButton.Click();
                Console.WriteLine($"Click button Query");
                Thread.Sleep(defaultLongWait);

                // รอให้ตารางผลลัพธ์แสดงขึ้นและดาวน์โหลดไฟล์
                DownloadAndConvertResultFiles(driver, wait, outputPath);

                // สลับกลับไปยัง context หลัก (หน้าเว็บหลัก)
                driver.SwitchTo().Window(originalWindowHandle);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Select Topic: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }

        }

        /// <summary>
        /// หาไฟล์ที่มีชื่อขึ้นต้นด้วย prefix ที่กำหนดในโฟลเดอร์ และคืนค่าเฉพาะชื่อไฟล์
        /// </summary>
        /// <param name="directoryPath">พาธของโฟลเดอร์ที่ต้องการค้นหา</param>
        /// <param name="filePrefix">Prefix ของชื่อไฟล์ที่ต้องการหา</param>
        /// <returns>ชื่อไฟล์ที่พบ หรือ null ถ้าไม่พบ</returns> 
        public string FindFileNameWithPrefix(string directoryPath, string filePrefix)
        {
            try
            {
                if (!Directory.Exists(directoryPath))
                {
                    Console.WriteLine($"โฟลเดอร์ไม่พบ: {directoryPath}");
                    return null;
                }

                // หาไฟล์ที่ขึ้นต้นด้วย prefix ที่กำหนด
                var matchingFiles = Directory.GetFiles(directoryPath, $"{filePrefix}*.xlsx")
                                             .OrderByDescending(File.GetLastWriteTime)
                                             .ToList();

                if (matchingFiles.Any())
                {
                    string foundFile = matchingFiles.First();
                    string fileName = Path.GetFileName(foundFile); // เอาเฉพาะชื่อไฟล์
                    //Console.WriteLine($"พบไฟล์: {fileName}");
                    return fileName;
                }
                else
                {
                    Console.WriteLine($"ไม่พบไฟล์ที่ขึ้นต้นด้วย '{filePrefix}' ในโฟลเดอร์ {directoryPath}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการค้นหาไฟล์: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// จัดการ File Picker Dialog ด้วย AutoIt สำหรับการอัพโหลดไฟล์ (ปรับปรุงความเข้ากันได้และแก้ปัญหา Firefox ผุดมา)
        /// </summary>
        /// <param name="filePath">พาธของไฟล์ที่ต้องการอัพโหลด</param>
        private void HandleFilePickerDialog(string filePath)
        {
            try
            {
                Console.WriteLine($"🔄 กำลังอัพโหลดไฟล์: {filePath}");

                // ตรวจสอบว่าไฟล์มีอยู่จริงก่อน
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException($"ไม่พบไฟล์: {filePath}");
                }

                // เก็บข้อมูล browser window ก่อนเปิด file dialog
                string originalActiveWindow = AutoItX.WinGetTitle("[ACTIVE]");
                Console.WriteLine($"🔍 หน้าต่างที่ active ก่อนเปิด dialog: {originalActiveWindow}");

                // รอให้ File Picker Dialog ปรากฏขึ้น (เพิ่มเวลารอและการตรวจสอบ)
                Thread.Sleep(2000);
                Console.WriteLine("⏳ รอ File Picker Dialog...");

                // หาหน้าต่าง File Picker Dialog (เพิ่มชื่อที่เป็นไปได้)
                string[] possibleTitles = { 
                    "Open", "เปิด", "Choose File", "File Upload", "เลือกไฟล์",
                    "Browse", "เรียกดู", "Select File", "Upload File", "เปิดไฟล์",
                    "Open File", "เลือกไฟล์ที่ต้องการอัพโหลด", "เลือกไฟล์เพื่ออัปโหลด"
                };
                string windowTitle = "";
                int maxAttempts = 10; // เพิ่มจาก 5 เป็น 10 ครั้ง

                // ลองหาหน้าต่าง File Picker (เพิ่ม retry และ timeout)
                for (int attempt = 0; attempt < maxAttempts; attempt++)
                {
                    foreach (string title in possibleTitles)
                    {
                        if (AutoItX.WinExists(title) == 1)
                        {
                            windowTitle = title;
                            Console.WriteLine($"✅ พบหน้าต่าง File Picker: {title}");
                            break;
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(windowTitle)) break;
                    
                    Console.WriteLine($"⏳ ครั้งที่ {attempt + 1}/{maxAttempts}: รอหาหน้าต่าง File Picker...");
                    Thread.Sleep(1500); // เพิ่มเวลารอ
                    
                    // ทุก 3 ครั้ง ให้ตรวจสอบว่า browser ยังอยู่หรือไม่
                    if (attempt % 3 == 2)
                    {
                        CheckBrowserStatus(originalActiveWindow);
                    }
                }

                // ถ้าหาไม่เจอ ให้แสดงหน้าต่างที่มีอยู่และลองหาแบบอื่น
                if (string.IsNullOrEmpty(windowTitle))
                {
                    Console.WriteLine("🔍 แสดงหน้าต่างที่มีอยู่ทั้งหมด:");
                    ShowAllWindows();
                    
                    // ลองหาด้วย class name
                    windowTitle = FindDialogByClass();
                    
                    if (string.IsNullOrEmpty(windowTitle))
                    {
                        // ใช้หน้าต่างที่ active อยู่เป็นตัวสุดท้าย
                        windowTitle = "[ACTIVE]";
                        Console.WriteLine("⚠️ ใช้หน้าต่าง ACTIVE แทน");
                    }
                }

                // เปิดใช้งานหน้าต่าง File Picker
                AutoItX.WinActivate(windowTitle);
                AutoItX.WinWaitActive(windowTitle, "", 5); // รอให้หน้าต่าง active
                Thread.Sleep(1000);

                // ลอง method หลายแบบตามลำดับความน่าเชื่อถือ
                bool success = TryFilePathInput_Method1(windowTitle, filePath) ||
                              TryFilePathInput_Method2(windowTitle, filePath) ||
                              TryFilePathInput_Method3(windowTitle, filePath);


                if (!success)
                {
                    Console.WriteLine("❌ ทุก method ล้มเหลว กำลังใช้วิธีสุดท้าย...");
                    // วิธีสุดท้าย: ใช้ Clipboard (ปรับปรุงแล้ว)
                    TryFilePathInput_Clipboard_Improved(filePath);
                }

                // กดปุ่ม Enter หรือ Open เพื่อยืนยันการเลือกไฟล์
                Thread.Sleep(500);
                AutoItX.Send("{ENTER}");
                Thread.Sleep(1000);

                // ตรวจสอบว่าหน้าต่าง File Picker ปิดแล้วหรือไม่
                WaitForDialogCloseImproved(windowTitle, originalActiveWindow);

                Console.WriteLine("✅ การอัพโหลดไฟล์เสร็จสิ้น");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ เกิดข้อผิดพลาดในการอัพโหลดไฟล์: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Method 1: ใช้ ControlSetText กับ Control ID ต่างๆ
        /// </summary>
        private bool TryFilePathInput_Method1(string windowTitle, string filePath)
        {
            try
            {
                Console.WriteLine("🔧 Method 1: ControlSetText");
                string[] possibleControlIds = { "Edit1", "1001", "ComboBoxEx32", "ComboBox", "1148", "1152" };
                
                foreach (string controlId in possibleControlIds)
                {
                    if (AutoItX.ControlFocus(windowTitle, "", controlId) == 1)
                    {
                        Console.WriteLine($"✅ พบ Control: {controlId}");
                        
                        //// ล้างข้อความเดิม
                        //AutoItX.ControlSetText(windowTitle, "", controlId, "");
                        //Thread.Sleep(300);
                        
                        // กรอกพาธไฟล์
                        AutoItX.ControlSetText(windowTitle, "", controlId, filePath);
                        Thread.Sleep(500);
                        
                        // ตรวจสอบว่ากรอกสำเร็จหรือไม่
                        string currentText = AutoItX.ControlGetText(windowTitle, "", controlId);
                        if (currentText.Contains(Path.GetFileName(filePath)))
                        {
                            Console.WriteLine($"✅ Method 1 สำเร็จ: {currentText}");
                            return true;
                        }
                    }
                }
                Console.WriteLine("❌ Method 1 ล้มเหลว");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Method 1 Error: {ex.Message}");
                return false;
            }
        }


        /// <summary>
        /// Method 3: ใช้ Send โดยตรง
        /// </summary>
        private bool TryFilePathInput_Method2(string windowTitle, string filePath)
        {
            try
            {
                Console.WriteLine("🔧 Method 3: Send โดยตรง");
                
                //AutoItX.Send("^a"); // Select All
                //Thread.Sleep(300);
                AutoItX.Send(filePath);
                Thread.Sleep(500);
                
                Console.WriteLine("✅ Method 3 เสร็จสิ้น");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Method 3 Error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Method 4: ใช้ Navigation ด้วยทาง Path ใน Address Bar
        /// </summary>
        private bool TryFilePathInput_Method3(string windowTitle, string filePath)
        {
            try
            {
                Console.WriteLine("🔧 Method 4: Navigation ด้วย Address Bar");
                
                // ลองกด Alt+D เพื่อไปที่ Address Bar
                AutoItX.Send("!d");
                Thread.Sleep(500);
                
                // กรอกพาธไฟล์ทั้งหมด
                AutoItX.Send(filePath);
                Thread.Sleep(500);
                
                // กด Enter
                AutoItX.Send("{ENTER}");
                Thread.Sleep(500);
                
                Console.WriteLine("✅ Method 4 เสร็จสิ้น");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Method 4 Error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ตรวจสอบสถานะ Browser เพื่อป้องกันปัญหา Firefox ผุดมา
        /// </summary>
        private void CheckBrowserStatus(string originalActiveWindow)
        {
            try
            {
                string currentActiveWindow = AutoItX.WinGetTitle("[ACTIVE]");
                
                if (!string.IsNullOrEmpty(originalActiveWindow) && 
                    !currentActiveWindow.Equals(originalActiveWindow, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"⚠️ หน้าต่าง active เปลี่ยนจาก '{originalActiveWindow}' เป็น '{currentActiveWindow}'");
                    
                    // ถ้าเป็น Firefox หรือ browser อื่นผุดมา ให้กลับไปหน้าต่างเดิม
                    if (currentActiveWindow.ToLower().Contains("firefox") || 
                        currentActiveWindow.ToLower().Contains("chrome") ||
                        currentActiveWindow.ToLower().Contains("browser"))
                    {
                        Console.WriteLine("🔄 ตรวจพบ browser ผุดมา - กำลังกลับไปหน้าต่างเดิม");
                        AutoItX.WinActivate(originalActiveWindow);
                        Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ ไม่สามารถตรวจสอบสถานะ browser ได้: {ex.Message}");
            }
        }

        /// <summary>
        /// หา Dialog ด้วย Class Name
        /// </summary>
        private string FindDialogByClass()
        {
            try
            {
                Console.WriteLine("🔍 กำลังหา Dialog ด้วย Class Name...");
                
                // class names ที่เป็นไปได้สำหรับ File Dialog
                string[] dialogClasses = { 
                    "#32770",           // Standard Windows Dialog
                    "Chrome_WidgetWin_1", // Chrome File Dialog
                    "MozillaDialogClass", // Firefox Dialog
                    "Qt5QFileDialog",     // Qt Application Dialog
                    "TFileOpenDialog"     // Delphi/C++ Builder Dialog
                };
                
                foreach (string className in dialogClasses)
                {
                    string classPattern = $"[CLASS:{className}]";
                    if (AutoItX.WinExists(classPattern) == 1)
                    {
                        string windowTitle = AutoItX.WinGetTitle(classPattern);
                        Console.WriteLine($"✅ พบ Dialog ด้วย class {className}: {windowTitle}");
                        return classPattern;
                    }
                }
                
                Console.WriteLine("❌ ไม่พบ Dialog ด้วย Class Name");
                return "";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error ในการหา Dialog ด้วย Class: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Method สุดท้าย: ใช้ Clipboard แบบปรับปรุง
        /// </summary>
        private void TryFilePathInput_Clipboard_Improved(string filePath)
        {
            try
            {
                Console.WriteLine("🔧 Method สุดท้าย: Clipboard แบบปรับปรุง");
                
                // วิธีที่ 1: ใช้ System.Windows.Clipboard (ถ้ามี)
                bool clipboardSet = false;
                try
                {
                    System.Windows.Clipboard.SetText(filePath);
                    clipboardSet = true;
                    Console.WriteLine("✅ ตั้งค่า Clipboard ด้วย System.Windows");
                }
                catch
                {
                    Console.WriteLine("⚠️ ไม่สามารถใช้ System.Windows.Clipboard ได้");
                }
                
                // วิธีที่ 2: ใช้ AutoIt (ถ้าวิธีแรกไม่ได้)
                if (!clipboardSet)
                {
                    try
                    {
                        AutoItX.ClipPut(filePath);
                        Console.WriteLine("✅ ตั้งค่า Clipboard ด้วย AutoIt");
                        clipboardSet = true;
                    }
                    catch
                    {
                        Console.WriteLine("⚠️ ไม่สามารถใช้ AutoIt.ClipPut ได้");
                    }
                }
                
                if (clipboardSet)
                {
                    // ลบข้อความเดิมและ paste
                    AutoItX.Send("^a"); // Select All
                    Thread.Sleep(300);
                    AutoItX.Send("^v"); // Paste
                    Thread.Sleep(500);
                    
                    // ตรวจสอบว่า paste สำเร็จหรือไม่
                    AutoItX.Send("^a"); // Select All เพื่อดูว่ามีข้อความไหม
                    Thread.Sleep(200);
                    
                    Console.WriteLine("✅ Method Clipboard แบบปรับปรุงเสร็จสิ้น");
                }
                else
                {
                    Console.WriteLine("❌ ไม่สามารถตั้งค่า Clipboard ได้ - ลองพิมพ์โดยตรง");
                    
                    // วิธีสุดท้าย: พิมพ์ทีละตัวอักษร
                    AutoItX.Send("^a"); // Select All
                    Thread.Sleep(300);
                    
                    // แบ่งเป็นชิ้นเล็กเพื่อป้องกันปัญหา
                    string[] parts = filePath.Split('\\');
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (i > 0) AutoItX.Send("\\"); // เพิ่ม backslash ระหว่าง path
                        
                        // พิมพ์แต่ละส่วนของ path
                        AutoItX.Send(parts[i]);
                        Thread.Sleep(100); // รอหน่อยระหว่างส่วน
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Method Clipboard แบบปรับปรุง Error: {ex.Message}");
            }
        }

        /// <summary>
        /// รอให้ Dialog ปิด แบบปรับปรุง
        /// </summary>
        private void WaitForDialogCloseImproved(string windowTitle, string originalActiveWindow)
        {
            int waitCount = 0;
            int maxWaitCount = 20; // เพิ่มจาก 15 เป็น 20
            
            while (AutoItX.WinExists(windowTitle) == 1 && waitCount < maxWaitCount)
            {
                Thread.Sleep(500);
                waitCount++;
                
                Console.WriteLine($"⏳ รอ Dialog ปิด... ({waitCount}/{maxWaitCount})");
                
                if (waitCount == 3)
                {
                    // ลองกดปุ่ม Open
                    Console.WriteLine("🔄 ลองกดปุ่ม Open...");
                    AutoItX.ControlClick(windowTitle, "", "Button1");
                    Thread.Sleep(500);
                }
                else if (waitCount == 6)
                {
                    // ลองกดปุ่ม OK
                    Console.WriteLine("🔄 ลองกดปุ่ม OK...");
                    AutoItX.ControlClick(windowTitle, "", "Button2");
                    Thread.Sleep(500);
                }
                else if (waitCount == 10)
                {
                    // ลองกด Enter อีกครั้ง
                    Console.WriteLine("🔄 ลองกด Enter อีกครั้ง...");
                    AutoItX.Send("{ENTER}");
                    Thread.Sleep(500);
                }
                else if (waitCount == 15)
                {
                    // ลองกด ESC เพื่อปิด dialog
                    Console.WriteLine("🔄 ลองกด ESC เพื่อปิด dialog...");
                    AutoItX.Send("{ESC}");
                    Thread.Sleep(500);
                }
                
                // ตรวจสอบว่า browser ยังอยู่หรือไม่
                if (waitCount % 5 == 0)
                {
                    CheckBrowserStatus(originalActiveWindow);
                }
            }
            
            if (AutoItX.WinExists(windowTitle) == 1)
            {
                Console.WriteLine("⚠️ หน้าต่างยังไม่ปิดหลังรอนาน - อาจมีปัญหา");
                
                // พยายาม force close dialog
                try
                {
                    AutoItX.WinClose(windowTitle);
                    Thread.Sleep(1000);
                    
                    if (AutoItX.WinExists(windowTitle) == 1)
                    {
                        Console.WriteLine("🔧 ลอง force kill dialog...");
                        AutoItX.WinKill(windowTitle);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ ไม่สามารถปิด dialog ได้: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("✅ Dialog ปิดเรียบร้อยแล้ว");
            }
        }

        /// <summary>
        /// แสดงหน้าต่างทั้งหมดที่มีอยู่ (สำหรับ debug) - ปรับปรุงแล้ว
        /// </summary>
        private void ShowAllWindows()
        {
            try
            {
                Console.WriteLine("📋 กำลังตรวจสอบหน้าต่างที่มีอยู่...");
                
                // แสดงหน้าต่าง active ปัจจุบัน
                string activeWindow = AutoItX.WinGetTitle("[ACTIVE]");
                Console.WriteLine($"🔍 หน้าต่าง Active: '{activeWindow}'");
                
                // ลองหาหน้าต่างที่มี title ว่าง
                if (AutoItX.WinExists("") == 1)
                {
                    Console.WriteLine("🔍 พบหน้าต่างที่ไม่มี title");
                }
                
                // ลองหาหน้าต่างด้วย class name
                string[] commonClasses = { 
                    "#32770",               // Standard Windows Dialog
                    "Chrome_WidgetWin_1",   // Chrome File Dialog
                    "MozillaDialogClass",   // Firefox Dialog
                    "Qt5QFileDialog",       // Qt Application Dialog
                    "TFileOpenDialog",      // Delphi/C++ Builder Dialog
                    "DirectUIHWND"          // Modern Windows Dialog
                };
                
                foreach (string className in commonClasses)
                {
                    string classPattern = $"[CLASS:{className}]";
                    if (AutoItX.WinExists(classPattern) == 1)
                    {
                        string windowTitle = AutoItX.WinGetTitle(classPattern);
                        Console.WriteLine($"🔍 พบหน้าต่าง class '{className}': '{windowTitle}'");
                    }
                }
                
                // ลองหา dialog ด้วยชื่อที่เป็นไปได้
                string[] dialogTitles = { 
                    "เปิด", "Open", "Browse", "Choose File", "Select File", 
                    "File Upload", "Upload", "เลือกไฟล์", "เรียกดู"
                };
                
                foreach (string title in dialogTitles)
                {
                    if (AutoItX.WinExists(title) == 1)
                    {
                        Console.WriteLine($"🔍 พบ dialog ชื่อ: '{title}'");
                    }
                }
                
                Console.WriteLine("📋 การตรวจสอบหน้าต่างเสร็จสิ้น");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ ไม่สามารถตรวจสอบหน้าต่าง: {ex.Message}");
            }
        }

        /// <summary>
        /// Method สุดท้าย: ใช้ Clipboard
        /// </summary>
        private void TryFilePathInput_Clipboard(string filePath)
        {
            try
            {
                Console.WriteLine("🔧 Method สุดท้าย: Clipboard");
                
                // Copy ไฟล์พาธไปที่ Clipboard ด้วย Win32 API
                SetClipboardText(filePath);
                Thread.Sleep(300);
                
                // Select All และ Paste
                AutoItX.Send("^a");
                Thread.Sleep(300);
                AutoItX.Send("^v");
                Thread.Sleep(500);
                
                Console.WriteLine("✅ Method Clipboard เสร็จสิ้น");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Method Clipboard Error: {ex.Message}");
            }
        }

        /// <summary>
        /// ตั้งค่า Clipboard ด้วยวิธีง่ายๆ
        /// </summary>
        private void SetClipboardText(string text)
        {
            try
            {
                // ใช้ System.Windows.Clipboard แทน
                System.Windows.Clipboard.SetText(text);
            }
            catch
            {
                // ถ้าไม่ได้ ใช้วิธีอื่น
                Console.WriteLine("⚠️ ไม่สามารถใช้ Clipboard ได้");
            }
        }

        /// <summary>
        /// รอให้ Dialog ปิด
        /// </summary>
        private void WaitForDialogClose(string windowTitle)
        {
            int waitCount = 0;
            while (AutoItX.WinExists(windowTitle) == 1 && waitCount < 15)
            {
                Thread.Sleep(500);
                waitCount++;
                
                if (waitCount == 5)
                {
                    // ลองกดปุ่ม Open
                    Console.WriteLine("🔄 ลองกดปุ่ม Open...");
                    AutoItX.ControlClick(windowTitle, "", "Button1");
                }
                else if (waitCount == 10)
                {
                    // ลองกด Enter อีกครั้ง
                    Console.WriteLine("🔄 ลองกด Enter อีกครั้ง...");
                    AutoItX.Send("{ENTER}");
                }
            }
            
            if (AutoItX.WinExists(windowTitle) == 1)
            {
                Console.WriteLine("⚠️ หน้าต่างยังไม่ปิด อาจมีปัญหา");
            }
        }

        /// <summary>
        /// แสดงหน้าต่างทั้งหมดที่มีอยู่ (สำหรับ debug) - Version เก่า [ลบได้]
        /// </summary>
        private void ShowAllWindows_OLD()
        {
            try
            {
                // แสดงข้อมูล debug พื้นฐาน
                Console.WriteLine("📋 กำลังตรวจสอบหน้าต่างที่มีอยู่...");
                
                // ลองหาหน้าต่างที่มี title ว่าง
                if (AutoItX.WinExists("") == 1)
                {
                    Console.WriteLine("🔍 พบหน้าต่างที่ไม่มี title");
                }
                
                // ลองหาหน้าต่างด้วย class name
                string[] commonClasses = { "#32770", "Chrome_WidgetWin_1", "MozillaDialogClass" };
                foreach (string className in commonClasses)
                {
                    if (AutoItX.WinExists($"[CLASS:{className}]") == 1)
                    {
                        Console.WriteLine($"� พบหน้าต่าง class: {className}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ ไม่สามารถตรวจสอบหน้าต่าง: {ex.Message}");
            }
        }

        /// <summary>
        /// เลือกเมนู My Reports ในระบบ IM
        /// </summary>
        public void SelectMyReportIM(IWebDriver driver, string titleReport)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            try
            {
                // ✅ รอจน iframe โหลดเสร็จ
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.XPath("//iframe[@id='PegaGadget0Ifr']")));
                Thread.Sleep(initialLoadWait);

                // ✅ หา element My Reports หลังเข้า iframe แล้ว
                IWebElement myReportsElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[@title='My Reports']")));
                myReportsElement.Click();
                Thread.Sleep(initialLoadWait);

                // ✅ หาเมนู ServiceRequest ภายใน iframe เดียวกัน
                IWebElement menuServiceRequest = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//div[@id='gridBody_right']//a[@title='{titleReport}']")));
                
                menuServiceRequest.Click();
                Thread.Sleep(initialLoadWait);

                // ✅ กลับสู่ context หลัก (optional ถ้า iframe เดียว)
                driver.SwitchTo().DefaultContent();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Select Menu Orders: {ex.Message}");
                throw;
            }
        }


        /// <summary>
        /// เลือกเมนู Manager Tools ในระบบ IM
        /// </summary>
        public void SelectManagerToolsIM(IWebDriver driver)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            try
            {
                // ✅ รอจน iframe โหลดเสร็จ
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.XPath("//iframe[@id='PegaGadget0Ifr']")));
                Thread.Sleep(initialLoadWait);

                // ✅ หา element Manager Tools หลังเข้า iframe แล้ว
                IWebElement ManagerToolsElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[@title='Manager Tools']")));
                ManagerToolsElement.Click();
                Thread.Sleep(initialLoadWait);

                // ✅ หาเมนู Maintain Case ภายใน iframe เดียวกัน
                IWebElement menuMaintainCase = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[text()='Maintain Case']")));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", menuMaintainCase);
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", menuMaintainCase);
                Thread.Sleep(defaultLongWait);

                // ✅ หาเมนู Import Case list ภายใน iframe เดียวกัน
                IWebElement menuImportCase = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//span[text()='Import Case list']")));
                menuImportCase.Click();
                Thread.Sleep(initialLoadWait);

                // ✅ กลับสู่ context หลัก (optional ถ้า iframe เดียว)
                driver.SwitchTo().DefaultContent();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Select Menu Manager Tools: {ex.Message}");
                throw;
            }
        }


        /// <summary>
        /// เลือกเมนู Import close case ในระบบ IM
        /// </summary>
        public void ImportCloseCase(IWebDriver driver, string fileCloseCase)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(600));

            try
            {
                // ✅ รอจน iframe โหลดเสร็จ
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.XPath("//iframe[@id='PegaGadget1Ifr']")));
                Thread.Sleep(defaultLongWait);

                // ✅ คลิก element Import หลังเข้า iframe แล้ว
                IWebElement ImportElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[@name='MaintainTask_pyWorkPage_1' and text()='Import']")));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ImportElement);
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", ImportElement);
                Thread.Sleep(defaultLongWait);

                // ✅ คลิก เลือกไฟล์ 
                IWebElement modal = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.Id("modalWrapper")));
                IWebElement fileInput = modal.FindElement(
                    By.XPath("//input[@type='file' and contains(@title,'No file chosen')]"));
                fileInput.SendKeys(fileCloseCase);
                IWebElement uploadButton = modal.FindElement(
                    By.XPath("//button[normalize-space(text())='Upload file']"));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", uploadButton);

                // ✅ คลิก Upload file
                IWebElement btnStartUpload = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//button[text()='Start Upload Data']")));
                btnStartUpload.Click();
                Thread.Sleep(defaultLongWait);
                Console.WriteLine("กำลังรอให้การดาวน์โหลดครบ 100%...");
                bool isDownloadComplete = false;
                // ใช้ while loop เพื่อรอจนกว่าจะเจอ element
                while (!isDownloadComplete)
                {
                    try
                    {
                        // ค้นหา element ด้วย XPath
                        IWebElement element = driver.FindElement(By.XPath("//div[@class='updates dataLabelWrite']/span"));

                        // ตรวจสอบข้อความใน element
                        if (element.Text == "100%")
                        {
                            isDownloadComplete = true;
                            Console.WriteLine("ดาวน์โหลดครบ 100%");
                        }
                        else
                        {
                            Thread.Sleep(1000); // รอ 1 วินาทีแล้วเช็คใหม่
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // ถ้าไม่พบ element ให้รอและเช็คใหม่
                        Thread.Sleep(1000);
                    }
                }

                // ✅ คลิก Submit
                IWebElement btnSubmit = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//button[text()='Submit']")));
                btnSubmit.Click();
                Thread.Sleep(defaultLongWait);

                // ✅ กลับสู่ context หลัก (optional ถ้า iframe เดียว)
                driver.SwitchTo().DefaultContent();

                Console.WriteLine($"Import close case success");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error import close case: {ex.Message}");
                throw;
            }
        }


        /// <summary>
        /// logout จากระบบ Myofiice
        /// </summary>
        public void LogoutMyofiice(IWebDriver driver)
        {
            try
            {
                IWebElement ClickLogout = driver.FindElement(By.XPath("//button/span/a[text()='Logout']"));
                ClickLogout.Click();
                Thread.Sleep(initialLoadWait);
                driver.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during logout: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }


        /// <summary>
        /// logout จากระบบ IM
        /// </summary>
        public void LogoutIM(IWebDriver driver)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            try
            {

                driver.SwitchTo().DefaultContent();
                IWebElement ClickProfile = driver.FindElement(By.XPath("//button[@data-test-id='px-opr-image-ctrl']"));
                ClickProfile.Click();
                Thread.Sleep(initialLoadWait);
                IWebElement ClickLogout = driver.FindElement(By.XPath("//li[@title='Logout']"));
                ClickLogout.Click();
                // รอให้ Alert ปรากฏขึ้น
                IAlert confirmationAlert = wait.Until(ExpectedConditions.AlertIsPresent());
                // ตรวจสอบข้อความใน Alert (ไม่บังคับ แต่ดีสำหรับการ Debug)
                string alertText = confirmationAlert.Text;
                // ยืนยันการ Logout โดยการ Accept (คลิกปุ่ม OK/Yes ใน Alert)
                confirmationAlert.Accept();
                Thread.Sleep(initialLoadWait);
                driver.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during logout: {ex.Message}");
                throw; // Re-throw the exception to be handled by the caller
            }
        }

        /// <summary>
        /// ค้นหาและ Export ข้อมูล Pending-Review ย้อนหลัง 6 วัน (รวมวันนี้)
        /// </summary>
        /// <param name="driver">WebDriver ที่ล็อกอินและอยู่หน้าค้นหาแล้ว</param>
        /// <param name="PathDownload">พาธโฟลเดอร์สำหรับเก็บไฟล์ที่ดาวน์โหลด</param>
        /// <param name="PathReport">พาธโฟลเดอร์สำหรับเก็บไฟล์รายงานที่จัดเก็บ</param>
        /// <param name="optionText">ข้อความของตัวเลือกใน dropdown (เช่น "Pending-Review")</param>
        /// <param name="numberDay">จำนวนวันที่ต้องการย้อนกลับไป (รวมวันนี้)</param>
        /// <param name="titleReport">ชื่อรายงานที่ต้องการ Export</param>
        public void ExportPendingReview(IWebDriver driver, string PathDownload, string PathReport, string optionText, string numberDay, string titleReport)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            int daysToCheck = int.Parse(numberDay);

            // วนลูปย้อนหลัง 6 วัน (รวมวันนี้)
            for (int i = 0; i < daysToCheck; i++)
            {

                // กลับสู่ context หลักก่อนทุกครั้ง
                driver.SwitchTo().DefaultContent();
                // สลับ context เข้าไปยัง iframe โดยใช้ XPath
                IWebElement iframeElement = wait.Until(d => d.FindElement(By.XPath("//iframe[@id='PegaGadget1Ifr']")));
                driver.SwitchTo().Frame(iframeElement);

                // 1. Scroll ไปด้านซ้ายสุดเพื่อหา datepicker ก่อนหา element อื่น
                IWebElement calendarIcon = wait.Until(ExpectedConditions.ElementExists(By.XPath("//input[@id='5f420ea9']")));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", calendarIcon);

                //// 2. กรอกวันที่ใน datepicker (ย้อนหลังตามรอบ)
                DateTime targetDate = DateTime.Now.AddDays(-i);
                string dateString = targetDate.ToString("d/M/yyyy HH:mm", CultureInfo.InvariantCulture); // เช่น 7/7/2025 15:50
                wait.Until(d =>
                {
                    try
                    {
                        IWebElement dateInput = d.FindElement(By.XPath("//input[@id='5f420ea9']"));
                        dateInput.Clear();
                        dateInput.SendKeys(dateString);
                        return true; // คืนค่า true เมื่อทำสำเร็จ
                    }
                    catch (StaleElementReferenceException)
                    {
                        return false; // เกิด StaleElementException ให้ลองใหม่ในรอบถัดไปของ wait.Until
                    }
                });

                // 3. เลือก Status "Pending-Review"
                IWebElement statusDropdown = wait.Until(d => d.FindElement(By.XPath("//select[@id='98ca3d8']")));
                Thread.Sleep(defaultShortWait);
                IWebElement pendingReviewOption = wait.Until(d => d.FindElement(By.XPath($"//option[normalize-space(text())='{optionText}']")));
                pendingReviewOption.Click();

                // 4. กดปุ่ม Apply filters เพื่อค้นหา
                IWebElement applyButton = wait.Until(d => d.FindElement(By.XPath("//button[@name='pyDefaultCustomFilterApplyCancel_pyReportContentPage_5']")));
                applyButton.Click();

                // 5. รอผลลัพธ์และเช็คว่ามีข้อมูลหรือไม่
                // XPath สำหรับเช็คจำนวนข้อมูล: //div[contains(@class,'standard(label)_dataLabelRead') and contains(text(),'Displaying')]
                Thread.Sleep(defaultLongWait); // รอโหลดข้อมูล

                var resultLabels = driver.FindElements(By.XPath("//div[@class='standard_(label)_dataLabelRead' and contains(text(),'Displaying')]"));
                bool hasData = false;
                int recordCount = 0;

                if (resultLabels.Any())
                {
                    string labelText = resultLabels.First().Text;
                    // ปรับ Regex ให้คำว่า 's' เป็น optional
                    var match = System.Text.RegularExpressions.Regex.Match(labelText, @"Displaying\s+(\d+)\s+records?");

                    if (match.Success)
                    {
                        recordCount = int.Parse(match.Groups[1].Value);
                        hasData = recordCount > 0;
                    }
                }

                if (hasData)
                {
                    // 6. คลิก Actions > Export To Excel
                    IWebElement fineActionBtn = wait.Until(ExpectedConditions.ElementExists(By.XPath("//button[text()='Actions']")));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", fineActionBtn);

                    Thread.Sleep(defaultShortWait);
                    IWebElement actionsButton = wait.Until(d => d.FindElement(By.XPath("//button[text()='Actions']")));
                    actionsButton.Click();

                    Thread.Sleep(defaultShortWait);
                    IWebElement exportExcelMenu = wait.Until(d => d.FindElement(By.XPath("//li/a[@class='menu-item-anchor ']/span/span[normalize-space()='Export To Excel']")));
                    exportExcelMenu.Click();

                    // รอให้ดาวน์โหลดเสร็จ (อาจต้องปรับเวลาตามขนาดไฟล์)
                    Thread.Sleep(initialLoadWait);
                    HandleDownloadedFiles(driver, PathDownload, PathReport, titleReport);

                    // กลับไปค้นหาวันถัดไป
                    continue;
                }
                else
                {
                    // ถ้าไม่พบข้อมูล ให้วนไปวันถัดไป
                    continue;
                }
            }

            try
            {
                // รอให้ element ปรากฏและสามารถคลิกได้
                IWebElement closeButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//div[@class='field-item dataValueWrite']/span/a[@aria-label='Close this item']")));

                // ใช้ JavaScript Click เพื่อความมั่นใจในการคลิก
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("arguments[0].click();", closeButton);

                Thread.Sleep(defaultShortWait);

                // รอให้ element ปรากฏและสามารถคลิกได้
                IWebElement confirmCloseButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//div[@id='modalWrapper']//button[text()='OK']")));
                js.ExecuteScript("arguments[0].click();", confirmCloseButton);

                Thread.Sleep(defaultShortWait);

                Console.WriteLine("Successfully clicked close item button");
            }
            catch (TimeoutException ex)
            {
                Console.WriteLine($"Timeout waiting for close button: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error clicking close button: {ex.Message}");
                throw;
            }

        }

        /// <summary>
        /// จัดการไฟล์ที่ดาวน์โหลดมา โดยสร้างโครงสร้างโฟลเดอร์ตามปี/เดือน และย้ายไฟล์ที่ตรงตามเงื่อนไขไป
        /// </summary>
        /// <param name="downloadPath">พาธของโฟลเดอร์ที่เบราว์เซอร์ดาวน์โหลดไฟล์ไป</param>
        /// <param name="baseTargetPath">พาธของโฟลเดอร์ปลายทางหลัก (เช่น "D:\Reports")</param>
        /// <param name="expectedFileNamePrefix">Prefix ของชื่อไฟล์ที่คาดหวัง เช่น "ReportSOSDPromotion_ServiceRequest_"</param>
        /// <returns>True ถ้าจัดการไฟล์สำเร็จ, False ถ้าไม่สำเร็จ</returns>
        public bool HandleDownloadedFiles(IWebDriver driver, string downloadPath, string baseTargetPath, string expectedFileNamePrefix)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                // รอไฟล์ดาวน์โหลดเสร็จสมบูรณ์
                // ใช้ WebDriverWait เพื่อรอไฟล์ที่ตรงตามเงื่อนไขและไม่ถูกล็อก
                string downloadedFilePath = wait.Until(d =>
                {
                    // รับรายการไฟล์ทั้งหมดในโฟลเดอร์ดาวน์โหลด
                    // และกรองเฉพาะไฟล์ .xlsx ที่ชื่อขึ้นต้นด้วย prefix
                    var files = Directory.GetFiles(downloadPath, "*.xlsx")
                                         .Where(file => Path.GetFileName(file).StartsWith(expectedFileNamePrefix, StringComparison.OrdinalIgnoreCase) &&
                                                        !file.EndsWith(".crdownload", StringComparison.OrdinalIgnoreCase) && // สำหรับ Chrome
                                                        !file.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase) &&        // สำหรับเบราว์เซอร์อื่น ๆ
                                                        !file.EndsWith(".part", StringComparison.OrdinalIgnoreCase))        // สำหรับ Firefox
                                         .OrderByDescending(File.GetLastWriteTime) // เรียงตามเวลาแก้ไขล่าสุด (ไฟล์ที่เพิ่งมาจะอยู่บนสุด)
                                         .ToList();

                    if (files.Any())
                    {
                        string latestFile = files.First();
                        // ตรวจสอบว่าไฟล์ถูกล็อกอยู่หรือไม่ (หมายถึงยังดาวน์โหลดไม่เสร็จ)
                        if (!IsFileLocked(latestFile))
                        {
                            Console.WriteLine($"Downloaded file found: {latestFile}");
                            return latestFile; // คืนค่าพาธไฟล์ที่สมบูรณ์แล้ว
                        }
                        Console.WriteLine($"File '{latestFile}' is still being downloaded or locked.");
                    }
                    return null; // ยังไม่เจอไฟล์ที่สมบูรณ์
                });

                if (string.IsNullOrEmpty(downloadedFilePath))
                {
                    Console.WriteLine($"Timed out waiting for file starting with '{expectedFileNamePrefix}' to download completely in '{downloadPath}'.");
                    return false;
                }

                // กำหนดชื่อไฟล์ปลายทาง (ใช้ชื่อเดิมของไฟล์ที่ดาวน์โหลดมา)
                string fileName = Path.GetFileName(downloadedFilePath);
                string finalFilePath = Path.Combine(baseTargetPath, fileName);

                // ถ้ามีไฟล์ชื่อเดียวกันอยู่ที่ปลายทางอยู่แล้ว ให้ลบก่อน
                if (File.Exists(finalFilePath))
                {
                    File.Delete(finalFilePath);
                    Console.WriteLine($"Deleted existing file at destination: {finalFilePath}");
                }

                // ย้ายไฟล์ไปยังโฟลเดอร์ปลายทาง
                File.Move(downloadedFilePath, finalFilePath);
                Console.WriteLine($"Moved file '{fileName}' to '{finalFilePath}'.");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error handling downloaded files: {ex.Message}");
                // Optional: Log the full stack trace for more detailed debugging
                // Console.WriteLine(ex.StackTrace); 
                return false;
            }
        }

        /// <summary>
        /// Helper method เพื่อตรวจสอบว่าไฟล์ถูกล็อกโดย Process อื่นอยู่หรือไม่
        /// </summary>
        /// <param name="filePath">พาธของไฟล์ที่ต้องการตรวจสอบ</param>
        /// <returns>True ถ้าไฟล์ถูกล็อก, False ถ้าไม่ได้ถูกล็อก</returns>
        private bool IsFileLocked(string filePath)
        {
            try
            {
                // ลองเปิดไฟล์แบบ ReadWrite โดยไม่ให้ Process อื่นเข้าถึงได้
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close(); // ปิด Stream ทันที
                }
            }
            catch (IOException)
            {
                // ถ้าเกิด IOException แสดงว่าไฟล์ถูกล็อกอยู่
                return true;
            }
            // ถ้าไม่เกิด Exception แสดงว่าไฟล์ไม่ได้ถูกล็อก
            return false;
        }

        /// <summary>
        /// จัดการการล็อกอินด้วยอีเมลและรหัสผ่าน พร้อมการยืนยันตัวตนที่โทรศัพท์
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        /// <param name="email">อีเมลผู้ใช้สำหรับล็อกอิน</param>
        /// <param name="password">รหัสผ่าน</param>
        private void HandleLoginWithEmailAuth(IWebDriver driver, string email, string password)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
                
                // ตรวจสอบว่ามีช่องกรอกอีเมลหรือไม่
                IList<IWebElement> emailFields = driver.FindElements(By.Id("i0116"));

                if (emailFields.Any())
                {
                    // กรอกอีเมล
                    IWebElement emailInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("i0116")));
                    emailInput.Clear();
                    emailInput.SendKeys(email);
                    Thread.Sleep(defaultShortWait);

                    // คลิกปุ่ม Next หลังจากกรอกอีเมล
                    IWebElement nextButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
                    nextButton.Click();
                    Thread.Sleep(defaultLongWait);

                    // กรอกรหัสผ่าน
                    IWebElement passwordInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("i0118")));
                    passwordInput.Clear();
                    passwordInput.SendKeys(password);
                    Thread.Sleep(defaultShortWait);

                    // คลิกปุ่ม Sign in
                    IWebElement signInButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
                    signInButton.Click();
                    Thread.Sleep(defaultLongWait);

                    // ตรวจสอบว่ามี element สำหรับส่งรหัสยืนยันใหม่หรือไม่
                    HandleAuthenticatorApprovalRequest(driver, wait);

                    Console.WriteLine("รหัสยืนยันถูกส่งไปยังโทรศัพท์ของคุณแล้ว");
                    Console.WriteLine("กรุณาทำการยืนยันตัวตนที่โทรศัพท์ของคุณ...");

                    //กดปุ่ม Stay signed in
                    IWebElement staySignedInButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("idBtn_Back")));
                    staySignedInButton.Click();
                    Thread.Sleep(defaultLongWait);

                    // รอให้ผู้ใช้ทำการยืนยันที่โทรศัพท์เสร็จสิ้น
                    WaitForPhoneAuthentication(driver, wait);

                    Console.WriteLine("การยืนยันตัวตนเสร็จสมบูรณ์ กำลังดำเนินการต่อ...");


                }
                else if (driver.FindElements(By.Id("i0118")).Any())
                {
                    // กรอกรหัสผ่าน
                    IWebElement passwordInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("i0118")));
                    passwordInput.Clear();
                    passwordInput.SendKeys(password);
                    Thread.Sleep(defaultShortWait);

                    // คลิกปุ่ม Sign in
                    IWebElement signInButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
                    signInButton.Click();
                    Thread.Sleep(defaultLongWait);

                    Console.WriteLine("รหัสยืนยันถูกส่งไปยังโทรศัพท์ของคุณแล้ว");
                    Console.WriteLine("กรุณาทำการยืนยันตัวตนที่โทรศัพท์ของคุณ...");

                    //กดปุ่ม Stay signed in
                    IWebElement staySignedInButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("idBtn_Back")));
                    staySignedInButton.Click();
                    Thread.Sleep(defaultLongWait);

                    // รอให้ผู้ใช้ทำการยืนยันที่โทรศัพท์เสร็จสิ้น
                    WaitForPhoneAuthentication(driver, wait);

                    Console.WriteLine("การยืนยันตัวตนเสร็จสมบูรณ์ กำลังดำเนินการต่อ...");



                }
                else if (driver.FindElements(By.Id("loginHeader")).Any())
                {

                    IWebElement pendingReviewOption = wait.Until(d => d.FindElement(By.XPath($"//div[@data-test-id='{email}']")));
                    pendingReviewOption.Click();
                    Thread.Sleep(defaultLongWait);
                    IWebElement passwordInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("i0118")));
                    passwordInput.Clear();
                    passwordInput.SendKeys(password);
                    Thread.Sleep(defaultShortWait);

                    // คลิกปุ่ม Sign in
                    IWebElement signInButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
                    signInButton.Click();
                    Thread.Sleep(defaultLongWait);

                    Console.WriteLine("รหัสยืนยันถูกส่งไปยังโทรศัพท์ของคุณแล้ว");
                    Console.WriteLine("กรุณาทำการยืนยันตัวตนที่โทรศัพท์ของคุณ...");

                    //กดปุ่ม Stay signed in
                    IWebElement staySignedInButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("idBtn_Back")));
                    staySignedInButton.Click();
                    Thread.Sleep(defaultLongWait);

                    // รอให้ผู้ใช้ทำการยืนยันที่โทรศัพท์เสร็จสิ้น
                    WaitForPhoneAuthentication(driver, wait);

                    Console.WriteLine("การยืนยันตัวตนเสร็จสมบูรณ์ กำลังดำเนินการต่อ...");


                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการล็อกอินด้วยอีเมล: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// รอให้ผู้ใช้ทำการยืนยันตัวตนที่โทรศัพท์เสร็จสิ้น
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        /// <param name="wait">WebDriverWait instance</param>
        private void WaitForPhoneAuthentication(IWebDriver driver, WebDriverWait wait)
        {
            try
            {
                Console.WriteLine("กำลังรอการยืนยันจากโทรศัพท์...");

                // รอจนกว่า URL จะเปลี่ยนหรือพบ element ที่บ่งบอกว่าการยืนยันสำเร็จ
                wait.Until(d =>
                {
                    try
                    {
                        // ตรวจสอบว่า URL เปลี่ยนไปเป็นหน้าหลักหรือไม่ (หลังจากล็อกอินสำเร็จ)
                        string currentUrl = d.Url.ToLower();
                        
                        // ตรวจสอบหลายเงื่อนไขที่บ่งบอกว่าล็อกอินสำเร็จแล้ว
                        if (currentUrl.Contains("dashboard") || 
                            currentUrl.Contains("home") || 
                            currentUrl.Contains("main") ||
                            !currentUrl.Contains("login.microsoftonline.com"))
                        {
                            return true;
                        }

                        // ตรวจสอบว่ามี element ที่บ่งบอกว่าเข้าสู่ระบบสำเร็จหรือไม่
                        var successElements = d.FindElements(By.XPath("//button/span/a[text()='Logout']"));
                        if (successElements.Any())
                        {
                            return true;
                        }

                        // ตรวจสอบว่าไม่มีหน้า authentication อีกแล้ว
                        var authElements = d.FindElements(By.XPath("//div[contains(@class, 'sign-in')] | //div[contains(@class, 'auth')] | //input[@type='tel']"));
                        if (!authElements.Any())
                        {
                            return true;
                        }

                        return false;
                    }
                    catch
                    {
                        return false;
                    }
                });

                Console.WriteLine("ตรวจพบการยืนยันสำเร็จแล้ว!");
            }
            catch (TimeoutException)
            {
                Console.WriteLine("หมดเวลารอการยืนยันจากโทรศัพท์ กรุณาลองใหม่อีกครั้ง");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการรอการยืนยัน: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// ดาวน์โหลดไฟล์ .ok และ .err จากตารางผลลัพธ์และแปลงเป็น .xlsx แล้วย้ายไปยังโฟลเดอร์ปลายทาง
        /// พร้อมการลูปรอไฟล์ถ้าไม่พบ และตรวจสอบ error page
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        /// <param name="wait">WebDriverWait instance</param>
        /// <param name="destinationPath">โฟลเดอร์ปลายทางสำหรับเก็บไฟล์</param>
        /// <param name="maxRetryMinutes">จำนวนนาทีสูงสุดที่จะลูปรอไฟล์ (default = 10 นาที)</param>
        public void DownloadAndConvertResultFiles(IWebDriver driver, WebDriverWait wait, string destinationPath, int maxRetryMinutes = 15)
        {
            try
            {
                // ตรวจสอบสถานะหน้าเว็บก่อนเริ่มดาวน์โหลด
                string currentUrl = driver.Url;
                string pageTitle = driver.Title;
                
                Console.WriteLine($"🔍 ตรวจสอบหน้าเว็บก่อนดาวน์โหลด: URL={currentUrl}, Title={pageTitle}");
                
                // ตรวจสอบว่าเป็นหน้า error หรือไม่
                if (currentUrl.Contains("error") || pageTitle.ToLower().Contains("error") || 
                    pageTitle.ToLower().Contains("bad request") || currentUrl.Contains("me.sh"))
                {
                    throw new Exception($"หน้าเว็บเป็นหน้า error: URL={currentUrl}, Title={pageTitle}");
                }

                // สร้างโฟลเดอร์ปลายทางถ้ายังไม่มี
                if (!Directory.Exists(destinationPath))
                {
                    Directory.CreateDirectory(destinationPath);
                    Console.WriteLine($"สร้างโฟลเดอร์: {destinationPath}");
                }

                // รอให้ตารางโหลดเสร็จ
                Thread.Sleep(defaultLongWait);
                
                // ลองหาตารางด้วย timeout ที่เพิ่มขึ้น
                IWebElement resultTable = null;
                try
                {
                    resultTable = wait.Until(ExpectedConditions.ElementExists(By.XPath("//table[@id='datatableResultChangePro']//td[@class='ng-star-inserted']")));
                    Thread.Sleep(defaultShortWait);
                }
                catch (WebDriverTimeoutException)
                {
                    // หากหาตารางไม่เจอ ให้ตรวจสอบว่าเป็นหน้า error หรือไม่
                    string errorCheckUrl = driver.Url;
                    string errorCheckTitle = driver.Title;
                    
                    if (errorCheckUrl.Contains("error") || errorCheckTitle.ToLower().Contains("error") || 
                        errorCheckTitle.ToLower().Contains("bad request") || errorCheckUrl.Contains("me.sh"))
                    {
                        throw new Exception($"หน้าเว็บเปลี่ยนเป็นหน้า error ระหว่างรอตาราง: URL={errorCheckUrl}, Title={errorCheckTitle}");
                    }
                    else
                    {
                        throw new Exception("ไม่พบตารางผลลัพธ์ในเวลาที่กำหนด");
                    }
                }

                bool filesFound = false;
                int retryCount = 0;
                DateTime startTime = DateTime.Now;

                Console.WriteLine($"เริ่มค้นหาไฟล์ .ok และ .err (รอสูงสุด {maxRetryMinutes} นาที)");

                while (!filesFound && retryCount < maxRetryMinutes)
                {
                    Console.WriteLine($"ครั้งที่ {retryCount + 1}: กำลังค้นหาไฟล์...");

                    // ตรวจสอบสถานะหน้าเว็บอีกครั้งในแต่ละรอบ
                    try
                    {
                        string loopUrl = driver.Url;
                        string loopTitle = driver.Title;
                        
                        if (loopUrl.Contains("error") || loopTitle.ToLower().Contains("error") || 
                            loopTitle.ToLower().Contains("bad request") || loopUrl.Contains("me.sh"))
                        {
                            throw new Exception($"หน้าเว็บเปลี่ยนเป็นหน้า error ระหว่างการค้นหาไฟล์: URL={loopUrl}, Title={loopTitle}");
                        }
                    }
                    catch (Exception urlEx)
                    {
                        Console.WriteLine($"⚠️ ไม่สามารถตรวจสอบ URL ได้: {urlEx.Message}");
                    }

                    // หาไฟล์ .ok และ .err โดยไม่รวมไฟล์ที่มีคำว่า RERUN
                    var allOkFiles = driver.FindElements(By.XPath("//table[@id='datatableResultChangePro']//td[contains(text(),'.ok')]"));
                    var allErrFiles = driver.FindElements(By.XPath("//table[@id='datatableResultChangePro']//td[contains(text(),'.err')]"));

                    // กรองไฟล์ที่ไม่มีคำว่า "RERUN"
                    var okFiles = allOkFiles.Where(file => !file.Text.Contains("RERUN")).ToList();
                    var errFiles = allErrFiles.Where(file => !file.Text.Contains("RERUN")).ToList();

                    Console.WriteLine($"พบไฟล์ทั้งหมด: .ok: {allOkFiles.Count} ไฟล์, .err: {allErrFiles.Count} ไฟล์");
                    Console.WriteLine($"ไฟล์ที่ไม่มี RERUN: .ok: {okFiles.Count} ไฟล์, .err: {errFiles.Count} ไฟล์");

                    if (okFiles.Any() || errFiles.Any())
                    {
                        filesFound = true;
                        Console.WriteLine($"พบไฟล์ที่เหมาะสม! .ok: {okFiles.Count} ไฟล์, .err: {errFiles.Count} ไฟล์");

                        // ดาวน์โหลดไฟล์ .ok
                        if (okFiles.Any())
                        {
                            Console.WriteLine($"กำลังดาวน์โหลดไฟล์ .ok จำนวน {okFiles.Count} ไฟล์");
                            foreach (var okFile in okFiles)
                            {
                                try
                                {
                                    string originalFileName = okFile.Text;
                                    Console.WriteLine($"กำลังดาวน์โหลดไฟล์: {originalFileName}");

                                    // ตรวจสอบอีกครั้งว่าไฟล์ไม่มีคำว่า RERUN
                                    if (originalFileName.Contains("RERUN"))
                                    {
                                        Console.WriteLine($"ข้ามไฟล์ที่มี RERUN: {originalFileName}");
                                        continue;
                                    }

                                    // คลิกเพื่อดาวน์โหลด
                                    okFile.Click();
                                    Thread.Sleep(defaultLongWait);

                                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalFileName);
                                    string FileOk = $"{fileNameWithoutExtension}_ok";
                                    
                                    // รอให้ดาวน์โหลดเสร็จและแปลงไฟล์
                                    ConvertDownloadedFileToXlsx(originalFileName, FileOk, destinationPath);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"เกิดข้อผิดพลาดในการดาวน์โหลดไฟล์ .ok: {ex.Message}");
                                }
                            }
                        }

                        // ดาวน์โหลดไฟล์ .err
                        if (errFiles.Any())
                        {
                            Console.WriteLine($"กำลังดาวน์โหลดไฟล์ .err จำนวน {errFiles.Count} ไฟล์");
                            foreach (var errFile in errFiles)
                            {
                                try
                                {
                                    string originalFileName = errFile.Text;
                                    Console.WriteLine($"กำลังดาวน์โหลดไฟล์: {originalFileName}");

                                    // ตรวจสอบอีกครั้งว่าไฟล์ไม่มีคำว่า RERUN
                                    if (originalFileName.Contains("RERUN"))
                                    {
                                        Console.WriteLine($"ข้ามไฟล์ที่มี RERUN: {originalFileName}");
                                        continue;
                                    }

                                    // คลิกเพื่อดาวน์โหลด
                                    errFile.Click();
                                    Thread.Sleep(defaultLongWait);

                                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalFileName);
                                    string FileErr = $"{fileNameWithoutExtension}_err";
                                    
                                    // รอให้ดาวน์โหลดเสร็จและแปลงไฟล์
                                    ConvertDownloadedFileToXlsx(originalFileName, FileErr, destinationPath);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"เกิดข้อผิดพลาดในการดาวน์โหลดไฟล์ .err: {ex.Message}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // ไม่พบไฟล์
                        retryCount++;
                        TimeSpan elapsed = DateTime.Now - startTime;
                        
                        Console.WriteLine($"ไม่พบไฟล์ .ok หรือ .err ในตาราง (เวลาผ่านไป: {elapsed.TotalMinutes:F1} นาที)");
                        
                        if (retryCount < maxRetryMinutes)
                        {
                            Console.WriteLine($"รอ 1 นาที แล้วลองใหม่... (เหลืออีก {maxRetryMinutes - retryCount} ครั้ง)");
                            
                            // รอ 1 นาที (60,000 มิลลิวินาที)
                            Thread.Sleep(60000);
                            
                            // รีเฟรชหน้าเว็บเพื่อตรวจสอบไฟล์ใหม่
                            try
                            {
                                driver.Navigate().Refresh();
                                Thread.Sleep(defaultLongWait);
                                
                                // ตรวจสอบสถานะหน้าเว็บหลังรีเฟรช
                                string refreshUrl = driver.Url;
                                string refreshTitle = driver.Title;
                                
                                if (refreshUrl.Contains("error") || refreshTitle.ToLower().Contains("error") || 
                                    refreshTitle.ToLower().Contains("bad request") || refreshUrl.Contains("me.sh"))
                                {
                                    throw new Exception($"หน้าเว็บเปลี่ยนเป็นหน้า error หลังรีเฟรช: URL={refreshUrl}, Title={refreshTitle}");
                                }
                                
                                // รอให้ตารางโหลดใหม่
                                resultTable = wait.Until(ExpectedConditions.ElementExists(By.XPath("//table[@id='datatableResultChangePro']//td[@class='ng-star-inserted']")));
                                Thread.Sleep(defaultShortWait);
                            }
                            catch (Exception refreshEx)
                            {
                                Console.WriteLine($"⚠️ เกิดปัญหาในการรีเฟรชหน้า: {refreshEx.Message}");
                                
                                // หากเป็นปัญหาเกี่ยวกับ error page ให้หยุดการลูป
                                if (refreshEx.Message.Contains("error") || refreshEx.Message.Contains("bad request"))
                                {
                                    throw new Exception($"หน้าเว็บเปลี่ยนเป็นหน้า error - หยุดการรอไฟล์: {refreshEx.Message}");
                                }
                            }
                        }
                    }
                }

                if (!filesFound)
                {
                    TimeSpan totalElapsed = DateTime.Now - startTime;
                    Console.WriteLine($"ไม่พบไฟล์ .ok หรือ .err หลังจากรอทั้งหมด {totalElapsed.TotalMinutes:F1} นาที");
                    Console.WriteLine("โปรดตรวจสอบว่าการประมวลผลในเว็บเสร็จสิ้นแล้วหรือไม่");
                }
                else
                {
                    //TimeSpan totalElapsed = DateTime.Now - startTime;
                    //Console.WriteLine($"ดาวน์โหลดไฟล์เสร็จสิ้น (ใช้เวลาทั้งหมด: {totalElapsed.TotalMinutes:F1} นาที)");
                    Console.WriteLine($"ดาวน์โหลดไฟล์เสร็จสิ้น");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการดาวน์โหลดไฟล์: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// แปลงไฟล์ที่ดาวน์โหลดมาเป็น .xlsx โดยอ่านข้อมูล pipe-delimited และย้ายไปยังโฟลเดอร์ปลายทาง
        /// </summary>
        /// <param name="originalFileName">ชื่อไฟล์เดิม</param>
        /// <param name="FileName">ชื่อไฟล์ใหม่ที่ต้องการ (ไม่ต้องมีนามสกุล)</param>
        /// <param name="destinationPath">โฟลเดอร์ปลายทางสำหรับเก็บไฟล์</param>
        public void ConvertDownloadedFileToXlsx(string originalFileName, string FileName, string destinationPath)
        {
            try
            {
                // กำหนดโฟลเดอร์ดาวน์โหลดเริ่มต้นของ Windows
                string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

                // รอให้ไฟล์ดาวน์โหลดเสร็จ
                Thread.Sleep(3000);

                // หาไฟล์ที่ดาวน์โหลดล่าสุด
                string downloadedFilePath = FindLatestDownloadedFile(downloadsPath, originalFileName);

                if (!string.IsNullOrEmpty(downloadedFilePath))
                {
                    // สร้างชื่อไฟล์ใหม่เป็น .xlsx
                    string xlsxFileName = Path.GetFileNameWithoutExtension(FileName) + ".xlsx";
                    string xlsxFilePath = Path.Combine(destinationPath, xlsxFileName);

                    // อ่านไฟล์และแปลงเป็น Excel
                    ConvertPipeDelimitedToExcel(downloadedFilePath, xlsxFilePath);

                    Console.WriteLine($"แปลงและย้ายไฟล์สำเร็จ: {xlsxFilePath}");

                    // ลบไฟล์เดิมจากโฟลเดอร์ Downloads
                    try
                    {
                        File.Delete(downloadedFilePath);
                        Console.WriteLine($"ลบไฟล์เดิมจาก Downloads: {Path.GetFileName(downloadedFilePath)}");
                    }
                    catch (Exception deleteEx)
                    {
                        Console.WriteLine($"ไม่สามารถลบไฟล์เดิมได้: {deleteEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"ไม่พบไฟล์ที่ดาวน์โหลด: {originalFileName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการแปลงและย้ายไฟล์: {ex.Message}");
            }
        }

        /// <summary>
        /// แปลงไฟล์ที่มีข้อมูล pipe-delimited (คั่นด้วย |) เป็นไฟล์ Excel
        /// </summary>
        /// <param name="sourceFilePath">พาธของไฟล์ต้นทาง</param>
        /// <param name="targetFilePath">พาธของไฟล์ Excel ปลายทาง</param>
        private void ConvertPipeDelimitedToExcel(string sourceFilePath, string targetFilePath)
        {
            try
            {
                // กำหนด License Context สำหรับ EPPlus (NonCommercial)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // อ่านข้อมูลจากไฟล์ต้นทาง
                string[] lines = File.ReadAllLines(sourceFilePath, Encoding.UTF8);

                if (lines.Length == 0)
                {
                    Console.WriteLine("ไฟล์ต้นทางไม่มีข้อมูล");
                    return;
                }

                // สร้างไฟล์ Excel ใหม่
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Data");

                    int rowIndex = 1;
                    foreach (string line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        // แยกข้อมูลด้วย pipe (|)
                        string[] columns = line.Split('|');

                        // ใส่ข้อมูลในแต่ละคอลัมน์
                        for (int colIndex = 0; colIndex < columns.Length; colIndex++)
                        {
                            string cellValue = columns[colIndex].Trim();
                            
                            // ลองแปลงเป็นตัวเลขก่อน หากไม่ได้ให้ใส่เป็น text
                            if (double.TryParse(cellValue, out double numericValue))
                            {
                                worksheet.Cells[rowIndex, colIndex + 1].Value = numericValue;
                            }
                            else if (DateTime.TryParse(cellValue, out DateTime dateValue))
                            {
                                worksheet.Cells[rowIndex, colIndex + 1].Value = dateValue;
                                worksheet.Cells[rowIndex, colIndex + 1].Style.Numberformat.Format = "dd/mm/yyyy";
                            }
                            else
                            {
                                worksheet.Cells[rowIndex, colIndex + 1].Value = cellValue;
                            }
                        }
                        rowIndex++;
                    }

                    // AutoFit columns
                    if (worksheet.Dimension != null)
                    {
                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    }

                    // บันทึกไฟล์
                    package.SaveAs(new FileInfo(targetFilePath));
                }

                Console.WriteLine($"แปลงไฟล์สำเร็จ: {lines.Length} บรรทัด -> {targetFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการแปลงไฟล์: {ex.Message}");
                throw;
            }
        }


        /// <summary>
        /// หาไฟล์ที่ดาวน์โหลดล่าสุดในโฟลเดอร์
        /// </summary>
        /// <param name="downloadsPath">พาธของโฟลเดอร์ดาวน์โหลด</param>
        /// <param name="fileName">ชื่อไฟล์ที่ค้นหา</param>
        /// <returns>พาธของไฟล์ที่พบ หรือ null</returns>
        private string FindLatestDownloadedFile(string downloadsPath, string fileName)
        {
            try
            {
                if (!Directory.Exists(downloadsPath))
                {
                    Console.WriteLine($"โฟลเดอร์ดาวน์โหลดไม่พบ: {downloadsPath}");
                    return null;
                }

                // หาไฟล์ที่ชื่อตรงกับที่ระบุ
                var files = Directory.GetFiles(downloadsPath)
                                    .Where(file => Path.GetFileName(file).Equals(fileName, StringComparison.OrdinalIgnoreCase) ||
                                                   Path.GetFileName(file).Contains(Path.GetFileNameWithoutExtension(fileName)))
                                    .OrderByDescending(File.GetLastWriteTime)
                                    .ToList();

                if (files.Any())
                {
                    string foundFile = files.First();
                    Console.WriteLine($"พบไฟล์: {foundFile}");
                    return foundFile;
                }
                else
                {
                    Console.WriteLine($"ไม่พบไฟล์: {fileName}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการค้นหาไฟล์: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// จัดการกับ element ที่ขึ้นมาเพื่อขอส่งรหัสยืนยันใหม่ผ่าน Microsoft Authenticator
        /// </summary>
        /// <param name="driver">WebDriver instance</param>
        /// <param name="wait">WebDriverWait instance</param>
        private void HandleAuthenticatorApprovalRequest(IWebDriver driver, WebDriverWait wait)
        {
            try
            {
                // รอ element ที่อาจขึ้นมาเพื่อส่งรหัสยืนยันใหม่
                var approvalElements = driver.FindElements(By.XPath("//div[text()='Approve a request on my Microsoft Authenticator app']"));
                
                if (approvalElements.Any())
                {
                    Console.WriteLine("พบตัวเลือกส่งรหัสยืนยันใหม่ผ่าน Microsoft Authenticator");
                    Console.WriteLine("กำลังคลิกเพื่อส่งรหัสยืนยันใหม่...");
                    
                    // คลิก element เพื่อส่งรหัสยืนยันใหม่
                    approvalElements.First().Click();
                    Thread.Sleep(defaultLongWait);
                    
                    Console.WriteLine("ส่งคำขอยืนยันใหม่แล้ว กรุณาตรวจสอบ Microsoft Authenticator app ของคุณ");
                }
                else
                {
                    // ตรวจสอบ element ทางเลือกอื่น ๆ ที่อาจขึ้นมา
                    var alternativeElements = driver.FindElements(By.XPath("//div[contains(text(), 'Microsoft Authenticator') or contains(text(), 'Approve') or contains(text(), 'request')]"));
                    
                    if (alternativeElements.Any())
                    {
                        foreach (var element in alternativeElements.Take(3)) // ตรวจสอบแค่ 3 element แรก
                        {
                            string elementText = element.Text;
                            if (!string.IsNullOrEmpty(elementText) && 
                                (elementText.Contains("Approve") || elementText.Contains("Microsoft Authenticator")))
                            {
                                Console.WriteLine($"พบ element ที่เกี่ยวข้อง: '{elementText}'");
                                Console.WriteLine("กำลังคลิกเพื่อส่งรหัสยืนยันใหม่...");
                                
                                element.Click();
                                Thread.Sleep(defaultLongWait);
                                break;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("ไม่พบ element สำหรับส่งรหัสยืนยันใหม่ ดำเนินการต่อตามปกติ");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"เกิดข้อผิดพลาดในการจัดการ Authenticator approval request: {ex.Message}");
                Console.WriteLine("ดำเนินการต่อตามปกติ...");
            }
        }
    }
    
}

