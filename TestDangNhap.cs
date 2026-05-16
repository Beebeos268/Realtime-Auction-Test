using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using OpenQA.Selenium.Interactions;
using ExcelDataReader;
using System.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;

namespace ICareAutomation.Tests
{
    [TestFixture]
    public class TestDangNhap
    {
        private IWebDriver _driver;
        private readonly string _url = "http://localhost:4200/login";

        [SetUp]
        public void Setup()
        {
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            var options = new ChromeOptions();
            options.AddArgument("--remote-allow-origins=*");
            options.AddArgument("--disable-web-security");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--disable-gpu");

            string tempProfilePath = Path.Combine(Path.GetTempPath(), "Selenium_Auction_" + Guid.NewGuid().ToString());
            options.AddArgument($"--user-data-dir={tempProfilePath}");

            _driver = new ChromeDriver(service, options);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            _driver.Manage().Window.Maximize();
        }

        [TearDown]
        public void Teardown()
        {
            if (_driver != null)
            {
                try
                {
                    _driver.Quit();
                }
                catch (Exception) { }
                finally
                {
                    _driver.Dispose();
                    _driver = null;
                }
            }
        }

        public static IEnumerable<TestCaseData> ReadExcelData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LoginTestData.xlsx");

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var table = reader.AsDataSet().Tables[0];
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    if (row[0] == null || string.IsNullOrEmpty(row[0].ToString().Trim())) continue;

                    yield return new TestCaseData(
                        row[2]?.ToString() == "(để trống)" ? "" : row[2]?.ToString()?.Trim(),
                        row[3]?.ToString() == "(để trống)" ? "" : row[3]?.ToString()?.Trim(),
                        row[4]?.ToString(), // Expected_Message
                        row[1]?.ToString()  // Scenario_Name
                    );
                }
            }
        }

        private string CleanTextAbsolute(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            string text = input.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
            text = Regex.Replace(text, @"[^\w\s]", "");
            text = Regex.Replace(text, @"\s+", " ").Trim();
            return text.ToLower();
        }

        [Test, TestCaseSource(nameof(ReadExcelData))]
        public void ExecuteLoginTest(string email, string password, string expectedMessage, string scenarioName)
        {
            try
            {
                _driver.Navigate().GoToUrl(_url);
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                // 1. Nhập trường Email
                var emailElem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@type='email'] | //input[contains(@placeholder, 'Email')] | //input[@formcontrolname='email']")));
                emailElem.Clear();
                if (!string.IsNullOrEmpty(email))
                {
                    emailElem.SendKeys(email);
                }

                // 2. Nhập trường Mật khẩu
                var passElem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//input[@type='password'] | //input[@formcontrolname='password']")));
                passElem.Clear();
                if (!string.IsNullOrEmpty(password))
                {
                    passElem.SendKeys(password);
                }

                // FIX TRIỆT ĐỂ LUỒNG CLICK CON MẮT HIỆN MẬT KHẨU
                if (scenarioName.ToLower().Contains("L08") || scenarioName.Contains("Ẩn/hiện") || scenarioName.Contains("mật khẩu"))
                {
                    try
                    {
                        // Lấy kích thước và vị trí của ô Input password thực tế trên màn hình
                        int inputWidth = passElem.Size.Width;

                        // Sử dụng Actions di chuyển chuột đến rìa bên phải của ô mật khẩu (nơi chứa icon mắt)
                        // Lùi lại 25 pixel để chắc chắn click trúng tâm của icon mắt
                        Actions actionClickEye = new Actions(_driver);
                        actionClickEye.MoveToElement(passElem, inputWidth / 2 - 25, 0).Click().Build().Perform();

                        Thread.Sleep(1000); // Chờ 1 giây cho hiệu ứng Angular chuyển text hoàn tất
                    }
                    catch (Exception eyeEx)
                    {
                        Console.WriteLine("Cảnh báo tương tác tọa độ mắt thất bại, thử lại bằng DOM: " + eyeEx.Message);
                        try
                        {
                            var eyeIconSelector = By.XPath("//input[@type='password']/following-sibling::*[local-name()='svg'] | //input[@type='password']/parent::div//*[contains(@class,'eye')] | //span[contains(@class, 'eye')]");
                            var eyeIcon = _driver.FindElement(eyeIconSelector);
                            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
                            js.ExecuteScript("arguments[0].click();", eyeIcon);
                            Thread.Sleep(1000);
                        }
                        catch { }
                    }
                }

                // 3. Click nút Đăng Nhập
                var loginBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//button[contains(., 'Đăng Nhập')] | //button[@type='submit']")));
                loginBtn.Click();

                // 4. CHUẨN HÓA PHÂN LUỒNG BẰNG CHUỖI MONG MUỐN (EXPECTED MESSAGE)
                string cleanExpected = CleanTextAbsolute(expectedMessage);

                // Cứ case nào mong muốn là thành công hoặc chào mừng thì chuyển thẳng sang luồng check URL nhảy trang
                bool isSuccessCase = cleanExpected.Contains("thanh cong") || cleanExpected.Contains("chao mung") || scenarioName.ToLower().Contains("l01") || scenarioName.ToLower().Contains("l08");

                if (isSuccessCase)
                {
                    // LUỒNG THÀNH CÔNG: Kiểm tra xem URL có thay đổi hoặc chứa endpoint trang chủ không
                    bool isNavigated = wait.Until(d => d.Url != _url || d.Url.Contains("/dashboard") || d.Url.Contains("/home"));
                    Assert.That(isNavigated, Is.True, $"Kịch bản [{scenarioName}] đăng nhập thành công nhưng URL không thực hiện điều hướng trang.");
                }
                else
                {
                    // LUỒNG THẤT BẠI: Đứng đợi Toastr lỗi render trên màn hình UI
                    var toastSelector = By.XPath("//div[contains(@class, 'toast-message')] | //div[contains(@class, 'toast-title')] | //div[contains(@id, 'toast-container')]//*");
                    string actualNotificationText = "";
                    try
                    {
                        var toastElem = wait.Until(ExpectedConditions.ElementIsVisible(toastSelector));
                        actualNotificationText = toastElem.Text;
                    }
                    catch (WebDriverTimeoutException)
                    {
                        try
                        {
                            actualNotificationText = _driver.FindElement(By.XPath("//*[contains(@class, 'toast')]")).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            actualNotificationText = _driver.FindElement(By.TagName("body")).Text;
                        }
                    }

                    string cleanActual = CleanTextAbsolute(actualNotificationText);

                    Assert.That(cleanActual, Does.Contain(cleanExpected),
                        $"\n[Thất bại kịch bản: {scenarioName}]\nThông báo không trùng khớp.\nMong muốn: '{cleanExpected}'\nThực tế bắt được trên UI: '{cleanActual}'\n");
                }
            }
            catch (Exception ex)
            {
                Assert.Fail($"Kịch bản [{scenarioName}] dừng do lỗi hệ thống: {ex.Message}");
            }
        }
    }
}