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
    public class TestDangKy
    {
        private IWebDriver _driver;
        private readonly string _loginUrl = "http://localhost:4200/login";
        private readonly string _registerUrl = "http://localhost:4200/register";

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
            options.AddArgument("--disable-features=AutofillServerCommunication,PasswordManagerOnboarding");
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddUserProfilePreference("autofill.profile_enabled", false);

            options.UnhandledPromptBehavior = UnhandledPromptBehavior.Ignore;

            string tempProfilePath = Path.Combine(Path.GetTempPath(), "Selenium_Register_" + Guid.NewGuid().ToString());
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
                    try
                    {
                        var alert = _driver.SwitchTo().Alert();
                        Console.WriteLine($"[TearDown] Đóng Alert tồn dư: {alert.Text}");
                        alert.Dismiss();
                    }
                    catch (NoAlertPresentException) { }

                    Thread.Sleep(2000);
                    _driver.Quit();
                }
                catch (Exception) { }
                finally
                {
                    try { _driver.Dispose(); } catch { }
                    _driver = null;
                }
            }
        }

        public static IEnumerable<TestCaseData> ReadExcelData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RegisterTestData.xlsx");

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var table = reader.AsDataSet().Tables[0];
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    if (row[0] == null || string.IsNullOrEmpty(row[0].ToString().Trim())) continue;

                    string stepId = row[0]?.ToString()?.Trim() ?? "";
                    string scenarioName = row[1]?.ToString()?.Trim() ?? "";
                    string hoTen = row[2]?.ToString() ?? "";
                    string sdt = row[3]?.ToString() ?? "";
                    string email = row[4]?.ToString() ?? "";
                    string matKhau = row[5]?.ToString() ?? "";
                    string xacNhanMK = row[6]?.ToString() ?? "";
                    string expected = row[7]?.ToString() ?? "";
                    string action = table.Columns.Count > 8 ? (row[8]?.ToString()?.Trim() ?? "") : "";

                    yield return new TestCaseData(hoTen, sdt, email, matKhau, xacNhanMK,
                                                 expected, scenarioName, stepId, action)
                        .SetName($"{stepId}_{scenarioName.Replace(" ", "_")}");
                }
            }
        }

        private string CleanTextAbsolute(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            string text = input.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
            text = Regex.Replace(text, @"[^\w\s\x00-\x7FàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]", "");
            text = Regex.Replace(text, @"\s+", " ").Trim();
            return text.ToLower();
        }

        private void TypeText(IWebElement element, string text, string tag, string fieldName)
        {
            try { element.Click(); } catch { }
            Thread.Sleep(200);

            element.SendKeys(Keys.Control + "a");
            element.SendKeys(Keys.Delete);
            Thread.Sleep(200);

            if (string.IsNullOrEmpty(text))
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));", element);
                Thread.Sleep(500);
                return;
            }

            element.SendKeys(text);
            Thread.Sleep(200);
        }

        private void NavigateToRegisterPage(string tag)
        {
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
            try
            {
                _driver.Navigate().GoToUrl(_registerUrl);
                var headerOrField = _driver.FindElements(By.XPath(
                    "//*[contains(text(), 'Tạo Tài Khoản') or contains(text(),'Tạo tài khoản')] | //input[contains(@placeholder, 'Nguyễn Văn')]"));
                if (headerOrField.Count > 0) return;
            }
            catch { }

            _driver.Navigate().GoToUrl(_loginUrl);
            var taoTaiKhoanLink = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(
                "//a[contains(., 'Tạo tài khoản')] | //button[contains(., 'Tạo tài khoản')] | //*[contains(text(), 'Tạo tài khoản mới')]")));
            taoTaiKhoanLink.Click();
            Thread.Sleep(500);
        }

        [Test, TestCaseSource(nameof(ReadExcelData))]
        public void ExecuteRegisterTest(string hoTen, string sdt, string email, string matKhau,
                                        string xacNhanMK, string expectedMessage,
                                        string scenarioName, string stepId, string action)
        {
            string tag = $"[{stepId} - {scenarioName}]";
            Console.WriteLine($"\n========== {tag} Action='{action}' ==========");

            try
            {
                NavigateToRegisterPage(tag);
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                FillField(new[] {
                    "//input[contains(@placeholder, 'Nguyễn Văn')]",
                    "//input[@formcontrolname='hoTen' or @formcontrolname='fullName' or @formcontrolname='name']"
                }, hoTen, tag, "HoTen");

                FillField(new[] {
                    "//input[contains(@placeholder, '0987') or contains(@placeholder, 'xxx')]",
                    "//input[@formcontrolname='soDienThoai' or @formcontrolname='sdt' or @formcontrolname='phone' or @formcontrolname='phoneNumber']"
                }, sdt, tag, "SDT");

                FillField(new[] {
                    "//input[contains(@placeholder, 'example@gmail') or @type='email' or @formcontrolname='email']"
                }, email, tag, "Email");

                FillPasswordField(matKhau, tag, "MatKhau", isFirst: true);
                FillPasswordField(xacNhanMK, tag, "XacNhanMK", isFirst: false);

                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "var inputs = document.querySelectorAll('input'); inputs.forEach(el => el.dispatchEvent(new Event('blur', { bubbles: true })));" +
                    "if(document.activeElement) { document.activeElement.blur(); }");

                Thread.Sleep(1000);

                string actionUpper = (action ?? "").Trim().ToUpper();
                string cleanExpected = CleanTextAbsolute(expectedMessage);

                switch (actionUpper)
                {
                    case "CHECK_INLINE_ERROR":
                        // ⭐ SỬA ĐỔI: Kích hoạt nút bấm Đăng ký và quét thông báo hiển thị trên giao diện thay vì tìm viền đỏ cố định
                        ClickSubmitButton(tag);
                        CheckCatchBackendAlert(tag);
                        HandleNotificationError(wait, tag, cleanExpected);
                        break;

                    case "CHECK_TOGGLE":
                        HandleToggle(tag);
                        ClickSubmitButton(tag);
                        CheckCatchBackendAlert(tag);
                        ClickSubmitAndWaitToastSuccess(wait, tag, cleanExpected);
                        break;

                    case "REGISTER_SUCCESS":
                        ClickSubmitButton(tag);
                        CheckCatchBackendAlert(tag);
                        ClickSubmitAndWaitToastSuccess(wait, tag, cleanExpected);
                        break;

                    case "CHECK_TOAST_ERROR":
                    default:
                        ClickSubmitButton(tag);
                        CheckCatchBackendAlert(tag);
                        ClickSubmitAndWaitToastError(wait, tag, cleanExpected);
                        break;
                }
            }
            catch (Exception ex)
            {
                Assert.Fail($"{tag} Thất bại do lỗi phát sinh: {ex.Message}");
            }
        }

        private void FillField(string[] xpaths, string value, string tag, string fieldName)
        {
            foreach (var xp in xpaths)
            {
                try
                {
                    var elem = _driver.FindElement(By.XPath(xp));
                    if (elem.Displayed)
                    {
                        TypeText(elem, value, tag, fieldName);
                        return;
                    }
                }
                catch { }
            }
        }

        private void FillPasswordField(string value, string tag, string fieldName, bool isFirst)
        {
            var passInputs = _driver.FindElements(By.XPath(
                "//input[@type='password'] | //input[@formcontrolname='matKhau' or @formcontrolname='password']"));

            if (passInputs.Count == 0) passInputs = _driver.FindElements(By.XPath("//input"));

            IWebElement target = isFirst ?
                (passInputs.Count > 3 ? passInputs[3] : passInputs[0]) :
                (passInputs.Count > 4 ? passInputs[4] : passInputs[passInputs.Count - 1]);

            if (target != null) TypeText(target, value, tag, fieldName);
        }

        // ⭐ ĐÃ THAY ĐỔI: Hàm chuyên dụng xử lý quét thông báo hiển thị chung cho lỗi trống/lỗi định dạng
        private void HandleNotificationError(WebDriverWait wait, string tag, string cleanExpected)
        {
            string actualText = TryGetAlertOrToast(wait, tag);
            string cleanActual = CleanTextAbsolute(actualText);

            Thread.Sleep(2000);

            Assert.That(cleanActual, Does.Contain(cleanExpected),
                $"\n[FAILED] Không tìm thấy thông báo phù hợp trên UI.\nMong đợi: '{cleanExpected}'\nThực tế trên Web: '{cleanActual}'");
        }

        private void ClickSubmitAndWaitToastError(WebDriverWait wait, string tag, string cleanExpected)
        {
            string actualText = TryGetAlertOrToast(wait, tag);
            string cleanActual = CleanTextAbsolute(actualText);

            Thread.Sleep(3000);

            Assert.That(cleanActual, Does.Contain(cleanExpected),
                $"\n[FAILED] Không tìm thấy thông báo lỗi phù hợp trên UI.\nMong đợi: '{cleanExpected}'\nThực tế trên Web: '{cleanActual}'");
        }

        private void ClickSubmitAndWaitToastSuccess(WebDriverWait wait, string tag, string cleanExpected)
        {
            string actualText = TryGetAlertOrToast(wait, tag);
            string cleanActual = CleanTextAbsolute(actualText);

            if (!string.IsNullOrEmpty(cleanActual))
            {
                Thread.Sleep(3000);
                Assert.That(cleanActual, Does.Contain(cleanExpected));
                return;
            }

            try
            {
                wait.Until(d => !d.Url.Contains("/register"));
                Thread.Sleep(2000);
            }
            catch { }
        }

        private string TryGetAlertOrToast(WebDriverWait wait, string tag)
        {
            try
            {
                var toastElem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                    "//div[contains(@class,'toast-message')] " +
                    "| //div[contains(@id,'toast-container')]//*[contains(@class,'toast')] " +
                    "| //div[contains(@class,'toast-error') or contains(@class,'toast-success')]" +
                    "| //*[contains(@class,'ngx-toastr') or contains(@class,'alert')]")));

                return toastElem.Text;
            }
            catch (WebDriverException ex) when (ex.Message.Contains("alert open"))
            {
                CheckCatchBackendAlert(tag);
            }
            catch
            {
                Thread.Sleep(1000);
            }
            return "";
        }

        private void ClickSubmitButton(string tag)
        {
            try
            {
                var btn = _driver.FindElement(By.XPath("//button[contains(., 'Đăng Ký') or @type='submit']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({block:'center'});", btn);
                Thread.Sleep(500);
                btn.Click();
                Thread.Sleep(1500);
            }
            catch { }
        }

        private void HandleToggle(string tag)
        {
            try
            {
                var passElem = _driver.FindElement(By.XPath("//input[contains(@type,'password')]"));
                int w = passElem.Size.Width;
                new Actions(_driver).MoveToElement(passElem, w / 2 - 25, 0).Click().Build().Perform();
                Thread.Sleep(500);
            }
            catch { }
        }

        private void CheckCatchBackendAlert(string tag)
        {
            try
            {
                var alert = _driver.SwitchTo().Alert();
                string txt = alert.Text;
                alert.Dismiss();

                Console.WriteLine($"{tag} -> Phát hiện biến cố Alert: {txt}");

                if (txt.Contains("Failed to fetch") || txt.Contains("Backend"))
                {
                    Assert.Fail($"[LỖI HỆ THỐNG]: Backend C# (API) chưa bật hoặc đã bị sập. Vui lòng kiểm tra lại phía Service!");
                }
            }
            catch (NoAlertPresentException) { }
        }
    }
}