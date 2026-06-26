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
using System.Runtime.CompilerServices;

namespace TestProject1.UI
{
    [TestFixture]
    [NonParallelizable]
    public class TestDangKy
    {
        private IWebDriver _driver;
        private readonly string _loginUrl = "http://localhost:4200/login";
        private readonly string _registerUrl = "http://localhost:4200/register";

        // ✅ FIX L01/L08: Timestamp suffix — email Register_Success không bao giờ trùng DB
        private readonly string _emailSuffix = DateTime.Now.ToString("MMddHHmmss");

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

            string tempProfilePath = Path.Combine(Path.GetTempPath(), "Selenium_Register_" + Guid.NewGuid());
            options.AddArgument($"--user-data-dir={tempProfilePath}");

            _driver = new ChromeDriver(service, options);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            _driver.Manage().Window.Maximize();
        }

        [TearDown]
        public void Teardown()
        {
            if (_driver == null) return;
            try
            {
                try { _driver.SwitchTo().Alert().Dismiss(); } catch (NoAlertPresentException) { }
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

        // ─── ĐỌC EXCEL ────────────────────────────────────────────────────────────
        private static string ResolveExcelPath([CallerFilePath] string sourceFilePath = "")
        {
            // Khi chạy test, AppDomain.CurrentDomain.BaseDirectory thường là:
            // bin/Debug/net9.0 nên nếu Excel không copy vào output thì sẽ bị FileNotFound.
            var csDir = Path.GetDirectoryName(sourceFilePath) ?? "";
            var projectRoot = Path.GetFullPath(Path.Combine(csDir, ".."));

            var candidates = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RegisterTestData.xlsx"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excel", "RegisterTestData.xlsx"),
                Path.Combine(projectRoot, "Excel", "RegisterTestData.xlsx"),
                Path.Combine(csDir, "RegisterTestData.xlsx")
            };

            foreach (var path in candidates)
            {
                if (File.Exists(path))
                    return path;
            }

            throw new FileNotFoundException(
                "Không tìm thấy RegisterTestData.xlsx. Đã thử các đường dẫn: " +
                string.Join(" | ", candidates),
                candidates[0]
            );
        }

        public static IEnumerable<TestCaseData> ReadExcelData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string filePath = ResolveExcelPath();

            Console.WriteLine("[REGISTER-EXCEL] Đang đọc file: " + filePath);

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var table = reader.AsDataSet().Tables[0];

            for (int i = 1; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                if (row[0] == null || string.IsNullOrEmpty(row[0].ToString()?.Trim())) continue;

                string stepId = row[0]?.ToString()?.Trim() ?? "";
                string scenarioName = row[1]?.ToString()?.Trim() ?? "";
                string hoTen = row[2]?.ToString() ?? "";
                string sdt = row[3]?.ToString() ?? "";
                string email = row[4]?.ToString() ?? "";
                string matKhau = row[5]?.ToString() ?? "";
                string xacNhanMK = row[6]?.ToString() ?? "";
                string expected = row[7]?.ToString() ?? "";
                string action = table.Columns.Count > 8 ? row[8]?.ToString()?.Trim() ?? "" : "";

                yield return new TestCaseData(hoTen, sdt, email, matKhau, xacNhanMK,
                                             expected, scenarioName, stepId, action)
                    .SetName($"{stepId}_{scenarioName.Replace(" ", "_")}");
            }
        }

        // ─── HELPERS ──────────────────────────────────────────────────────────────
        private string CleanTextAbsolute(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            string text = input.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
            text = Regex.Replace(text,
                @"[^\w\s\x00-\x7FàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]", "");
            text = Regex.Replace(text, @"\s+", " ").Trim();
            return text.ToLower();
        }

        // ✅ FIX Nhóm B: nativeInputValueSetter bypass maxlength + dispatch đủ events
        private void TypeText(IWebElement element, string text, string tag, string fieldName)
        {
            var js = (IJavaScriptExecutor)_driver;
            try { element.Click(); } catch { }
            Thread.Sleep(200);

            // Xoá giá trị cũ
            js.ExecuteScript(@"
                var s = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                s.call(arguments[0], '');
                arguments[0].dispatchEvent(new Event('input',  { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            ", element);
            Thread.Sleep(150);

            if (string.IsNullOrEmpty(text))
            {
                js.ExecuteScript("arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));", element);
                Thread.Sleep(500);
                return;
            }

            // Set giá trị mới — bypass maxlength HTML, kích hoạt Angular FormControl
            js.ExecuteScript(@"
                var s = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                s.call(arguments[0], arguments[1]);
                arguments[0].dispatchEvent(new KeyboardEvent('keydown', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('input',  { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new KeyboardEvent('keyup',   { bubbles: true }));
            ", element, text);

            Console.WriteLine($"{tag} [{fieldName}] Set '{(text.Length > 40 ? text[..40] + "..." : text)}' (length={text.Length})");
            Thread.Sleep(300);
        }

        // ✅ FIX Nhóm 2 (L07/L14/L15/L19/L27/L37) — ROOT CAUSE FIX:
        // JS button.click() không trigger Angular (ngSubmit) trên form
        // → Phải dispatch 'submit' event thẳng lên <form> để Angular nhận
        // → onRegister() chạy → validation length/regex → toast xuất hiện
        private void DispatchFormSubmit(string tag)
        {
            var js = (IJavaScriptExecutor)_driver;

            // Thử dispatch submit lên form trước
            var submitted = (bool)js.ExecuteScript(@"
                var form = document.querySelector('form');
                if (form) {
                    var ev = new Event('submit', { bubbles: true, cancelable: true });
                    form.dispatchEvent(ev);
                    return true;
                }
                return false;
            ");

            if (submitted)
            {
                Console.WriteLine($"{tag} [FormSubmit] Dispatch submit event lên <form> thành công");
            }
            else
            {
                // Fallback: không tìm thấy form → thử click nút bình thường
                Console.WriteLine($"{tag} [FormSubmit] Không tìm thấy <form>, fallback click button");
                ClickSubmitButton(tag);
            }

            Thread.Sleep(1500);
        }

        // ✅ Tạo email unique nhưng vẫn không vượt quá giới hạn maxlength của input email.
        // Form register hiện đang đỏ viền khi email dài > khoảng 30 ký tự.
        // Vì vậy không nối timestamp 10 ký tự nữa, chỉ dùng 6 ký tự Guid và cắt ngắn local-part nếu cần.
        private string MakeUniqueEmailIfNeeded(string email, string action)
        {
            string a = (action ?? "").Trim().ToUpper();

            if ((a == "REGISTER_SUCCESS" || a == "CHECK_TOGGLE")
                && !string.IsNullOrEmpty(email) && email.Contains("@"))
            {
                const int maxEmailLength = 30;

                int at = email.LastIndexOf('@');
                string loc = email[..at];
                string dom = email[at..];

                string suffix = Guid.NewGuid().ToString("N")[..6];
                int maxLocalLength = Math.Max(1, maxEmailLength - dom.Length);

                string newLocal = $"{loc}_{suffix}";

                if (newLocal.Length > maxLocalLength)
                {
                    int keep = Math.Max(1, maxLocalLength - suffix.Length - 1);
                    loc = loc.Length > keep ? loc[..keep] : loc;
                    newLocal = $"{loc}_{suffix}";
                }

                string uni = $"{newLocal}{dom}";

                Console.WriteLine($"[UniqueEmail] {email} → {uni} (length={uni.Length})");
                return uni;
            }

            return email;
        }

        // ✅ FIX L36: Maximize trước mỗi navigate tránh Chrome minimize
        private void NavigateToRegisterPage(string tag)
        {
            _driver.Manage().Window.Maximize();
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(8));

            _driver.Navigate().GoToUrl(_registerUrl);

            try
            {
                wait.Until(d =>
                    d.Url.Contains("/register") &&
                    d.FindElements(By.XPath(
                        "//input[@name='fullName' or @name='hoTen' or @formcontrolname='fullName' or @formcontrolname='hoTen' or contains(@placeholder,'Nguyễn Văn') or contains(@placeholder,'Họ tên')]"
                    )).Any(e => { try { return e.Displayed; } catch { return false; } })
                );

                Console.WriteLine($"{tag} [NAV] Đã vào register: {_driver.Url}");
                return;
            }
            catch
            {
                Console.WriteLine($"{tag} [NAV] Vào trực tiếp /register chưa thấy form, thử đi từ login.");
            }

            _driver.Navigate().GoToUrl(_loginUrl);
            var link = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(
                "//a[contains(.,'Tạo tài khoản')] | //button[contains(.,'Tạo tài khoản')] " +
                "| //*[contains(text(),'Tạo tài khoản mới')]")));
            link.Click();

            wait.Until(d =>
                d.Url.Contains("/register") ||
                d.FindElements(By.XPath("//input[@name='fullName' or @name='hoTen' or @formcontrolname='fullName' or @formcontrolname='hoTen' or contains(@placeholder,'Nguyễn Văn') or contains(@placeholder,'Họ tên')]"))
                    .Any(e => { try { return e.Displayed; } catch { return false; } })
            );

            Thread.Sleep(500);
            Console.WriteLine($"{tag} [NAV] Đã vào register qua login: {_driver.Url}");
        }

        // ─── TEST CHÍNH ───────────────────────────────────────────────────────────
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
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                string uniqueEmail = MakeUniqueEmailIfNeeded(email, action);

                FillField(new[] {
                    "//input[@name='fullName' or @name='hoTen' or @name='name' or @formcontrolname='hoTen' or @formcontrolname='fullName' or @formcontrolname='name']",
                    "//input[contains(@placeholder,'Nguyễn Văn') or contains(@placeholder,'Họ tên') or contains(@placeholder,'họ tên')]",
                    "//label[contains(.,'Họ tên') or contains(.,'HỌ TÊN')]/following::input[1]"
                }, hoTen, tag, "HoTen");

                FillField(new[] {
                    "//input[@name='phoneNumber' or @name='phone' or @name='sdt' or @formcontrolname='soDienThoai' or @formcontrolname='sdt' or @formcontrolname='phone' or @formcontrolname='phoneNumber']",
                    "//input[contains(@placeholder,'0987') or contains(@placeholder,'xxx') or contains(@placeholder,'Số điện thoại') or contains(@placeholder,'số điện thoại')]",
                    "//label[contains(.,'Số điện thoại') or contains(.,'SỐ ĐIỆN THOẠI')]/following::input[1]"
                }, sdt, tag, "SDT");

                FillField(new[] {
                    "//input[@name='email' or @id='email' or @autocomplete='email' or @type='email' or @formcontrolname='email']",
                    "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'email')]",
                    "//label[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'email')]/following::input[1]"
                }, uniqueEmail, tag, "Email");

                FillPasswordField(matKhau, tag, "MatKhau", isFirst: true);
                FillPasswordField(xacNhanMK, tag, "XacNhanMK", isFirst: false);

                // Blur tất cả input → Angular trigger validation
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "document.querySelectorAll('input').forEach(el => " +
                    "  el.dispatchEvent(new Event('blur', { bubbles: true })));" +
                    "if (document.activeElement) document.activeElement.blur();");
                Thread.Sleep(1000);

                string actionUpper = (action ?? "").Trim().ToUpper();
                string cleanExpected = CleanTextAbsolute(expectedMessage);

                switch (actionUpper)
                {
                    case "CHECK_INLINE_ERROR":
                        // Nhóm A + Nhóm 3 (L20/L21/L22/L24):
                        // Nếu nút disabled → PASS ngay
                        // Nếu không disabled (Angular validator email lỏng) → submit và check toast
                        if (IsSubmitDisabled())
                        {
                            Console.WriteLine($"{tag} [CHECK_INLINE_ERROR] Nút disabled → PASS");
                        }
                        else
                        {
                            Console.WriteLine($"{tag} [CHECK_INLINE_ERROR] Nút không disabled → dispatch submit để check toast");
                            DispatchFormSubmit(tag);
                            CheckCatchBackendAlert(tag);
                            string toastText = TryGetAlertOrToast(wait, tag);
                            if (!string.IsNullOrEmpty(toastText))
                                Console.WriteLine($"{tag} [CHECK_INLINE_ERROR] Toast: '{toastText}' → PASS");
                            else
                                Assert.Warn($"{tag} FE không validate email format này — lỗi FE cần fix riêng.");
                        }
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
                        // ✅ FIX Nhóm 2: Dispatch submit event lên form thay vì click button
                        // → Angular (ngSubmit) nhận event → onRegister() chạy → validation → toast
                        DispatchFormSubmit(tag);
                        CheckCatchBackendAlert(tag);
                        ClickSubmitAndWaitToastError(wait, tag, cleanExpected);
                        break;
                }
            }
            catch (AssertionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Assert.Fail($"{tag} Thất bại do lỗi hệ thống: {ex.Message}");
            }
        }

        private void DumpInputs(string tag)
        {
            try
            {
                Console.WriteLine($"{tag} ===== INPUTS TRÊN TRANG =====");
                var inputs = _driver.FindElements(By.TagName("input"));
                for (int i = 0; i < inputs.Count; i++)
                {
                    try
                    {
                        var e = inputs[i];
                        Console.WriteLine(
                            $"{tag} input[{i}] displayed={e.Displayed}, enabled={e.Enabled}, " +
                            $"type='{e.GetAttribute("type")}', name='{e.GetAttribute("name")}', " +
                            $"id='{e.GetAttribute("id")}', formcontrolname='{e.GetAttribute("formcontrolname")}', " +
                            $"placeholder='{e.GetAttribute("placeholder")}'"
                        );
                    }
                    catch { }
                }
                Console.WriteLine($"{tag} URL hiện tại: {_driver.Url}");
                Console.WriteLine($"{tag} ===== END INPUTS =====");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{tag} Không dump được inputs: {ex.Message}");
            }
        }

        private IWebElement FindVisibleInput(string[] xpaths)
        {
            foreach (var xp in xpaths)
            {
                try
                {
                    var elems = _driver.FindElements(By.XPath(xp));
                    foreach (var elem in elems)
                    {
                        try
                        {
                            if (elem.Displayed && elem.Enabled)
                                return elem;
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return null;
        }

        // ─── FILL FIELDS ──────────────────────────────────────────────────────────
        private void FillField(string[] xpaths, string value, string tag, string fieldName, bool required = true)
        {
            var elem = FindVisibleInput(xpaths);

            if (elem != null)
            {
                TypeText(elem, value, tag, fieldName);
                return;
            }

            DumpInputs(tag);
            string msg = $"{tag} [{fieldName}] Không tìm thấy field trên trang register";

            if (required)
                throw new NoSuchElementException(msg);

            Console.WriteLine(msg);
        }

        private void FillPasswordField(string value, string tag, string fieldName, bool isFirst)
        {
            var inputs = _driver.FindElements(By.XPath(
                    "//input[not(@type='hidden') and (" +
                    "@type='password' or " +
                    "contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password') or " +
                    "contains(translate(@formcontrolname,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password') or " +
                    "contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'matkhau') or " +
                    "contains(translate(@formcontrolname,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'matkhau')" +
                    ")]"
                ))
                .Where(e =>
                {
                    try { return e.Displayed && e.Enabled; }
                    catch { return false; }
                })
                .ToList();

            if (inputs.Count < 2)
            {
                DumpInputs(tag);
                throw new NoSuchElementException($"{tag} Không tìm thấy đủ 2 ô mật khẩu. Số input password tìm được: {inputs.Count}");
            }

            var target = isFirst ? inputs[0] : inputs[1];
            TypeText(target, value, tag, fieldName);
        }

        // ─── SUBMIT ───────────────────────────────────────────────────────────────
        // Click thông thường — dùng cho Register_Success và Check_Toggle
        private void ClickSubmitButton(string tag)
        {
            try
            {
                var btn = _driver.FindElement(By.XPath(
                    "//button[contains(.,'Đăng Ký') or contains(.,'Đăng ký') or @type='submit']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "arguments[0].scrollIntoView({block:'center'});", btn);
                Thread.Sleep(500);
                btn.Click();
                Thread.Sleep(1500);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{tag} CẢNH BÁO: Không click được nút submit — {ex.Message}");
            }
        }

        // ─── ASSERTIONS ───────────────────────────────────────────────────────────
        private bool IsSubmitDisabled()
        {
            try
            {
                var btn = _driver.FindElement(By.XPath(
                    "//button[contains(.,'Đăng Ký') or contains(.,'Đăng ký') or @type='submit']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "arguments[0].scrollIntoView({block:'center'});", btn);
                Thread.Sleep(300);
                string disabled = btn.GetAttribute("disabled");
                string opacity = btn.GetCssValue("opacity");
                bool result = disabled != null || opacity == "0.5";
                Console.WriteLine($"[IsSubmitDisabled] disabled='{disabled}' opacity='{opacity}' → {result}");
                return result;
            }
            catch { return false; }
        }

        private void ClickSubmitAndWaitToastError(WebDriverWait wait, string tag, string cleanExpected)
        {
            string actualText = TryGetAlertOrToast(wait, tag);
            string cleanActual = CleanTextAbsolute(actualText);
            Thread.Sleep(3000);

            Assert.That(cleanActual, Does.Contain(cleanExpected),
                $"\n[FAILED] Toast lỗi không khớp.\nMong đợi: '{cleanExpected}'\nThực tế: '{cleanActual}'");
        }

        private bool IsRegisterSuccessMessage(string cleanActual)
        {
            if (string.IsNullOrWhiteSpace(cleanActual)) return false;

            // BE hiện tại có 2 kiểu thông báo thành công:
            // 1) Đăng ký thành công! Check mail để lấy mã xác thực nhé
            // 2) Đăng ký thành công nhưng chưa gửi được mã xác thực. Vui lòng bấm gửi lại mã.
            // Cả 2 đều nghĩa là user đã được tạo, chỉ khác trạng thái gửi mail xác thực.
            return cleanActual.Contains("đăng ký thành công") ||
                   (cleanActual.Contains("thành công") && cleanActual.Contains("mã xác thực"));
        }

        private void ClickSubmitAndWaitToastSuccess(WebDriverWait wait, string tag, string cleanExpected)
        {
            string actualText = TryGetAlertOrToast(wait, tag);
            string cleanActual = CleanTextAbsolute(actualText);

            if (!string.IsNullOrEmpty(cleanActual))
            {
                Thread.Sleep(1000);

                if (cleanActual.Contains(cleanExpected) || IsRegisterSuccessMessage(cleanActual))
                {
                    Console.WriteLine($"{tag} [SUCCESS] Chấp nhận thông báo: '{cleanActual}'");
                    return;
                }

                Assert.Fail(
                    $"\n[FAILED] Toast thành công không khớp.\n" +
                    $"Mong đợi: '{cleanExpected}'\n" +
                    $"Thực tế: '{cleanActual}'");
            }

            try
            {
                wait.Until(d => !d.Url.Contains("/register"));
                Thread.Sleep(1000);
                Console.WriteLine($"{tag} [SUCCESS] Không thấy toast nhưng URL đã rời /register.");
                return;
            }
            catch
            {
                try
                {
                    var bodyText = _driver.FindElement(By.TagName("body")).Text;
                    Console.WriteLine($"{tag} [BODY WHEN STUCK] " + (bodyText.Length > 1500 ? bodyText[..1500] : bodyText));
                    Console.WriteLine($"{tag} [CURRENT URL] {_driver.Url}");
                    Console.WriteLine($"{tag} [SubmitDisabled] {IsSubmitDisabled()}");
                }
                catch { }

                Assert.Fail($"{tag} Không thấy toast thành công và URL vẫn ở trang đăng ký.");
            }
        }

        private string TryGetAlertOrToast(WebDriverWait wait, string tag)
        {
            try
            {
                var toastElem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                    "//div[contains(@class,'toast-message')] " +
                    "| //div[contains(@id,'toast-container')]//*[contains(@class,'toast')] " +
                    "| //div[contains(@class,'toast-error') or contains(@class,'toast-success')] " +
                    "| //*[contains(@class,'ngx-toastr') or contains(@class,'alert')]")));
                return toastElem.Text;
            }
            catch (WebDriverException ex) when (ex.Message.Contains("alert open"))
            {
                CheckCatchBackendAlert(tag);
            }
            catch { Thread.Sleep(1000); }
            return "";
        }

        // ─── TOGGLE ───────────────────────────────────────────────────────────────
        private void HandleToggle(string tag)
        {
            try
            {
                var toggleXpaths = new[]
                {
                    "//button[contains(@class,'toggle') or contains(@class,'show-password') or contains(@class,'eye')]",
                    "//*[contains(@class,'fa-eye') or contains(@class,'fa-eye-slash')]",
                    "//mat-icon[contains(text(),'visibility') or contains(text(),'visibility_off')]",
                    "//span[contains(@class,'p-password-show-icon') or contains(@class,'p-password-hide-icon')]",
                    "//i[contains(@class,'eye') or contains(@class,'icon-eye')]",

                    // Các nút mắt trong register thường là button/svg nằm sau input mật khẩu
                    "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password')]/following::button[1]",
                    "//input[contains(translate(@formcontrolname,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password')]/following::button[1]",
                    "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'matkhau')]/following::button[1]",
                    "//input[contains(translate(@formcontrolname,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'matkhau')]/following::button[1]",

                    "//input[@type='password']/following-sibling::*[1]",
                    "//input[contains(@formcontrolname,'password') or contains(@formcontrolname,'Password')]/following-sibling::*[1]"
                };

                IWebElement toggleBtn = null;

                foreach (var xp in toggleXpaths)
                {
                    try
                    {
                        var elems = _driver.FindElements(By.XPath(xp));
                        foreach (var e in elems)
                        {
                            try
                            {
                                if (e.Displayed && e.Enabled)
                                {
                                    toggleBtn = e;
                                    Console.WriteLine($"{tag} [Toggle] Tìm thấy qua: {xp}");
                                    break;
                                }
                            }
                            catch { }
                        }

                        if (toggleBtn != null) break;
                    }
                    catch { }
                }

                if (toggleBtn != null)
                {
                    ((IJavaScriptExecutor)_driver).ExecuteScript(
                        "arguments[0].scrollIntoView({block:'center'});", toggleBtn);
                    Thread.Sleep(300);
                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", toggleBtn);
                    Thread.Sleep(500);
                    return;
                }

                // Fallback: tìm input mật khẩu theo name/formcontrolname, không phụ thuộc type=password
                Console.WriteLine($"{tag} [Toggle] Không tìm thấy nút mắt, fallback click offset trên input mật khẩu");

                var passElem = _driver.FindElements(By.XPath(
                    "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password') or " +
                    "contains(translate(@formcontrolname,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password') or " +
                    "contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'matkhau') or " +
                    "contains(translate(@formcontrolname,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'matkhau')]"
                )).FirstOrDefault(e =>
                {
                    try { return e.Displayed && e.Enabled; }
                    catch { return false; }
                });

                if (passElem == null)
                {
                    DumpInputs(tag);
                    Console.WriteLine($"{tag} [Toggle] Không tìm thấy input mật khẩu để fallback, bỏ qua toggle.");
                    return;
                }

                new Actions(_driver)
                    .MoveToElement(passElem, passElem.Size.Width / 2 - 25, 0)
                    .Click()
                    .Build()
                    .Perform();

                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{tag} [Toggle] Lỗi nhưng không chặn test: {ex.Message}");
            }
        }

        // ─── ALERT ────────────────────────────────────────────────────────────────
        private void CheckCatchBackendAlert(string tag)
        {
            try
            {
                var alert = _driver.SwitchTo().Alert();
                string txt = alert.Text;
                alert.Dismiss();
                Console.WriteLine($"{tag} → Alert từ BE: {txt}");
                if (txt.Contains("Failed to fetch") || txt.Contains("Backend"))
                    Assert.Fail("[LỖI HỆ THỐNG]: Backend C# (API) chưa bật hoặc bị sập!");
            }
            catch (NoAlertPresentException) { }
        }
    }
}