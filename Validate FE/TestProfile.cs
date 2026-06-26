using ExcelDataReader;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;

// ================================================================
//  TestProfile.cs — Kiểm thử localhost:4200/profile (POM)
//  - Login tắt autofill + gõ-verify; chụp ảnh khi login fail.
//  - Lấy thông điệp từ CẢ toast LẪN lỗi inline (chữ đỏ dưới ô).
//  - Tự KHÔI PHỤC mật khẩu khi có lần đổi thành công.
//  - Tự lưu & khôi phục Họ tên / SĐT về ban đầu khi kết thúc.
// ================================================================
namespace RealtimeAuctionTest.ValidateFE
{
    public class ProfilePage
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        public ProfilePage(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
        }

        private IWebElement HoTenInput => FindFirst(
            By.CssSelector("input[formcontrolname='fullName']"),
            By.CssSelector("input[formcontrolname='hoTen']"),
            By.CssSelector("input[formcontrolname='name']"),
            By.Name("fullName"), By.Name("name"),
            By.XPath("//label[contains(.,'Họ tên') or contains(.,'họ tên') or contains(.,'HỌ TÊN')]/following::input[1]"));

        private IWebElement SdtInput => FindFirst(
            By.CssSelector("input[formcontrolname='phone']"),
            By.CssSelector("input[formcontrolname='phoneNumber']"),
            By.CssSelector("input[formcontrolname='soDienThoai']"),
            By.Name("phone"), By.Name("phoneNumber"),
            By.XPath("//label[contains(.,'iện thoại')]/following::input[1]"));

        private IWebElement LuuThayDoiBtn => FindFirstVisible(
            By.XPath("//button[contains(normalize-space(.),'Lưu thay đổi')]"));

        private IWebElement MkHienTaiInput => FindFirst(
            By.CssSelector("input[name='current-password-field']"),
            By.Name("currentPassword"),
            By.XPath("//label[contains(.,'hiện tại')]/following::input[1]"));

        private IWebElement MkMoiInput => FindFirst(
            By.CssSelector("input[name='new-password-field']"),
            By.Name("newPassword"),
            By.XPath("//label[contains(.,'ật khẩu mới')]/following::input[1]"));

        private IWebElement XacNhanMkInput => FindFirst(
            By.CssSelector("input[name='confirm-password-field']"),
            By.Name("confirmPassword"),
            By.XPath("//label[contains(.,'Xác nhận') or contains(.,'ác nhận')]/following::input[1]"));

        private IWebElement DoiMatKhauBtn => FindFirstVisible(
            By.XPath("//button[contains(normalize-space(.),'Đổi mật khẩu')]"));

        private IWebElement FindFirst(params By[] locators)
        {
            return FindFirstInternal(requireEnabled: true, locators);
        }

        private IWebElement FindFirstVisible(params By[] locators)
        {
            return FindFirstInternal(requireEnabled: false, locators);
        }

        private IWebElement FindFirstPresent(params By[] locators)
        {
            try
            {
                return _wait.Until(d =>
                {
                    foreach (var by in locators)
                    {
                        try
                        {
                            var el = d.FindElements(by).FirstOrDefault();
                            if (el != null) return el;
                        }
                        catch (StaleElementReferenceException) { }
                    }

                    return null;
                });
            }
            catch (WebDriverTimeoutException)
            {
                throw new NoSuchElementException(
                    "Không tìm thấy phần tử với các locator: " +
                    string.Join(" | ", locators.Select(l => l.ToString())));
            }
        }

        private IWebElement FindFirstInternal(bool requireEnabled, params By[] locators)
        {
            try
            {
                return _wait.Until(d =>
                {
                    foreach (var by in locators)
                    {
                        try
                        {
                            var el = d.FindElements(by)
                                .FirstOrDefault(e => e.Displayed && (!requireEnabled || e.Enabled));

                            if (el != null) return el;
                        }
                        catch (StaleElementReferenceException) { }
                    }

                    return null;
                });
            }
            catch (WebDriverTimeoutException)
            {
                throw new NoSuchElementException(
                    "Không tìm thấy phần tử với các locator: " +
                    string.Join(" | ", locators.Select(l => l.ToString())));
            }
        }

        private void Type(IWebElement el, string text)
        {
            text ??= "";

            ((IJavaScriptExecutor)_driver).ExecuteScript(
                "arguments[0].scrollIntoView({block:'center', inline:'nearest'});", el);

            el.Click();
            Thread.Sleep(120);

            try
            {
                el.SendKeys(Keys.Control + "a");
                el.SendKeys(Keys.Delete);
            }
            catch { }

            if (!string.IsNullOrEmpty(text))
                el.SendKeys(text);

            var actual = el.GetAttribute("value") ?? "";
            if (actual != text)
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "const el = arguments[0];" +
                    "const value = arguments[1] ?? '';" +
                    "const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" +
                    "setter.call(el, value);" +
                    "el.dispatchEvent(new Event('input', { bubbles: true }));" +
                    "el.dispatchEvent(new Event('change', { bubbles: true }));" +
                    "el.dispatchEvent(new Event('blur', { bubbles: true }));",
                    el, text);
            }
            else
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));" +
                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));" +
                    "arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));",
                    el);
            }

            var finalVal = el.GetAttribute("value") ?? "";
            Console.WriteLine($"[TYPE] name='{el.GetAttribute("name")}' expectedLen={text.Length}, actualLen={finalVal.Length}");
        }

        private void ForceSetInputByName(string inputName, string value)
        {
            value ??= "";

            var el = FindFirstVisible(By.CssSelector($"input[name='{inputName}']"));

            ((IJavaScriptExecutor)_driver).ExecuteScript(
                "const el = arguments[0];" +
                "const value = arguments[1] ?? '';" +
                "el.scrollIntoView({block:'center', inline:'nearest'});" +
                "el.focus();" +
                "const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;" +
                "setter.call(el, value);" +
                "el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: value }));" +
                "el.dispatchEvent(new Event('change', { bubbles: true }));" +
                "el.dispatchEvent(new Event('blur', { bubbles: true }));",
                el, value);

            var actual = el.GetAttribute("value") ?? "";
            Console.WriteLine($"[PW-DOM-SET] {inputName}: expectedLen={value.Length}, actualLen={actual.Length}");
        }

        private void DumpPasswordDomLengths()
        {
            try
            {
                var result = ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "return [" +
                    "document.querySelector(\"input[name='current-password-field']\")?.value?.length ?? -1," +
                    "document.querySelector(\"input[name='new-password-field']\")?.value?.length ?? -1," +
                    "document.querySelector(\"input[name='confirm-password-field']\")?.value?.length ?? -1" +
                    "].join(',');");

                Console.WriteLine("[PW-DOM-LENS] current,new,confirm = " + result);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[PW-DOM-LENS] Không đọc được DOM password: " + ex.Message);
            }
        }

        private (string Message, bool Success) ClickAndCollect(IWebElement btn, string buttonName)
        {
            if (!btn.Enabled)
            {
                var before = CollectResult(1).Message;
                var msg = string.IsNullOrWhiteSpace(before)
                    ? $"[BLOCKED_BY_DISABLED_BUTTON] Nút '{buttonName}' đang bị vô hiệu hóa, form đã chặn dữ liệu không hợp lệ."
                    : $"[BLOCKED_BY_DISABLED_BUTTON] {before}";
                Console.WriteLine(msg);
                return (msg, false);
            }

            try
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "arguments[0].scrollIntoView({block:'center', inline:'nearest'});", btn);
                btn.Click();
            }
            catch (ElementClickInterceptedException)
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", btn);
            }

            return CollectResult();
        }

        public string GetHoTen() { try { return HoTenInput.GetAttribute("value") ?? ""; } catch { return ""; } }
        public string GetSoDienThoai() { try { return SdtInput.GetAttribute("value") ?? ""; } catch { return ""; } }

        private void DismissToasts()
        {
            try
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "document.querySelectorAll('.toast,[class*=\"toast\"],[class*=\"Toastify\"],.snackbar,[class*=\"snack-bar\"]')" +
                    ".forEach(function(t){try{t.remove();}catch(e){}});");
            }
            catch { }
        }

        // Thu thập thông điệp (toast + lỗi inline). Success = có thông điệp chứa "thành công"
        // hoặc phần tử có class chứa "success" (nhận diện theo NỘI DUNG, không phụ thuộc tên class cụ thể).
        private (string Message, bool Success) CollectResult(int timeoutSeconds = 6)
        {
            string[] msgSel =
            {
                ".toast-success", ".toast-error", ".toast-message", ".toast", ".toast-body",
                ".Toastify__toast--success", ".Toastify__toast--error", ".Toastify__toast",
                ".ngx-toastr", ".alert", ".snackbar", ".mat-snack-bar-label", ".mat-mdc-snack-bar-label",
                ".p-toast-detail", ".swal2-html-container", ".swal2-title",
                "[class*='text-red']", "mat-error", ".mat-error",
                ".invalid-feedback", ".error-message", ".text-danger", ".field-error",
                ".validation-error", "[role='alert']"
            };

            var seen = new List<string>();
            bool success = false;
            var end = DateTime.Now.AddSeconds(timeoutSeconds);

            while (DateTime.Now < end)
            {
                foreach (var s in msgSel)
                {
                    try
                    {
                        foreach (var e in _driver.FindElements(By.CssSelector(s)))
                        {
                            try
                            {
                                if (!e.Displayed) continue;
                                var t = (e.Text ?? "").Trim();
                                if (t.Length < 2 || t == "*") continue;
                                var cls = (e.GetAttribute("class") ?? "").ToLower();
                                if (t.ToLower().Contains("thành công") || cls.Contains("success"))
                                    success = true;
                                if (!seen.Contains(t)) seen.Add(t);
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                if (success) break;          // đã thấy thành công -> chốt ngay
                Thread.Sleep(250);
            }

            foreach (var m in seen) Console.WriteLine($"[MSG] {m}");
            return (string.Join(" | ", seen), success);
        }

        private static string CreateTempAvatarPng()
        {
            // Ảnh PNG 1x1 pixel hợp lệ, tạo runtime để test không phụ thuộc file ngoài.
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

            var path = Path.Combine(Path.GetTempPath(), "selenium_profile_avatar.png");
            File.WriteAllBytes(path, Convert.FromBase64String(base64Png));
            return path;
        }

        public (string Message, bool Success) ChangeAvatar()
        {
            DismissToasts();

            var avatarPath = CreateTempAvatarPng();

            var input = FindFirstPresent(
                By.XPath("//input[@type='file' and (contains(@accept,'image') or contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'avatar') or contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'avatar') or contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'anh') or contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'anh'))]"),
                By.CssSelector("input[type='file'][accept*='image']"),
                By.CssSelector("input[type='file']")
            );

            try
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "arguments[0].style.display='block';" +
                    "arguments[0].style.visibility='visible';" +
                    "arguments[0].style.opacity='1';" +
                    "arguments[0].removeAttribute('hidden');",
                    input);
            }
            catch { }

            input.SendKeys(avatarPath);
            Console.WriteLine("[AVATAR] Uploaded file: " + avatarPath);

            var result = CollectResult(10);

            if (!result.Success)
            {
                // Một số UI đổi avatar xong chỉ hiện preview, không hiện toast. Dump text để debug.
                try
                {
                    var body = _driver.FindElement(By.TagName("body")).Text ?? "";
                    Console.WriteLine("[AVATAR BODY] " + (body.Length > 2000 ? body.Substring(0, 2000) : body));
                }
                catch { }
            }

            return result;
        }

        public (string Message, bool Success) UpdateProfile(string hoTen, string sdt)
        {
            DismissToasts();
            Type(HoTenInput, hoTen);
            Type(SdtInput, sdt);
            return ClickAndCollect(LuuThayDoiBtn, "Lưu thay đổi");
        }

        public (string Message, bool Success) ChangePassword(string cur, string moi, string xn)
        {
            DismissToasts();

            ForceSetInputByName("current-password-field", cur);
            ForceSetInputByName("new-password-field", moi);
            ForceSetInputByName("confirm-password-field", xn);

            Thread.Sleep(3000); // nhìn 3 ô đã có dữ liệu chưa

            DumpPasswordDomLengths();

            var result = ClickAndCollect(DoiMatKhauBtn, "Đổi mật khẩu");

            Thread.Sleep(5000); // nhìn toast sau khi bấm

            return result;
        }
    }

    [TestFixture]
    public class TestProfile
    {
        private static ChromeDriver _driver = null!;
        private static WebDriverWait _wait = null!;
        private ProfilePage _profile = null!;

        private const string BaseUrl = "http://localhost:4200";
        private const string LoginPath = "/login";

        // ⚠️ Đặt ĐÚNG mật khẩu hiện tại của tài khoản (mật khẩu khôi phục cũng lấy từ đây).
        private const string Email = "vietanhdd268@gmail.com";
        private const string Password = "Vietanh268@";

        private const int WaitTimeout = 25;

        // Lưu thông tin gốc để khôi phục cuối phiên
        private static string _origHoTen = "";
        private static string _origSdt = "";

        private static readonly string ExcelPath = ResolveExcelPath();
        private static string ResolveExcelPath([CallerFilePath] string src = "")
        {
            var csDir = Path.GetDirectoryName(src)!;
            var rootDir = Path.GetFullPath(Path.Combine(csDir, ".."));
            return Path.Combine(rootDir, "Excel", "ProfileTestData.xlsx");
        }

        private static IJavaScriptExecutor Js => (IJavaScriptExecutor)_driver;

        static TestProfile()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            var opt = new ChromeOptions();
            opt.AddArgument("--no-sandbox");
            opt.AddArgument("--disable-dev-shm-usage");
            opt.AddArgument("--window-size=1440,900");
            opt.AddUserProfilePreference("credentials_enable_service", false);
            opt.AddUserProfilePreference("profile.password_manager_enabled", false);

            _driver = new ChromeDriver(opt);
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(WaitTimeout));

            Console.WriteLine("=== TESTPROFILE_VERSION: REAL_FIXED_20260626_MK01_DOM_SET ===");

            Login();

            // Đọc & lưu thông tin cá nhân gốc để khôi phục khi kết thúc
            try
            {
                OpenProfileViaMenu();
                var p = new ProfilePage(_driver);
                _origHoTen = p.GetHoTen();
                _origSdt = p.GetSoDienThoai();
                Console.WriteLine($"[ORIG] HoTen='{_origHoTen}' | SDT='{_origSdt}'");
            }
            catch (Exception ex) { Console.WriteLine("[ORIG] không đọc được thông tin gốc: " + ex.Message); }
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            // Khôi phục họ tên / SĐT về ban đầu (best effort) trước khi đóng
            try
            {
                if (_driver != null && (!string.IsNullOrEmpty(_origHoTen) || !string.IsNullOrEmpty(_origSdt)))
                {
                    _driver.Navigate().GoToUrl(BaseUrl);
                    WaitForPageLoad();
                    if (_driver.Url.Contains(LoginPath)) LoginWith(Email, Password);
                    OpenProfileViaMenu();
                    new ProfilePage(_driver).UpdateProfile(_origHoTen, _origSdt);
                    Console.WriteLine("[RESTORE] Đã khôi phục Họ tên / SĐT về ban đầu.");
                }
            }
            catch (Exception ex) { Console.WriteLine("[RESTORE] không khôi phục được thông tin: " + ex.Message); }
            finally
            {
                _driver?.Quit();
                _driver?.Dispose();
            }
        }

        [SetUp]
        public void SetUp()
        {
            OpenProfileViaMenu();
            _profile = new ProfilePage(_driver);
        }

        private static void WaitForPageLoad() =>
            _wait.Until(d => ((IJavaScriptExecutor)d)
                .ExecuteScript("return document.readyState").Equals("complete"));

        private static void TypeAndVerify(IWebElement el, string value, string fieldName)
        {
            el.Clear();
            el.SendKeys(value);
            var actual = el.GetAttribute("value") ?? "";
            if (actual != value)
            {
                Js.ExecuteScript(
                    "arguments[0].value=arguments[1];" +
                    "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));" +
                    "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));" +
                    "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));", el, value);
                actual = el.GetAttribute("value") ?? "";
            }
            Console.WriteLine($"   [{fieldName}] giá trị cuối: '{actual}' (len={actual.Length})");
        }

        private static void Login() => LoginWith(Email, Password);

        private static void LoginWith(string email, string password)
        {
            _driver.Navigate().GoToUrl(BaseUrl + LoginPath);
            WaitForPageLoad();

            var emailEl = _wait.Until(ExpectedConditions.ElementIsVisible(By.Name("email")));
            TypeAndVerify(emailEl, email, "email");

            var pwd = _wait.Until(ExpectedConditions.ElementIsVisible(By.Name("password")));
            TypeAndVerify(pwd, password, "password");

            _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//button[contains(.,'Đăng Nhập')]"))).Click();

            try { _wait.Until(d => !d.Url.Contains(LoginPath)); }
            catch (WebDriverTimeoutException)
            {
                throw new Exception("❌ ĐĂNG NHẬP KHÔNG THÀNH CÔNG với '" + email + "'." + DumpLoginFailure());
            }
            WaitForPageLoad();
            Console.WriteLine("✅ LOGIN SUCCESS - " + _driver.Url);
        }

        private void RestorePassword(string fromPassword, string toPassword)
        {
            try
            {
                _driver.Navigate().GoToUrl(BaseUrl);
                WaitForPageLoad();
                if (_driver.Url.Contains(LoginPath)) LoginWith(Email, fromPassword);
                OpenProfileViaMenu();
                var r = new ProfilePage(_driver).ChangePassword(fromPassword, toPassword, toPassword);
                Console.WriteLine($"   [RESTORE-PW] Đổi mật khẩu về ban đầu → '{r.Message}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine("   ⚠️⚠️ [RESTORE-PW] KHÔNG khôi phục được mật khẩu! " +
                    "Mật khẩu tài khoản hiện đang là: '" + fromPassword + "'. Đổi lại bằng tay. Lỗi: " + ex.Message);
            }
        }

        private static string DumpLoginFailure()
        {
            var sb = new StringBuilder();
            sb.AppendLine();
            try { sb.AppendLine("   URL: " + _driver.Url); } catch { }
            try
            {
                var ev = Js.ExecuteScript("var e=document.querySelector(\"input[name='email']\");return e?e.value:'(no)';");
                var pl = Js.ExecuteScript("var e=document.querySelector(\"input[name='password']\");return e?(e.value.length+' ký tự'):'(no)';");
                sb.AppendLine("   Ô email: '" + ev + "' | Ô password: " + pl);
            }
            catch { }
            try
            {
                var errs = _driver.FindElements(By.CssSelector(
                        "[class*='text-red'],[class*='error'],mat-error,[role='alert'],.toast-error,[class*='Toastify']"))
                    .Where(e => { try { return e.Displayed && !string.IsNullOrWhiteSpace(e.Text); } catch { return false; } })
                    .Select(e => e.Text.Trim()).Distinct().ToList();
                sb.AppendLine("   Text lỗi: " + (errs.Count > 0 ? string.Join(" | ", errs) : "(không có)"));
            }
            catch { }
            try
            {
                var dir = Path.GetDirectoryName(typeof(TestProfile).Assembly.Location) ?? ".";
                ((ITakesScreenshot)_driver).GetScreenshot().SaveAsFile(Path.Combine(dir, "login_fail.png"));
                sb.AppendLine("   📸 Ảnh: " + Path.Combine(dir, "login_fail.png"));
            }
            catch { }
            return sb.ToString();
        }

        private void OpenProfileViaMenu()
        {
            _driver.Navigate().GoToUrl(BaseUrl);
            WaitForPageLoad();

            var btnMenu = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//button[contains(.,'Menu')]")));
            Js.ExecuteScript("arguments[0].click();", btnMenu);

            var item = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(.,'Quản lý tài khoản')]")));
            Js.ExecuteScript("arguments[0].click();", item);

            _wait.Until(d => d.Url.Contains("/profile"));
            WaitForPageLoad();
            _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//button[contains(normalize-space(.),'Đổi mật khẩu')]")));
        }

        private static List<Dictionary<string, string>> ReadSheet(string filePath, string sheetName)
        {
            var rows = new List<Dictionary<string, string>>();
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            do
            {
                if (!string.Equals(reader.Name, sheetName, StringComparison.OrdinalIgnoreCase)) continue;
                if (!reader.Read()) break;
                var headers = new List<string>();
                for (int i = 0; i < reader.FieldCount; i++)
                    headers.Add(reader.GetValue(i)?.ToString()?.Trim() ?? "");
                while (reader.Read())
                {
                    var dict = new Dictionary<string, string>();
                    bool allEmpty = true;
                    for (int i = 0; i < headers.Count; i++)
                    {
                        if (string.IsNullOrEmpty(headers[i])) continue;
                        // Không Trim dữ liệu test, vì có case cố ý nhập toàn dấu cách.
                        var val = reader.GetValue(i)?.ToString() ?? "";
                        dict[headers[i]] = val;
                        if (!string.IsNullOrWhiteSpace(val)) allEmpty = false;
                    }
                    if (!allEmpty) rows.Add(dict);
                }
                break;
            } while (reader.NextResult());
            return rows;
        }

        public static IEnumerable<TestCaseData> ThongTinCases()
        {
            foreach (var r in ReadSheet(ExcelPath, "ThongTinCaNhan"))
                yield return new TestCaseData(r).SetName($"TT_{r["TestCaseID"]}_{r["MoTa"]}");
        }

        public static IEnumerable<TestCaseData> DoiMatKhauCases()
        {
            foreach (var r in ReadSheet(ExcelPath, "DoiMatKhau"))
                yield return new TestCaseData(r).SetName($"MK_{r["TestCaseID"]}_{r["MoTa"]}");
        }

        [TestCaseSource(nameof(ThongTinCases))]
        public void Test_ThongTinCaNhan(Dictionary<string, string> data)
        {
            string testCaseId = data.ContainsKey("TestCaseID") ? data["TestCaseID"].Trim() : "";
            string expected = data.ContainsKey("ExpectedToast") ? data["ExpectedToast"] : "";
            string ket = data.ContainsKey("KetQuaMongDoi") ? data["KetQuaMongDoi"] : "";

            (string Message, bool Success) r;

            if (testCaseId.Equals("TT19", StringComparison.OrdinalIgnoreCase))
            {
                // TT19 là case thay ảnh đại diện, không dùng HoTen/SoDienThoai.
                r = _profile.ChangeAvatar();
            }
            else
            {
                r = _profile.UpdateProfile(data["HoTen"], data["SoDienThoai"]);
            }

            Console.WriteLine($"[{data["TestCaseID"]}] success={r.Success} | message='{r.Message}' | ket={ket}");
            AssertOutcome(r, expected, ket);
            // Họ tên/SĐT sẽ được khôi phục về gốc ở OneTimeTearDown
        }

        [TestCaseSource(nameof(DoiMatKhauCases))]
        public void Test_DoiMatKhau(Dictionary<string, string> data)
        {
            string ket = data.ContainsKey("KetQuaMongDoi") ? data["KetQuaMongDoi"].Trim().ToUpper() : "";
            bool validChange = ket == "PASS";

            // Case hợp lệ: ép mật khẩu hiện tại = Password (chắc chắn đổi được). Case khác: theo Excel.
            string current = validChange ? Password : data["MatKhauHienTai"];
            string newPass = data["MatKhauMoi"];
            string confirm = data["XacNhanMatKhau"];

            var r = _profile.ChangePassword(current, newPass, confirm);
            string expected = data["ExpectedToast"];
            Console.WriteLine($"[{data["TestCaseID"]}] success={r.Success} | message='{r.Message}' | ket={ket}");

            try
            {
                AssertOutcome(r, expected, ket);
            }
            finally
            {
                // KHÔI PHỤC nếu BẤT KỲ lần đổi nào thật sự thành công (an toàn cả khi case "FAIL" lỡ thành công)
                if (r.Success)
                    RestorePassword(fromPassword: newPass, toPassword: Password);
            }
        }

        // Ra PASS/FAIL ngay dựa trên KetQuaMongDoi (không cần điền ExpectedToast).
        //  - PASS  => thao tác phải THÀNH CÔNG (có toast success)
        //  - FAIL  => thao tác phải BỊ CHẶN  (không có toast success)
        //  - ExpectedToast (nếu điền) => kiểm tra thêm nội dung thông điệp.
        private static void AssertOutcome((string Message, bool Success) r, string expectedToast, string ketQua)
        {
            string k = (ketQua ?? "").Trim().ToUpper();

            if (k == "PASS")
            {
                Assert.That(r.Success, Is.True,
                    $"Mong đợi THÀNH CÔNG nhưng không thấy toast thành công. Thông điệp: '{r.Message}'");
            }
            else if (k == "FAIL")
            {
                Assert.That(r.Success, Is.False,
                    $"Mong đợi BỊ TỪ CHỐI nhưng thao tác lại THÀNH CÔNG. Thông điệp: '{r.Message}'");
            }
            else
            {
                Assert.Inconclusive("Thiếu KetQuaMongDoi (PASS/FAIL) trong Excel cho case này.");
                return;
            }

            // Kiểm tra thêm nội dung nếu đã điền ExpectedToast
            if (!string.IsNullOrWhiteSpace(expectedToast) &&
                !r.Message.Contains("[BLOCKED_BY_DISABLED_BUTTON]"))
            {
                Assert.That(r.Message, Does.Contain(expectedToast.Trim()),
                    $"Thông điệp không khớp ExpectedToast. Mong đợi chứa: '{expectedToast}', thực tế: '{r.Message}'.");
            }
        }
    }
}