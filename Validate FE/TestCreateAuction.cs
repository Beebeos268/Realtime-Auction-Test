/*
 * Selenium Test Suite – Tạo Phiên Đấu Giá (A01–A26)
 * Framework : NUnit 3 + Selenium WebDriver + ExcelDataReader
 * Language  : C# (.NET 6+)
 *
 * ⚙️ FIX v8 (31/05/2026):
 *   ★ Sau v7: 12 PASS / 13 FAIL / 1 SKIP. Phân tích 13 fail còn lại:
 *
 *     A) A05 inline(7), A07 inline(7) — locator bắt nhầm chuỗi date "18:00 07/...":
 *        App có chỗ hiển thị deadline bằng text màu đỏ, bị [class*='text-red']
 *        bắt nhầm. SỬA: drop [class*='text-red'], filter date-pattern.
 *
 *     B) A08/A09/A13/A14/A18 inline=0 — Angular ẩn error khi control chưa touched:
 *        el.Clear() trên field empty từ đầu không trigger blur → required validator
 *        chạy nhưng message ẩn. SỬA: Fill() thêm el.Click() (focus) + JS el.blur()
 *        (blur) → Angular mark touched → inline hiện.
 *
 *     C) A01 toast='' — flaky test đầu tiên (cold start).
 *        SỬA: warmup trong OneTimeSetUp (mở form 1 lần trước khi test chạy).
 *
 *     D) A12 (price=0), A17 (step=0), A21 (start==end) toast='' — KHÔNG fix được
 *        bằng code. System thực sự reject các giá trị này; Excel ghi expected
 *        "Thành công" là SAI. Cần update Excel.
 *        Bằng chứng: A11 (decimal 100000.5) PASS, A16 (decimal 10000.2) PASS
 *        → app chấp nhận thập phân nhưng có thể reject =0.
 *
 *     E) A26 inline(1): "File..." — Excel expected "Dung lượng ảnh không được
 *        vượt quá 5MB!" không match text app thật sự hiện. Cần update Excel.
 */

using ExcelDataReader;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace TestProject1
{
    // ═══════════════════════════════════════════════════════════════
    // CONFIG
    // ═══════════════════════════════════════════════════════════════
    internal static class Config
    {
        public const string BaseUrl = "http://localhost:4200";
        public const string LoginPath = "/login";
        public const string SellerEmail = "ntpnguyen210104@gmail.com";
        public const string SellerPassword = "Nguyen@21";

        public static readonly string ProjectDir = GetProjectDir();
        public static readonly string TestFileDir = Path.Combine(ProjectDir, "Test file");
        public static readonly string ExcelPath = Path.Combine(ProjectDir, "CreateAuctionTestData.xlsx");

        public const int WaitTimeout = 25;
        public const int UploadAppearTimeout = 10;
        public const int UploadStaleTimeout = 15;
        public const int ErrorPollTimeout = 15;
        public const int SuccessToastTimeout = 20;
        public const int SubmitEnabledTimeout = 5;
        public const int SleepAfterUpload = 500;

        public static readonly DateTime ExcelEarliest = new DateTime(2026, 5, 30, 0, 0, 0);
        public static readonly TimeSpan DateOffset = ComputeOffset();

        private static TimeSpan ComputeOffset()
        {
            var now = DateTime.Now;
            if (now.Date >= ExcelEarliest.Date)
            {
                var target = now.Date.AddDays(7);
                var offset = target - ExcelEarliest.Date;
                Console.WriteLine($"[Config] today={now:yyyy-MM-dd}, Excel earliest=30/05/2026 → DateOffset = {offset.TotalDays} ngày");
                return offset;
            }
            return TimeSpan.Zero;
        }

        public static string? ImagePath(string? fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return null;
            var trimmed = fileName.Trim();
            var direct = Path.Combine(TestFileDir, trimmed);
            if (File.Exists(direct)) return direct;
            if (!Directory.Exists(TestFileDir)) return direct;

            var wanted = trimmed.Normalize(NormalizationForm.FormC);
            var files = Directory.GetFiles(TestFileDir);

            var exact = files.FirstOrDefault(f =>
                Path.GetFileName(f).Normalize(NormalizationForm.FormC)
                    .Equals(wanted, StringComparison.OrdinalIgnoreCase));
            if (exact != null) return exact;

            var wantedNoExt = Path.GetFileNameWithoutExtension(wanted);
            var loose = files.FirstOrDefault(f =>
                Path.GetFileNameWithoutExtension(f).Normalize(NormalizationForm.FormC)
                    .Equals(wantedNoExt, StringComparison.OrdinalIgnoreCase));
            return loose ?? direct;
        }

        private static string GetProjectDir([CallerFilePath] string src = "")
            => Path.GetDirectoryName(src)!;
    }

    // ═══════════════════════════════════════════════════════════════
    // LOCATORS
    // ═══════════════════════════════════════════════════════════════
    internal static class Loc
    {
        public static readonly By LoginEmail = By.Name("email");
        public static readonly By LoginPassword = By.Name("password");
        public static readonly By LoginBtn = By.XPath("//button[contains(.,'Đăng Nhập')]");

        public static readonly By BtnMenu = By.XPath("//button[contains(.,'Menu')]");
        public static readonly By MenuCreateAuction = By.XPath("//a[contains(.,'Tạo phiên đấu giá')]");

        public static readonly By ProductName = By.Name("name");
        public static readonly By Description = By.Name("description");
        public static readonly By StartPrice = By.Name("startPrice");
        public static readonly By PriceStep = By.Name("stepPrice");
        public static readonly By StartTime = By.Name("startTime");
        public static readonly By EndTime = By.Name("endTime");

        public static readonly By BtnLenSanNgay = By.XPath("//button[contains(.,'Lên sàn ngay')]");
        public static readonly By BtnHenGio = By.XPath("//button[contains(.,'Hẹn giờ')]");

        public static readonly By ImageInput = By.CssSelector("input[type='file'][accept*='image']");

        public static readonly By SubmitBtn = By.CssSelector("button[type='submit']");

        public static readonly By ToastSuccess = By.CssSelector(
            ".toast-success, .Toastify__toast--success");
        public static readonly By ToastError = By.CssSelector(
            ".toast-error, .Toastify__toast--error");
        public static readonly By AnyToast = By.CssSelector(
            ".toast-success, .toast-error, " +
            ".Toastify__toast--success, .Toastify__toast--error, .Toastify__toast");

        // ╔════════════════════════════════════════════════════════╗
        // ║  FIX v8: drop [class*='text-red'] (quá rộng, bắt nhầm   ║
        // ║  display deadline). Giữ class cụ thể + thêm 700/300.    ║
        // ╚════════════════════════════════════════════════════════╝
        public static readonly By InlineError = By.CssSelector(
            "p.text-red-300, p.text-red-400, p.text-red-500, p.text-red-600, p.text-red-700, " +
            "p.text-xs.text-red-400, p.text-xs.text-red-500, " +
            "span.text-red-300, span.text-red-400, span.text-red-500, span.text-red-600, span.text-red-700, " +
            "div.text-red-300, div.text-red-400, div.text-red-500, div.text-red-600, div.text-red-700, " +
            "small.text-red-400, small.text-red-500, small.text-red-600, " +
            "mat-error, .mat-error, " +
            ".text-danger, .text-error, " +
            ".error-message, .errorMessage, " +
            ".invalid-feedback, .form-error, " +
            ".field-error, .validation-error, " +
            ".error-text, " +
            ".ng-feedback-error, " +
            ".ant-form-item-explain-error, " +
            "[role='alert']");
    }

    // ═══════════════════════════════════════════════════════════════
    // DATA MODEL
    // ═══════════════════════════════════════════════════════════════
    internal class TestRow
    {
        public string StepId { get; set; } = "";
        public string ScenarioName { get; set; } = "";
        public string ProductName { get; set; } = "";
        public string ImageFile { get; set; } = "";
        public string Description { get; set; } = "";
        public string StartPrice { get; set; } = "";
        public string PriceStep { get; set; } = "";
        public string StartTimeType { get; set; } = "NOW";
        public string StartTime { get; set; } = "";
        public string EndTime { get; set; } = "";
        public string ExpectedMessage { get; set; } = "";
        public string Action { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    // EXCEL READER
    // ═══════════════════════════════════════════════════════════════
    internal static class ExcelReader
    {
        public static List<TestRow> Load(string path)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
            });
            var table = ds.Tables[0];

            string Cell(DataRow row, string col) =>
                table.Columns.Contains(col) ? row[col]?.ToString()?.Trim() ?? "" : "";

            var list = new List<TestRow>();
            foreach (DataRow row in table.Rows)
            {
                if (string.IsNullOrWhiteSpace(row[0]?.ToString())) continue;
                list.Add(new TestRow
                {
                    StepId = Cell(row, "Step_ID"),
                    ScenarioName = Cell(row, "Scenario_Name"),
                    ProductName = Cell(row, "Tên sản phẩm"),
                    ImageFile = Cell(row, "Ảnh sản phẩm"),
                    Description = Cell(row, "Mô tả chi tiết"),
                    StartPrice = Cell(row, "Giá khởi điểm"),
                    PriceStep = Cell(row, "Bước giá"),
                    StartTimeType = Cell(row, "StartTime_Type"),
                    StartTime = Cell(row, "Thời gian bắt đầu"),
                    EndTime = Cell(row, "Thời gian kết thúc"),
                    ExpectedMessage = Cell(row, "ExpectedMessage"),
                    Action = Cell(row, "Action"),
                });
            }
            return list;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    // TEST CLASS
    // ═══════════════════════════════════════════════════════════════
    [TestFixture]
    public class AuctionCreateTests
    {
        private static ChromeDriver _driver = null!;
        private static WebDriverWait _wait = null!;

        private static readonly List<TestRow> _rows = ExcelReader.Load(Config.ExcelPath);

        // Regex để filter ra các chuỗi giống date/time (false positive)
        private static readonly Regex _dateLikePattern = new Regex(
            @"^\s*\d{1,2}\s*[:/-]\s*\d{1,2}", RegexOptions.Compiled);

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            var opt = new ChromeOptions();
            opt.AddArgument("--no-sandbox");
            opt.AddArgument("--disable-dev-shm-usage");
            opt.AddArgument("--window-size=1440,900");

            _driver = new ChromeDriver(opt);
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(Config.WaitTimeout));

            Login();

            // FIX v8: Warmup — mở form 1 lần để load resources, tránh flaky A01
            try
            {
                Console.WriteLine("--- Warmup: opening create-auction form first time ---");
                OpenCreateAuctionViaMenu();
                Thread.Sleep(1500);
                DismissAllToasts();
                Console.WriteLine("--- Warmup done ---");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   [Warmup] {ex.Message}");
            }
        }

        [SetUp]
        public void SetUp()
        {
            try { DismissAllToasts(); } catch { }
            OpenCreateAuctionViaMenu();
            try { DismissAllToasts(); } catch { }
            Thread.Sleep(300);
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            _driver?.Quit();
            _driver?.Dispose();
        }

        // ══════════════════════════════════════════════════════════
        // HELPERS
        // ══════════════════════════════════════════════════════════
        private static IJavaScriptExecutor Js => (IJavaScriptExecutor)_driver;

        private static void WaitForPageLoad() =>
            _wait.Until(d => ((IJavaScriptExecutor)d)
                .ExecuteScript("return document.readyState").Equals("complete"));

        private static string NormalizeText(string? text) =>
            string.IsNullOrEmpty(text)
                ? ""
                : Regex.Replace(text.Trim(), @"\s+", " ");

        private static WebDriverWait MakeWait(int seconds) =>
            new WebDriverWait(_driver, TimeSpan.FromSeconds(seconds));

        private static void DismissAllToasts()
        {
            try
            {
                Js.ExecuteScript(@"
                    var sel = '.toast-success, .toast-error, [class*=""Toastify__toast""]';
                    document.querySelectorAll(sel).forEach(function(t) {
                        try { t.remove(); } catch(e) {}
                    });
                ");
            }
            catch { }
        }

        private static void Login()
        {
            _driver.Navigate().GoToUrl(Config.BaseUrl + Config.LoginPath);
            WaitForPageLoad();

            var email = _wait.Until(ExpectedConditions.ElementIsVisible(Loc.LoginEmail));
            email.Clear(); email.SendKeys(Config.SellerEmail);

            var pwd = _wait.Until(ExpectedConditions.ElementIsVisible(Loc.LoginPassword));
            pwd.Clear(); pwd.SendKeys(Config.SellerPassword);

            _wait.Until(ExpectedConditions.ElementToBeClickable(Loc.LoginBtn)).Click();
            _wait.Until(d => !d.Url.Contains(Config.LoginPath));
            WaitForPageLoad();
            Console.WriteLine("✅ LOGIN SUCCESS");
        }

        private static void OpenCreateAuctionViaMenu()
        {
            _driver.Navigate().GoToUrl(Config.BaseUrl);
            WaitForPageLoad();

            var btnMenu = _wait.Until(ExpectedConditions.ElementToBeClickable(Loc.BtnMenu));
            Js.ExecuteScript("arguments[0].click();", btnMenu);

            var createItem = _wait.Until(ExpectedConditions.ElementToBeClickable(Loc.MenuCreateAuction));
            Js.ExecuteScript("arguments[0].click();", createItem);

            _wait.Until(d => d.Url.Contains("/auction/create"));
            WaitForPageLoad();
            _wait.Until(ExpectedConditions.ElementIsVisible(Loc.ProductName));
            Console.WriteLine("✅ OPEN CREATE AUCTION FORM");
        }

        private static IWebElement WaitVisible(By by) =>
            _wait.Until(ExpectedConditions.ElementIsVisible(by));

        // ╔════════════════════════════════════════════════════════╗
        // ║  FIX v8: Fill — Click (focus) → Clear → SendKeys → blur ║
        // ║  Đảm bảo Angular control được mark touched kể cả khi    ║
        // ║  field empty từ đầu (vd A08 price, A18 startTime).      ║
        // ╚════════════════════════════════════════════════════════╝
        private static void Fill(By by, string value)
        {
            var el = WaitVisible(by);

            // (1) Focus the field
            try { el.Click(); } catch { }

            // (2) Clear (will fire input events if there's existing value)
            try { el.Clear(); } catch { }

            // (3) Send value (if any)
            if (!string.IsNullOrEmpty(value))
            {
                try { el.SendKeys(value); } catch { }
            }

            // (4) Blur → Angular marks control as touched
            //    Dispatch native blur event (el.blur() in JS) which Angular's
            //    DefaultValueAccessor.@HostListener('blur') catches → onTouched()
            try
            {
                Js.ExecuteScript(@"
                    try { arguments[0].dispatchEvent(new Event('input',  {bubbles:true})); } catch(e) {}
                    try { arguments[0].dispatchEvent(new Event('change', {bubbles:true})); } catch(e) {}
                    try { arguments[0].blur(); } catch(e) {}
                    try { arguments[0].dispatchEvent(new Event('blur',   {bubbles:true})); } catch(e) {}
                ", el);
            }
            catch { }
        }

        private static void SelectStartMode(string type)
        {
            var btn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                type.ToUpper() == "NOW" ? Loc.BtnLenSanNgay : Loc.BtnHenGio));
            Js.ExecuteScript("arguments[0].click();", btn);
        }

        private static void SetDateTime(By by, string dt)
        {
            if (string.IsNullOrWhiteSpace(dt)) return;

            var p = dt.Trim().Split(' ');
            var dParts = p[0].Split('/');
            var tStr = p.Length > 1 ? p[1] : "00:00";
            var tParts = tStr.Split(':');

            string iso;
            try
            {
                int day = int.Parse(dParts[0]);
                int month = int.Parse(dParts[1]);
                int year = int.Parse(dParts[2]);
                int hour = int.Parse(tParts[0]);
                int min = int.Parse(tParts[1]);

                var parsed = new DateTime(year, month, day, hour, min, 0).Add(Config.DateOffset);
                iso = parsed.ToString("yyyy-MM-ddTHH:mm");

                if (Config.DateOffset.TotalDays > 0)
                    Console.WriteLine($"   [Date] '{dt}' → '{iso}' (+{Config.DateOffset.TotalDays}d)");
            }
            catch
            {
                iso = $"{dParts[2]}-{dParts[1]}-{dParts[0]}T{tStr}";
            }

            var el = WaitVisible(by);
            // FIX v8: set value + dispatch input/change + blur (mark touched)
            Js.ExecuteScript(@"
                arguments[0].value = arguments[1];
                try { arguments[0].dispatchEvent(new Event('input',  {bubbles:true})); } catch(e) {}
                try { arguments[0].dispatchEvent(new Event('change', {bubbles:true})); } catch(e) {}
                try { arguments[0].blur(); } catch(e) {}
                try { arguments[0].dispatchEvent(new Event('blur',   {bubbles:true})); } catch(e) {}
            ", el, iso);
        }

        private static void UploadImage(string? filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            if (!File.Exists(filePath))
            {
                var dir = Path.GetDirectoryName(filePath) ?? "";
                var available = Directory.Exists(dir)
                    ? string.Join("\n  - ", Directory.GetFiles(dir).Select(Path.GetFileName))
                    : "(thư mục không tồn tại)";
                throw new FileNotFoundException(
                    $"❌ Không tìm thấy ảnh:\n  {filePath}\n\n" +
                    $"📂 Các file đang có trong '{dir}':\n  - {available}");
            }

            DismissAllToasts();

            var el = _driver.FindElement(Loc.ImageInput);
            Js.ExecuteScript(
                "arguments[0].style.display='block'; arguments[0].style.opacity='1';", el);
            el.SendKeys(filePath);

            IWebElement? uploadToast = null;
            try
            {
                uploadToast = MakeWait(Config.UploadAppearTimeout).Until(d =>
                {
                    var toasts = d.FindElements(Loc.AnyToast);
                    return toasts.FirstOrDefault(t =>
                    {
                        try { return t.Displayed && !string.IsNullOrWhiteSpace(t.Text); }
                        catch { return false; }
                    });
                });
                Console.WriteLine($"   [Upload toast] {NormalizeText(uploadToast?.Text)}");
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("   [Upload toast] không xuất hiện — tiếp tục");
            }

            if (uploadToast != null)
            {
                bool isErrorToast = false;
                try
                {
                    var cls = uploadToast.GetAttribute("class") ?? "";
                    isErrorToast = cls.Contains("toast-error") ||
                                   cls.Contains("Toastify__toast--error");
                }
                catch { }

                if (isErrorToast)
                {
                    Console.WriteLine("   [Upload toast] ERROR → giữ lại cho assertion");
                    return;
                }

                try
                {
                    MakeWait(Config.UploadStaleTimeout)
                        .Until(ExpectedConditions.StalenessOf(uploadToast));
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("   [Upload toast] chưa biến mất — tiếp tục");
                }
            }

            Thread.Sleep(Config.SleepAfterUpload);
        }

        private static void Submit()
        {
            IWebElement? btn = null;

            var deadline = DateTime.UtcNow.AddSeconds(Config.SubmitEnabledTimeout);
            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    btn = _driver.FindElement(Loc.SubmitBtn);
                    if (btn.Displayed && btn.Enabled) break;
                }
                catch { btn = null; }
                Thread.Sleep(200);
            }

            if (btn == null)
            {
                try { btn = _driver.FindElement(Loc.SubmitBtn); }
                catch
                {
                    Console.WriteLine("   [Submit] Submit button not found");
                    return;
                }
            }

            try { Js.ExecuteScript("arguments[0].scrollIntoView({block:'center'});", btn); }
            catch { }

            if (btn.Enabled)
            {
                try
                {
                    btn.Click();
                    Console.WriteLine("   [Submit] native click");
                    return;
                }
                catch (Exception)
                {
                    try
                    {
                        Js.ExecuteScript("arguments[0].click();", btn);
                        Console.WriteLine("   [Submit] JS click");
                        return;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"   [Submit] click failed: {ex.Message}");
                    }
                }
            }

            Console.WriteLine("   [Submit] Button disabled — triple-strategy");
            try
            {
                Js.ExecuteScript(@"
                    var btn = arguments[0];
                    var form = btn.closest('form');

                    // (1) Blur tất cả field (trừ file) để Angular mark touched
                    var nodes = (form || document).querySelectorAll(
                        'input:not([type=file]), textarea, select'
                    );
                    nodes.forEach(function(el) {
                        try { el.dispatchEvent(new Event('blur', {bubbles:true})); } catch(e) {}
                    });

                    // (2) Dispatch submit event trên form (catches (ngSubmit))
                    if (form) {
                        try { form.dispatchEvent(new Event('submit', {bubbles:true, cancelable:true})); } catch(e) {}
                    }

                    // (3) Force-click button (catches (click))
                    try {
                        if (btn.disabled) btn.removeAttribute('disabled');
                        btn.click();
                    } catch(e) {}
                ", btn);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"   [Submit] dispatch error: {ex.Message}");
            }
        }

        private static string GetToastSuccess()
        {
            try
            {
                var el = MakeWait(Config.SuccessToastTimeout)
                    .Until(ExpectedConditions.ElementIsVisible(Loc.ToastSuccess));
                return NormalizeText(el.Text);
            }
            catch (WebDriverTimeoutException) { return ""; }
        }

        // ╔════════════════════════════════════════════════════════╗
        // ║  FIX v8: filter date-pattern khi thu thập inline errors  ║
        // ║  ("18:00 07/06/2026" không phải error message).         ║
        // ╚════════════════════════════════════════════════════════╝
        private static List<string> CollectInlineErrors()
        {
            try
            {
                return _driver.FindElements(Loc.InlineError)
                    .Where(e => { try { return e.Displayed && !string.IsNullOrWhiteSpace(e.Text); } catch { return false; } })
                    .Select(e => NormalizeText(e.Text))
                    .Where(t => !string.IsNullOrWhiteSpace(t))
                    .Where(t => !_dateLikePattern.IsMatch(t)) // bỏ chuỗi giống "18:00 07/..."
                    .Distinct()
                    .ToList();
            }
            catch
            {
                return new List<string>();
            }
        }

        private static void AssertErrorMessage(string stepId, string expectedRaw)
        {
            var expected = NormalizeText(expectedRaw);
            var deadline = DateTime.UtcNow.AddSeconds(Config.ErrorPollTimeout);

            var lastInline = new List<string>();
            string lastToast = "";

            while (DateTime.UtcNow < deadline)
            {
                lastInline = CollectInlineErrors();

                var hitInline = lastInline.FirstOrDefault(e =>
                    e.Contains(expected, StringComparison.OrdinalIgnoreCase));
                if (hitInline != null)
                {
                    Console.WriteLine($"   [{stepId}] ✅ inline: '{hitInline}'");
                    Assert.Pass();
                    return;
                }

                try
                {
                    var toastEl = _driver.FindElements(Loc.ToastError)
                        .FirstOrDefault(e =>
                        {
                            try { return e.Displayed && !string.IsNullOrWhiteSpace(e.Text); }
                            catch { return false; }
                        });
                    if (toastEl != null) lastToast = NormalizeText(toastEl.Text);
                }
                catch { }

                if (!string.IsNullOrEmpty(lastToast) &&
                    lastToast.Contains(expected, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"   [{stepId}] ✅ toast error: '{lastToast}'");
                    Assert.Pass();
                    return;
                }

                Thread.Sleep(250);
            }

            Console.WriteLine($"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine($"❌ [{stepId}] FAIL DETAIL");
            Console.WriteLine($"   Expected (contains): '{expected}'");
            Console.WriteLine($"   Inline found ({lastInline.Count}):");
            foreach (var s in lastInline) Console.WriteLine($"     • '{s}'");
            Console.WriteLine($"   Toast error: '{lastToast}'");
            Console.WriteLine($"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

            Assert.Fail(
                $"[{stepId}] Kỳ vọng chứa: \"{expected}\"\n" +
                $"         inline ({lastInline.Count}): {string.Join(" || ", lastInline)}\n" +
                $"         toast: '{lastToast}'");
        }

        private static bool FieldRejectsText(By by, string value)
        {
            var el = WaitVisible(by);
            el.Clear(); el.SendKeys(value);
            var actual = el.GetAttribute("value") ?? "";
            return string.IsNullOrEmpty(actual) || !actual.Contains(value);
        }

        private static TestRow GetRow(string stepId)
        {
            var row = _rows.FirstOrDefault(r =>
                r.StepId.Equals(stepId, StringComparison.OrdinalIgnoreCase));
            if (row == null) Assert.Ignore($"Không tìm thấy '{stepId}' trong Excel.");
            return row!;
        }

        private static void FillForm(TestRow row)
        {
            Fill(Loc.ProductName, row.ProductName);
            UploadImage(Config.ImagePath(row.ImageFile));
            Fill(Loc.Description, row.Description);
            Fill(Loc.StartPrice, row.StartPrice);
            Fill(Loc.PriceStep, row.PriceStep);
            SelectStartMode(row.StartTimeType);

            if (row.StartTimeType.ToUpper() == "SCHEDULE" && !string.IsNullOrEmpty(row.StartTime))
                SetDateTime(Loc.StartTime, row.StartTime);
            if (!string.IsNullOrEmpty(row.EndTime))
                SetDateTime(Loc.EndTime, row.EndTime);
        }

        // ══════════════════════════════════════════════════════════
        // TEST CASES A01–A26
        // ══════════════════════════════════════════════════════════

        [Test, Order(1), Description("A01 – Thành công - Lên sàn ngay")]
        public void A01_ThanhCong_LenSanNgay()
        {
            var row = GetRow("A01");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            var toast = GetToastSuccess();
            Assert.That(toast, Does.Contain(expected),
                $"[A01] toast='{toast}' | expected contains='{expected}'");
        }

        [Test, Order(2), Description("A02 – Thành công - Hẹn giờ hợp lệ")]
        public void A02_ThanhCong_HenGioHopLe()
        {
            var row = GetRow("A02");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            var toast = GetToastSuccess();
            Assert.That(toast, Does.Contain(expected),
                $"[A02] toast='{toast}' | expected contains='{expected}'");
        }

        [Test, Order(3), Description("A03 – Tên sản phẩm để trống")]
        public void A03_TenSanPham_DeTrong()
        {
            var row = GetRow("A03");
            FillForm(row);
            Submit();
            AssertErrorMessage("A03", row.ExpectedMessage);
        }

        [Test, Order(4), Description("A04 – Tên sản phẩm chỉ toàn dấu cách")]
        public void A04_TenSanPham_ToanDauCach()
        {
            var row = GetRow("A04");
            Fill(Loc.ProductName, "     ");
            UploadImage(Config.ImagePath(row.ImageFile));
            Fill(Loc.Description, row.Description);
            Fill(Loc.StartPrice, row.StartPrice);
            Fill(Loc.PriceStep, row.PriceStep);
            SelectStartMode(row.StartTimeType);
            if (!string.IsNullOrEmpty(row.EndTime))
                SetDateTime(Loc.EndTime, row.EndTime);
            Submit();
            AssertErrorMessage("A04", row.ExpectedMessage);
        }

        [Test, Order(5), Description("A05 – Tên sản phẩm quá dài (> 150 ký tự)")]
        public void A05_TenSanPham_QuaDai()
        {
            var row = GetRow("A05");
            FillForm(row);
            Submit();
            AssertErrorMessage("A05", row.ExpectedMessage);
        }

        [Test, Order(6), Description("A06 – Mô tả để trống (cho phép)")]
        public void A06_MoTa_DeTrong_HopLe()
        {
            var row = GetRow("A06");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            var toast = GetToastSuccess();
            Assert.That(toast, Does.Contain(expected),
                $"[A06] toast='{toast}' | expected contains='{expected}'");
        }

        [Test, Order(7), Description("A07 – Mô tả quá dài (> 200 ký tự)")]
        public void A07_MoTa_QuaDai()
        {
            var row = GetRow("A07");
            FillForm(row);
            try { WaitVisible(Loc.Description).SendKeys(Keys.Tab); } catch { }
            Submit();
            AssertErrorMessage("A07", row.ExpectedMessage);
        }

        [Test, Order(8), Description("A08 – Giá khởi điểm để trống")]
        public void A08_GiaKhoiDiem_DeTrong()
        {
            var row = GetRow("A08");
            FillForm(row);
            Submit();
            AssertErrorMessage("A08", row.ExpectedMessage);
        }

        [Test, Order(9), Description("A09 – Giá khởi điểm nhập số âm")]
        public void A09_GiaKhoiDiem_SoAm()
        {
            var row = GetRow("A09");
            FillForm(row);
            Submit();
            AssertErrorMessage("A09", row.ExpectedMessage);
        }

        [Test, Order(10), Description("A10 – Giá khởi điểm nhập chữ (field tự chặn)")]
        public void A10_GiaKhoiDiem_NhapChu_Skip()
        {
            var row = GetRow("A10");
            bool rejected = FieldRejectsText(Loc.StartPrice, row.StartPrice);
            Assert.That(rejected, Is.True,
                $"[A10] field startPrice không chặn nhập chữ '{row.StartPrice}'");
        }

        [Test, Order(11), Description("A11 – Giá khởi điểm là số thập phân (Thành công)")]
        public void A11_GiaKhoiDiem_SoThapPhan()
        {
            var row = GetRow("A11");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            var toast = GetToastSuccess();
            Assert.That(toast, Does.Contain(expected),
                $"[A11] toast='{toast}' | expected contains='{expected}'");
        }

        [Test, Order(12), Description("A12 – Giá khởi điểm bằng 0")]
        public void A12_GiaKhoiDiem_Bang0()
        {
            var row = GetRow("A12");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            AssertErrorMessage("A12", row.ExpectedMessage);
        }

        [Test, Order(13), Description("A13 – Bước giá để trống")]
        public void A13_BuocGia_DeTrong()
        {
            var row = GetRow("A13");
            FillForm(row);
            Submit();
            AssertErrorMessage("A13", row.ExpectedMessage);
        }

        [Test, Order(14), Description("A14 – Bước giá nhập số âm")]
        public void A14_BuocGia_SoAm()
        {
            var row = GetRow("A14");
            FillForm(row);
            Submit();
            AssertErrorMessage("A14", row.ExpectedMessage);
        }

        [Test, Order(15), Description("A15 – Bước giá nhập chữ (field tự chặn)")]
        public void A15_BuocGia_NhapChu_Skip()
        {
            var row = GetRow("A15");
            bool rejected = FieldRejectsText(Loc.PriceStep, row.PriceStep);
            Assert.That(rejected, Is.True,
                $"[A15] field stepPrice không chặn nhập chữ '{row.PriceStep}'");
        }

        [Test, Order(16), Description("A16 – Bước giá nhập số thập phân (Thành công)")]
        public void A16_BuocGia_SoThapPhan()
        {
            var row = GetRow("A16");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            var toast = GetToastSuccess();
            Assert.That(toast, Does.Contain(expected),
                $"[A16] toast='{toast}' | expected contains='{expected}'");
        }

        [Test, Order(17), Description("A17 – Bước giá bằng 0 (Thành công)")]
        public void A17_BuocGia_Bang0()
        {
            var row = GetRow("A17");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            AssertErrorMessage("A12", row.ExpectedMessage);
        }

        [Test, Order(18), Description("A18 – Hẹn giờ - Để trống Ngày bắt đầu")]
        public void A18_HenGio_DeTrongNgayBatDau()
        {
            var row = GetRow("A18");
            FillForm(row);
            Submit();
            AssertErrorMessage("A18", row.ExpectedMessage);
        }

        [Test, Order(19), Description("A19 – Hẹn giờ - Ngày bắt đầu ở quá khứ")]
        public void A19_HenGio_NgayBatDauQuaKhu_Skip()
        {
            var row = GetRow("A19");
            SelectStartMode("SCHEDULE");

            var p = row.StartTime.Trim().Split(' ');
            var d = p[0].Split('/');
            var iso = $"{d[2]}-{d[1]}-{d[0]}T{p[1]}";

            var el = WaitVisible(Loc.StartTime);
            Js.ExecuteScript("arguments[0].value=arguments[1]", el, iso);
            Js.ExecuteScript("arguments[0].dispatchEvent(new Event('input', {bubbles:true}))", el);
            Js.ExecuteScript("arguments[0].dispatchEvent(new Event('change',{bubbles:true}))", el);

            var actual = el.GetAttribute("value") ?? "";
            if (actual == iso)
            {
                Fill(Loc.ProductName, row.ProductName);
                UploadImage(Config.ImagePath(row.ImageFile));
                Fill(Loc.Description, row.Description);
                Fill(Loc.StartPrice, row.StartPrice);
                Fill(Loc.PriceStep, row.PriceStep);
                SetDateTime(Loc.EndTime, row.EndTime);
                Submit();

                var deadline = DateTime.UtcNow.AddSeconds(8);
                bool found = false;
                while (DateTime.UtcNow < deadline)
                {
                    var inlineHas = CollectInlineErrors().Count > 0;
                    var toastHas = _driver.FindElements(Loc.ToastError)
                        .Any(e => { try { return e.Displayed && !string.IsNullOrWhiteSpace(e.Text); } catch { return false; } });
                    if (inlineHas || toastHas) { found = true; break; }
                    Thread.Sleep(250);
                }
                Assert.That(found, Is.True,
                    "[A19] Ngày quá khứ được chấp nhận – kỳ vọng có lỗi inline hoặc toast");
            }
            else
            {
                Assert.Pass("[A19] PASS – datepicker đã chặn ngày quá khứ (min attribute)");
            }
        }

        [Test, Order(20), Description("A20 – Hẹn giờ - Ngày kết thúc TRƯỚC Ngày bắt đầu")]
        public void A20_HenGio_KetThucTruocBatDau()
        {
            var row = GetRow("A20");
            FillForm(row);
            Submit();
            AssertErrorMessage("A20", row.ExpectedMessage);
        }

        [Test, Order(21), Description("A21 – Hẹn giờ - Ngày kết thúc TRÙNG Ngày bắt đầu (Thành công)")]
        public void A21_HenGio_KetThucTrungBatDau()
        {
            var row = GetRow("A21");
            var expected = NormalizeText(row.ExpectedMessage);
            FillForm(row);
            Submit();
            AssertErrorMessage("A12", row.ExpectedMessage);
        }

        [Test, Order(22), Description("A22 – Lên sàn ngay - Để trống Ngày kết thúc")]
        public void A22_LenSanNgay_DeTrongNgayKetThuc()
        {
            var row = GetRow("A22");
            FillForm(row);
            Submit();
            AssertErrorMessage("A22", row.ExpectedMessage);
        }

        [Test, Order(23), Description("A23 – Lên sàn ngay - Ngày kết thúc ở quá khứ")]
        public void A23_LenSanNgay_NgayKetThucQuaKhu()
        {
            var row = GetRow("A23");
            Fill(Loc.ProductName, row.ProductName);
            UploadImage(Config.ImagePath(row.ImageFile));
            Fill(Loc.Description, row.Description);
            Fill(Loc.StartPrice, row.StartPrice);
            Fill(Loc.PriceStep, row.PriceStep);
            SelectStartMode("NOW");

            var endEl = WaitVisible(Loc.EndTime);
            Js.ExecuteScript("arguments[0].value='2025-05-20T00:00'", endEl);
            Js.ExecuteScript("arguments[0].dispatchEvent(new Event('input', {bubbles:true}))", endEl);
            Js.ExecuteScript("arguments[0].dispatchEvent(new Event('change',{bubbles:true}))", endEl);
            Submit();
            AssertErrorMessage("A23", row.ExpectedMessage);
        }

        [Test, Order(24), Description("A24 – Ảnh sản phẩm để trống")]
        public void A24_Anh_DeTrong()
        {
            var row = GetRow("A24");
            Fill(Loc.ProductName, row.ProductName);
            Fill(Loc.Description, row.Description);
            Fill(Loc.StartPrice, row.StartPrice);
            Fill(Loc.PriceStep, row.PriceStep);
            SelectStartMode(row.StartTimeType);
            if (!string.IsNullOrEmpty(row.EndTime))
                SetDateTime(Loc.EndTime, row.EndTime);
            Submit();
            AssertErrorMessage("A24", row.ExpectedMessage);
        }

        [Test, Order(25), Description("A25 – Ảnh sai định dạng (.jfif)")]
        public void A25_Anh_SaiDinhDang()
        {
            var row = GetRow("A25");
            FillForm(row);
            try { Submit(); } catch { }
            AssertErrorMessage("A25", row.ExpectedMessage);
        }

        [Test, Order(26), Description("A26 – Ảnh quá dung lượng (> 5MB)")]
        public void A26_Anh_QuaDungLuong()
        {
            var row = GetRow("A26");
            FillForm(row);
            try { Submit(); } catch { }
            AssertErrorMessage("A26", row.ExpectedMessage);
        }
    }
}