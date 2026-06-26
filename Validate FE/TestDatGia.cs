using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TestProject1.UI
{
    [TestFixture]
    public class TestDatGia : IDisposable
    {
        private ChromeDriver _driver;
        private WebDriverWait _wait;

        private readonly string _baseUrl = "http://localhost:4200";

        // =========================
        // LOGIN ACCOUNT TEST
        // =========================
        private readonly string _email = "vietanhdd268@gmail.com";
        private readonly string _password = "Vietanh268@";
        private string ReadBidInputAmount(IWebElement input)
        {
            var value = input.GetAttribute("value") ?? "";
            return DigitsOnly(value);
        }
        private static string NormalizeVietnamese(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            text = text.ToLowerInvariant()
                .Replace("đ", "d")
                .Replace("Đ", "d")
                .Replace("•", " ")
                .Replace("-", " ");

            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (var c in normalized)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(c);
                if (category != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }

            return Regex.Replace(sb.ToString().Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();
        }

        // Số tiền đặt giá ở BID_05 — chỉnh nếu cần
        private string _bidAmount = "";


        [OneTimeSetUp]
        public void SetupOnce()
        {
            _driver = new ChromeDriver();

            _driver.Manage().Window.Maximize();

            _wait = new WebDriverWait(
                _driver,
                TimeSpan.FromSeconds(20)
            );

            // =========================
            // GO TO LOGIN PAGE
            // =========================
            _driver.Navigate().GoToUrl(_baseUrl + "/login");

            WaitForPageLoad();

            Console.WriteLine("Current URL: " + _driver.Url);

            // =========================
            // INPUT EMAIL
            // =========================
            var txtEmail = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.Name("email")
                )
            );

            txtEmail.Clear();
            txtEmail.SendKeys(_email);

            // =========================
            // INPUT PASSWORD
            // =========================
            var txtPassword = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.Name("password")
                )
            );

            txtPassword.Clear();
            txtPassword.SendKeys(_password);

            // =========================
            // CLICK LOGIN
            // =========================
            var btnLogin = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(.,'Đăng Nhập')]")
                )
            );

            btnLogin.Click();

            // =========================
            // WAIT LOGIN SUCCESS
            // =========================
            _wait.Until(d =>
                !d.Url.Contains("/login"));

            WaitForPageLoad();

            Console.WriteLine("Login success!");
            Console.WriteLine("After login URL: " + _driver.Url);
        }

        private void WaitForPageLoad()
        {
            _wait.Until(driver =>
                ((IJavaScriptExecutor)driver)
                    .ExecuteScript("return document.readyState")
                    .Equals("complete"));
        }

        [OneTimeTearDown]
        public void TeardownOnce()
        {
            _driver?.Quit();
            _driver?.Dispose();
        }

        public void Dispose()
        {
            _driver?.Dispose();
        }

        // ======================================================
        // BID_01
        // Mở trang chi tiết phiên đấu giá
        // (Menu -> Danh sách đấu giá -> click phiên đầu tiên)
        // ======================================================
        [Test, Order(1)]
        public void BID_01_MoChiTietPhienDauGia()
        {
            // Click Menu
            var btnMenu = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(.,'Menu')]")
                )
            );

            btnMenu.Click();

            // Click "Danh sách đấu giá"
            var linkDanhSach = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(.,'Danh sách đấu giá')]")
                )
            );

            linkDanhSach.Click();

            WaitForPageLoad();

            // Đợi danh sách load xong
            _wait.Until(
                ExpectedConditions.ElementExists(
                    By.CssSelector(".browse-card")
                )
            );

            // Click vào phiên đấu giá đầu tiên
            var btnThamGia = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("(//a[contains(.,'Tham gia đấu giá')])[1]")
                )
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    btnThamGia
                );

            _wait.Until(d =>
                d.Url.Contains("/auction/"));

            WaitForPageLoad();

            Assert.That(
                _driver.Url.Contains("/auction/"),
                "Không vào được trang chi tiết phiên đấu giá"
            );

            Console.WriteLine("BID_01 - OK: " + _driver.Url);
        }

        // ======================================================
        // BID_02
        // Kiểm tra hiển thị nút CHỐT GIÁ
        // ======================================================
        [Test, Order(2)]
        public void BID_02_HienThiNutChotGia()
        {
            var btnChotGia = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//button[contains(.,'CHỐT GIÁ')]")
                )
            );

            Assert.That(
                btnChotGia.Displayed,
                "Không thấy nút CHỐT GIÁ"
            );

            Assert.That(
                btnChotGia.Enabled,
                "Nút CHỐT GIÁ bị disable"
            );

            Console.WriteLine("BID_02 - OK");
        }

        // ======================================================
        // BID_03
        // Kiểm tra hiển thị "GIÁ KHỞI ĐIỂM"
        // ======================================================
        [Test, Order(3)]
        public void BID_03_HienThiGiaKhoiDiem()
        {
            // Tìm label "GIÁ KHỞI ĐIỂM" hoặc "Giá khởi điểm" (không phụ thuộc hoa/thường)
            var lblGiaKhoiDiem = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(
                        "//*[contains(translate(normalize-space(.), " +
                        "'GIÁKHỞIĐIỂM', 'giákhởiđiểm'), 'giá khởi điểm')]"
                    )
                )
            );

            Assert.That(
                lblGiaKhoiDiem.Displayed,
                "Không thấy 'Giá khởi điểm' trên trang chi tiết"
            );

            Console.WriteLine("BID_03 - OK");
        }

        // ======================================================
        // BID_04
        // Kiểm tra hiển thị "BƯỚC GIÁ"
        // ======================================================
        [Test, Order(4)]
        public void BID_04_HienThiBuocGia()
        {
            var lblBuocGia = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(
                        "//*[contains(translate(normalize-space(.), " +
                        "'BƯỚCGIÁ', 'bướcgiá'), 'bước giá')]"
                    )
                )
            );

            Assert.That(
                lblBuocGia.Displayed,
                "Không thấy 'Bước giá' trên trang chi tiết"
            );

            Console.WriteLine("BID_04 - OK");
        }

        // ======================================================
        // BID_05
        // Đặt giá thầu hợp lệ
        // (Nhập số tiền -> Click CHỐT GIÁ -> Xác nhận)
        // ======================================================
        [Test, Order(5)]
        public void BID_05_DatGiaThauHopLe()
        {
            var inputGia = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.CssSelector(
                        "input[type='number'], " +
                        "input[name*='price'], " +
                        "input[name*='bid'], " +
                        "input[placeholder*='giá']"
                    )
                )
            );

            inputGia.Click();
            inputGia.SendKeys(Keys.Control + "a");
            inputGia.SendKeys(Keys.Delete);

            // Bấm nút Giá tối thiểu để luôn lấy giá hợp lệ hiện tại
            var btnGiaToiThieu = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(normalize-space(.),'Giá tối thiểu')]")
                )
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript("arguments[0].click();", btnGiaToiThieu);

            Thread.Sleep(500);

            // Bấm thêm + Bước giá để chắc chắn vượt tối thiểu
            var btnCongBuocGia = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(normalize-space(.),'+ Bước giá') or (contains(.,'+') and contains(.,'Bước giá'))]")
                )
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript("arguments[0].click();", btnCongBuocGia);

            Thread.Sleep(500);

            _bidAmount = ReadBidInputAmount(inputGia);

            Console.WriteLine("BID_05 - Giá sẽ đặt: " + _bidAmount);

            Assert.That(
                string.IsNullOrWhiteSpace(_bidAmount),
                Is.False,
                "Không đọc được giá đặt từ input"
            );

            var btnChotGia = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(.,'CHỐT GIÁ')]")
                )
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript("arguments[0].click();", btnChotGia);

            try
            {
                var btnConfirm = new WebDriverWait(_driver, TimeSpan.FromSeconds(5))
                    .Until(
                        ExpectedConditions.ElementToBeClickable(
                            By.XPath(
                                "//button[contains(.,'Xác nhận') or " +
                                "contains(.,'Đồng ý') or " +
                                "contains(.,'OK')]"
                            )
                        )
                    );

                btnConfirm.Click();
                Console.WriteLine("BID_05 - Đã xác nhận dialog");
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("BID_05 - Không có dialog confirm");
            }

            Thread.Sleep(2000);

            Console.WriteLine("BID_05 - OK, đã đặt giá: " + _bidAmount);
        }


        private string GetBodyText()
        {
            try
            {
                return _driver.FindElement(By.TagName("body")).Text ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string DigitsOnly(string text)
        {
            return Regex.Replace(text ?? "", @"\D", "");
        }

        private static bool PageTextContainsMoney(string bodyText, string amount)
        {
            var expectedDigits = DigitsOnly(amount);
            if (string.IsNullOrWhiteSpace(expectedDigits)) return false;

            // Chấp nhận 2.200.000 / 2,200,000 / 2 200 000 / 2200000 / 2.200.000 VNĐ
            var bodyDigits = DigitsOnly(bodyText);
            if (bodyDigits.Contains(expectedDigits)) return true;

            return bodyText.Contains(amount);
        }

        private void DumpBidPageText(string title)
        {
            try
            {
                var bodyText = GetBodyText();
                Console.WriteLine("========== " + title + " ==========");
                Console.WriteLine("URL: " + _driver.Url);
                Console.WriteLine("Expected bid amount: " + _bidAmount);
                Console.WriteLine("Body contains expected digits: " + PageTextContainsMoney(bodyText, _bidAmount));
                Console.WriteLine(bodyText.Length > 5000 ? bodyText.Substring(0, 5000) : bodyText);
                Console.WriteLine("========== END DUMP ==========");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Không dump được text trang: " + ex.Message);
            }
        }

        // ======================================================
        // BID_06
        // Kiểm tra giá hiện tại cập nhật sau khi đặt giá
        // ======================================================
        [Test, Order(6)]
        public void BID_06_KiemTraGiaHienTaiCapNhat()
        {
            Assert.That(
                string.IsNullOrWhiteSpace(_bidAmount),
                Is.False,
                "BID_06 không có _bidAmount. Hãy chạy từ BID_01 đến BID_06, không chạy riêng BID_06."
            );

            bool found = _wait.Until(d =>
            {
                var bodyText = d.FindElement(By.TagName("body")).Text ?? "";
                var bodyDigits = DigitsOnly(bodyText);

                return bodyDigits.Contains(_bidAmount);
            });

            Assert.That(
                found,
                Is.True,
                $"Không thấy giá vừa đặt {_bidAmount} trên trang"
            );

            Console.WriteLine("BID_06 - OK, giá hiện tại đã cập nhật: " + _bidAmount);
        }

        // ======================================================
        // BID_07
        // Kiểm tra hiển thị lịch sử trả giá có "đang dẫn đầu"
        //
        // FIX: UI hiển thị "đang dẫn đầu" (chữ đ thường) chứ không
        // phải "Đang dẫn đầu" (chữ Đ hoa) -> dùng translate để khớp
        // không phụ thuộc hoa/thường
        // ======================================================
        [Test, Order(7)]
        public void BID_07_HienThiLichSuTraGia_DangDanDau()
        {
            ((IJavaScriptExecutor)_driver)
                .ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");

            Thread.Sleep(800);

            try
            {
                bool found = _wait.Until(d =>
                {
                    var bodyText = d.FindElement(By.TagName("body")).Text ?? "";
                    var normalized = NormalizeVietnamese(bodyText);

                    return
                        normalized.Contains("ban dan dau") ||
                        normalized.Contains("ban dang dan dau") ||
                        normalized.Contains("dang dan dau") ||
                        normalized.Contains("dan dau") ||
                        bodyText.Contains("Bạn • dẫn đầu") ||
                        bodyText.Contains("Bạn - dẫn đầu") ||
                        bodyText.Contains("Bạn đang dẫn đầu");
                });

                Assert.That(found, Is.True, "Không thấy trạng thái dẫn đầu trong bảng xếp hạng.");
                Console.WriteLine("BID_07 - OK: đã thấy người dùng đang dẫn đầu.");
            }
            catch (WebDriverTimeoutException)
            {
                var bodyText = _driver.FindElement(By.TagName("body")).Text ?? "";

                Console.WriteLine("========== BID_07 FAIL - BODY TEXT ==========");
                Console.WriteLine(bodyText);
                Console.WriteLine("========== END ==========");

                Assert.Fail("Không thấy trạng thái 'Bạn • dẫn đầu' hoặc 'dẫn đầu' trên trang.");
            }
        }

        [Test, Order(8)]
        public void BID_08_ClickHome()
        {
            var linkTrangChu = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("(//a[contains(.,'Trang chủ')])[1]")
                )
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    linkTrangChu
                );

            _wait.Until(d =>
                d.Url == _baseUrl ||
                d.Url == _baseUrl + "/");

            Assert.That(
                _driver.Url == _baseUrl ||
                _driver.Url == _baseUrl + "/",
                "Không quay về trang chủ"
            );

            Console.WriteLine("LIST_04 - OK");
        }

    }
}