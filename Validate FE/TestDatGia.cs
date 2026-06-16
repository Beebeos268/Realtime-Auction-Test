using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;

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
        private readonly string _email = "doanthiennhi210104@gmail.com";
        private readonly string _password = "1234N21@";

        // Số tiền đặt giá ở BID_05 — chỉnh nếu cần
        private readonly string _bidAmount = "2200000";

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
            // Tìm ô nhập giá thầu (input number trong form chốt giá)
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

            inputGia.Clear();
            inputGia.SendKeys(_bidAmount);

            // Click CHỐT GIÁ
            var btnChotGia = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(.,'CHỐT GIÁ')]")
                )
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    btnChotGia
                );

            // Nếu có dialog confirm thì click xác nhận
            try
            {
                var btnConfirm = new WebDriverWait(
                    _driver,
                    TimeSpan.FromSeconds(5)
                ).Until(
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
                // Không có dialog -> bỏ qua
                Console.WriteLine("BID_05 - Không có dialog confirm");
            }

            // Đợi UI cập nhật
            Thread.Sleep(2000);

            Console.WriteLine("BID_05 - OK (đã đặt giá " + _bidAmount + ")");
        }

        // ======================================================
        // BID_06
        // Kiểm tra giá hiện tại cập nhật sau khi đặt giá
        // ======================================================
        [Test, Order(6)]
        public void BID_06_KiemTraGiaHienTaiCapNhat()
        {
            // Hiển thị số tiền vừa đặt — chấp nhận cả format
            // "2.200.000", "2,200,000" hoặc "2200000"
            var lblGiaMoi = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(
                        "//*[contains(text(),'2.200.000') or " +
                        "contains(text(),'2,200,000') or " +
                        "contains(text(),'2200000')]"
                    )
                )
            );

            Assert.That(
                lblGiaMoi.Displayed,
                "Giá hiện tại không cập nhật sau khi đặt giá"
            );

            Console.WriteLine("BID_06 - OK (giá hiện tại đã cập nhật)");
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
            var lblDangDanDau = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(
                        "//*[contains(translate(normalize-space(.), " +
                        "'ĐANGDẪĐẦU', 'đangdẫđầu'), 'đang dẫn đầu')]"
                    )
                )
            );

            Assert.That(
                lblDangDanDau.Displayed,
                "Không thấy 'đang dẫn đầu' trong lịch sử trả giá"
            );

            Console.WriteLine("BID_07 - OK");
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