using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;

namespace TestProject1
{
    [TestFixture]
    public class TestDanhSachDauGia : IDisposable
    {
        private ChromeDriver _driver;
        private WebDriverWait _wait;

        private readonly string _baseUrl = "http://localhost:4200";

        // =========================
        // LOGIN ACCOUNT TEST
        // =========================
        private readonly string _email = "vietanhdd2608@gmail.com";
        private readonly string _password = "Vietanh268@";

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

            // đợi page load
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
        // LIST_01
        // Click menu -> click danh sách đấu giá
        // ======================================================
        [Test, Order(1)]
        public void LIST_01_Navigation()
        {
            var btnMenu = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(.,'Menu')]")
                )
            );

            btnMenu.Click();

            var linkDanhSach = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(.,'Danh sách đấu giá')]")
                )
            );

            linkDanhSach.Click();

            WaitForPageLoad();

            var title = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'Danh sách đấu giá')]")
                )
            );

            Assert.That(
                title.Displayed,
                "Không vào được trang Danh sách đấu giá"
            );

            Console.WriteLine("LIST_01 - OK");
        }

        // ======================================================
        // LIST_02
        // Click tham gia đấu giá
        // ======================================================
        [Test, Order(2)]
        public void LIST_02_ThamGiaDauGia()
        {
            _wait.Until(
                ExpectedConditions.ElementExists(
                    By.CssSelector(".browse-card")
                )
            );

            var btnThamGia = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("(//a[contains(.,'Tham gia đấu giá')])[1]")
                )
            );

            Assert.That(
                btnThamGia.Displayed,
                "Không thấy nút Tham gia đấu giá"
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    btnThamGia
                );

            _wait.Until(d =>
                d.Url.Contains("/auction/"));

            WaitForPageLoad();

            Console.WriteLine("LIST_02 - OK");
        }

        // ======================================================
        // LIST_03
        // Check nút chốt giá
        // ======================================================
        [Test, Order(3)]
        public void LIST_03_CheckUI_ChiTiet()
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

            Console.WriteLine("LIST_03 - OK");
        }

        // ======================================================
        // LIST_04
        // Click Trang chủ
        // ======================================================
        [Test, Order(4)]
        public void LIST_04_ClickHome()
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

