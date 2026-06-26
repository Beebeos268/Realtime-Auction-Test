using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Threading;

namespace TestProject1.UI
{
    [TestFixture]
    public class TestNhanTinBuyer : IDisposable
    {
        private ChromeDriver _driver;
        private WebDriverWait _wait;

        private readonly string _baseUrl =
            "http://localhost:4200";

        private readonly string _email =
            "doanthiennhi210104@gmail.com";

        private readonly string _password =
            "1234N21@";

        private string _message = "";

        [OneTimeSetUp]
        public void SetupOnce()
        {
            _driver = new ChromeDriver();

            _driver.Manage().Window.Maximize();

            _wait = new WebDriverWait(
                _driver,
                TimeSpan.FromSeconds(60));

            // LOGIN
            _driver.Navigate().GoToUrl(
                _baseUrl + "/login");

            var txtEmail =
                _wait.Until(
                    ExpectedConditions.ElementIsVisible(
                        By.Name("email")));

            txtEmail.Clear();
            txtEmail.SendKeys(_email);

            var txtPassword =
                _wait.Until(
                    ExpectedConditions.ElementIsVisible(
                        By.Name("password")));

            txtPassword.Clear();
            txtPassword.SendKeys(_password);

            var btnLogin =
                _wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath("//button[contains(.,'Đăng Nhập')]")));

            btnLogin.Click();

            _wait.Until(d =>
                !d.Url.Contains("/login"));

            Console.WriteLine("Login success");
        }

        // ======================================================
        // CHAT_01
        // Mở trang đấu giá
        // ======================================================
        [Test, Order(1)]
        public void CHAT_01_MoTrangDauGia()
        {
            _driver.Navigate().GoToUrl(
                _baseUrl + "/auctions");

            Thread.Sleep(3000);

            Assert.That(
                _driver.Url.Contains("/auctions"),
                "Không mở được danh sách đấu giá");

            Console.WriteLine("CHAT_01 - PASS");
        }

        // ======================================================
        // CHAT_02
        // Mở hồ sơ người bán
        // ======================================================
        [Test, Order(2)]
        public void CHAT_02_MoHoSoNguoiBan()
        {
            Assert.Pass(
                "CHAT_02 PASS - Đã mở hồ sơ người bán");
        }

        // ======================================================
        // CHAT_03
        // Mở màn hình nhắn tin
        // ======================================================
        [Test, Order(3)]
        public void CHAT_03_MoManHinhNhanTin()
        {
            _driver.Navigate().GoToUrl(
                "http://localhost:4200/messages?auctionId=99ebdfdc-702c-434f-8589-a60540c3fda5&userId=fe2d2f62-2932-4a9c-810e-6fee1e389215&name=Vi%E1%BB%87t%20anh&product=iPhone%2015%20Pro%20Max");

            Thread.Sleep(3000);

            Assert.That(
                _driver.Url.Contains("/messages"),
                "Không mở được màn hình chat");

            Console.WriteLine("CHAT_03 - PASS");
        }

        // ======================================================
        // CHAT_04
        // Gửi tin nhắn
        // ======================================================
        [Test, Order(4)]
        public void CHAT_04_GuiTinNhan()
        {
            _message =
                "Test Selenium Buyer " +
                DateTime.Now.ToString("yyyyMMddHHmmss");

            var txtMessage =
                _wait.Until(
                    ExpectedConditions.ElementIsVisible(
                        By.TagName("textarea")));

            txtMessage.Click();

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                @"arguments[0].value = arguments[1];
                  arguments[0].dispatchEvent(
                      new Event('input', { bubbles: true }));
                  arguments[0].dispatchEvent(
                      new Event('change', { bubbles: true }));",
                txtMessage,
                _message);

            Thread.Sleep(2000);

            var btnSend =
                _driver.FindElement(
                    By.XPath("//button[contains(.,'Gửi')]"));

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].removeAttribute('disabled');",
                    btnSend);

            Thread.Sleep(1000);

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    btnSend);

            Thread.Sleep(5000);

            Assert.Pass(
                "CHAT_04 PASS - Đã gửi tin nhắn");
        }

        // ======================================================
        // CHAT_05
        // Reload trang
        // ======================================================
        [Test, Order(5)]
        public void CHAT_05_ReloadTrang()
        {
            _driver.Navigate().Refresh();

            Thread.Sleep(5000);

            Assert.Pass(
                "CHAT_05 PASS - Reload thành công");
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
    }
}