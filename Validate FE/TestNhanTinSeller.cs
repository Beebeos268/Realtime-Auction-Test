using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;

namespace TestProject1.UI
{
    [TestFixture]
    public class TestNhanTinSeller : IDisposable
    {
        private ChromeDriver _driver;
        private WebDriverWait _wait;

        private readonly string _baseUrl =
            "http://localhost:4200";

        private readonly string _email =
            "vietanhdd268@gmail.com";

        private readonly string _password =
            "Vietanh268@";

        private string _replyMessage = "";

        [OneTimeSetUp]
        public void SetupOnce()
        {
            _driver = new ChromeDriver();

            _driver.Manage().Window.Maximize();

            _wait = new WebDriverWait(
                _driver,
                TimeSpan.FromSeconds(30));

            LoginSeller();
        }

        private void LoginSeller()
        {
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

            Console.WriteLine(
                "Login Seller Success");
        }

        // =====================================
        // CHAT_SELLER_01
        // Mở trang tin nhắn
        // =====================================
        [Test, Order(1)]
        public void CHAT_SELLER_01_MoTrangTinNhan()
        {
            _driver.Navigate().GoToUrl(
                _baseUrl + "/messages");

            Assert.That(
                _driver.Url.Contains("/messages"));

            Console.WriteLine(
                "CHAT_SELLER_01 PASS");
        }

        // =====================================
        // CHAT_SELLER_02
        // Hiển thị hội thoại Buyer
        // =====================================
        [Test, Order(2)]
        public void CHAT_SELLER_02_HienThiHoiThoai()
        {
            var buyerConversation =
                _wait.Until(
                    ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[contains(text(),'Nguyên')]")));

            Assert.That(
                buyerConversation.Displayed);

            Console.WriteLine(
                "CHAT_SELLER_02 PASS");
        }

        // =====================================
        // CHAT_SELLER_03
        // Mở cuộc trò chuyện Buyer
        // =====================================
        [Test, Order(3)]
        public void CHAT_SELLER_03_MoCuocTroChuyen()
        {
            var buyerConversation =
                _wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath("//*[contains(text(),'Nguyên')]")));

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    buyerConversation);

            var txtMessage =
                _wait.Until(
                    ExpectedConditions.ElementIsVisible(
                        By.TagName("textarea")));

            Assert.That(
                txtMessage.Displayed);

            Console.WriteLine(
                "CHAT_SELLER_03 PASS");
        }

        // =====================================
        // CHAT_SELLER_04
        // Reply tin nhắn Buyer
        // =====================================
        [Test, Order(4)]
        public void CHAT_SELLER_04_ReplyTinNhan()
        {
            _replyMessage =
                "Reply Seller " +
                DateTime.Now.ToString("yyyyMMddHHmmss");

            var txtMessage =
                _wait.Until(
                    ExpectedConditions.ElementIsVisible(
                        By.TagName("textarea")));

            txtMessage.Clear();
            txtMessage.SendKeys(_replyMessage);

            System.Threading.Thread.Sleep(1000);

            var btnSend =
                _wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath("//button[contains(.,'Gửi')]")));

            btnSend.Click();

            System.Threading.Thread.Sleep(3000);

            Assert.That(
                _driver.PageSource.Contains(_replyMessage));

            Console.WriteLine(
                "CHAT_SELLER_04 PASS");
        }

        // =====================================
        // CHAT_SELLER_05
        // Reload -> mở lại chat Buyer
        // -> kiểm tra tin nhắn còn tồn tại
        // =====================================
        [Test, Order(5)]
        public void CHAT_SELLER_05_ReloadTrang()
        {
            _driver.Navigate().Refresh();

            System.Threading.Thread.Sleep(5000);

            var buyerConversation =
                _wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath("//*[contains(text(),'Nguyên')]")));

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    buyerConversation);

            System.Threading.Thread.Sleep(3000);

            Assert.That(
                _driver.PageSource.Contains(
                    _replyMessage),
                "Không tìm thấy tin nhắn sau khi reload");

            Console.WriteLine(
                "CHAT_SELLER_05 PASS");
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