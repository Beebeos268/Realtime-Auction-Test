using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Globalization;

namespace TestProject1
{
    [TestFixture]
    public class TestDatGia : IDisposable
    {
        private ChromeDriver _driver;
        private WebDriverWait _wait;

        private readonly string _baseUrl = "http://localhost:4200";

        // account test
        private readonly string _email = "ntpmm268255@gmail.com";
        private readonly string _password = "Lam255@";

        [OneTimeSetUp]
        public void SetupOnce()
        {
            _driver = new ChromeDriver();

            _driver.Manage().Window.Maximize();

            _wait = new WebDriverWait(
                _driver,
                TimeSpan.FromSeconds(20)
            );

            Login();

            OpenAuctionDetail();
        }

        // ======================================================
        // LOGIN
        // ======================================================
        private void Login()
        {
            _driver.Navigate().GoToUrl(_baseUrl + "/login");

            WaitForPageLoad();

            var txtEmail = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.Name("email")
                )
            );

            txtEmail.Clear();
            txtEmail.SendKeys(_email);

            var txtPassword = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.Name("password")
                )
            );

            txtPassword.Clear();
            txtPassword.SendKeys(_password);

            var btnLogin = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(.,'Đăng Nhập')]")
                )
            );

            btnLogin.Click();

            _wait.Until(d =>
                !d.Url.Contains("/login"));

            WaitForPageLoad();

            Console.WriteLine("LOGIN SUCCESS");
        }

        // ======================================================
        // OPEN AUCTION DETAIL
        // ======================================================
        private void OpenAuctionDetail()
        {
            _driver.Navigate()
                .GoToUrl(_baseUrl + "/auctions");

            WaitForPageLoad();

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

            Console.WriteLine("OPEN DETAIL SUCCESS");
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
        // Check input nhập giá
        // ======================================================
        [Test, Order(1)]
        public void BID_01_CheckBidInput()
        {
            var inputBid = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//input[@type='number']")
                )
            );

            Assert.That(
                inputBid.Displayed,
                "Input nhập giá không hiển thị"
            );

            Console.WriteLine("BID_01 - OK");
        }

        // ======================================================
        // BID_02
        // Check button chốt giá
        // ======================================================
        [Test, Order(2)]
        public void BID_02_CheckBidButton()
        {
            var btnBid = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//button[contains(.,'CHỐT GIÁ')]")
                )
            );

            Assert.That(
                btnBid.Displayed,
                "Không thấy nút CHỐT GIÁ"
            );

            Assert.That(
                btnBid.Enabled,
                "Nút CHỐT GIÁ bị disabled"
            );

            Console.WriteLine("BID_02 - OK");
        }

        // ======================================================
        // BID_03
        // Check text giá tối thiểu
        // ======================================================
        [Test, Order(3)]
        public void BID_03_CheckMinimumBidText()
        {
            var minBidText = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'Tối thiểu')]")
                )
            );

            Assert.That(
                minBidText.Text.Contains("Tối thiểu"),
                "Không hiển thị giá tối thiểu"
            );

            Console.WriteLine("BID_03 - OK");
        }

        // ======================================================
        // BID_04
        // Nhập giá hợp lệ
        // ======================================================
        [Test, Order(4)]
        public void BID_04_InputValidBid()
        {
            var inputBid = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//input[@type='number']")
                )
            );

            // lấy placeholder minBidRequired
            string placeholder =
                inputBid.GetAttribute("placeholder");

            decimal minBid =
                decimal.Parse(
                    placeholder,
                    CultureInfo.InvariantCulture
                );

            decimal validBid = minBid + 100000;

            inputBid.Clear();

            inputBid.SendKeys(
                validBid.ToString()
            );

            Assert.That(
                inputBid.GetAttribute("value")
                    == validBid.ToString(),
                "Không nhập được giá"
            );

            Console.WriteLine("BID_04 - OK");
        }

        // ======================================================
        // BID_05
        // Click đặt giá thành công
        // ======================================================

        [Test, Order(5)]
        public void BID_05_ClickBidButton()
                {
            var btnBid = _wait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[contains(.,'CHỐT GIÁ')]")
                )
            );

            ((IJavaScriptExecutor)_driver)
                .ExecuteScript(
                    "arguments[0].click();",
                    btnBid
                );

            // chờ toastr success
            var successToast = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.CssSelector(".toast-success")
                )
            );

            Assert.That(
                successToast.Displayed,
                "Không hiển thị toast success"
            );

            Console.WriteLine("BID_05 - OK");
        }


        // ======================================================
        // BID_06
        // Check button đổi trạng thái
        // ======================================================
        [Test, Order(6)]
        public void BID_06_CheckButtonStateAfterBid()
        {
            var bidButton = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//button[contains(.,'trả giá')]")
                )
            );

            Assert.That(
                bidButton.Text.Contains("trả giá"),
                "Button không đổi trạng thái"
            );

            Console.WriteLine("BID_06 - OK");
        }


        // ======================================================
        // BID_07
        // Check history update
        // ======================================================
        [Test, Order(7)]
        public void BID_07_CheckBidHistory()
        {
            var historyItem = _wait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'Đang dẫn đầu')]")
                )
            );

            Assert.That(
                historyItem.Displayed,
                "History không update"
            );

            Console.WriteLine("BID_07 - OK");
        }
    }
}
