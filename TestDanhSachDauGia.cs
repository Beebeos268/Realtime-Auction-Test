#nullable disable
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Threading;

namespace TestProject1
{
    [TestFixture]
    public class TestDanhSachDauGia : IDisposable
    {
        private ChromeDriver _driver;
        private readonly string _baseUrl = "http://localhost:4200";
        private WebDriverWait _wait;

        [OneTimeSetUp]
        public void SetupOnce()
        {
            _driver = new ChromeDriver();
            _driver.Manage().Window.Maximize();
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            _driver.Navigate().GoToUrl(_baseUrl);
            Thread.Sleep(2000);
        }

        [OneTimeTearDown]
        public void TeardownOnce()
        {
            _driver?.Dispose();
            _driver = null;
        }

        public void Dispose() => _driver?.Dispose();

        // LIST_01: Click nút Menu → click link "Danh sách đấu giá"
        //          → assert "Đang diễn ra" hiển thị
        [Test, Order(1)]
        public void LIST_01_Navigation()
        {
            var btnMenu = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//button[.//span[text()='Menu']]")));
            btnMenu.Click();

            var linkDanhSach = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[@href='/auctions']")));
            linkDanhSach.Click();

            Assert.That(
                _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'Đang diễn ra')]"))).Displayed,
                "LIST_01: Trang danh sách không hiển thị 'Đang diễn ra'");

            Console.WriteLine("LIST_01 - OK");
        }

        // LIST_02: Assert link "Tham gia đấu giá" hiển thị
        //          → click vào để vào trang Chi tiết
        [Test, Order(2)]
        public void LIST_02_ThamGiaDauGia()
        {
            // HTML thực tế: <a href="/auction/...">Tham gia đấu giá</a>
            var linkThamGia = _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("(//a[normalize-space(text())='Tham gia đấu giá'])[1]")));

            Assert.That(linkThamGia.Displayed,
                "LIST_02: Link 'Tham gia đấu giá' không hiển thị");

            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", linkThamGia);
            Thread.Sleep(1500);

            Console.WriteLine("LIST_02 - OK");
        }

        // LIST_03: Tại trang Chi tiết
        //          → assert nút "CHỐT GIÁ NGAY" hiển thị và enabled
        [Test, Order(3)]
        public void LIST_03_CheckUI_ChiTiet()
        {
            // HTML thực tế: <button class="... bg-amber-600 ...">🔨 CHỐT GIÁ NGAY</button>
            var btnChotGia = _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//button[contains(normalize-space(text()),'CHỐT GIÁ NGAY')]")));

            Assert.That(btnChotGia.Displayed,
                "LIST_03: Nút 'CHỐT GIÁ NGAY' không hiển thị");
            Assert.That(btnChotGia.Enabled,
                "LIST_03: Nút 'CHỐT GIÁ NGAY' bị disabled");

            Console.WriteLine("LIST_03 - OK");
        }

        // LIST_04: Click link "Trang chủ" trong breadcrumb → assert URL về trang chủ
        [Test, Order(4)]
        public void LIST_04_ClickHome()
        {
            // HTML thực tế: <a routerlink="/" href="/">Trang chủ</a> trong breadcrumb
            var linkTrangChu = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[@href='/' and normalize-space(text())='Trang chủ']")));
            linkTrangChu.Click();

            Assert.That(
                _wait.Until(d => d.Url == _baseUrl + "/" || d.Url.EndsWith('/')),
                "LIST_04: URL không trỏ về trang chủ");

            Console.WriteLine("LIST_04 - OK");
        }
    }
}