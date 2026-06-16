using NUnit.Framework;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TestProject1.API
{
    [TestFixture]
    public class AuthApiTests
    {
        private const string RegisterFile = "RegisterTestData.xlsx";
        private const string LoginFile = "LoginTestData.xlsx";

        private const string RegisterUrl = "api/Auth/register";
        private const string VerifyEmailUrl = "api/Auth/verify-email";
        private const string LoginUrl = "api/Auth/login";
        private const string ForgotPasswordUrl = "api/Auth/forgot-password";
        private const string ResetPasswordUrl = "api/Auth/reset-password";

        private static readonly string RunId = Guid.NewGuid().ToString("N").Substring(0, 6);

        // API chỉ validate email (định dạng + không trùng)
        private static readonly Regex EmailRx =
            new Regex(@"^[A-Za-z0-9][A-Za-z0-9._%+\-]*@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$");

        static AuthApiTests()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        // Bắt đầu fixture -> làm mới dữ liệu báo cáo
        [OneTimeSetUp]
        public void ResetReport() => TestReporter.Reset();

        // Xuất báo cáo HTML sau khi chạy xong cả fixture + tự mở trình duyệt
        [OneTimeTearDown]
        public void ExportReport()
        {
            var path = TestReporter.Write("BÁO CÁO KIỂM THỬ API – XÁC THỰC (Auth)", "AuthApiReport.html");
            TestContext.Progress.WriteLine($"\n>>> Báo cáo: {path}\n");
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch { /* không mở được trình duyệt thì bỏ qua, file vẫn đã ghi */ }
        }

        // ===================== ĐỌC EXCEL =====================
        private static List<Dictionary<string, string>> ReadExcel(string fileName)
        {
            var rows = new List<Dictionary<string, string>>();
            var path = Locate(fileName);
            if (path == null)
            {
                TestContext.WriteLine($"[ReadExcel] KHÔNG tìm thấy {fileName} -> Copy to Output Directory.");
                return rows;
            }
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                });
                if (ds.Tables.Count == 0) return rows;
                var table = ds.Tables[0];
                foreach (System.Data.DataRow dr in table.Rows)
                {
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (System.Data.DataColumn col in table.Columns)
                        dict[col.ColumnName.Trim()] = dr[col]?.ToString()?.Trim() ?? "";
                    if (dict.Values.Any(v => !string.IsNullOrEmpty(v))) rows.Add(dict);
                }
            }
            return rows;
        }

        private static string Locate(string fileName)
        {
            var direct = Path.Combine(AppContext.BaseDirectory, fileName);
            if (File.Exists(direct)) return direct;
            try
            {
                var f = Directory.GetFiles(AppContext.BaseDirectory, fileName, SearchOption.AllDirectories);
                if (f.Length > 0) return f[0];
            }
            catch { }
            return null;
        }

        // ===================== MAP CỘT =====================
        private static string V(Dictionary<string, string> r, params string[] keys)
        {
            foreach (var k in keys) if (r.TryGetValue(k, out var v)) return v;
            return "";
        }
        private static string FullName(Dictionary<string, string> r) => V(r, "Họ và tên", "FullName", "HoTen");
        private static string Phone(Dictionary<string, string> r) => V(r, "SĐT", "PhoneNumber", "Phone");
        private static string ExcelEmail(Dictionary<string, string> r) => V(r, "Địa chỉ Email", "Email", "Mail");
        private static string Pwd(Dictionary<string, string> r) => V(r, "Mật khẩu", "Password", "MatKhau");
        private static string Scenario(Dictionary<string, string> r) => V(r, "Scenario_Name");
        private static string Action(Dictionary<string, string> r) => V(r, "Action");
        private static string Id(Dictionary<string, string> r) => V(r, "Step_ID", "STT");
        private static string DynEmail(Dictionary<string, string> r) =>
            $"auto_{RunId}_{Id(r)}@test.com".ToLowerInvariant();

        private static bool IsDuplicateCase(Dictionary<string, string> r) =>
            Scenario(r).IndexOf("tồn tại", StringComparison.OrdinalIgnoreCase) >= 0;
        private static bool RegisterShouldSucceed(Dictionary<string, string> r) =>
            EmailRx.IsMatch(ExcelEmail(r)) && !IsDuplicateCase(r);
        private static bool LoginShouldSucceed(Dictionary<string, string> r)
        {
            var a = Action(r);
            return a.IndexOf("Success", StringComparison.OrdinalIgnoreCase) >= 0
                || a.IndexOf("Toggle", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // Tên test: "L05: Để trống Họ tên"
        private static string Label(Dictionary<string, string> r) => $"{Id(r)}: {Scenario(r)}";

        // ===================== HTTP + GHI NHẬN =====================
        private static async Task<(int status, string body)> PostJsonAsync(string url, object payload)
        {
            var json = JsonSerializer.Serialize(payload);
            using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
            {
                var resp = await ApiHelper.Client.PostAsync(url, content);
                var body = await resp.Content.ReadAsStringAsync();
                return ((int)resp.StatusCode, body);
            }
        }

        // Ghi nhận vào báo cáo RỒI mới assert (để case fail vẫn hiện trong báo cáo)
        private static void Eval(string group, Dictionary<string, string> data, string endpoint,
            string input, bool expectSuccess, int status, string body)
        {
            bool ok = expectSuccess ? (status >= 200 && status < 300) : (status >= 400);
            string expected = expectSuccess ? "Thành công (HTTP 2xx)" : "API từ chối (HTTP 4xx)";
            TestReporter.Add(group, Id(data), Scenario(data), endpoint, input, expected, status, body, ok);

            if (expectSuccess)
                Assert.That(status, Is.InRange(200, 299), $"Mong 2xx. Body: {body}");
            else
                Assert.That(status, Is.GreaterThanOrEqualTo(400), $"Mong 4xx. Body: {body}");
        }

        private static string ResolveCode(string email)
        {
            var code = DbHelper.GetEmailVerificationToken(email);
            TestContext.WriteLine($"[DB] code của {email} = {code}");
            return code;
        }

        private static string ExtractToken(string body)
        {
            try
            {
                using (var doc = JsonDocument.Parse(body))
                {
                    var root = doc.RootElement;
                    foreach (var k in new[] { "token", "accessToken", "access_token", "jwt" })
                        if (root.TryGetProperty(k, out var el)) return el.GetString();
                    if (root.TryGetProperty("data", out var d))
                        foreach (var k in new[] { "token", "accessToken", "access_token" })
                            if (d.TryGetProperty(k, out var el)) return el.GetString();
                }
            }
            catch { }
            return "";
        }

        // ===================== CASE SOURCES =====================
        public static IEnumerable<TestCaseData> RegisterCases()
        {
            foreach (var r in ReadExcel(RegisterFile))
                yield return new TestCaseData(r).SetName($"Register_{Label(r)}");
        }
        private static IEnumerable<TestCaseData> SuccessRows(string prefix)
        {
            foreach (var r in ReadExcel(RegisterFile).Where(RegisterShouldSucceed))
                yield return new TestCaseData(r).SetName($"{prefix}_{Label(r)}");
        }
        public static IEnumerable<TestCaseData> VerifyEmailCases() => SuccessRows("VerifyEmail");
        public static IEnumerable<TestCaseData> ForgotPasswordCases() => SuccessRows("ForgotPassword");
        public static IEnumerable<TestCaseData> ResetPasswordCases() => SuccessRows("ResetPassword");
        public static IEnumerable<TestCaseData> LoginCases()
        {
            foreach (var r in ReadExcel(LoginFile))
                yield return new TestCaseData(r).SetName($"Login_{Label(r)}");
        }

        // ===================== 1) REGISTER =====================
        [Test, Order(1), TestCaseSource(nameof(RegisterCases))]
        public async Task Register_Test(Dictionary<string, string> data)
        {
            ApiHelper.ClearToken();
            bool success = RegisterShouldSucceed(data);
            string email = success ? DynEmail(data) : ExcelEmail(data);
            var payload = new { fullName = FullName(data), email, phoneNumber = Phone(data), password = Pwd(data) };

            var (status, body) = await PostJsonAsync(RegisterUrl, payload);
            string input = $"Tên='{FullName(data)}', SĐT='{Phone(data)}', Email='{email}', MK='{Pwd(data)}'";
            TestContext.WriteLine($"[Register] {Label(data)} status={status} body={body}");
            Eval("1. Đăng ký (Register)", data, RegisterUrl, input, success, status, body);
        }

        // ===================== 2) VERIFY EMAIL =====================
        [Test, Order(2), TestCaseSource(nameof(VerifyEmailCases))]
        public async Task VerifyEmail_Test(Dictionary<string, string> data)
        {
            var email = DynEmail(data);
            var code = ResolveCode(email);
            var (status, body) = await PostJsonAsync(VerifyEmailUrl, new { email, code });
            TestContext.WriteLine($"[VerifyEmail] {email} status={status} body={body}");
            Eval("2. Xác thực email (Verify)", data, VerifyEmailUrl, $"Email='{email}', Code='{code}'", true, status, body);
        }

        // ===================== 3) LOGIN =====================
        [Test, Order(3), TestCaseSource(nameof(LoginCases))]
        public async Task Login_Test(Dictionary<string, string> data)
        {
            ApiHelper.ClearToken();
            bool success = LoginShouldSucceed(data);
            var (status, body) = await PostJsonAsync(LoginUrl, new { email = ExcelEmail(data), password = Pwd(data) });
            string input = $"Email='{ExcelEmail(data)}', MK='{Pwd(data)}'";
            TestContext.WriteLine($"[Login] {Label(data)} status={status} body={body}");
            Eval("3. Đăng nhập (Login)", data, LoginUrl, input, success, status, body);

            if (success && status >= 200 && status < 300)
            {
                var token = ExtractToken(body);
                if (!string.IsNullOrEmpty(token)) { ApiHelper.SetToken(token); ApiHelper.SaveToken(token); }
                else TestContext.WriteLine("[Login] ⚠️ Không thấy token - gửi mình body để sửa ExtractToken().");
            }
        }

        // ===================== 4) FORGOT PASSWORD =====================
        [Test, Order(4), TestCaseSource(nameof(ForgotPasswordCases))]
        public async Task ForgotPassword_Test(Dictionary<string, string> data)
        {
            var email = DynEmail(data);
            var (status, body) = await PostJsonAsync(ForgotPasswordUrl, new { email });
            TestContext.WriteLine($"[ForgotPassword] {email} status={status} body={body}");
            Eval("4. Quên mật khẩu (Forgot)", data, ForgotPasswordUrl, $"Email='{email}'", true, status, body);
        }

        // ===================== 5) RESET PASSWORD =====================
        [Test, Order(5), TestCaseSource(nameof(ResetPasswordCases))]
        public async Task ResetPassword_Test(Dictionary<string, string> data)
        {
            var email = DynEmail(data);
            var code = ResolveCode(email);
            var (status, body) = await PostJsonAsync(ResetPasswordUrl, new { email, code, newPassword = "NewPass@123" });
            TestContext.WriteLine($"[ResetPassword] {email} status={status} body={body}");
            Eval("5. Đặt lại mật khẩu (Reset)", data, ResetPasswordUrl,
                $"Email='{email}', Code='{code}', MK mới='NewPass@123'", true, status, body);
        }
    }
}