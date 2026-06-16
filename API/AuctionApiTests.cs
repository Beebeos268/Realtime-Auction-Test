using NUnit.Framework;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace TestProject1.API
{
    // Controller Auctions đầy đủ: tạo phiên + danh sách + search + suggest
    // + chi tiết + lịch sử + ĐẶT GIÁ (dat-gia, data từ "đặt giá.xlsx" sheet "Test Data").
    [TestFixture]
    public class AuctionApiTests
    {
        private const string CreateFile = "CreateAuctionTestData.xlsx";
        // file bid có thể tên có dấu cách hoặc gạch dưới -> thử nhiều tên
        private static readonly string[] BidDataFiles = { "đặt giá.xlsx", "đặt_giá.xlsx", "dat gia.xlsx", "dat_gia.xlsx" };
        private const string BidDataSheet = "Test Data";

        private const string CreateUrl = "api/Auctions/tao-moi";
        private const string ListUrl = "api/Auctions/danh-sach";
        private const string SearchUrl = "api/Auctions/search";
        private const string SuggestUrl = "api/Auctions/suggest";
        private const string HistoryUrl = "api/Auctions/lich-su";
        private const string LoginUrl = "api/Auth/login";
        private static string DetailUrl(string id) => $"api/Auctions/{id}";
        private static string BidUrl(string id) => $"api/Auctions/{id}/dat-gia";

        // ⚠️ Tài khoản TẠO phiên (seller). Sửa cho khớp DB.
        private const string SellerEmail = "ntpnguyen210104@gmail.com";
        private const string SellerPassword = "Nguyen@21";
        // Bidder mặc định (nếu Excel chưa điền BUYER_EMAIL/PASSWORD)
        private const string BidderEmailDefault = "test1_0523183347@gmail.com";
        private const string BidderPassDefault = "CHANGE_ME";

        private static readonly string RunId = Guid.NewGuid().ToString("N").Substring(0, 6);

        private static string _sellerToken, _bidderToken;
        private static string _createdId, _createdName;
        private static string _bidArenaId;
        private static decimal _bidStart = 1_800_000m, _bidStep = 100_000m;

        static AuctionApiTests()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        // ===================== SETUP =====================
        [OneTimeSetUp]
        public async Task Setup()
        {
            TestReporter.Reset();

            // đọc data đặt giá từ Excel (sheet Test Data: Variable | Value)
            var m = ReadDataMap(BidDataFiles, BidDataSheet);
            _bidStart = ParseNum(Get(m, "CURRENT_HIGHEST", "1800000"), 1_800_000m);
            _bidStep = ParseNum(Get(m, "BID_STEP", "100000"), 100_000m);

            // đăng nhập seller -> tạo phiên đang mở để đặt giá
            var (sls, slb) = await PostJsonAsync(LoginUrl, new { email = SellerEmail, password = SellerPassword });
            _sellerToken = Field(slb, "token", "accessToken", "access_token", "jwt");
            TestContext.Progress.WriteLine($"[setup] login seller status={sls} token={(string.IsNullOrEmpty(_sellerToken) ? "<FAIL>" : "<OK>")}");
            if (!string.IsNullOrEmpty(_sellerToken)) ApiHelper.SetToken(_sellerToken);

            var arena = new
            {
                name = $"BidArena_{RunId}",
                description = "Phiên dùng để test đặt giá",
                imageUrl = "https://example.com/a.jpg",
                startPrice = _bidStart,
                stepPrice = _bidStep,
                startTime = DateTime.UtcNow.AddSeconds(-30),
                endTime = DateTime.UtcNow.AddDays(7)
            };
            var (cs, cb) = await PostJsonAsync(CreateUrl, arena);
            _bidArenaId = Field(cb, "id", "auctionId", "Id");
            TestContext.Progress.WriteLine($"[setup] tạo phiên bid status={cs} start={_bidStart} step={_bidStep} id={_bidArenaId}");

            // Tạo MỚI 1 bidder (đăng ký + xác thực, code lấy từ DB) -> token chắc chắn hợp lệ,
            // khác seller nên không vướng luật "không tự đấu giá phiên của mình".
            var bidderEmail = $"bidder_{RunId}@test.com";
            var bidderPass = "Bidder@123";
            await PostJsonAsync("api/Auth/register", new { fullName = "Bidder Test", email = bidderEmail, phoneNumber = "0987000000", password = bidderPass });
            var vcode = DbHelper.GetEmailVerificationToken(bidderEmail);
            await PostJsonAsync("api/Auth/verify-email", new { email = bidderEmail, code = vcode });
            var (bls, blb) = await PostJsonAsync(LoginUrl, new { email = bidderEmail, password = bidderPass });
            _bidderToken = Field(blb, "token", "accessToken", "access_token", "jwt");
            TestContext.Progress.WriteLine($"[setup] bidder mới {bidderEmail} (code={vcode}) login status={bls} token={(string.IsNullOrEmpty(_bidderToken) ? "<FAIL>" : "<OK>")}");
        }

        [OneTimeTearDown]
        public void ExportReport()
        {
            var path = TestReporter.Write("BÁO CÁO KIỂM THỬ API – PHIÊN ĐẤU GIÁ (Auctions + Bid)", "AuctionApiReport.html");
            TestContext.Progress.WriteLine($"\n>>> Báo cáo: {path}\n");
            try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); } catch { }
        }

        // ===================== HTTP =====================
        private static async Task<(int status, string body)> PostJsonAsync(string url, object payload)
        {
            var json = JsonSerializer.Serialize(payload);
            using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
            {
                var resp = await ApiHelper.Client.PostAsync(url, content);
                return ((int)resp.StatusCode, await resp.Content.ReadAsStringAsync());
            }
        }
        private static async Task<(int status, string body)> GetAsync(string url)
        {
            var resp = await ApiHelper.Client.GetAsync(url);
            return ((int)resp.StatusCode, await resp.Content.ReadAsStringAsync());
        }

        private static void Eval(string group, string caseId, string scenario, string endpoint,
            string input, bool expectSuccess, int status, string body)
        {
            bool ok = expectSuccess ? (status >= 200 && status < 300) : (status >= 400);
            string expected = expectSuccess ? "Thành công (HTTP 2xx)" : "API từ chối (HTTP 4xx)";
            TestReporter.Add(group, caseId, scenario, endpoint, input, expected, status, body, ok);
            if (expectSuccess) Assert.That(status, Is.InRange(200, 299), $"Mong 2xx. Body: {body}");
            else Assert.That(status, Is.GreaterThanOrEqualTo(400), $"Mong 4xx. Body: {body}");
        }

        private static string Field(string body, params string[] keys)
        {
            try
            {
                using (var doc = JsonDocument.Parse(body))
                {
                    var root = doc.RootElement;
                    foreach (var k in keys)
                        if (root.TryGetProperty(k, out var el))
                            return el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString();
                    if (root.TryGetProperty("data", out var d) && d.ValueKind == JsonValueKind.Object)
                        foreach (var k in keys)
                            if (d.TryGetProperty(k, out var el))
                                return el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString();
                }
            }
            catch { }
            return "";
        }
        private static string FirstIdFromList(string body)
        {
            try
            {
                using (var doc = JsonDocument.Parse(body))
                {
                    JsonElement arr = doc.RootElement;
                    if (arr.ValueKind == JsonValueKind.Object)
                        foreach (var k in new[] { "data", "items", "result", "auctions" })
                            if (arr.TryGetProperty(k, out var inner) && inner.ValueKind == JsonValueKind.Array) { arr = inner; break; }
                    if (arr.ValueKind == JsonValueKind.Array && arr.GetArrayLength() > 0)
                        foreach (var k in new[] { "id", "auctionId", "Id" })
                            if (arr[0].TryGetProperty(k, out var el))
                                return el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString();
                }
            }
            catch { }
            return "";
        }

        // ===================== ĐỌC EXCEL =====================
        private static List<Dictionary<string, string>> ReadExcel(string fileName)
        {
            var rows = new List<Dictionary<string, string>>();
            var path = Locate(fileName);
            if (path == null) { TestContext.WriteLine($"[ReadExcel] KHÔNG tìm thấy {fileName}."); return rows; }
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet(new ExcelDataSetConfiguration
                { ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true } });
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

        // đọc sheet dạng key-value (cột 0 = tên biến, cột 1 = giá trị)
        private static Dictionary<string, string> ReadDataMap(string[] fileNames, string sheetName)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var path = LocateAny(fileNames);
            if (path == null) { TestContext.WriteLine($"[ReadDataMap] KHÔNG tìm thấy file bid trong {string.Join(", ", fileNames)}."); return map; }
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var ds = reader.AsDataSet(); // không dùng header -> đọc thô từng ô
                if (!ds.Tables.Contains(sheetName)) { TestContext.WriteLine($"[ReadDataMap] Không có sheet '{sheetName}'."); return map; }
                var t = ds.Tables[sheetName];
                foreach (System.Data.DataRow dr in t.Rows)
                {
                    var key = dr[0]?.ToString()?.Trim() ?? "";
                    var val = t.Columns.Count > 1 ? (dr[1]?.ToString()?.Trim() ?? "") : "";
                    if (!string.IsNullOrEmpty(key) && !map.ContainsKey(key)) map[key] = val;
                }
            }
            return map;
        }

        private static string Locate(string fileName)
        {
            var direct = Path.Combine(AppContext.BaseDirectory, fileName);
            if (File.Exists(direct)) return direct;
            try { var f = Directory.GetFiles(AppContext.BaseDirectory, fileName, SearchOption.AllDirectories); if (f.Length > 0) return f[0]; }
            catch { }
            return null;
        }
        private static string LocateAny(string[] names)
        { foreach (var n in names) { var p = Locate(n); if (p != null) return p; } return null; }

        private static string Get(Dictionary<string, string> m, string key, string def)
            => m.TryGetValue(key, out var v) && !string.IsNullOrEmpty(v) ? v : def;

        // ===================== MAP CỘT (create) =====================
        private static string V(Dictionary<string, string> r, params string[] keys)
        { foreach (var k in keys) if (r.TryGetValue(k, out var v)) return v; return ""; }
        private static string CName(Dictionary<string, string> r) => V(r, "Tên sản phẩm", "Name");
        private static string CImg(Dictionary<string, string> r) => V(r, "Ảnh sản phẩm", "ImageUrl");
        private static string CDesc(Dictionary<string, string> r) => V(r, "Mô tả chi tiết", "Description");
        private static string CStart(Dictionary<string, string> r) => V(r, "Giá khởi điểm", "StartPrice");
        private static string CStep(Dictionary<string, string> r) => V(r, "Bước giá", "StepPrice");
        private static string CType(Dictionary<string, string> r) => V(r, "StartTime_Type");
        private static string CT1(Dictionary<string, string> r) => V(r, "Thời gian bắt đầu", "StartTime");
        private static string CT2(Dictionary<string, string> r) => V(r, "Thời gian kết thúc", "EndTime");
        private static string Scenario(Dictionary<string, string> r) => V(r, "Scenario_Name");
        private static string Action(Dictionary<string, string> r) => V(r, "Action");
        private static string Id(Dictionary<string, string> r) => V(r, "Step_ID", "STT");
        private static string Label(Dictionary<string, string> r) => $"{Id(r)}: {Scenario(r)}";

        private static bool IsSuccess(Dictionary<string, string> r) =>
            Action(r).IndexOf("Success", StringComparison.OrdinalIgnoreCase) >= 0;
        private static bool IsSkip(Dictionary<string, string> r) =>
            Action(r).IndexOf("Skip", StringComparison.OrdinalIgnoreCase) >= 0;
        private static bool IsLengthCase(Dictionary<string, string> r)
        {
            var sc = Scenario(r);
            return sc.IndexOf("quá dài", StringComparison.OrdinalIgnoreCase) >= 0
                || sc.IndexOf("vượt quá", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static decimal ParseNum(string s, decimal def)
        {
            if (string.IsNullOrWhiteSpace(s)) return def;
            var clean = new string(s.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());
            return decimal.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : def;
        }
        private static decimal? ParseNumOrNull(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            var clean = new string(s.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());
            return decimal.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : (decimal?)null;
        }
        private static DateTime? ParseVnTime(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Replace("//", "/").Trim();
            string[] fmts = { "dd/MM/yyyy HH:mm", "d/M/yyyy H:mm", "dd/MM/yyyy", "d/M/yyyy" };
            if (DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                return DateTime.SpecifyKind(dt, DateTimeKind.Utc);
            return null;
        }

        public static IEnumerable<TestCaseData> CreateCases()
        {
            foreach (var r in ReadExcel(CreateFile))
                yield return new TestCaseData(r).SetName($"CreateAuction_{Label(r)}");
        }

        // ===================== 1) TẠO PHIÊN =====================
        [Test, Order(1), TestCaseSource(nameof(CreateCases))]
        public async Task Create_Auction_Test(Dictionary<string, string> data)
        {
            if (IsSkip(data))
                Assert.Ignore("FE chặn nhập (vd nhập chữ vào ô số) - không kiểm được ở tầng API.");
            ApiHelper.SetToken(_sellerToken);

            bool success = IsSuccess(data) || IsLengthCase(data);
            bool isNow = CType(data).Equals("NOW", StringComparison.OrdinalIgnoreCase);
            string name = success
                ? $"{(string.IsNullOrEmpty(CName(data)) ? "Phien" : CName(data))}_{RunId}_{Id(data)}"
                : CName(data);

            DateTime? start, end;
            if (success) { start = isNow ? DateTime.UtcNow.AddMinutes(1) : DateTime.UtcNow.AddHours(2); end = DateTime.UtcNow.AddDays(7); }
            else { start = isNow ? DateTime.UtcNow : ParseVnTime(CT1(data)); end = ParseVnTime(CT2(data)); }

            var payload = new
            {
                name,
                description = CDesc(data),
                imageUrl = CImg(data),
                startPrice = ParseNumOrNull(CStart(data)),
                stepPrice = ParseNumOrNull(CStep(data)),
                startTime = start,
                endTime = end
            };
            var (status, body) = await PostJsonAsync(CreateUrl, payload);
            TestContext.WriteLine($"[Create] {Label(data)} status={status} body={body}");

            if (success && status >= 200 && status < 300 && string.IsNullOrEmpty(_createdId))
            { _createdId = Field(body, "id", "auctionId", "Id"); _createdName = name; }

            string input = $"name='{name}', startPrice={payload.startPrice}, stepPrice={payload.stepPrice}, " +
                           $"start={(start?.ToString("u") ?? "null")}, end={(end?.ToString("u") ?? "null")}";
            Eval("1. Tạo phiên (tao-moi)", Id(data), Scenario(data), CreateUrl, input, success, status, body);
        }

        // ===================== 2) DANH SÁCH =====================
        [Test, Order(2)]
        public async Task GetList_Test()
        {
            var (status, body) = await GetAsync(ListUrl);
            if (string.IsNullOrEmpty(_createdId)) _createdId = FirstIdFromList(body);
            Eval("2. Danh sách (danh-sach)", "AUC-LIST", "Lấy danh sách phiên đấu giá", ListUrl, "(không tham số)", true, status, body);
        }

        // ===================== 3) SEARCH =====================
        [Test, Order(3)]
        public async Task Search_Test()
        {
            var keyword = string.IsNullOrEmpty(_createdName) ? "Rolex" : _createdName;
            var url = $"{SearchUrl}?keyword={Uri.EscapeDataString(keyword)}";
            var (status, body) = await GetAsync(url);
            Eval("3. Tìm kiếm (search)", "AUC-SEARCH", $"Tìm phiên theo keyword='{keyword}'", url, $"keyword={keyword}", true, status, body);
        }

        // ===================== 4) SUGGEST =====================
        [Test, Order(4)]
        public async Task Suggest_Test()
        {
            var keyword = string.IsNullOrEmpty(_createdName) ? "Rol" : _createdName.Substring(0, Math.Min(4, _createdName.Length));
            var url = $"{SuggestUrl}?keyword={Uri.EscapeDataString(keyword)}";
            var (status, body) = await GetAsync(url);
            Eval("4. Gợi ý (suggest)", "AUC-SUGGEST", $"Gợi ý theo keyword='{keyword}'", url, $"keyword={keyword}", true, status, body);
        }

        // ===================== 5) CHI TIẾT =====================
        [Test, Order(5)]
        public async Task GetDetail_Valid_Test()
        {
            var id = string.IsNullOrEmpty(_createdId) ? _bidArenaId : _createdId;
            if (string.IsNullOrEmpty(id)) Assert.Ignore("Chưa có id phiên hợp lệ.");
            var url = DetailUrl(id);
            var (status, body) = await GetAsync(url);
            Eval("5. Chi tiết ({id})", "AUC-DETAIL-OK", "Xem chi tiết với id hợp lệ", url, $"id={id}", true, status, body);
        }

        [Test, Order(6)]
        public async Task GetDetail_Invalid_Test()
        {
            var fakeId = Guid.NewGuid().ToString();
            var url = DetailUrl(fakeId);
            var (status, body) = await GetAsync(url);
            Eval("5. Chi tiết ({id})", "AUC-DETAIL-NF", "Xem chi tiết với id không tồn tại", url, $"id={fakeId}", false, status, body);
        }

        // ===================== 6) LỊCH SỬ =====================
        [Test, Order(7)]
        public async Task History_Test()
        {
            var (status, body) = await GetAsync(HistoryUrl);
            Eval("6. Lịch sử (lich-su)", "AUC-HISTORY", "Lấy lịch sử đấu giá", HistoryUrl, "(không tham số)", true, status, body);
        }

        // ===================== 7) ĐẶT GIÁ (data từ "Test Data") =====================
        // (caseId, scenario, biến trong Excel, mặc định, kỳ vọng thành công)
        public static IEnumerable<TestCaseData> BidCases()
        {
            var m = ReadDataMap(BidDataFiles, BidDataSheet);
            string G(string k, string d) => m.TryGetValue(k, out var v) && !string.IsNullOrEmpty(v) ? v : d;
            yield return new TestCaseData("BID-V1", "Đặt giá hợp lệ #1", G("VALID_BID_1", "2000000"), true).SetName("Bid_VALID_BID_1");
            yield return new TestCaseData("BID-V2", "Đặt giá hợp lệ #2 (cao hơn)", G("VALID_BID_2", "2100000"), true).SetName("Bid_VALID_BID_2");
            yield return new TestCaseData("BID-LOW", "Đặt giá thấp hơn tối thiểu", G("INVALID_BID_LOW", "1800000"), false).SetName("Bid_INVALID_LOW");
            yield return new TestCaseData("BID-ZERO", "Đặt giá bằng 0", G("INVALID_BID_ZERO", "0"), false).SetName("Bid_INVALID_ZERO");
            yield return new TestCaseData("BID-NEG", "Đặt giá số âm", G("INVALID_BID_NEG", "-100000"), false).SetName("Bid_INVALID_NEG");
            yield return new TestCaseData("BID-STR", "Đặt giá không phải số", G("INVALID_BID_STR", "abc123"), false).SetName("Bid_INVALID_STR");
        }

        [Test, Order(8), TestCaseSource(nameof(BidCases))]
        public async Task Bid_Test(string caseId, string scenario, string amountStr, bool expectSuccess)
        {
            if (string.IsNullOrEmpty(_bidArenaId))
                Assert.Ignore("Chưa tạo được phiên để đặt giá (xem log [setup]).");
            if (!decimal.TryParse(amountStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var amount))
                Assert.Ignore($"'{amountStr}' không phải số - API kiểu số không nhận (giống FE chặn).");

            ApiHelper.SetToken(_bidderToken);
            var (status, body) = await PostJsonAsync(BidUrl(_bidArenaId), new { bidAmount = amount });
            TestContext.WriteLine($"[Bid] {caseId} amount={amount} status={status} body={body}");
            Eval("7. Đặt giá (dat-gia)", caseId, scenario, BidUrl("{id}"), $"bidAmount={amount}", expectSuccess, status, body);
        }

        // ===================== 8) ĐẶT GIÁ - phiên không tồn tại / chưa đăng nhập =====================
        [Test, Order(9)]
        public async Task Bid_FakeAuction_Test()
        {
            ApiHelper.SetToken(_bidderToken);
            var fakeId = Guid.NewGuid().ToString();
            var (status, body) = await PostJsonAsync(BidUrl(fakeId), new { bidAmount = _bidStart + _bidStep });
            TestContext.WriteLine($"[Bid fake] id={fakeId} status={status} body={body}");
            Eval("7. Đặt giá (dat-gia)", "BID-NF", "Đặt giá trên phiên không tồn tại -> 404", BidUrl("{id}"), $"id={fakeId}", false, status, body);
        }

        [Test, Order(10)]
        public async Task Bid_NoAuth_Test()
        {
            if (string.IsNullOrEmpty(_bidArenaId)) Assert.Ignore("Chưa có phiên để đặt giá.");
            ApiHelper.ClearToken();
            var (status, body) = await PostJsonAsync(BidUrl(_bidArenaId), new { bidAmount = _bidStart + _bidStep });
            TestContext.WriteLine($"[Bid noauth] status={status} body={body}");
            Eval("7. Đặt giá (dat-gia)", "BID-401", "Đặt giá khi chưa đăng nhập -> 401", BidUrl("{id}"), "(không token)", false, status, body);
            if (!string.IsNullOrEmpty(_bidderToken)) ApiHelper.SetToken(_bidderToken);
        }
    }
}