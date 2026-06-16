using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;

namespace TestProject1.API
{
    // Thu thập kết quả từng test rồi xuất báo cáo HTML để nộp đồ án.
    public static class TestReporter
    {
        public class Row
        {
            public string Group, CaseId, Scenario, Endpoint, Input, Expected, Actual, Message, Result;
        }

        private static readonly object _lock = new object();
        private static readonly List<Row> _rows = new List<Row>();

        public static void Reset() { lock (_lock) { _rows.Clear(); } }

        public static void Add(string group, string caseId, string scenario, string endpoint,
            string input, string expected, int actualStatus, string body, bool passed)
        {
            lock (_lock)
            {
                _rows.Add(new Row
                {
                    Group = group,
                    CaseId = caseId,
                    Scenario = scenario,
                    Endpoint = endpoint,
                    Input = Trunc(input, 240),
                    Expected = expected,
                    Actual = $"{actualStatus} {(HttpStatusCode)actualStatus}",
                    Message = Trunc(ServerMessage(body), 160),
                    Result = passed ? "ĐẠT" : "KHÔNG ĐẠT"
                });
            }
        }

        public static string ServerMessage(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return "";
            try
            {
                using (var doc = JsonDocument.Parse(body))
                {
                    var root = doc.RootElement;
                    foreach (var k in new[] { "message", "Message", "error", "title", "detail" })
                        if (root.TryGetProperty(k, out var el) && el.ValueKind == JsonValueKind.String)
                            return el.GetString();
                }
            }
            catch { }
            return body;
        }

        private static string Trunc(string s, int n)
            => string.IsNullOrEmpty(s) ? "" : (s.Length > n ? s.Substring(0, n) + "…" : s);

        public static string Write(string title, string fileName)
        {
            lock (_lock)
            {
                var dir = Path.Combine(AppContext.BaseDirectory, "TestReports");
                Directory.CreateDirectory(dir);
                var path = Path.Combine(dir, fileName);

                int total = _rows.Count;
                int pass = _rows.Count(r => r.Result == "ĐẠT");
                int fail = total - pass;
                double rate = total == 0 ? 0 : pass * 100.0 / total;

                var sb = new StringBuilder();
                sb.Append($@"<!DOCTYPE html><html lang='vi'><head><meta charset='utf-8'>
<title>{E(title)}</title><style>
*{{box-sizing:border-box}}
body{{font-family:Segoe UI,Arial,sans-serif;margin:16px;color:#1f2937}}
.wrap{{max-width:1200px;margin:0 auto}}
h1{{font-size:18px;margin:0 0 4px}}
.meta{{color:#6b7280;font-size:12px;margin-bottom:14px}}
.cards{{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap}}
.card{{border:1px solid #e5e7eb;border-radius:10px;padding:8px 14px;min-width:80px}}
.card span{{font-size:12px;color:#6b7280}} .card b{{display:block;font-size:20px}}
.pass{{color:#15803d}} .fail{{color:#b91c1c}}
table{{border-collapse:collapse;width:100%;font-size:11px;table-layout:fixed}}
th,td{{border:1px solid #e5e7eb;padding:5px 6px;text-align:left;vertical-align:top;
       overflow-wrap:anywhere;word-break:break-word;white-space:normal}}
th{{background:#1e293b;color:#fff;font-size:11px}}
tr:nth-child(even){{background:#f9fafb}}
.r-pass{{color:#15803d;font-weight:700}} .r-fail{{color:#b91c1c;font-weight:700}}
.grp td{{background:#eef2ff;font-weight:700;font-size:12px}}
code{{font-size:10px;background:#f1f5f9;padding:1px 3px;border-radius:3px}}
</style></head><body><div class='wrap'>");
                sb.Append($"<h1>{E(title)}</h1>");
                sb.Append($"<div class='meta'>Thời điểm chạy: {DateTime.Now:dd/MM/yyyy HH:mm:ss}</div>");
                sb.Append("<div class='cards'>");
                sb.Append($"<div class='card'><span>Tổng</span><b>{total}</b></div>");
                sb.Append($"<div class='card pass'><span>Đạt</span><b>{pass}</b></div>");
                sb.Append($"<div class='card fail'><span>Không đạt</span><b>{fail}</b></div>");
                sb.Append($"<div class='card'><span>Tỉ lệ đạt</span><b>{rate:0.0}%</b></div>");
                sb.Append("</div>");

                // table-layout: fixed -> định nghĩa bề rộng từng cột (tổng = 100%)
                sb.Append("<table><colgroup>" +
                    "<col style='width:3%'><col style='width:5%'><col style='width:14%'>" +
                    "<col style='width:11%'><col style='width:26%'><col style='width:9%'>" +
                    "<col style='width:8%'><col style='width:19%'><col style='width:5%'>" +
                    "</colgroup><thead><tr>" +
                    "<th>#</th><th>Mã TC</th><th>Kịch bản</th><th>Endpoint</th><th>Dữ liệu vào</th>" +
                    "<th>Mong đợi</th><th>HTTP thực tế</th><th>Message API</th><th>KQ</th>" +
                    "</tr></thead><tbody>");

                int i = 0; string curGroup = null;
                foreach (var r in _rows.OrderBy(x => x.Group).ThenBy(x => x.CaseId))
                {
                    if (r.Group != curGroup)
                    {
                        curGroup = r.Group;
                        sb.Append($"<tr class='grp'><td colspan='9'>{E(curGroup)}</td></tr>");
                    }
                    var cls = r.Result == "ĐẠT" ? "r-pass" : "r-fail";
                    sb.Append($"<tr><td>{++i}</td><td>{E(r.CaseId)}</td><td>{E(r.Scenario)}</td>" +
                        $"<td><code>{E(r.Endpoint)}</code></td><td>{E(r.Input)}</td>" +
                        $"<td>{E(r.Expected)}</td><td>{E(r.Actual)}</td><td>{E(r.Message)}</td>" +
                        $"<td class='{cls}'>{E(r.Result)}</td></tr>");
                }
                sb.Append("</tbody></table></div></body></html>");

                File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
                return path;
            }
        }

        private static string E(string s) => WebUtility.HtmlEncode(s ?? "");
    }
}