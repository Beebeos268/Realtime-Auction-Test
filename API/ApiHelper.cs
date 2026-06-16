using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;

namespace TestProject1.API
{
    public static class ApiHelper
    {
        private static readonly HttpClient _client;

        static ApiHelper()
        {
            var handler = new HttpClientHandler
            {
                // Chỉ dùng cho test local với self-signed cert
                ServerCertificateCustomValidationCallback =
                    (msg, cert, chain, errors) => true
            };
            _client = new HttpClient(handler);
            _client.BaseAddress = new Uri("https://localhost:7018/");
            _client.Timeout = TimeSpan.FromSeconds(30);
        }

        public static HttpClient Client
        {
            get { return _client; }
        }

        public static void SetToken(string token)
        {
            _client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue(
                    "Bearer",
                    token);
        }

        public static void ClearToken()
        {
            _client.DefaultRequestHeaders.Authorization =
                null;
        }

        public static void SaveToken(string token)
        {
            File.WriteAllText(
                "seller_token.txt",
                token);
        }

        public static string LoadToken()
        {
            if (!File.Exists("seller_token.txt"))
                return "";

            return File.ReadAllText(
                "seller_token.txt");
        }
    }
}