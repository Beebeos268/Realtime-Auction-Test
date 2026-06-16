using System;
using Microsoft.Data.SqlClient;   // NuGet: Microsoft.Data.SqlClient

namespace TestProject1.API
{
    public static class DbHelper
    {
        // Lấy từ screenshot SSMS: (localdb)\MSSQLLocalDB - RealtimeAuctionDB
        private const string ConnStr =
            @"Server=(localdb)\MSSQLLocalDB;Database=RealtimeAuctionDB;" +
            @"Integrated Security=true;TrustServerCertificate=true;";

        /// <summary>
        /// Đọc EmailVerificationToken (code 6 số) theo Email.
        /// Dùng cho cả verify-email và reset-password (forgot-password ghi
        /// code mới vào cùng cột này).
        /// </summary>
        public static string GetEmailVerificationToken(string email)
        {
            const string query =
                "SELECT EmailVerificationToken FROM [dbo].[Users] WHERE Email = @Email";

            using (var conn = new SqlConnection(ConnStr))
            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@Email", email);
                conn.Open();
                var result = cmd.ExecuteScalar();
                return (result == null || result == DBNull.Value)
                    ? ""
                    : result.ToString();
            }
        }

        // Tiện ích phụ: xoá user theo email để test register lặp lại được
        // (tránh lỗi "email đã tồn tại"). Gọi trong [OneTimeSetUp] nếu cần.
        public static void DeleteUserByEmail(string email)
        {
            const string query = "DELETE FROM [dbo].[Users] WHERE Email = @Email";
            using (var conn = new SqlConnection(ConnStr))
            using (var cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@Email", email);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }
    }
}