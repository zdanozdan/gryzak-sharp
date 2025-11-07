using System;

namespace Gryzak.Models
{
    public class SubiektConfig
    {
        public string ServerAddress { get; set; } = "";
        public string DatabaseName { get; set; } = "";
        public string ServerUsername { get; set; } = "";
        public string ServerPassword { get; set; } = "";
        public string User { get; set; } = "Szef";
        public string Password { get; set; } = "zdanoszef123";
        public int AutoReleaseLicenseTimeoutMinutes { get; set; } = 0; // 0 = wyłączone, inna wartość = minuty nieaktywności
        public string DiscountCalculationMode { get; set; } = "percent"; // "percent" lub "amount"
    }
}

