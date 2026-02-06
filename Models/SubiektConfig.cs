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
        public int GtProdukt { get; set; } = 1; // 1 - gtaProduktSubiekt, 2 - gtaProduktGestor, etc.
        public int AuthenticationMode { get; set; } = 0; // 0 - Mieszana, 1 - Windows
        public int LaunchDopasujOperatora { get; set; } = 2; // 0 - gtaUruchomDopasuj, 1 - gtaUruchomDopasujUzytkownika, 2 - gtaUruchomDopasujOperatora
        public int LaunchTryb { get; set; } = 0; // 0 - gtaUruchomNormalnie, 1 - gtaUruchomWTle
        public int AutoReleaseLicenseTimeoutMinutes { get; set; } = 0; // 0 = wyłączone, inna wartość = minuty nieaktywności
        public string DiscountCalculationMode { get; set; } = "percent"; // "percent" lub "amount"
        public bool CalculateFromGrossPrices { get; set; } = false; // true = liczenie od cen brutto, false = liczenie od cen netto
        public string DiscountRoundingMode { get; set; } = "percent"; // "none" = bez zaokrąglania, "percent" = do pełnych procentów (1%), "tens" = do dziesiątek procentów (10%)
    }
}

