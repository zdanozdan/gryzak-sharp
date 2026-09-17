using System;
using System.Text;

namespace Gryzak.Services
{
    /// <summary>
    /// Lekka obfuskacja sekretów w pliku eksportu ustawień (nie jest kryptografią).
    /// Format: gryzak-obf1: + Base64(XOR(UTF-8)).
    /// </summary>
    internal static class SettingsSecretCodec
    {
        private const string Prefix = "gryzak-obf1:";
        private static readonly byte[] XorKey = Encoding.UTF8.GetBytes("GryzakSettingsExport.v1");

        public static bool IsEncoded(string? value) =>
            !string.IsNullOrEmpty(value)
            && value.StartsWith(Prefix, StringComparison.Ordinal);

        public static string Encode(string? plain)
        {
            if (string.IsNullOrEmpty(plain))
                return plain ?? "";
            if (IsEncoded(plain))
                return plain;

            var bytes = Encoding.UTF8.GetBytes(plain);
            XorInPlace(bytes);
            return Prefix + Convert.ToBase64String(bytes);
        }

        public static string Decode(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return value ?? "";
            if (!IsEncoded(value))
                return value; // stary eksport / plain text

            try
            {
                var bytes = Convert.FromBase64String(value.Substring(Prefix.Length));
                XorInPlace(bytes);
                return Encoding.UTF8.GetString(bytes);
            }
            catch (FormatException)
            {
                return value;
            }
        }

        private static void XorInPlace(byte[] data)
        {
            for (var i = 0; i < data.Length; i++)
                data[i] ^= XorKey[i % XorKey.Length];
        }
    }
}
