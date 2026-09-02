using System;

namespace Gryzak.Models
{
    /// <summary>Snapshot ustawień aplikacji do przenoszenia między stacjami.</summary>
    public class GryzakSettingsExport
    {
        public const int CurrentVersion = 2;

        public int Version { get; set; } = CurrentVersion;
        public DateTime ExportedAt { get; set; } = DateTime.UtcNow;
        public AppEnvironmentConfig Environment { get; set; } = new();
        public ApiConfig Shop { get; set; } = new();
        public SubiektConfig Subiekt { get; set; } = new();
        public GlsConfig Gls { get; set; } = new();
    }
}
