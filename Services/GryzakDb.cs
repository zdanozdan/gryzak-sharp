using System;
using System.IO;
using Microsoft.Data.Sqlite;

namespace Gryzak.Services
{
    /// <summary>Wspólna lokalna baza aplikacji (%AppData%\Gryzak\gryzak.db).</summary>
    public sealed class GryzakDb
    {
        private static readonly object Gate = new();
        private static bool _initialized;

        public string DatabasePath { get; }

        public GryzakDb()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(appData, "Gryzak");
            Directory.CreateDirectory(dir);
            DatabasePath = Path.Combine(dir, "gryzak.db");
        }

        public SqliteConnection OpenConnection()
        {
            EnsureCreated();
            var connection = new SqliteConnection($"Data Source={DatabasePath}");
            connection.Open();
            return connection;
        }

        public void EnsureCreated()
        {
            lock (Gate)
            {
                if (_initialized && File.Exists(DatabasePath))
                {
                    return;
                }

                using var connection = new SqliteConnection($"Data Source={DatabasePath}");
                connection.Open();
                using var cmd = connection.CreateCommand();
                cmd.CommandText =
                    """
                    CREATE TABLE IF NOT EXISTS gls_shipment (
                        dok_id INTEGER NOT NULL,
                        dok_typ INTEGER NOT NULL,
                        nr_pelny TEXT NOT NULL DEFAULT '',
                        carrier TEXT NOT NULL DEFAULT 'GLS',
                        box_id INTEGER NOT NULL DEFAULT 0,
                        nr_przyg TEXT NOT NULL DEFAULT '',
                        nr_nad TEXT NOT NULL DEFAULT '',
                        updated_at TEXT NOT NULL,
                        PRIMARY KEY (dok_id, dok_typ)
                    );
                    CREATE INDEX IF NOT EXISTS ix_gls_shipment_nr_pelny
                        ON gls_shipment (nr_pelny COLLATE NOCASE);

                    CREATE TABLE IF NOT EXISTS document_ai_address (
                        dok_id INTEGER NOT NULL,
                        dok_typ INTEGER NOT NULL,
                        nr_pelny TEXT NOT NULL DEFAULT '',
                        name1 TEXT NOT NULL DEFAULT '',
                        name2 TEXT NOT NULL DEFAULT '',
                        name3 TEXT NOT NULL DEFAULT '',
                        street TEXT NOT NULL DEFAULT '',
                        zip_code TEXT NOT NULL DEFAULT '',
                        city TEXT NOT NULL DEFAULT '',
                        country TEXT NOT NULL DEFAULT 'PL',
                        phone TEXT NOT NULL DEFAULT '',
                        contact TEXT NOT NULL DEFAULT '',
                        notes TEXT NOT NULL DEFAULT '',
                        uwagi_hash TEXT NOT NULL DEFAULT '',
                        raw_json TEXT NOT NULL DEFAULT '',
                        confidence TEXT NOT NULL DEFAULT '',
                        updated_at TEXT NOT NULL,
                        PRIMARY KEY (dok_id, dok_typ)
                    );
                    CREATE INDEX IF NOT EXISTS ix_document_ai_address_nr_pelny
                        ON document_ai_address (nr_pelny COLLATE NOCASE);
                    """;
                cmd.ExecuteNonQuery();
                _initialized = true;
            }
        }
    }
}
