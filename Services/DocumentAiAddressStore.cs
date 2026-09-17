using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Gryzak.Models;
using Microsoft.Data.Sqlite;

namespace Gryzak.Services
{
    public sealed class DocumentAiAddressStore
    {
        private readonly GryzakDb _db;

        public DocumentAiAddressStore(GryzakDb? db = null)
        {
            _db = db ?? new GryzakDb();
        }

        public Task<DocumentAiAddress?> GetAsync(
            int dokId,
            int dokTyp,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => Get(dokId, dokTyp), cancellationToken);

        /// <summary>Własny dok_id, potem related ZK → FS → WZ → PA.</summary>
        public async Task<(DocumentAiAddress? Record, int SourceDokId, int SourceDokTyp)> GetForDocumentAsync(
            SubiektDocument document,
            IEnumerable<SubiektRelatedDocument>? related,
            CancellationToken cancellationToken = default)
        {
            var own = await GetAsync(document.DokId, document.DokTyp, cancellationToken)
                .ConfigureAwait(false);
            if (own is { IsEmpty: false })
            {
                return (own, document.DokId, document.DokTyp);
            }

            if (related == null)
            {
                return (null, document.DokId, document.DokTyp);
            }

            var candidates = related
                .Where(r => r.DokId > 0 && !(r.DokId == document.DokId && r.DokTyp == document.DokTyp))
                .OrderBy(r => r.IsZk ? 0 : r.IsFs ? 1 : r.IsWz ? 2 : r.IsPa ? 3 : 4)
                .ThenBy(r => r.DokId);

            foreach (var candidate in candidates)
            {
                var fromRelated = await GetAsync(candidate.DokId, candidate.DokTyp, cancellationToken)
                    .ConfigureAwait(false);
                if (fromRelated is { IsEmpty: false })
                {
                    return (fromRelated, candidate.DokId, candidate.DokTyp);
                }
            }

            return (null, document.DokId, document.DokTyp);
        }

        public Task UpsertAsync(DocumentAiAddress record, CancellationToken cancellationToken = default) =>
            Task.Run(() => Upsert(record), cancellationToken);

        public async Task PropagateToRelatedAsync(
            DocumentAiAddress source,
            IEnumerable<SubiektRelatedDocument>? related,
            CancellationToken cancellationToken = default)
        {
            if (related == null || source.DokId <= 0)
            {
                return;
            }

            foreach (var item in related)
            {
                if (item.DokId <= 0)
                {
                    continue;
                }

                if (item.DokId == source.DokId && item.DokTyp == source.DokTyp)
                {
                    continue;
                }

                if (!item.IsFs && !item.IsWz && !item.IsZk && !item.IsPa)
                {
                    continue;
                }

                var copy = CloneForDocument(source, item.DokId, item.DokTyp, item.NrPelny ?? "");
                await UpsertAsync(copy, cancellationToken).ConfigureAwait(false);
            }
        }

        public Task DeleteAsync(int dokId, int dokTyp, CancellationToken cancellationToken = default) =>
            Task.Run(() => Delete(dokId, dokTyp), cancellationToken);

        public async Task DeleteForDocumentAndRelatedAsync(
            SubiektDocument document,
            IEnumerable<SubiektRelatedDocument>? related,
            CancellationToken cancellationToken = default)
        {
            if (document.DokId > 0)
            {
                await DeleteAsync(document.DokId, document.DokTyp, cancellationToken).ConfigureAwait(false);
            }

            if (related == null)
            {
                return;
            }

            foreach (var item in related)
            {
                if (item.DokId <= 0)
                {
                    continue;
                }

                if (item.DokId == document.DokId && item.DokTyp == document.DokTyp)
                {
                    continue;
                }

                if (!item.IsFs && !item.IsWz && !item.IsZk && !item.IsPa)
                {
                    continue;
                }

                await DeleteAsync(item.DokId, item.DokTyp, cancellationToken).ConfigureAwait(false);
            }
        }

        public void Delete(int dokId, int dokTyp)
        {
            if (dokId <= 0)
            {
                return;
            }

            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                DELETE FROM document_ai_address
                WHERE dok_id = $dok_id AND dok_typ = $dok_typ;
                """;
            cmd.Parameters.AddWithValue("$dok_id", dokId);
            cmd.Parameters.AddWithValue("$dok_typ", dokTyp);
            cmd.ExecuteNonQuery();
        }

        public DocumentAiAddress? Get(int dokId, int dokTyp)
        {
            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                SELECT dok_id, dok_typ, nr_pelny, name1, name2, name3, street, zip_code, city,
                       country, phone, contact, notes, uwagi_hash, raw_json, confidence, updated_at
                FROM document_ai_address
                WHERE dok_id = $dok_id AND dok_typ = $dok_typ
                LIMIT 1;
                """;
            cmd.Parameters.AddWithValue("$dok_id", dokId);
            cmd.Parameters.AddWithValue("$dok_typ", dokTyp);

            using var reader = cmd.ExecuteReader();
            if (!reader.Read())
            {
                return null;
            }

            return ReadRow(reader);
        }

        public void Upsert(DocumentAiAddress record)
        {
            if (record.DokId <= 0)
            {
                return;
            }

            record.UpdatedAt = DateTime.UtcNow;

            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                INSERT INTO document_ai_address (
                    dok_id, dok_typ, nr_pelny, name1, name2, name3, street, zip_code, city,
                    country, phone, contact, notes, uwagi_hash, raw_json, confidence, updated_at)
                VALUES (
                    $dok_id, $dok_typ, $nr_pelny, $name1, $name2, $name3, $street, $zip_code, $city,
                    $country, $phone, $contact, $notes, $uwagi_hash, $raw_json, $confidence, $updated_at)
                ON CONFLICT(dok_id, dok_typ) DO UPDATE SET
                    nr_pelny = excluded.nr_pelny,
                    name1 = excluded.name1,
                    name2 = excluded.name2,
                    name3 = excluded.name3,
                    street = excluded.street,
                    zip_code = excluded.zip_code,
                    city = excluded.city,
                    country = excluded.country,
                    phone = excluded.phone,
                    contact = excluded.contact,
                    notes = excluded.notes,
                    uwagi_hash = excluded.uwagi_hash,
                    raw_json = excluded.raw_json,
                    confidence = excluded.confidence,
                    updated_at = excluded.updated_at;
                """;
            Bind(cmd, record);
            cmd.ExecuteNonQuery();
        }

        public static DocumentAiAddress CloneForDocument(
            DocumentAiAddress source,
            int dokId,
            int dokTyp,
            string nrPelny)
        {
            return new DocumentAiAddress
            {
                DokId = dokId,
                DokTyp = dokTyp,
                NrPelny = nrPelny ?? "",
                Name1 = source.Name1,
                Name2 = source.Name2,
                Name3 = source.Name3,
                Street = source.Street,
                ZipCode = source.ZipCode,
                City = source.City,
                Country = source.Country,
                Phone = source.Phone,
                Contact = source.Contact,
                Notes = source.Notes,
                UwagiHash = source.UwagiHash,
                RawJson = source.RawJson,
                Confidence = source.Confidence,
                UpdatedAt = source.UpdatedAt
            };
        }

        private static void Bind(SqliteCommand cmd, DocumentAiAddress record)
        {
            cmd.Parameters.AddWithValue("$dok_id", record.DokId);
            cmd.Parameters.AddWithValue("$dok_typ", record.DokTyp);
            cmd.Parameters.AddWithValue("$nr_pelny", record.NrPelny ?? "");
            cmd.Parameters.AddWithValue("$name1", record.Name1 ?? "");
            cmd.Parameters.AddWithValue("$name2", record.Name2 ?? "");
            cmd.Parameters.AddWithValue("$name3", record.Name3 ?? "");
            cmd.Parameters.AddWithValue("$street", record.Street ?? "");
            cmd.Parameters.AddWithValue("$zip_code", record.ZipCode ?? "");
            cmd.Parameters.AddWithValue("$city", record.City ?? "");
            cmd.Parameters.AddWithValue("$country", string.IsNullOrWhiteSpace(record.Country) ? "PL" : record.Country);
            cmd.Parameters.AddWithValue("$phone", record.Phone ?? "");
            cmd.Parameters.AddWithValue("$contact", record.Contact ?? "");
            cmd.Parameters.AddWithValue("$notes", record.Notes ?? "");
            cmd.Parameters.AddWithValue("$uwagi_hash", record.UwagiHash ?? "");
            cmd.Parameters.AddWithValue("$raw_json", record.RawJson ?? "");
            cmd.Parameters.AddWithValue("$confidence", record.Confidence ?? "");
            cmd.Parameters.AddWithValue(
                "$updated_at",
                record.UpdatedAt.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture));
        }

        private static DocumentAiAddress ReadRow(SqliteDataReader reader)
        {
            var updatedAtRaw = reader.GetString(16);
            DateTime.TryParse(
                updatedAtRaw,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out var updatedAt);

            return new DocumentAiAddress
            {
                DokId = reader.GetInt32(0),
                DokTyp = reader.GetInt32(1),
                NrPelny = reader.GetString(2),
                Name1 = reader.GetString(3),
                Name2 = reader.GetString(4),
                Name3 = reader.GetString(5),
                Street = reader.GetString(6),
                ZipCode = reader.GetString(7),
                City = reader.GetString(8),
                Country = reader.GetString(9),
                Phone = reader.GetString(10),
                Contact = reader.GetString(11),
                Notes = reader.GetString(12),
                UwagiHash = reader.GetString(13),
                RawJson = reader.GetString(14),
                Confidence = reader.GetString(15),
                UpdatedAt = updatedAt == default ? DateTime.UtcNow : updatedAt
            };
        }
    }
}
