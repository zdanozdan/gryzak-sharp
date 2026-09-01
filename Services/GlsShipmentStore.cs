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
    public sealed class GlsShipmentStore
    {
        private readonly GryzakDb _db;

        public GlsShipmentStore(GryzakDb? db = null)
        {
            _db = db ?? new GryzakDb();
        }

        public Task<GlsShipmentRecord?> GetAsync(
            int dokId,
            int dokTyp,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => Get(dokId, dokTyp), cancellationToken);

        /// <summary>
        /// Własny dok_id, potem related WZ, ZK, FS.
        /// </summary>
        public async Task<(GlsShipmentRecord? Record, int SourceDokId, int SourceDokTyp)> GetForDocumentAsync(
            SubiektDocument document,
            IEnumerable<SubiektRelatedDocument>? related,
            CancellationToken cancellationToken = default)
        {
            var own = await GetAsync(document.DokId, document.DokTyp, cancellationToken)
                .ConfigureAwait(false);
            if (own is { IsEmpty: false } complete
                && (complete.HasBoxId || !string.IsNullOrWhiteSpace(complete.NrPrzyg)))
            {
                return (complete, document.DokId, document.DokTyp);
            }

            SubiektRelatedDocument? wz = null;
            SubiektRelatedDocument? zk = null;
            SubiektRelatedDocument? fs = null;
            if (related != null)
            {
                foreach (var item in related)
                {
                    if (item.DokId <= 0 || item.DokId == document.DokId)
                    {
                        continue;
                    }

                    if (item.IsWz && wz == null)
                    {
                        wz = item;
                    }
                    else if (item.IsZk && zk == null)
                    {
                        zk = item;
                    }
                    else if (item.IsFs && fs == null)
                    {
                        fs = item;
                    }
                }
            }

            foreach (var candidate in new[] { wz, zk, fs })
            {
                if (candidate == null)
                {
                    continue;
                }

                var fromRelated = await GetAsync(candidate.DokId, candidate.DokTyp, cancellationToken)
                    .ConfigureAwait(false);
                if (fromRelated is { IsEmpty: false })
                {
                    return (MergePreferPreparing(own, fromRelated), candidate.DokId, candidate.DokTyp);
                }
            }

            if (own is { IsEmpty: false })
            {
                return (own, document.DokId, document.DokTyp);
            }

            return (null, document.DokId, document.DokTyp);
        }

        public Task<GlsShipmentRecord?> GetByNrPelnyAsync(
            string? nrPelny,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => GetByNrPelny(nrPelny), cancellationToken);

        private static GlsShipmentRecord MergePreferPreparing(
            GlsShipmentRecord? primary,
            GlsShipmentRecord secondary)
        {
            if (primary == null || primary.IsEmpty)
            {
                return secondary;
            }

            if (secondary == null || secondary.IsEmpty)
            {
                return primary;
            }

            var preparingFromPrimary = primary.HasBoxId || !string.IsNullOrWhiteSpace(primary.NrPrzyg);
            var preparingFromSecondary = secondary.HasBoxId || !string.IsNullOrWhiteSpace(secondary.NrPrzyg);

            return new GlsShipmentRecord
            {
                DokId = primary.DokId,
                DokTyp = primary.DokTyp,
                NrPelny = string.IsNullOrWhiteSpace(primary.NrPelny) ? secondary.NrPelny : primary.NrPelny,
                Carrier = string.IsNullOrWhiteSpace(primary.Carrier) ? secondary.Carrier : primary.Carrier,
                BoxId = preparingFromPrimary
                    ? primary.BoxId
                    : preparingFromSecondary
                        ? secondary.BoxId
                        : 0,
                NrPrzyg = preparingFromPrimary ? primary.NrPrzyg : secondary.NrPrzyg,
                NrNad = !string.IsNullOrWhiteSpace(primary.NrNad) ? primary.NrNad : secondary.NrNad,
                UpdatedAt = primary.UpdatedAt > secondary.UpdatedAt ? primary.UpdatedAt : secondary.UpdatedAt
            };
        }

        /// <summary>
        /// Kopiuje (lub czyści) cache GLS na powiązane ZK/WZ/FS — np. po etykiecie z ZK.
        /// </summary>
        public async Task PropagateToRelatedAsync(
            GlsShipmentRecord? source,
            int sourceDokId,
            int sourceDokTyp,
            IEnumerable<SubiektRelatedDocument>? related,
            CancellationToken cancellationToken = default)
        {
            if (related == null || sourceDokId <= 0)
            {
                return;
            }

            foreach (var item in related)
            {
                if (item.DokId <= 0)
                {
                    continue;
                }

                if (item.DokId == sourceDokId && item.DokTyp == sourceDokTyp)
                {
                    continue;
                }

                if (!item.IsFs && !item.IsWz && !item.IsZk)
                {
                    continue;
                }

                if (source == null || source.IsEmpty)
                {
                    await DeleteAsync(item.DokId, item.DokTyp, cancellationToken).ConfigureAwait(false);
                    continue;
                }

                await UpsertAsync(
                        new GlsShipmentRecord
                        {
                            DokId = item.DokId,
                            DokTyp = item.DokTyp,
                            NrPelny = item.NrPelny ?? "",
                            Carrier = string.IsNullOrWhiteSpace(source.Carrier) ? "GLS" : source.Carrier,
                            BoxId = source.BoxId,
                            NrPrzyg = source.NrPrzyg,
                            NrNad = source.NrNad
                        },
                        cancellationToken)
                    .ConfigureAwait(false);
            }
        }

        public Task UpsertAsync(GlsShipmentRecord record, CancellationToken cancellationToken = default) =>
            Task.Run(() => Upsert(record), cancellationToken);

        public Task DeleteAsync(int dokId, int dokTyp, CancellationToken cancellationToken = default) =>
            Task.Run(() => Delete(dokId, dokTyp), cancellationToken);

        public Task<List<GlsShipmentRecord>> GetAllAsync(CancellationToken cancellationToken = default) =>
            Task.Run(GetAll, cancellationToken);

        public List<GlsShipmentRecord> GetAllSnapshot() => GetAll();

        public Task<GlsShipmentRecord?> GetByBoxIdAsync(
            int boxId,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => GetByBoxId(boxId), cancellationToken);

        /// <summary>
        /// Usuwa wpisy tylko-przygotowalnia, których dok_id nie ma na bieżącej liście dokumentów.
        /// </summary>
        public Task<int> CleanupPreparingOrphansAsync(
            IReadOnlyCollection<int> knownDokIds,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => CleanupPreparingOrphans(knownDokIds), cancellationToken);

        /// <summary>
        /// Usuwa cache GLS, którego dok_id nie występuje już na bieżącej liście (np. po usunięciu dokumentu w Subiekcie).
        /// </summary>
        public Task<int> CleanupOrphansAsync(
            IReadOnlyCollection<int> knownDokIds,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => CleanupOrphans(knownDokIds), cancellationToken);

        public Task<List<GlsShipmentRecord>> GetPageAsync(
            int offset,
            int limit,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => GetPage(offset, limit), cancellationToken);

        public Task<int> CountAsync(CancellationToken cancellationToken = default) =>
            Task.Run(Count, cancellationToken);

        public string DatabasePath => _db.DatabasePath;

        /// <summary>Usuwa cache tylko-przygotowalnia starszy niż maxAge (domyślnie 14 dni).</summary>
        public Task CleanupStalePreparingAsync(
            TimeSpan? maxAge = null,
            CancellationToken cancellationToken = default) =>
            Task.Run(() => CleanupStalePreparing(maxAge ?? TimeSpan.FromDays(14)), cancellationToken);

        private List<GlsShipmentRecord> GetAll() => GetPage(0, int.MaxValue);

        private List<GlsShipmentRecord> GetPage(int offset, int limit)
        {
            if (offset < 0)
            {
                offset = 0;
            }

            if (limit <= 0)
            {
                return new List<GlsShipmentRecord>();
            }

            var list = new List<GlsShipmentRecord>();
            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                SELECT dok_id, dok_typ, nr_pelny, carrier, box_id, nr_przyg, nr_nad, updated_at
                FROM gls_shipment
                ORDER BY updated_at DESC, nr_pelny COLLATE NOCASE
                LIMIT $limit OFFSET $offset;
                """;
            cmd.Parameters.AddWithValue("$limit", limit);
            cmd.Parameters.AddWithValue("$offset", offset);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(ReadRow(reader));
            }

            return list;
        }

        private int Count()
        {
            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM gls_shipment;";
            var result = cmd.ExecuteScalar();
            return result is long n ? (int)n : Convert.ToInt32(result);
        }

        private GlsShipmentRecord? Get(int dokId, int dokTyp)
        {
            if (dokId <= 0)
            {
                return null;
            }

            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                SELECT dok_id, dok_typ, nr_pelny, carrier, box_id, nr_przyg, nr_nad, updated_at
                FROM gls_shipment
                WHERE dok_id = $dok_id AND dok_typ = $dok_typ
                LIMIT 1;
                """;
            cmd.Parameters.AddWithValue("$dok_id", dokId);
            cmd.Parameters.AddWithValue("$dok_typ", dokTyp);
            using var reader = cmd.ExecuteReader();
            return reader.Read() ? ReadRow(reader) : null;
        }

        private GlsShipmentRecord? GetByBoxId(int boxId)
        {
            if (boxId <= 0)
            {
                return null;
            }

            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                SELECT dok_id, dok_typ, nr_pelny, carrier, box_id, nr_przyg, nr_nad, updated_at
                FROM gls_shipment
                WHERE box_id = $box_id
                ORDER BY updated_at DESC
                LIMIT 1;
                """;
            cmd.Parameters.AddWithValue("$box_id", boxId);
            using var reader = cmd.ExecuteReader();
            return reader.Read() ? ReadRow(reader) : null;
        }

        private GlsShipmentRecord? GetByNrPelny(string? nrPelny)
        {
            var nr = (nrPelny ?? "").Trim();
            if (nr.Length == 0)
            {
                return null;
            }

            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                SELECT dok_id, dok_typ, nr_pelny, carrier, box_id, nr_przyg, nr_nad, updated_at
                FROM gls_shipment
                WHERE nr_pelny = $nr_pelny COLLATE NOCASE
                ORDER BY updated_at DESC
                LIMIT 1;
                """;
            cmd.Parameters.AddWithValue("$nr_pelny", nr);
            using var reader = cmd.ExecuteReader();
            return reader.Read() ? ReadRow(reader) : null;
        }

        private int CleanupPreparingOrphans(IReadOnlyCollection<int> knownDokIds)
        {
            if (knownDokIds.Count == 0)
            {
                return 0;
            }

            var known = knownDokIds as HashSet<int> ?? knownDokIds.ToHashSet();
            using var connection = _db.OpenConnection();
            using var select = connection.CreateCommand();
            select.CommandText =
                """
                SELECT dok_id, dok_typ
                FROM gls_shipment
                WHERE box_id > 0
                  AND IFNULL(nr_nad, '') = '';
                """;
            var toDelete = new List<(int DokId, int DokTyp)>();
            using (var reader = select.ExecuteReader())
            {
                while (reader.Read())
                {
                    var dokId = reader.GetInt32(0);
                    if (!known.Contains(dokId))
                    {
                        toDelete.Add((dokId, reader.GetInt32(1)));
                    }
                }
            }

            var removed = 0;
            foreach (var (dokId, dokTyp) in toDelete)
            {
                Delete(dokId, dokTyp);
                removed++;
            }

            return removed;
        }

        private int CleanupOrphans(IReadOnlyCollection<int> knownDokIds)
        {
            if (knownDokIds.Count == 0)
            {
                return 0;
            }

            var known = knownDokIds as HashSet<int> ?? knownDokIds.ToHashSet();
            using var connection = _db.OpenConnection();
            using var select = connection.CreateCommand();
            select.CommandText =
                """
                SELECT dok_id, dok_typ
                FROM gls_shipment;
                """;
            var toDelete = new List<(int DokId, int DokTyp)>();
            using (var reader = select.ExecuteReader())
            {
                while (reader.Read())
                {
                    var dokId = reader.GetInt32(0);
                    if (!known.Contains(dokId))
                    {
                        toDelete.Add((dokId, reader.GetInt32(1)));
                    }
                }
            }

            var removed = 0;
            foreach (var (dokId, dokTyp) in toDelete)
            {
                Delete(dokId, dokTyp);
                removed++;
            }

            return removed;
        }

        private void Upsert(GlsShipmentRecord record)
        {
            if (record.DokId <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(record), "DokId must be > 0.");
            }

            record.NrPrzyg = GlsShipmentRecord.NormalizeNrListu(record.NrPrzyg);
            record.NrNad = GlsShipmentRecord.NormalizeNrListu(record.NrNad);
            record.Carrier = string.IsNullOrWhiteSpace(record.Carrier) ? "GLS" : record.Carrier.Trim().ToUpperInvariant();
            record.UpdatedAt = DateTime.Now;

            if (record.IsEmpty)
            {
                Delete(record.DokId, record.DokTyp);
                return;
            }

            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                INSERT INTO gls_shipment
                    (dok_id, dok_typ, nr_pelny, carrier, box_id, nr_przyg, nr_nad, updated_at)
                VALUES
                    ($dok_id, $dok_typ, $nr_pelny, $carrier, $box_id, $nr_przyg, $nr_nad, $updated_at)
                ON CONFLICT(dok_id, dok_typ) DO UPDATE SET
                    nr_pelny = excluded.nr_pelny,
                    carrier = excluded.carrier,
                    box_id = excluded.box_id,
                    nr_przyg = excluded.nr_przyg,
                    nr_nad = excluded.nr_nad,
                    updated_at = excluded.updated_at;
                """;
            cmd.Parameters.AddWithValue("$dok_id", record.DokId);
            cmd.Parameters.AddWithValue("$dok_typ", record.DokTyp);
            cmd.Parameters.AddWithValue("$nr_pelny", record.NrPelny ?? "");
            cmd.Parameters.AddWithValue("$carrier", record.Carrier);
            cmd.Parameters.AddWithValue("$box_id", record.BoxId);
            cmd.Parameters.AddWithValue("$nr_przyg", record.NrPrzyg ?? "");
            cmd.Parameters.AddWithValue("$nr_nad", record.NrNad ?? "");
            cmd.Parameters.AddWithValue(
                "$updated_at",
                record.UpdatedAt.ToString("o", CultureInfo.InvariantCulture));
            cmd.ExecuteNonQuery();
        }

        private void Delete(int dokId, int dokTyp)
        {
            if (dokId <= 0)
            {
                return;
            }

            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                DELETE FROM gls_shipment
                WHERE dok_id = $dok_id AND dok_typ = $dok_typ;
                """;
            cmd.Parameters.AddWithValue("$dok_id", dokId);
            cmd.Parameters.AddWithValue("$dok_typ", dokTyp);
            cmd.ExecuteNonQuery();
        }

        private void CleanupStalePreparing(TimeSpan maxAge)
        {
            var cutoff = DateTime.Now.Subtract(maxAge)
                .ToString("o", CultureInfo.InvariantCulture);
            using var connection = _db.OpenConnection();
            using var cmd = connection.CreateCommand();
            cmd.CommandText =
                """
                DELETE FROM gls_shipment
                WHERE box_id > 0
                  AND IFNULL(nr_nad, '') = ''
                  AND updated_at < $cutoff;
                """;
            cmd.Parameters.AddWithValue("$cutoff", cutoff);
            cmd.ExecuteNonQuery();
        }

        private static GlsShipmentRecord ReadRow(SqliteDataReader reader)
        {
            var updatedRaw = reader.GetString(7);
            DateTime updatedAt = DateTime.TryParse(
                    updatedRaw,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind,
                    out var parsed)
                ? parsed
                : DateTime.Now;

            return new GlsShipmentRecord
            {
                DokId = reader.GetInt32(0),
                DokTyp = reader.GetInt32(1),
                NrPelny = reader.IsDBNull(2) ? "" : reader.GetString(2),
                Carrier = reader.IsDBNull(3) ? "GLS" : reader.GetString(3),
                BoxId = reader.GetInt32(4),
                NrPrzyg = reader.IsDBNull(5) ? "" : reader.GetString(5),
                NrNad = reader.IsDBNull(6) ? "" : reader.GetString(6),
                UpdatedAt = updatedAt
            };
        }
    }
}
