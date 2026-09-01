using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Gryzak.Models;
using static Gryzak.Services.Logger;

namespace Gryzak.Services
{
    public class GlsLocalSyncResult
    {
        public int? CacheDokId { get; set; }
        public int CacheDokTyp { get; set; }
        public string CacheTypKod { get; set; } = "";
        public GlsShipmentRecord? Shipment { get; set; }
        public bool Changed { get; set; }
        public bool Cleared { get; set; }
        public bool Skipped { get; set; }
    }

    /// <summary>Indeks lokalnego cache GLS do dopasowania po box_id / dok_id.</summary>
    public sealed class GlsShipmentCacheIndex
    {
        private readonly Dictionary<int, GlsShipmentRecord> _byBoxId = new();
        private readonly Dictionary<(int DokId, int DokTyp), GlsShipmentRecord> _byDokKey = new();

        public static GlsShipmentCacheIndex From(IEnumerable<GlsShipmentRecord>? records)
        {
            var index = new GlsShipmentCacheIndex();
            if (records == null)
            {
                return index;
            }

            foreach (var record in records)
            {
                index.Add(record);
            }

            return index;
        }

        public bool IsEmpty => _byDokKey.Count == 0;

        public bool TryGetByBoxId(int boxId, out GlsShipmentRecord? record)
        {
            if (boxId > 0 && _byBoxId.TryGetValue(boxId, out var found))
            {
                record = found;
                return true;
            }

            record = null;
            return false;
        }

        public bool TryGetByDokId(int dokId, int dokTyp, out GlsShipmentRecord? record)
        {
            if (dokId > 0 && _byDokKey.TryGetValue((dokId, dokTyp), out var found))
            {
                record = found;
                return true;
            }

            record = null;
            return false;
        }

        private void Add(GlsShipmentRecord record)
        {
            if (record.DokId <= 0)
            {
                return;
            }

            _byDokKey[(record.DokId, record.DokTyp)] = record;
            if (record.HasBoxId)
            {
                _byBoxId[record.BoxId] = record;
            }
        }
    }

    /// <summary>
    /// Sync GLS → lokalny cache SQLite (przygotowalnia efemeryczna, nadania read-only merge).
    /// </summary>
    public static class GlsSubiektSync
    {
        public static bool NeedsGlsSync(SubiektDocument? document)
        {
            var shipment = document?.GlsShipment;
            return shipment == null || shipment.IsEmpty;
        }

        /// <summary>
        /// Po odświeżeniu przygotowalni: nowe wpisy + dokumenty z box_id w cache
        /// (żeby zdjąć id, gdy przesyłka opuściła przygotowalnię).
        /// ZK/WZ/FS z dopasowanym live box_id też synchronizujemy (także bez własnego cache).
        /// </summary>
        public static bool NeedsPreparingBoxReconcile(SubiektDocument? document)
        {
            if (document == null)
            {
                return false;
            }

            if (document.GlsPreparingBoxId is > 0)
            {
                return true;
            }

            if (NeedsGlsSync(document))
            {
                return false;
            }

            return document.GlsShipment is { HasBoxId: true };
        }

        public static List<GlsPreparingBoxItem> MatchPreparingBox(
            SubiektDocument document,
            IEnumerable<GlsPreparingBoxItem> items,
            GlsShipmentCacheIndex? cacheIndex = null)
        {
            var list = items as IList<GlsPreparingBoxItem> ?? items.ToList();

            var boxId = document.GlsShipment is { HasBoxId: true } cached ? cached.BoxId : 0;
            if (boxId <= 0
                && cacheIndex != null
                && cacheIndex.TryGetByDokId(document.DokId, document.DokTyp, out var fromIndex)
                && fromIndex is { HasBoxId: true })
            {
                boxId = fromIndex.BoxId;
            }

            if (boxId > 0)
            {
                var byId = list.Where(item => item.Id == boxId).ToList();
                if (byId.Count > 0)
                {
                    return byId;
                }
            }

            var ownerDokId = ResolveGlsOwnerDokId(document);
            var byDokIdRef = list
                .Where(item =>
                    ownerDokId is int owner
                    && GlsConsignment.TryParseReferenceDokId(item.References, out var refDokId)
                    && refDokId == owner)
                .ToList();
            if (byDokIdRef.Count > 0)
            {
                return byDokIdRef;
            }

            return MatchByParcelNumbers(
                list,
                document.GlsShipment?.NrPrzyg,
                item => item.ParcelNumber);
        }

        public static List<GlsPickupItem> MatchPickups(
            SubiektDocument document,
            IEnumerable<GlsPickupItem> items,
            GlsShipmentCacheIndex? cacheIndex = null)
        {
            var list = items as IList<GlsPickupItem> ?? items.ToList();
            var ownerDokId = ResolveGlsOwnerDokId(document);

            var byDokIdRef = list
                .Where(item =>
                    ownerDokId is int owner
                    && GlsConsignment.TryParseReferenceDokId(item.References, out var refDokId)
                    && refDokId == owner)
                .ToList();
            if (byDokIdRef.Count > 0)
            {
                return byDokIdRef;
            }

            var knownParcels = GlsShipmentRecord.CombineNrListu(
                document.GlsShipment?.NrNad,
                document.GlsShipment?.NrPrzyg,
                document.GlsPickupParcelNumber);

            return MatchByParcelNumbers(
                list,
                knownParcels,
                item => item.ParcelNumber);
        }

        private static int? ResolveGlsOwnerDokId(SubiektDocument document)
        {
            if (document.GlsShipmentDokId is int owner && owner > 0)
            {
                return owner;
            }

            return document.DokId > 0 ? document.DokId : null;
        }

        public static SubiektDocument? ResolveDocumentForGlsItem(
            string? references,
            int glsBoxId,
            IEnumerable<SubiektDocument> documents,
            GlsShipmentCacheIndex? cacheIndex = null,
            IReadOnlyDictionary<int, SubiektDocument>? relatedOwnerByDokId = null)
        {
            var docList = documents as IList<SubiektDocument> ?? documents.ToList();

            if (glsBoxId > 0
                && cacheIndex != null
                && cacheIndex.TryGetByBoxId(glsBoxId, out var owner)
                && owner != null)
            {
                var byBox = docList.FirstOrDefault(d =>
                    d.DokId == owner.DokId && d.DokTyp == owner.DokTyp);
                if (byBox != null)
                {
                    return byBox;
                }

                var byCachedOwner = docList.FirstOrDefault(d => d.GlsShipmentDokId == owner.DokId);
                if (byCachedOwner != null)
                {
                    return byCachedOwner;
                }

                if (relatedOwnerByDokId != null
                    && relatedOwnerByDokId.TryGetValue(owner.DokId, out var viaCachedOwner)
                    && viaCachedOwner != null)
                {
                    return viaCachedOwner;
                }
            }

            var byRef = GlsConsignment.FindDocumentForReferenceDokId(references, docList);
            if (byRef != null)
            {
                return byRef;
            }

            if (GlsConsignment.TryParseReferenceDokId(references, out var refDokId)
                && relatedOwnerByDokId != null
                && relatedOwnerByDokId.TryGetValue(refDokId, out var viaRefRelated)
                && viaRefRelated != null)
            {
                return viaRefRelated;
            }

            // Referencja z (#dok_id) wskazuje konkretny dokument — nie dopasowuj po samym numerze FS/WZ.
            if (GlsConsignment.TryParseReferenceDokId(references, out _))
            {
                return null;
            }

            return GlsConsignment.FindDocumentForReferenceNrPelny(references, docList);
        }

        private static List<T> MatchByParcelNumbers<T>(
            IEnumerable<T> items,
            string? knownNrListu,
            Func<T, string?> parcelSelector)
        {
            var known = GlsShipmentRecord.SplitNrListu(knownNrListu)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (known.Count == 0)
            {
                return new List<T>();
            }

            return items
                .Where(item => GlsShipmentRecord.SplitNrListu(parcelSelector(item))
                    .Any(nr => known.Contains(nr)))
                .ToList();
        }

        public static string CombinedParcelNumbers(IEnumerable<GlsPreparingBoxItem> items)
        {
            var numbers = items
                .SelectMany(item => GlsShipmentRecord.SplitNrListu(item.ParcelNumber))
                .ToList();
            return GlsShipmentRecord.NormalizeNrListu(string.Join(",", numbers));
        }

        public static string CombinedParcelNumbers(IEnumerable<GlsPickupItem> pickups)
        {
            var numbers = pickups
                .SelectMany(item => GlsShipmentRecord.SplitNrListu(item.ParcelNumber))
                .ToList();
            return GlsShipmentRecord.NormalizeNrListu(string.Join(",", numbers));
        }

        public static string CombineNrListu(params string?[] values) =>
            GlsShipmentRecord.CombineNrListu(values);

        /// <summary>
        /// Zapis do lokalnego cache. matchedPickupNrListu == null → nie ruszaj nr_nad (sync box).
        /// Non-null → merge/ustaw nr_nad (sync pickup, read-only lustro GLS).
        /// </summary>
        public static async Task<GlsLocalSyncResult> ApplyToLocalAsync(
            GlsShipmentStore store,
            SubiektDocument document,
            IReadOnlyCollection<int> liveBoxIds,
            int? matchedPreparingBoxId,
            string? matchedPreparingNrListu,
            string? matchedPickupNrListu,
            CancellationToken cancellationToken = default)
        {
            var result = new GlsLocalSyncResult { Skipped = true };
            var live = liveBoxIds as HashSet<int>
                ?? liveBoxIds.Where(id => id > 0).ToHashSet();

            // Sync zapisujemy pod dokument z listy (match).
            var cacheDokId = document.DokId;
            var cacheDokTyp = document.DokTyp;

            result.CacheDokId = cacheDokId;
            result.CacheDokTyp = cacheDokTyp;
            result.CacheTypKod = SubiektApiDocumentTypes.FromDokTyp(cacheDokTyp);

            var persisted = await store.GetAsync(cacheDokId, cacheDokTyp, cancellationToken)
                .ConfigureAwait(false);
            var current = persisted;
            if (current == null && document.GlsShipment is { IsEmpty: false } mem)
            {
                current = mem;
            }

            if (current != null
                && current.HasBoxId
                && live.Contains(current.BoxId)
                && matchedPreparingBoxId is not > 0)
            {
                matchedPreparingBoxId = current.BoxId;
            }

            var desiredBoxId = matchedPreparingBoxId is int boxId && live.Contains(boxId)
                ? boxId
                : 0;
            document.GlsPreparingBoxId = desiredBoxId > 0 ? desiredBoxId : null;
            if (desiredBoxId <= 0)
            {
                document.GlsPreparingBoxParcelNumber = "";
            }

            var desiredPreparingNr = desiredBoxId > 0
                ? GlsShipmentRecord.NormalizeNrListu(matchedPreparingNrListu)
                : "";

            // Nadania: null = zachowaj cache; non-null = merge z GLS (nie kasuj przy pustym oknie poza explicit "").
            string desiredNadaneNr;
            if (matchedPickupNrListu is null)
            {
                desiredNadaneNr = ResolveExistingNadaneNr(current, document);
            }
            else if (matchedPickupNrListu.Length == 0)
            {
                // Jawne puste z callera — i tak zachowaj istniejące (pickup read-only, okno API krótkie).
                desiredNadaneNr = ResolveExistingNadaneNr(current, document);
            }
            else
            {
                desiredNadaneNr = GlsShipmentRecord.CombineNrListu(
                    ResolveExistingNadaneNr(current, document),
                    matchedPickupNrListu);
            }

            var staleId = current is { HasBoxId: true } && !live.Contains(current.BoxId);

            if (desiredBoxId > 0
                || !string.IsNullOrEmpty(desiredPreparingNr)
                || !string.IsNullOrEmpty(desiredNadaneNr))
            {
                if (current is { IsEmpty: false }
                    && current.BoxId == desiredBoxId
                    && GlsShipmentRecord.SameNrListu(current.NrPrzyg, desiredPreparingNr)
                    && GlsShipmentRecord.SameNrListu(current.NrNad, desiredNadaneNr)
                    && persisted is { IsEmpty: false }
                    && persisted.BoxId == desiredBoxId
                    && GlsShipmentRecord.SameNrListu(persisted.NrPrzyg, desiredPreparingNr)
                    && GlsShipmentRecord.SameNrListu(persisted.NrNad, desiredNadaneNr))
                {
                    result.Shipment = persisted;
                    ApplyCacheToDocument(document, persisted);
                    return result;
                }

                var next = new GlsShipmentRecord
                {
                    DokId = cacheDokId,
                    DokTyp = cacheDokTyp,
                    NrPelny = document.NrPelny ?? "",
                    Carrier = "GLS",
                    BoxId = desiredBoxId,
                    NrPrzyg = desiredPreparingNr,
                    NrNad = desiredNadaneNr
                };

                await store.UpsertAsync(next, cancellationToken).ConfigureAwait(false);
                result.Shipment = next;
                ApplyCacheToDocument(document, next);
                result.Changed = true;
                result.Skipped = false;
                Info(
                    desiredBoxId > 0
                        ? $"Cache GLS {document.NrPelny}: box={desiredBoxId}, przyg={desiredPreparingNr}, nad={desiredNadaneNr}."
                        : $"Cache GLS {document.NrPelny}: zdjęto box, przyg={desiredPreparingNr}, nad={desiredNadaneNr}.",
                    "GlsLocalSync");
                return result;
            }

            if (current == null || current.IsEmpty || (!staleId && !current.IsPreparingBoxOnly))
            {
                result.Shipment = current;
                ApplyCacheToDocument(document, current);
                return result;
            }

            if (!string.IsNullOrEmpty(current.NrNad))
            {
                var next = new GlsShipmentRecord
                {
                    DokId = cacheDokId,
                    DokTyp = cacheDokTyp,
                    NrPelny = document.NrPelny ?? "",
                    Carrier = "GLS",
                    BoxId = 0,
                    NrPrzyg = "",
                    NrNad = GlsShipmentRecord.NormalizeNrListu(current.NrNad)
                };

                await store.UpsertAsync(next, cancellationToken).ConfigureAwait(false);
                result.Shipment = next;
                ApplyCacheToDocument(document, next);
                document.GlsPreparingBoxId = null;
                document.GlsPreparingBoxParcelNumber = "";
                result.Changed = true;
                result.Skipped = false;
                Info(
                    $"Cache GLS {document.NrPelny}: zdjęto box, zachowano nad={next.NrNad}.",
                    "GlsLocalSync");
                return result;
            }

            await store.DeleteAsync(cacheDokId, cacheDokTyp, cancellationToken).ConfigureAwait(false);
            result.Shipment = null;
            ApplyCacheToDocument(document, null);
            document.GlsPreparingBoxId = null;
            document.GlsPreparingBoxParcelNumber = "";
            result.Cleared = true;
            result.Skipped = false;
            Info(
                $"Wyczyszczono cache GLS {document.NrPelny} — box={current.BoxId} nie ma w przygotowalni.",
                "GlsLocalSync");
            return result;
        }

        public static void ApplyCacheToDocument(
            SubiektDocument document,
            GlsShipmentRecord? shipment,
            int? sourceDokId = null,
            int? sourceDokTyp = null)
        {
            if (shipment == null || shipment.IsEmpty)
            {
                document.GlsShipment = null;
                document.GlsShipmentDokId = null;
                document.GlsShipmentDokTyp = null;
                return;
            }

            document.GlsShipment = shipment;
            document.GlsShipmentDokId = sourceDokId ?? shipment.DokId;
            document.GlsShipmentDokTyp = sourceDokTyp ?? shipment.DokTyp;
        }

        public static (GlsShipmentRecord? Record, int SourceDokId, int SourceDokTyp) ResolveCacheForDocument(
            SubiektDocument document,
            IReadOnlyList<GlsShipmentRecord> allCache,
            IReadOnlyDictionary<int, SubiektDocument>? relatedOwnerByDokId = null)
        {
            if (document.DokId <= 0 || allCache.Count == 0)
            {
                return (null, document.DokId, document.DokTyp);
            }

            GlsShipmentRecord? own = null;
            foreach (var row in allCache)
            {
                if (row.DokId == document.DokId && row.DokTyp == document.DokTyp)
                {
                    own = row;
                    break;
                }
            }

            if (own is { IsEmpty: false } complete && HasPreparingData(complete))
            {
                return (CopyCacheForDocument(document, complete), complete.DokId, complete.DokTyp);
            }

            if (relatedOwnerByDokId != null)
            {
                foreach (var row in allCache)
                {
                    if (row.IsEmpty || !HasPreparingData(row))
                    {
                        continue;
                    }

                    if (relatedOwnerByDokId.TryGetValue(row.DokId, out var owner)
                        && owner.DokId == document.DokId)
                    {
                        return (CopyCacheForDocument(document, row), row.DokId, row.DokTyp);
                    }
                }
            }

            if (own is { IsEmpty: false })
            {
                return (CopyCacheForDocument(document, own), own.DokId, own.DokTyp);
            }

            return (null, document.DokId, document.DokTyp);
        }

        public static GlsShipmentRecord CopyCacheForDocument(
            SubiektDocument document,
            GlsShipmentRecord source) =>
            new()
            {
                DokId = document.DokId,
                DokTyp = document.DokTyp,
                NrPelny = document.NrPelny ?? "",
                Carrier = string.IsNullOrWhiteSpace(source.Carrier) ? "GLS" : source.Carrier,
                BoxId = source.BoxId,
                NrPrzyg = source.NrPrzyg,
                NrNad = source.NrNad,
                UpdatedAt = source.UpdatedAt
            };

        private static bool HasPreparingData(GlsShipmentRecord record) =>
            record.HasBoxId || !string.IsNullOrWhiteSpace(record.NrPrzyg);

        private static string ResolveExistingNadaneNr(GlsShipmentRecord? current, SubiektDocument document)
        {
            if (current is { IsEmpty: false } live && !string.IsNullOrWhiteSpace(live.NrNad))
            {
                return GlsShipmentRecord.NormalizeNrListu(live.NrNad);
            }

            if (document.GlsShipment is { IsEmpty: false } stored && !string.IsNullOrWhiteSpace(stored.NrNad))
            {
                return GlsShipmentRecord.NormalizeNrListu(stored.NrNad);
            }

            return "";
        }
    }
}
