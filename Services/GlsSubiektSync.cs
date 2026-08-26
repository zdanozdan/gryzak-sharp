using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Gryzak.Models;
using static Gryzak.Services.Logger;

namespace Gryzak.Services
{
    public class GlsSubiektSyncResult
    {
        public int? FsDokId { get; set; }
        public SubiektPrzesylka? Przesylka { get; set; }
        public bool Changed { get; set; }
        public bool Cleared { get; set; }
        public bool Skipped { get; set; }
    }

    public static class GlsSubiektSync
    {
        public static bool NeedsGlsSync(SubiektDocument? document)
        {
            var przesylka = document?.Przesylka;
            return przesylka == null
                || przesylka.IsDeleted
                || (!przesylka.HasId && !przesylka.HasAnyNrListu);
        }

        /// <summary>
        /// Po odświeżeniu przygotowalni: nowe wpisy + dokumenty z id w Subiekcie
        /// (żeby zdjąć id, gdy przesyłka opuściła przygotowalnię).
        /// </summary>
        public static bool NeedsPreparingBoxReconcile(SubiektDocument? document)
        {
            if (document == null)
            {
                return false;
            }

            if (NeedsGlsSync(document))
            {
                return true;
            }

            if (document.GlsPreparingBoxId is > 0)
            {
                return true;
            }

            return document.Przesylka is { IsDeleted: false, HasId: true };
        }

        public static List<GlsPreparingBoxItem> MatchPreparingBox(
            SubiektDocument document,
            IEnumerable<GlsPreparingBoxItem> items)
        {
            var list = items as IList<GlsPreparingBoxItem> ?? items.ToList();
            var byRef = list
                .Where(item => GlsConsignment.MatchesDocumentReference(item.References, document))
                .ToList();
            if (byRef.Count > 0)
            {
                return byRef;
            }

            var przesylkaId = document.Przesylka is { IsDeleted: false, HasId: true } p ? p.Id : 0;
            if (przesylkaId > 0)
            {
                var byId = list.Where(item => item.Id == przesylkaId).ToList();
                if (byId.Count > 0)
                {
                    return byId;
                }
            }

            return MatchByParcelNumbers(
                list,
                document.Przesylka?.NrListuPrzygotowalnia,
                item => item.ParcelNumber);
        }

        public static List<GlsPickupItem> MatchPickups(
            SubiektDocument document,
            IEnumerable<GlsPickupItem> items)
        {
            var list = items as IList<GlsPickupItem> ?? items.ToList();
            var byRef = list
                .Where(item => GlsConsignment.MatchesDocumentReference(item.References, document))
                .ToList();
            if (byRef.Count > 0)
            {
                return byRef;
            }

            return MatchByParcelNumbers(
                list,
                document.Przesylka?.NrListuNadane,
                item => item.ParcelNumber);
        }

        private static List<T> MatchByParcelNumbers<T>(
            IEnumerable<T> items,
            string? knownNrListu,
            Func<T, string?> parcelSelector)
        {
            var known = SubiektPrzesylka.SplitNrListu(knownNrListu)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (known.Count == 0)
            {
                return new List<T>();
            }

            return items
                .Where(item => SubiektPrzesylka.SplitNrListu(parcelSelector(item))
                    .Any(nr => known.Contains(nr)))
                .ToList();
        }

        public static string CombinedParcelNumbers(IEnumerable<GlsPreparingBoxItem> items)
        {
            var numbers = items
                .SelectMany(item => SubiektPrzesylka.SplitNrListu(item.ParcelNumber))
                .ToList();
            return SubiektPrzesylka.NormalizeNrListu(string.Join(",", numbers));
        }

        public static string CombinedParcelNumbers(IEnumerable<GlsPickupItem> pickups)
        {
            var numbers = pickups
                .SelectMany(item => SubiektPrzesylka.SplitNrListu(item.ParcelNumber))
                .ToList();
            return SubiektPrzesylka.NormalizeNrListu(string.Join(",", numbers));
        }

        public static string CombineNrListu(params string?[] values) =>
            SubiektPrzesylka.CombineNrListu(values);

        public static async Task<GlsSubiektSyncResult> ApplyToSubiektAsync(
            SubiektApiService api,
            SubiektDocument document,
            IReadOnlyCollection<int> liveBoxIds,
            int? matchedPreparingBoxId,
            string? matchedPreparingNrListu,
            string? matchedPickupNrListu,
            CancellationToken cancellationToken = default)
        {
            // matchedPickupNrListu == null → nie aktualizuj nadań (sync tylko przygotowalni).
            // Non-null → merge/aktualizacja nadań (sync pickup GLS).
            var result = new GlsSubiektSyncResult { Skipped = true };
            var live = liveBoxIds as HashSet<int>
                ?? liveBoxIds.Where(id => id > 0).ToHashSet();

            var (fsId, current) = await api.ResolveInvoiceShipmentAsync(
                document, cancellationToken: cancellationToken).ConfigureAwait(false);
            if (fsId is not int fsDokId || fsDokId <= 0)
            {
                Warning(
                    $"Brak powiązanego FS dla {document.NrPelny} — pominięto zapis Przesylka.",
                    "GlsSubiektSync");
                return result;
            }

            result.FsDokId = fsDokId;

            if (current == null
                && (string.Equals(document.TypKod, "fs", StringComparison.OrdinalIgnoreCase)
                    || document.DokTyp == SubiektApiDocumentTypes.Fs))
            {
                var fsDoc = await api.GetDocumentByIdAsync(
                    "fs", fsDokId, cancellationToken: cancellationToken).ConfigureAwait(false);
                current = fsDoc?.Przesylka;
            }

            if (current != null
                && current.HasId
                && live.Contains(current.Id)
                && matchedPreparingBoxId is not > 0)
            {
                matchedPreparingBoxId = current.Id;
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
                ? SubiektPrzesylka.NormalizeNrListu(matchedPreparingNrListu)
                : "";
            var desiredNadaneNr = matchedPickupNrListu is null
                ? ResolveExistingNadaneNr(current, document)
                : SubiektPrzesylka.NormalizeNrListu(matchedPickupNrListu);
            var staleId = current is { HasId: true } && !live.Contains(current.Id);

            if (desiredBoxId <= 0
                && string.IsNullOrEmpty(desiredPreparingNr)
                && current != null
                && !current.IsDeleted)
            {
                desiredPreparingNr = "";
            }

            if (desiredBoxId > 0
                || !string.IsNullOrEmpty(desiredPreparingNr)
                || !string.IsNullOrEmpty(desiredNadaneNr))
            {
                if (current != null
                    && !current.IsDeleted
                    && current.Id == desiredBoxId
                    && SubiektPrzesylka.SameNrListu(current.NrListuPrzygotowalnia, desiredPreparingNr)
                    && SubiektPrzesylka.SameNrListu(current.NrListuNadane, desiredNadaneNr))
                {
                    result.Przesylka = current;
                    return result;
                }

                var next = new SubiektPrzesylka
                {
                    Typ = "GLS",
                    Id = desiredBoxId,
                    NrListuPrzygotowalnia = desiredPreparingNr,
                    NrListuNadane = desiredNadaneNr,
                    Status = current is { HasId: true, IsDeleted: false }
                        || !string.IsNullOrEmpty(desiredPreparingNr)
                        || !string.IsNullOrEmpty(desiredNadaneNr)
                        ? SubiektPrzesylka.StatusEdytowano
                        : SubiektPrzesylka.StatusUtworzono,
                    Data = DateTime.Now
                };

                var saved = await api.PutPrzesylkaAsync(fsDokId, next, cancellationToken: cancellationToken)
                    .ConfigureAwait(false);
                result.Przesylka = saved ?? next;
                document.Przesylka = result.Przesylka;
                result.Changed = true;
                result.Skipped = false;
                Info(
                    desiredBoxId > 0
                        ? $"Zapisano Przesylka na FS {fsDokId} z GLS: id={desiredBoxId}, przygotowalnia={desiredPreparingNr}, nadane={desiredNadaneNr}."
                        : $"Usunięto id z Przesylka na FS {fsDokId} (brak w przygotowalni GLS), przygotowalnia={desiredPreparingNr}, nadane={desiredNadaneNr}.",
                    "GlsSubiektSync");
                return result;
            }

            if (current == null || current.IsDeleted || (!staleId && !current.IsPreparingBoxOnly))
            {
                result.Przesylka = current;
                return result;
            }

            if (!string.IsNullOrEmpty(current.NrListuNadane))
            {
                var next = new SubiektPrzesylka
                {
                    Typ = "GLS",
                    Id = 0,
                    NrListuPrzygotowalnia = "",
                    NrListuNadane = SubiektPrzesylka.NormalizeNrListu(current.NrListuNadane),
                    Status = SubiektPrzesylka.StatusEdytowano,
                    Data = DateTime.Now
                };

                var saved = await api.PutPrzesylkaAsync(fsDokId, next, cancellationToken: cancellationToken)
                    .ConfigureAwait(false);
                result.Przesylka = saved ?? next;
                document.Przesylka = result.Przesylka;
                document.GlsPreparingBoxId = null;
                document.GlsPreparingBoxParcelNumber = "";
                result.Changed = true;
                result.Skipped = false;
                Info(
                    $"Usunięto id z Przesylka na FS {fsDokId} (brak w przygotowalni GLS), zachowano nadane={next.NrListuNadane}.",
                    "GlsSubiektSync");
                return result;
            }

            var cleared = await api.ClearPrzesylkaAsync(fsDokId, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            result.Przesylka = cleared ?? new SubiektPrzesylka
            {
                Typ = "GLS",
                Status = SubiektPrzesylka.StatusUsuniete,
                Data = DateTime.Now
            };
            document.Przesylka = result.Przesylka;
            document.GlsPreparingBoxId = null;
            document.GlsPreparingBoxParcelNumber = "";
            result.Cleared = true;
            result.Skipped = false;
            Info(
                $"Wyczyszczono Przesylka na FS {fsDokId} — id={current.Id} nie ma już w przygotowalni GLS.",
                "GlsSubiektSync");
            return result;
        }

        private static string ResolveExistingNadaneNr(SubiektPrzesylka? current, SubiektDocument document)
        {
            if (current is { IsDeleted: false } live && !string.IsNullOrWhiteSpace(live.NrListuNadane))
            {
                return SubiektPrzesylka.NormalizeNrListu(live.NrListuNadane);
            }

            if (document.Przesylka is { IsDeleted: false } stored && !string.IsNullOrWhiteSpace(stored.NrListuNadane))
            {
                return SubiektPrzesylka.NormalizeNrListu(stored.NrListuNadane);
            }

            return "";
        }
    }
}
