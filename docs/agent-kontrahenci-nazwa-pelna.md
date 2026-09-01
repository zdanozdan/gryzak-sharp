# `GET /kontrahenci?nazwaPelna=` — dopasowanie pełnej nazwy firmy

Instrukcja dla agenta / kontrakt Gryzak ↔ Subiekt REST API.

## Status

**Naprawione** (2026-09-01). Kryterium akceptacji spełnione — `nazwaPelna` dla kh 17066 / zamówienia 55396 zwraca 1 wynik (LF, CRLF, trailing CRLF, wartość z bazy). Gryzak używa `nazwaPelna` bez zmian w kodzie; fallback po liniach (`nazwa=`) zostaje jako rezerwa.

## Kontekst

Gryzak wyszukuje kontrahenta przy edycji ze szczegółów zamówienia (OpenCart → Subiekt). Od wersji API z parametrem **`nazwaPelna`** firma z zamówienia powinna trafiać dokładnie w `adr_NazwaPelna`, zamiast luźnego `search=` (który zwraca dziesiątki „Pracownia Protetyczno-Ortodontyczna …”).

Gryzak woła:

```
GET /api/v1/kontrahenci?nazwaPelna={payment_company}&pageSize=20
```

z normalizacją końców linii po stronie klienta: `\r\n` / `\r` → `\n`, `Trim()` / `TrimEnd('\n')`.

## Oczekiwane zachowanie (wg dokumentacji)

Parametr `nazwaPelna` = **dokładne** dopasowanie do kolumny `adr_NazwaPelna` po normalizacji końców linii (`\r\n`, `\r` → `\n`).

Przykład z docs:

```bash
curl -G "http://127.0.0.1:5082/api/v1/kontrahenci" \
  --data-urlencode $'nazwaPelna=Linijka 1\r\nLinijka 2'
```

## Przypadek reprodukcji (zamówienie 55396)

### Dane w Subiekcie (kh_Id = 17066)

| Pole | Wartość |
|------|---------|
| `kh_Id` | `17066` |
| `kh_Symbol` | `KOPIJ` |
| `adr_Nazwa` | `Kopij Aleksandra` |
| `adr_NazwaPelna` | `Pracownia Protetyczno-Ortodontyczna` + CRLF + `Kopij Aleksandra` + CRLF |
| `adr_NIP` | `747-175-29-50` |

Bajty `adr_NazwaPelna` (UTF-8):

```
50-72-61-63-6F-77-6E-69-61-20-50-72-6F-74-65-74-79-63-7A-6E-6F-2D-4F-72-74-6F-64-6F-6E-74-79-63-7A-6E-61-0D-0A
4B-6F-70-69-6A-20-41-6C-65-6B-73-61-6E-64-72-61-0D-0A
```

Tekstowo:

```
Pracownia Protetyczno-Ortodontyczna\r\nKopij Aleksandra\r\n
```

### Firma w zamówieniu OpenCart (payment_company)

W UI Gryzaka wyświetla się jako dwie linie:

```
Pracownia Protetyczno-Ortodontyczna
Kopij Aleksandra
```

To powinno mapować się na `adr_NazwaPelna` tego kontrahenta.

## Historia buga (naprawione)

Wcześniej wszystkie poniższe zwracały **`totalCount: 0`**:

```bash
# dokładna wartość z GET /kontrahenci/17066
curl -G "http://127.0.0.1:5082/api/v1/kontrahenci" \
  --data-urlencode "nazwaPelna=Pracownia Protetyczno-Ortodontyczna
Kopij Aleksandra
"

# warianty końców linii
curl -G ".../kontrahenci" --data-urlencode $'nazwaPelna=Pracownia Protetyczno-Ortodontyczna\nKopij Aleksandra'
curl -G ".../kontrahenci" --data-urlencode $'nazwaPelna=Pracownia Protetyczno-Ortodontyczna\r\nKopij Aleksandra'
curl -G ".../kontrahenci" --data-urlencode $'nazwaPelna=Pracownia Protetyczno-Ortodontyczna\r\nKopij Aleksandra\r\n'

# jedna linia ze spacją
curl -G ".../kontrahenci" --data-urlencode "nazwaPelna=Pracownia Protetyczno-Ortodontyczna Kopij Aleksandra"
```

Odpowiedź:

```json
{
  "data": [],
  "pagination": { "page": 1, "pageSize": 5, "totalCount": 0, "totalPages": 0 }
}
```

## Co działa (kontrola)

Te same dane, inne parametry — OK:

```bash
# dokładne adr_Nazwa
curl "http://127.0.0.1:5082/api/v1/kontrahenci?nazwa=Kopij%20Aleksandra"
# → 1 wynik, kh_Id=17066

# luźne search po fragmencie (stare zachowanie — za dużo wyników)
curl "http://127.0.0.1:5082/api/v1/kontrahenci?search=Pracownia%20Protetyczno-Ortodontyczna"
# → ~20 wyników (cała „rodzina” pracowni)

curl "http://127.0.0.1:5082/api/v1/kontrahenci?search=Kopij%20Aleksandra"
# → 1 wynik, kh_Id=17066

curl "http://127.0.0.1:5082/api/v1/kontrahenci/17066"
# → pełna kartoteka z adr_NazwaPelna jak wyżej
```

## Hipoteza buga (do weryfikacji w API)

1. **Filtr `nazwaPelna` nie jest podpięty do SQL** albo porównuje inną kolumnę niż `adr_NazwaPelna` z `adr__Ewid` (`adr_TypAdresu = 1`).
2. **Normalizacja po stronie API** nie jest symetryczna z wartością w bazie (np. brak `RTRIM` końcowych `\r\n` w SQL, albo porównanie bez `REPLACE(adr_NazwaPelna, CHAR(13)+CHAR(10), CHAR(10))`).
3. **Kodowanie / collation** — mało prawdopodobne przy ASCII, ale warto sprawdzić czy parametr w ogóle dociera do warunku WHERE (log SQL).
4. **Trailing newline** — w bazie `adr_NazwaPelna` kończy się `\r\n`; zapytanie bez trailing newline powinno i tak trafić po normalizacji obu stron.

Sugerowane SQL (pseudokod):

```sql
WHERE REPLACE(REPLACE(LTRIM(RTRIM(adr_NazwaPelna)), CHAR(13)+CHAR(10), CHAR(10)), CHAR(13), CHAR(10))
    = @nazwaPelnaNormalized
```

gdzie `@nazwaPelnaNormalized` to parametr z requestu po tej samej normalizacji.

## Obejście po stronie Gryzaka (rezerwowe)

Gdy dokładne `nazwa=` / `nazwaPelna=` zwracają 0 wyników, Gryzak woła **`search=`** z tą samą frazą (luźniejsze dopasowanie, AND po tokenach, kolejność słów nieważna — po naprawie API 2026-09-01).

## Powiązane pliki Gryzaka

- `Services/SubiektApiService.cs` — `SearchKontrahenciAsync`, `NormalizeKontrahentField`
- `Views/OrderDetailsDialog.xaml.cs` — edycja kontrahenta ze szczegółów zamówienia
- Dokumentacja API: `GET /kontrahenci` — parametry `nazwa`, `nazwaPelna`, `search`
