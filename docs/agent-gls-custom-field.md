# Zadanie: pole własne Przesylka na FS (JSON)

Instrukcja dla agenta implementującego **Subiekt REST API** (nie Gryzaka).

## Kontekst

Konsument: aplikacja desktop Gryzak. Po utworzeniu listu kuriera (na start GLS `adePreparingBox_Insert`) API zwraca **id przesyłki**. Później (pickup) pojawi się **numer listu / paczki**. Gryzak musi to zapisać w Subiekcie, żeby:

- operator widział to na FS (zakładka Własne),
- Gryzak przy ponownym otwarciu FS wiedział, że list już jest.

Gryzak NIE łączy się do MSSQL. Wszystko przez to REST API.

Paczka realnie wychodzi z **FS**. Źródło prawdy = pole własne **`Przesylka`** na **FS**. Nie kopiować na ZK/WZ. Gryzak na ZK/WZ dociągnie FS przez `GET /documents/{id}/related`, potem pełne FS.

Kurier jest **w JSON** (`typ`), nie w nazwie pola — jedno pole na GLS i ewentualnie kolejnych przewoźników.

## Stan obecny (sprawdzić na żywo)

GET `/documents/fs/{id}` zwraca nagłówek + pozycje + kontrahentów. **Brak pól własnych.**

Sfera (`gta.txt`): `SuDokument.PoleWlasne("Nazwa_pola") = wartosc`, zapis do `dbo.pw_Dane`. Wymaga **Niebieskiego Plus**. Pole trzeba **najpierw zdefiniować w Subiekcie** — API go nie tworzy.

Konwencje API (zachować 1:1):

- Envelope pojedynczego: `{ "data": { ... } }`
- Błąd: `{ "error": "...", "code": "..." }`
- Nazwy pól SQL-like (`dok_Id`), property names case-insensitive OK
- Auth: opcjonalny `X-Api-Key`
- Docs: `/docs?format=md` oraz `/examples` — zaktualizować oba
- Zapis SQL bez Sfery już jest (WyPLUwacz: `POST /products/fix-plu`). Ten feature też **SQL na `pw_Dane`**, nie Sfera (nie spalaj licencji na jedną komórkę).

## Pole w Subiekcie (przed implementacją / testami)

W **lokalnej kopii** (`MIKRAN_kopia`), nie na produkcji:

Administracja → Parametry → Pola własne → obiekt **Dokument** → dodaj pole rozszerzone:

- nazwa: **`Przesylka`** (dokładnie tak: bez spacji, bez polskich znaków — nie `Przesyłka`)
- typ: **tekst**
- nie wymagane
- wartość domyślna: pusta

Jeśli pola nie ma — API ma zwracać błąd (patrz niżej), **nie** INSERT-ować definicji pola.

## Kontrakt JSON w polu `Przesylka`

W bazie: **jeden string** — kompaktowy JSON (bez pretty-print, UTF-8). Operator czyta to w zakładce Własne.

W API (GET/PUT): **obiekt**, nie string.

```json
{
  "typ": "GLS",
  "id": 3283,
  "nr_listu_przygotowalnia": "",
  "nr_listu_nadane": ""
}
```

| Klucz | Typ | Wymagany | Znaczenie |
|-------|-----|----------|-----------|
| `typ` | string | tak przy zapisie | przewoźnik. P0: tylko `"GLS"` (przyjmij case-insensitive, **zapisuj wielkimi**: `GLS`) |
| `id` | number (int ≥ 0) | nie | id przesyłki u przewoźnika (GLS: `adePreparingBox_Insert`, przygotowalnia). `0` lub brak = przesyłka opuściła przygotowalnię |
| `nr_listu_przygotowalnia` | string | nie | numery listów/paczek **w przygotowalni GLS** (przechowalnia — po wygenerowaniu etykiety, przed nadaniem). Lista CSV (`,`). Pusty `""` gdy brak etykiet |
| `nr_listu_nadane` | string | nie | numery listów/paczek **już nadanych** (pickup GLS). Lista CSV. Pusty `""` dopóki kurier nie odebrał |
| `status` | string | nie | `utworzono` / `edytowano` / `usuniete` — opcjonalne metadane od Gryzaka |
| `data` | string | nie | ostatnia zmiana, format `yyyy-MM-dd HH:mm:ss` |

### Cykl życia (GLS)

1. **Przygotowalnia bez etykiety** — tylko `id`, oba `nr_listu_*` puste.
2. **Przygotowalnia z etykietą (przechowalnia)** — `id` + numery w `nr_listu_przygotowalnia`.
3. **Po nadaniu** — numery przechodzą z `nr_listu_przygotowalnia` do `nr_listu_nadane`; `id` znika gdy przesyłka opuściła przygotowalnię.

Numery **nie mogą** być jednocześnie w obu listach — po nadaniu zdejmij je z `nr_listu_przygotowalnia`.

**Niezależność list:** operacje na przygotowalni (nowa paczka, etykieta, sync przygotowalni) **nie wolno** czyścić ani nadpisywać `nr_listu_nadane`. Nadania aktualizuje wyłącznie sync pickup GLS.

Zasady:

- Tylko powyższe klucze (+ opcjonalnie `status`, `data`). Nie dodawaj `nr_listu` (stary format — ignoruj). Nie dodawaj `utworzono` / `dok_Id` / gniazd.
- `typ` inny niż `GLS` w P0 → 400 `validation_error` (zostaw miejsce na DPD itd. później, nie implementuj teraz).
- Wartości `nr_listu_*`: trim; `null` w body traktuj jak `""`; normalizacja CSV bez duplikatów po stronie konsumenta.
- Nieznane klucze w **istniejącym** stringu w bazie: przy PUT **nie zachowuj** (pełna podmiana P0). Prościej i przewidywalnie.
- Max długość stringu: zmieść się w typie tekstowym pola (JSON może urosnąć przy wielu paczkach). Jeśli SQL obetnie — 400 `przesylka_value_too_long`.

Przykłady stringu w `pw_Dane`:

```
{"typ":"GLS","id":3283,"nr_listu_przygotowalnia":"","nr_listu_nadane":""}
{"typ":"GLS","id":3283,"nr_listu_przygotowalnia":"000111,000222","nr_listu_nadane":""}
{"typ":"GLS","id":0,"nr_listu_przygotowalnia":"","nr_listu_nadane":"000111,000222"}
```

## Cel (P0)

### 1) Odczyt na szczegółach FS

Na `GET /documents/fs/{id}` oraz `GET /documents/{id}` gdy `dok_Typ = 2`:

```
pw_Przesylka   object | null
```

`null` gdy brak wiersza w `pw_Dane` albo pusty tekst.

Jeśli w bazie jest śmieć (nie JSON / brak `typ` / brak `id`) — nie wal 500. Zwróć `pw_Przesylka: null`. Możesz zalogować warning.

Na ZK/WZ **nie** dokładaj `pw_Przesylka` z powiązanego FS. Gryzak sam idzie w `related`.

Lista `GET /documents/fs` — P0 **nie musi** mieć `pw_Przesylka` (join na listę). P1: dodać, jeśli tanie (LEFT JOIN), do badge w Gryzaku.

### 2) Zapis

```
PUT /api/v1/documents/fs/{id}/przesylka
PUT /api/v1/documents/{id}/przesylka
```

`{id}` = `dok_Id`. Alias bez `fs` działa tylko gdy dokument **jest FS** (`dok_Typ = 2`).

Body (bez envelope):

```json
{
  "typ": "GLS",
  "id": 3283,
  "nr_listu_przygotowalnia": "",
  "nr_listu_nadane": ""
}
```

`typ` wymagane. `id` opcjonalne (≥ 0). `nr_listu_przygotowalnia` i `nr_listu_nadane` opcjonalne.

Sukces HTTP **200**:

```json
{
  "data": {
    "dok_Id": 1841100,
    "dok_Typ": 2,
    "dok_NrPelny": "FS 440/M/08/2026",
    "pw_Przesylka": {
      "typ": "GLS",
      "id": 3283,
      "nr_listu_przygotowalnia": "000111",
      "nr_listu_nadane": ""
    }
  }
}
```

Zachowanie zapisu: **upsert** całego JSON-a w polu `Przesylka` tego FS. Drugi PUT nadpisuje (np. uzupełnienie numerów po etykiecie lub przeniesienie do `nr_listu_nadane`).

### 3) Błędy

| HTTP | `code` | Kiedy |
|------|--------|--------|
| 404 | `not_found` | brak dokumentu |
| 400 | `not_fs` | `PUT /documents/{id}/przesylka` a `dok_Typ ≠ 2` |
| 400 | `validation_error` | brak `typ`, `typ` ≠ `GLS`, zły typ pól |
| 409 | `przesylka_field_not_defined` | w bazie nie ma definicji pola własnego nazwanego `Przesylka` dla dokumentu |
| 400 | `przesylka_value_too_long` | JSON dłuższy niż kolumna tekstowa |
| 409 | `conflict` | (opcjonalnie) gdy chcesz zablokować nadpisanie istniejącego `id` innym — **nie w P0**. P0 zawsze nadpisuje. |

Body błędu jak reszta API: `{ "error": "...", "code": "..." }`. Komunikat `przesylka_field_not_defined` ma mówić: dodać pole w Administracja → Pola własne → Dokument, nazwa `Przesylka`, typ tekst.

## SQL — najpierw odkryj schemat

Kolumny `pw_*` różnią się między wersjami InsERT GT. **Nie zgaduj nazw.** Na `MIKRAN_kopia`:

```sql
SELECT TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = 'dbo'
  AND (TABLE_NAME LIKE 'pw%' OR TABLE_NAME LIKE '%Wlasn%')
ORDER BY TABLE_NAME;
```

Potem kolumny `pw_Dane` i tabeli definicji (często `pw__Definicja` / `pw__Pole` / podobne). GET `/schema` API jeśli pomaga.

Szukaj definicji:

- nazwa pola = `Przesylka` (case-insensitive, `LTRIM/RTRIM`; nie mylić z `Przesyłka`),
- obiekt = dokument (`SuDokument` / dokument handlowy — ustal po danych, nie po zgadywaniu enum).

Odczyt wartości: join definicja ↔ `pw_Dane` WHERE obiekt_id = `@dok_Id`.

Zapis: `UPDATE` tekstu gdy wiersz jest; `INSERT` gdy nie ma. Nie ruszaj innych pól własnych tego dokumentu.

Typ pola w definicji musi być **tekst**. Jeśli `Przesylka` istnieje jako flaga/liczba — 409 z jasnym błędem (`przesylka_field_wrong_type`), nie pchaj JSON-a w `pwd_Flaga`.

Transakcja na INSERT/UPDATE. Parametryzowane zapytania.

## Czego NIE robić

- Nie twórz definicji pola `Przesylka` w SQL / Sferze. Tylko używaj istniejącej.
- Nie nazywaj pola `GLS` / `Przesyłka`. Kurier jest w JSON `typ`.
- Nie zapisuj przez Sferę (`PoleWlasne`) w P0 — za drogie (licencja). SQL jak fix-plu.
- Nie kopiuj JSON-a na ZK/WZ ani na `pochodne` / `zrodlowy`.
- Nie zgaduj FS po `dok_NrPelnyOryg`.
- Nie zmieniaj istniejących nazw pól dokumentów ani paginacji.
- Nie zwracaj surowego stringa jako `pw_Przesylka` gdy da się sparsować — Gryzak chce obiekt.
- Nie rób DELETE endpointu w P0. Brak kasowania w P0 (Gryzak na razie tylko ustawia).
- Nie zapisuj na produkcji. Testy tylko `MIKRAN_kopia`.

## Dokumentacja

Dopisać do `/docs?format=md` i `/examples`:

- `pw_Przesylka` na GET szczegółu FS (`null` vs obiekt z `typ`/`id`/`nr_listu_przygotowalnia`/`nr_listu_nadane`)
- `PUT /documents/fs/{id}/przesylka` + alias `/documents/{id}/przesylka`
- przykłady JSON (przygotowalnia z etykietą / po nadaniu)
- 404 / 400 `not_fs` / 409 `przesylka_field_not_defined`
- notka: pole `Przesylka` dodać ręcznie w Subiekcie; źródło prawdy = FS; ZK/WZ nie dziedziczą; P0 tylko `typ=GLS`

## Testy (lokalna kopia)

**Baza tylko lokalna / kopia — nigdy produkcja.**

Przed testami: `/health` albo connection string — `Initial Catalog` to kopia (`MIKRAN_kopia`), nie `MIKRAN`. Jeśli produkcja — **zatrzymaj się**.

PUT **wolno** odpalać na kopii (to jest ten feature). Nie na produkcji.

1. Upewnij się, że pole `Przesylka` (tekst, Dokument) istnieje na kopii. Jeśli nie — dodaj w Subiekcie albo poproś operatora; bez tego 409 jest poprawnym wynikiem testu „brak definicji”.
2. Weź prawdziwe FS z kopii:

```sql
SELECT TOP 10 dok_Id, dok_NrPelny, dok_Typ
FROM dbo.dok__Dokument
WHERE dok_Typ = 2
ORDER BY dok_Id DESC;
```

3. GET `/documents/fs/{id}` — 200, jest `pw_Przesylka` (null albo obiekt).
4. PUT `/documents/fs/{id}/przesylka` body `{ "typ": "GLS", "id": 3283, "nr_listu_przygotowalnia": "", "nr_listu_nadane": "" }` — 200, GET zwraca ten sam obiekt (`typ` wielkimi).
5. PUT ponownie `{ "typ": "gls", "id": 3283, "nr_listu_przygotowalnia": "000111", "nr_listu_nadane": "" }` — nadpisanie, GET ma `typ: "GLS"` i numery w przygotowalni.
6. PUT `{ "typ": "GLS", "id": 0, "nr_listu_przygotowalnia": "", "nr_listu_nadane": "000111" }` — po nadaniu, GET ma puste przygotowalnia i numery w nadane.
7. PUT `{ "typ": "DPD", "id": 1 }` → 400 `validation_error` (P0).
8. PUT na ZK id — alias `/documents/{zkId}/przesylka` → 400 `not_fs`.
9. GET nieistniejące id → 404 `{ "error", "code": "not_found" }`.
10. PUT `{ "typ": "GLS" }` bez `typ` poprawnego → 400 `validation_error`.
11. Smoke: GET `/documents/fs?page=1&pageSize=1` nadal 200 (lista bez regresji).
12. W Subiekcie (kopia) otwórz to FS → zakładka Własne → pole **Przesylka** pokazuje JSON. Jeśli nie widać: sprawdź czy UPDATE trafił w ten sam obiekt/definicję co UI (typ obiektu).

Po testach PUT zostaw JSON na kopii albo wróć poprzednią wartość — kopia, nie produkcja, ale nie śmieć setek FS.

## Kryterium ukończenia

- [ ] `pw_Przesylka` na GET szczegółu FS (`/documents/fs/{id}` i `/documents/{id}` dla FS)
- [ ] `PUT /documents/fs/{id}/przesylka` upsert JSON `{ typ, id, nr_listu_przygotowalnia, nr_listu_nadane }`
- [ ] alias `PUT /documents/{id}/przesylka` tylko dla FS, inaczej `not_fs`
- [ ] P0: `typ` tylko `GLS` (normalizacja do wielkich)
- [ ] 409 gdy brak definicji pola `Przesylka`
- [ ] brak zapisu na ZK/WZ, brak kopiowania powiązań
- [ ] docs + examples
- [ ] brak regresji GET list/szczegółów
- [ ] testy na `MIKRAN_kopia`, nie na produkcji
- [ ] lista FS z `pw_Przesylka` zrobiona albo jawnie follow-up w docs (P1)

Po merge Gryzak: po udanym GLS insert → `PUT .../przesylka` z `typ: GLS`; przy otwarciu FS czyta `pw_Przesylka`; jeśli `id` jest — nie tworzy drugiego listu. Na ZK/WZ podgląd przez related → FS.
