# Zadanie: powiązane / pochodne dokumenty (ZK → WZ/FS)

Instrukcja dla agenta implementującego **Subiekt REST API** (nie Gryzaka).

## Kontekst

Konsument: aplikacja desktop Gryzak. Ma listę ZK/WZ/FS i okno szczegółów dokumentu.
W Subiekcie GT na dokumencie widać powiązania (zakładka „Powiązane”): z ZK wystawione WZ/FS, z WZ faktura itd.

Gryzak NIE łączy się do MSSQL. Wszystko przez to REST API.

## Stan obecny (sprawdzone na żywo)

Baza ma pola, API ich NIE zwraca.

Tabela `dbo.dok__Dokument`:

- `dok_DoDokId` — id dokumentu źródłowego
- `dok_DoDokNrPelny` — numer źródłowego (np. `ZK 8230/M/08/2026`)
- `dok_DoDokDataWyst` — data źródłowego

Tabela `dbo.dok_Pozycja` (powiązanie FS↔WZ na liniach, faktury zbiorcze):

- `ob_DokHanId` / `ob_DokHanLp`
- `ob_DokMagId` / `ob_DokMagLp`

Typy `dok_Typ`: FS=2, WZ=11, ZD=15, ZK=16.

GET `/documents`, `/documents/{id}`, `/documents/zk|wz|fs`, `/documents/zk|wz|fs/{id}`
zwracają m.in. `dok_Id`, `dok_NrPelny`, `dok_Typ`, status, kwoty, kontrahentów, `dok_Pozycja[]`
(`ob_Id`, `ob_TowId`, `tw_Symbol`, `tw_Nazwa`, `ob_Ilosc`, `ob_CenaNetto`, `ob_CenaBrutto`).

Brak w payloadzie: `dok_DoDokId`, `dok_DoDokNrPelny`.
Brak endpointu related/powiązania. Docs/examples: 0 trafień na DoDok / powiąz.

Konwencje API (zachować 1:1):

- Envelope listy: `{ "data": [...], "pagination": { page, pageSize, totalCount, totalPages } }`
- Envelope pojedynczego: `{ "data": { ... } }`
- Błąd: `{ "error": "...", "code": "..." }`
- Nazwy pól = kolumny SQL (`dok_Id`, nie camelCase)
- Property names case-insensitive OK
- Auth: opcjonalny `X-Api-Key`
- Docs: `/docs?format=md` oraz `/examples` — zaktualizować oba

## Cel (P0)

Dwa elementy. Oba potrzebne Gryzakowi.

### 1) Dodać źródło do istniejących DTO dokumentów

Na KAŻDYM GET dokumentu (lista i szczegół: `/documents`, `/documents/{id}`, `/zk`, `/wz`, `/fs`, `/zd` i ich `/{id}`):

```
dok_DoDokId        int?     (null / 0 = brak źródła)
dok_DoDokNrPelny   string?
dok_DoDokDataWyst  datetime?
```

Nie łam istniejących klientów — tylko nowe pola.

### 2) Nowy endpoint: dokumenty pochodne + źródłowy

```
GET /api/v1/documents/{id}/related
GET /api/v1/documents/{typ}/{id}/related
```

`typ` opcjonalny alias jak reszta API: `zk` | `wz` | `fs` | `zd`.
`{id}` = `dok_Id`. 404 gdy dokument nie istnieje.

Odpowiedź:

```json
{
  "data": {
    "dokument": {
      "dok_Id": 1840914,
      "dok_Typ": 16,
      "dok_NrPelny": "ZK 8230/M/08/2026",
      "dok_DataWyst": "2026-08-17T00:00:00",
      "dok_WartBrutto": 123.45
    },
    "zrodlowy": null,
    "pochodne": [
      {
        "dok_Id": 1841001,
        "dok_Typ": 11,
        "dok_NrPelny": "WZ 990/M/08/2026",
        "dok_DataWyst": "2026-08-17T00:00:00",
        "dok_WartBrutto": 123.45,
        "dok_Status": 0,
        "dok_StatusNazwa": "...",
        "zrodloPowiazania": "dok_DoDokId"
      },
      {
        "dok_Id": 1841100,
        "dok_Typ": 2,
        "dok_NrPelny": "FS 440/M/08/2026",
        "dok_DataWyst": "2026-08-18T00:00:00",
        "dok_WartBrutto": 123.45,
        "dok_Status": 0,
        "dok_StatusNazwa": "...",
        "zrodloPowiazania": "dok_DoDokId"
      }
    ]
  }
}
```

`zrodlowy`: jeden dokument albo `null` — ten z `dok_DoDokId` bieżącego.
`pochodne`: tablica (może być pusta), posortowana `dok_DataWyst DESC`, potem `dok_Id DESC`.

`zrodloPowiazania`:

- `"dok_DoDokId"` — nagłówek (główne, P0)
- `"ob_DokHanId"` / `"ob_DokMagId"` — z pozycji (P1, patrz niżej)

Deduplikacja po `dok_Id`. Nie zwracaj samego siebie.

## SQL — P0 (nagłówek, obowiązkowe)

Źródłowy:

```sql
SELECT src.dok_Id, src.dok_Typ, src.dok_NrPelny, src.dok_DataWyst,
       src.dok_WartBrutto, src.dok_Status
FROM dbo.dok__Dokument AS cur
INNER JOIN dbo.dok__Dokument AS src ON src.dok_Id = cur.dok_DoDokId
WHERE cur.dok_Id = @id
  AND cur.dok_DoDokId IS NOT NULL
  AND cur.dok_DoDokId <> 0;
```

Pochodne:

```sql
SELECT dok_Id, dok_Typ, dok_NrPelny, dok_DataWyst, dok_WartBrutto, dok_Status
FROM dbo.dok__Dokument
WHERE dok_DoDokId = @id;
```

Status nazwa: tak samo jak na istniejących listach dokumentów (`dok_StatusNazwa` — ten sam mapping co GET `/documents`).

Slim DTO related NIE musi mieć kontrahenta ani pozycji. Gryzak i tak pobierze pełny dokument po kliknięciu (`GET /documents/{typ}/{id}`).

## SQL — P1 (FS↔WZ z pozycji, zrób jeśli P0 jest gotowe)

W Subiekcie faktura zbiorcza z wielu WZ często NIE ma `dok_DoDokId` na nagłówku. Wtedy:

```sql
-- z dokumentu magazynowego (WZ) → powiązane handlowe (FS)
SELECT DISTINCT d.dok_Id, d.dok_Typ, d.dok_NrPelny, d.dok_DataWyst,
       d.dok_WartBrutto, d.dok_Status
FROM dbo.dok_Pozycja AS p
INNER JOIN dbo.dok__Dokument AS d ON d.dok_Id = p.ob_DokHanId
WHERE p.ob_DokMagId = @id
  AND p.ob_DokHanId IS NOT NULL
  AND p.ob_DokHanId <> 0
  AND d.dok_Id <> @id;

-- z dokumentu handlowego (FS) → powiązane magazynowe (WZ)
SELECT DISTINCT d.dok_Id, d.dok_Typ, d.dok_NrPelny, d.dok_DataWyst,
       d.dok_WartBrutto, d.dok_Status
FROM dbo.dok_Pozycja AS p
INNER JOIN dbo.dok__Dokument AS d ON d.dok_Id = p.ob_DokMagId
WHERE p.ob_DokHanId = @id
  AND p.ob_DokMagId IS NOT NULL
  AND p.ob_DokMagId <> 0
  AND d.dok_Id <> @id;
```

Złącz wyniki z P0, DISTINCT po `dok_Id`. Ustaw `zrodloPowiazania` odpowiednio. Dokument będący już w `zrodlowy` nie duplikuj w `pochodne`.

## Czego NIE robić

- Nie zgaduj powiązań po `dok_NrPelnyOryg` (numer ze sklepu). WZ/FS często go nie mają — fałszywe zapałki.
- Nie zwracaj pełnego drzewa rekurencyjnie (ZK→WZ→FS→KFS). Tylko 1 poziom: źródłowy + bezpośrednie pochodne. Gryzak sam dociągnie kolejny poziom po kliknięciu.
- Nie dodawaj POST/tworzenia powiązań. Tylko odczyt.
- Nie zmieniaj istniejących nazw pól ani kontraktu paginacji.
- Nie wymagaj `typ` do działania `GET /documents/{id}/related` — id jest unikalne.

## Dokumentacja

Dopisać do `/docs?format=md` i `/examples`:

- nowe pola `dok_DoDok*` na listach/szczegółach
- `GET /documents/{id}/related` z przykładem JSON (pusty `pochodne` + niepusty)
- 404 gdy brak dokumentu
- krótka notka: P0 = `dok_DoDokId`; P1 = pozycje FS/WZ

## Testy (obowiązkowe, readonly)

**Baza tylko lokalna / kopia — nigdy produkcja.**

Agent testujący ma używać bazy z **lokalnej konfiguracji API** (connection string / ustawienia aplikacji API na tej maszynie). Nie łącz się do produkcyjnego SQL, nie uruchamiaj zapytań na bazie produkcyjnej.

Zwykle nazwa bazy to `MIKRAN_kopia`.

Przed testami sprawdź w konfiguracji API (`/health` albo ustawienia SQL), że `Initial Catalog` / nazwa bazy to kopia (np. `MIKRAN_kopia`), a nie produkcja (`MIKRAN` bez `_kopia` albo inny host produkcyjny). Jeśli konfiguracja wskazuje na produkcję — **zatrzymaj się** i poproś o przełączenie na kopię.

Testy wyłącznie odczyt (SELECT / GET). Żadnych UPDATE/DELETE/POST tworzących dokumenty na potrzeby testu, chyba że na jawnie lokalnej kopii i po potwierdzeniu.

1. Dokument bez powiązań: `zrodlowy: null`, `pochodne: []`, HTTP 200.
2. WZ/FS z wypełnionym `dok_DoDokId` w SQL → `zrodlowy.dok_Id` się zgadza, `dok_DoDokNrPelny` na GET listy/szczegółu też.
3. ZK, z którego wystawiono WZ: WZ jest w `pochodne` tego ZK; na WZ `zrodlowy` = to ZK.
4. Nieistniejące id → 404 `{ "error": "...", "code": "..." }` w konwencji API.
5. Alias `/documents/zk/{id}/related` = to samo co `/documents/{id}/related`.
6. Smoke: istniejące GET `/documents/zk?page=1&pageSize=1` nadal 200, tylko doszły pola DoDok (mogą być null).

Znajdź w **lokalnej kopii** (`MIKRAN_kopia`) realny przykład:

```sql
SELECT TOP 20
  d.dok_Id, d.dok_Typ, d.dok_NrPelny, d.dok_DoDokId, d.dok_DoDokNrPelny
FROM dbo.dok__Dokument d
WHERE d.dok_DoDokId IS NOT NULL AND d.dok_DoDokId <> 0
  AND d.dok_Typ IN (2, 11, 16)
ORDER BY d.dok_Id DESC;
```

Użyj tych id w testach ręcznych (curl), nie wymyślaj.

## Kryterium ukończenia

- [ ] `dok_DoDokId` / `dok_DoDokNrPelny` / `dok_DoDokDataWyst` na wszystkich GET dokumentów
- [ ] `GET /documents/{id}/related` z `zrodlowy` + `pochodne`
- [ ] alias `/documents/{zk|wz|fs|zd}/{id}/related`
- [ ] docs + examples
- [ ] brak regresji list/szczegółów
- [ ] testy na lokalnej kopii (`MIKRAN_kopia`), nie na produkcji
- [ ] P1 (`ob_DokHanId`/`ob_DokMagId`) zrobiony albo jawnie opisany jako follow-up w docs, jeśli nie w tym PR

Po merge Gryzak doda w oknie szczegółów listę pochodnych i klik → GET pełnego dokumentu. Bez tego endpointu Gryzak nie implementuje UI.
