# Zadanie: uwagi dokumentu (`dok_Uwagi`) na GET dokumentów

Instrukcja dla agenta implementującego **Subiekt REST API** (nie Gryzaka).

## Kontekst

Konsument: aplikacja desktop Gryzak. Ma listę ZK/WZ/FS i okno szczegółów dokumentu.
W Subiekcie GT na dokumencie jest pole **Uwagi** (opis dokumentu). Gryzak przy tworzeniu ZK zapisuje tam m.in. dane z zamówienia (Sfera: `ZmienOpisDokumentu` / `dokument.Uwagi`).

Gryzak NIE łączy się do MSSQL. Wszystko przez to REST API.
Chce **czytać** uwagi dla każdego typu dokumentu (nie tylko ZK).

## Stan obecny (sprawdzone na żywo, API test `…:5082`)

### Baza

Tabela `dbo.dok__Dokument` — kolumna:

| Kolumna SQL | Znaczenie |
|-------------|-----------|
| `dok_Uwagi` | Tekst uwag / opisu dokumentu (to, co widać w Subiekcie jako Uwagi) |

(Jeśli w danej instalacji uwagi siedzą też w osobnej tabeli opisów — mapować **to samo**, co pokazuje UI Subiekta jako Uwagi; źródło prawdy = wartość widoczna operatorowi.)

### API — odczyt

`GET /documents`, `/documents/{id}`, `/documents/zk|wz|fs|zd|…`, `/{typ}/{id}`:

- zwracają m.in. `dok_Id`, `dok_NrPelny`, `dok_NrPelnyOryg`, status, kwoty, kontrahentów, `dok_DoDok*`, pozycje (na szczegółach)
- **brak** w payloadzie: `dok_Uwagi` / `uwagi`

Sprawdzone na liście i szczególe ZK oraz liście FS/WZ (prod `…:5080`): klucze odpowiedzi **nie zawierają** `dok_Uwagi` ani `uwagi`.

### API — zapis (już jest, nie ruszać kontraktu)

`POST /documents/zk/create` oraz `POST /documents/zd/create` przyjmują opcjonalne body:

```json
{
  "kontrahentId": 1,
  "uwagi": "opcjonalnie",
  "pozycje": [ … ]
}
```

To jest **write** przy tworzeniu. Po `201` odpowiedź ma być jak `GET /documents/zk/{id}` — czyli po tej zmianie **powinna już zawierać** `dok_Uwagi`.

Konwencje API (zachować 1:1):

- Envelope listy: `{ "data": [...], "pagination": { … } }`
- Envelope pojedynczego: `{ "data": { … } }`
- Nazwy pól = kolumny SQL (`dok_Uwagi`, nie camelCase `uwagi` w GET)
- Property names case-insensitive OK
- Auth: opcjonalny `X-Api-Key`
- Docs: `/docs?format=md` oraz `/examples` — zaktualizować oba

## Cel (P0)

### Dodać `dok_Uwagi` do DocumentDto na KAŻDYM GET dokumentu

Na liście i szczegółach — te same ścieżki co reszta pól dokumentu:

```
GET /api/v1/documents
GET /api/v1/documents/{id}
GET /api/v1/documents/zk|wz|fs|zd|pa|…   (wszystkie aliasy typów już obsługiwane)
GET /api/v1/documents/{typ}/{id}
```

oraz wszędzie, gdzie zwracany jest pełny `DocumentDto` (np. odpowiedź po `POST …/create`, lookup `?numerOryginalny=`).

Nowe pole:

```
dok_Uwagi    string?    (null lub "" = brak uwag)
```

Wymagania:

- **Wszystkie typy dokumentów** obsługiwane przez API dokumentów (ZK, WZ, FS, ZD, PA, …) — nie tylko ZK.
- Lista i szczegół — ta sama mapa pól (lista może mieć uwagi; nie ukrywać ich tylko na `{id}`).
- Nie łamać istniejących klientów — tylko **dodać** pole.
- Nie zmieniać semantyki body `uwagi` przy `POST …/create` (zostaje `uwagi` w JSON body; w odpowiedzi GET: `dok_Uwagi`).
- Encoding UTF-8; nie ucinać treści bez potrzeby (jak w bazie).

### Poza zakresem (nie robić w tym zadaniu)

- `PUT` / patch uwag na istniejącym dokumencie (osobne zadanie, jeśli kiedyś potrzebne).
- Wyszukiwanie `?search=` po treści uwag (opcjonalne później — nie wymagane teraz).
- Zmiany w Gryzaku (konsument dociągnie pole po wdrożeniu API).

## Przykład odpowiedzi (fragment)

```json
{
  "data": {
    "dok_Id": 1840914,
    "dok_Typ": 16,
    "dok_NrPelny": "ZK 8230/M/08/2026",
    "dok_NrPelnyOryg": "141325",
    "dok_Uwagi": "Zamówienie sklep #141325 …",
    "dok_WartBrutto": 123.45,
    "kh__Kontrahent_Odbiorca": { }
  }
}
```

Lista:

```json
{
  "data": [
    {
      "dok_Id": 1840914,
      "dok_NrPelny": "ZK 8230/M/08/2026",
      "dok_Uwagi": "…",
      "…": "…"
    }
  ],
  "pagination": { "page": 1, "pageSize": 20, "totalCount": 1, "totalPages": 1 }
}
```

## Jak sprawdzić (acceptance)

1. W Subiekcie GT otwórz ZK/WZ/FS z niepustymi Uwagami; zapisz `dok_Id`.
2. `GET /documents/zk/{id}` (oraz wz/fs) — w `data` jest `dok_Uwagi` z tą samą treścią co w UI.
3. `GET /documents/zk?page=1&pageSize=20` — pozycje mają `dok_Uwagi` (null/`""` gdy puste).
4. Dokument bez uwag: pole obecne, wartość `null` lub `""` (wybrać jedną konwencję i trzymać ją w docs).
5. `POST /documents/zk/create` z `"uwagi":"test-api-uwagi"` → `201` → odpowiedź i kolejny GET zawierają `dok_Uwagi` = `"test-api-uwagi"`.
6. Zaktualizowane `/docs?format=md` i `/examples`: lista pól DocumentDto zawiera `dok_Uwagi`; w przykładach JSON widać pole.

## Kryteria ukończenia

- [ ] `dok_Uwagi` na wszystkich GET dokumentów (lista + szczegół + wszystkie typy aliasów)
- [ ] `dok_Uwagi` w DocumentDto zwracanym po create / lookupach zwracających pełny dokument
- [ ] Brak regresji istniejących pól
- [ ] `/docs` i `/examples` zaktualizowane
- [ ] Ręczny smoke na ZK + FS + WZ z niepustymi uwagami
