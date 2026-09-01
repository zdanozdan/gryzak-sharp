# Pole własne Przesylka na FS i WZ (JSON)

Instrukcja dla agenta / kontrakt Gryzak ↔ Subiekt REST API.

## Kontekst

Konsument: aplikacja desktop Gryzak. Po utworzeniu listu kuriera (GLS `adePreparingBox_Insert`) API zwraca **id przesyłki**. Później (pickup) pojawi się **numer listu / paczki**. Gryzak zapisuje to w Subiekcie, żeby:

- operator widział to na dokumencie (zakładka Własne),
- Gryzak przy ponownym otwarciu wiedział, że list już jest.

Gryzak NIE łączy się do MSSQL. Wszystko przez to REST API.

## Źródło prawdy (flow biznesowy)

| Skąd wygenerowano list | Gdzie zapis `Przesylka` |
|------------------------|-------------------------|
| **WZ** | pole własne na **WZ** |
| **FS** (list z faktury) | pole własne na **FS** |
| **ZK** | **nie zapisujemy** na ZK — Gryzak szuka powiązanego **WZ** (preferowane), potem **FS** |

- Towar często wychodzi na **WZ bez FS** (faktura zbiorcza / na koniec miesiąca).
- Przy wystawieniu FS z WZ **nie kopiujemy** JSON-a na FS — FS czyta przez `GET .../related` → WZ.
- Legacy: stare wpisy mogą siedzieć na FS; odczyt z WZ bez własnego pola sprawdza też powiązane FS.

Kurier jest **w JSON** (`typ`), nie w nazwie pola — jedno pole na GLS i ewentualnie kolejnych przewoźników.

## Pole w Subiekcie

Administracja → Parametry → Pola własne → obiekt **Dokument** → pole rozszerzone:

- nazwa: **`Przesylka`** (bez spacji, bez polskich znaków)
- typ: **tekst**
- nie wymagane

API **nie** tworzy definicji pola.

## Kontrakt JSON

W bazie: jeden string — kompaktowy JSON (UTF-8). W API (GET/PUT): **obiekt**.

```json
{
  "typ": "GLS",
  "id": 3283,
  "nr_przyg": "",
  "nr_nad": ""
}
```

| Klucz | Znaczenie |
|-------|-----------|
| `typ` | przewoźnik (`GLS`) |
| `id` | id przygotowalni GLS; `0` gdy opuściła przygotowalnię |
| `nr_przyg` | CSV numerów w przygotowalni |
| `nr_nad` | CSV numerów po nadaniu (pickup) |

Numery nie mogą być jednocześnie w obu listach. Sync przygotowalni **nie** czyści `nr_nad`.

Uwaga: kolumna tekstowa bywa `varchar(255)` — stąd krótkie klucze; JSON musi się zmieścić.

## API

### Odczyt

Na `GET /documents/fs/{id}`, `GET /documents/wz/{id}` oraz `GET /documents/{id}` gdy FS/WZ:

```
pw_Przesylka   object | null
```

Na ZK API **nie** dociąga `pw_Przesylka`. Gryzak idzie w `related`.

Lista FS (i ewentualnie WZ) może zwracać `pw_Przesylka` do badge.

### Zapis

```
PUT /api/v1/documents/fs/{id}/przesylka
PUT /api/v1/documents/wz/{id}/przesylka
PUT /api/v1/documents/{id}/przesylka          # alias — tylko FS lub WZ
```

Body: obiekt JSON (`typ` wymagane przy zapisie GLS). Upsert całego JSON-a.

| HTTP | `code` | Kiedy |
|------|--------|--------|
| 404 | `not_found` | brak dokumentu |
| 400 | `not_fs_or_wz` | alias `/documents/{id}/przesylka` a typ ≠ FS/WZ |
| 400 | `validation_error` | zły body |
| 409 | `przesylka_field_not_defined` | brak definicji pola w Subiekcie |
| 400 | `przesylka_value_too_long` | JSON dłuższy niż kolumna |

## Gryzak — zachowanie

1. Po GLS insert / etykiecie / sync → `PUT .../przesylka` na **właścicielu** (WZ lub FS).
2. Hydracja listy: własne `pw_Przesylka`; gdy puste na FS → related WZ; na ZK → related WZ, potem FS.
3. Nie kopiować Przesylka przy tworzeniu FS z WZ.
4. Sync GLS→Subiekt nie nadpisuje FS, które tylko **wyświetla** przesyłkę z powiązanego WZ (`PrzesylkaOwnerDokId` ≠ bieżący dokument).

## Czego nie robić

- Nie zapisywać `Przesylka` na ZK.
- Nie kopiować automatycznie WZ → FS.
- Nie tworzyć definicji pola w SQL.
- Nie używać Sfery (`PoleWlasne`) do zapisu — SQL jak w API.
