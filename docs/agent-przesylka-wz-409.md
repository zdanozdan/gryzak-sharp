# Bug: PUT Przesylka na WZ → 409 `przesylka_field_not_defined`

## Objaw (2026-08-26, `MIKRAN_kopia`, API `127.0.0.1:5082`)

To samo pole własne **`Przesylka`** (Dokument, tekst):

| Request | Wynik |
|---------|--------|
| `PUT /documents/fs/{id}/przesylka` | **200** — zapis OK |
| `PUT /documents/wz/{id}/przesylka` | **409** `przesylka_field_not_defined` |
| `PUT /documents/{wzId}/przesylka` | **409** to samo |

Przykład WZ: `dok_Id=1782146` (`WZ 11/M/08/2026`). GET `/documents/wz/{id}` zwraca `pw_Przesylka: null` (odczyt ścieżki działa / pole mapowane), ale upsert pada na lookup definicji.

Gryzak zapisuje list z WZ na WZ — poprawnie. Bez naprawy API sync WZ→Przesylka zawsze failuje 409.

## Oczekiwane

Lookup definicji pola `Przesylka` dla obiektu **Dokument** ma być **wspólny** dla FS (`dok_Typ=2`) i WZ (`dok_Typ=11`). W Subiekcie GT to jedno pole własne dokumentu, nie osobne per typ.

Nie wymagaj osobnej definicji „Przesylka dla WZ”. Jeśli definicja istnieje i działa dla FS — PUT WZ musi użyć tej samej definicji / tego samego `pw_*` id pola.

## Reprodukcja

```http
PUT /api/v1/documents/wz/1782146/przesylka
Content-Type: application/json

{"typ":"GLS","id":1,"nr_przyg":"","nr_nad":""}
```

Oczekiwane: 200 + `pw_Przesylka` w body (jak na FS).  
Aktualne: 409 `przesylka_field_not_defined`.

Kontrola: ten sam body na `PUT /documents/fs/{fsId}/przesylka` → 200.

## Kryterium ukończenia

- [ ] PUT WZ i PUT FS używają tej samej definicji `Przesylka`
- [ ] 409 tylko gdy w bazie **naprawdę** brak definicji (wtedy FS też 409)
- [ ] GET WZ po PUT zwraca zapisany obiekt
