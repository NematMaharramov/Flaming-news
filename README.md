# Flaming News Generator — Fairmont Baku

Gündəlik "Flaming News" hesabatını sürətləndirmək üçün veb-app. HTML editor
əsl Excel şablonunun (Sheet 1 "Accomodation" + Sheet 2 "F&B Section")
strukturunu, sətir/sütunlarını və VIP rəng kodlamasını birə-bir əks etdirir,
bütün xanalar birbaşa brauzerdə redaktə olunur, və "Download Excel" düyməsi
eyni strukturda .xlsx faylı yaradır.

## Necə işləyir

```
Əvvəlki günün datası (avtomatik yüklənir)
        ↓
Bugünkü hesabatları yüklə (GIH, Departures, Outlet revenues)
        ↓
Data avtomatik yenilənir (VIP qaydaları tətbiq olunur)
        ↓
HTML-də lazım olan yerləri redaktə et
        ↓
Excel-i endir (Save düyməsi + Download Excel)
```

Tətbiqi hər gün açanda, əgər həmin gün üçün hələ heç nə saxlanmayıbsa, ən son
əvvəlki günün datası avtomatik yüklənir (MOD adları, hava, All Enrollments,
Fairmont Goals, Events/Birthday/Anniversary, Quote of the Day və s. saxlanılır)
— yalnız VIP qonaq siyahıları, Site Inspections və F&B Performance təzədən
başlayır ki, günün hesabatlarından təzə doldurulsun. Beləliklə hər gün
sıfırdan yazmaq lazım deyil, yalnız fərqli olanı dəyişmək kifayətdir.

## VIP qaydaları (avtomatik tətbiq olunur)
- **DV** qonaqlar heç vaxt göstərilmir.
- **Booking.com** mənbəli **V1** qonaqlar VIP kimi göstərilmir (əsl VIP deyil).
- Yalnız əsl VIP kodları saxlanılır: T3 / T4 / T5 / T6 / SA / V1.
- Kod xanası rəngi avtomatik: T3 qara, T4 qırmızı, T5 ağ, T6 qara, SA yaşıl,
  V1 qırmızı — orijinal Excel şablonundakı rənglərlə eynidir.
- VIP Code Legend bölməsi silinib (tələbə uyğun).

## Yerli test (lokal)
```
pip install -r requirements.txt
python app.py
```
Sonra brauzerdə: http://localhost:5000

## Render.com-da deploy
1. Bu qovluğu GitHub repo-suna yükləyin
2. Render.com > New > Web Service > repo-nu seçin
3. Build Command: `pip install -r requirements.txt`
4. Start Command: `gunicorn app:app`
   (və ya `render.yaml` faylı avtomatik tanınacaq — "Blueprint" olaraq import edə bilərsiniz)
5. Deploy edin

**Qeyd:** `data/` qovluğu (gündəlik JSON faylları) Render-in fayl sistemi
"ephemeral" olduğu üçün restart-da silinə bilər — uzunmüddətli istifadə üçün
bir gün kalıcı disk (Render Disk) və ya verilənlər bazası qoşmaq tövsiyə olunur.

## Fayl strukturu
- `app.py` — Flask marşrutları: `/` (HTML editor), `/api/save`, `/api/upload/*`, `/api/export`
- `data_store.py` — gündəlik JSON saxlama + əvvəlki günün datasının avtomatik yüklənməsi
- `parsers.py` — hesabat parserləri: Arrivals / GIH (In-House) / Departures (PDF, regex əsaslı,
  pdfplumber ilə) və Outlet revenues (.xlsx, sütun adına görə)
- `rules.py` — VIP kod biznes qaydaları (filtr + Remarks + rəng)
- `excel_writer.py` — nəticə Excel faylını orijinal şablonun strukturuna görə qurur (openpyxl)
- `templates/index.html` — tam redaktə olunan HTML editor (hər iki vərəq, yükləmə düymələri)

## Hazırkı vəziyyət / bilinən məhdudiyyətlər
- Arrivals/Departures/GIH PDF parserləri əvvəlki versiyadan qalıb — PMS export
  formatı dəyişərsə, `parsers.py`-dəki regex-lər yenidən kalibrlənməli ola bilər.
- Outlet revenues (.xlsx) parseri real nümunə faylı ilə test olunub və sütun
  başlıqlarına görə uyğunlaşır (Outlet name / In house guests / Outside guests / Revenue).
- SA üçün "Specials" bəzən tam söz (Birthday), bəzən qısaldılmış kod (BD) formatında gəlir
  — hər ikisi dəstəklənir, yeni kodlar aşkar olunsa `rules.py`-də `SA_ABBREV` lüğətinə əlavə edin.

Test etdikcə problemli sətirləri mənə göndərin, parserləri dəqiqləşdirim.
