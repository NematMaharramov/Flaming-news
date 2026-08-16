# Flaming News Generator — Fairmont Baku

PMS reportlarından (PDF) avtomatik "Flaming News" Excel faylı yaradan veb-app.

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

## Fayl strukturu
- `app.py` — Flask marşrutları (upload forma + /generate)
- `parsers.py` — 4 PDF report parseri (regex əsaslı, pdfplumber ilə)
- `rules.py` — VIP kod biznes qaydaları (Remarks + rəng)
- `excel_writer.py` — Nəticə Excel faylını qurur (openpyxl)
- `templates/index.html` — Upload forması

## Hazırkı vəziyyət / bilinən məhdudiyyətlər
- Arrivals: 14/26 sətir tapılır (~54%) — bəzi qonaqların otaq nömrəsi PDF-də oxunmur, belə hallarda "Room" xanası boş qalır və qırmızı işarələnir (əl ilə doldurulmalıdır)
- Departures: 26/29 sətir tapılır (~90%)
- In-House: 62/66 sətir tapılır (~94%)
- F&B Section hələ avtomatlaşdırılmayıb (statik boş vərəq kimi yaradılır)
- SA üçün "Specials" bəzən tam söz (Birthday), bəzən qısaldılmış kod (BD) formatında gəlir — hər ikisi dəstəklənir, amma yeni kodlar aşkar olunsa `rules.py`-də `SA_ABBREV` lüğətinə əlavə edilməlidir

Test etdikcə problemli sətirləri mənə göndərin, parserləri dəqiqləşdirim.
