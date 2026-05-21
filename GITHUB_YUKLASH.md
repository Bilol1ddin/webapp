# GitHub Pages ga yuklash yo'riqnomasi

## GitHub ga yuklash kerak bo'lgan fayllar:
1. webapp.html   ← Dashboard (shu papkada)
2. users.json    ← Avtomatik yaratiladi (bot ishlangandan keyin)
3. davomat.json  ← Avtomatik yaratiladi
4. analitika.json← Avtomatik yaratiladi

## Bosqichlar:

### 1. GitHub repoga kiring
https://github.com/bilol1ddin/webapp

### 2. webapp.html ni yangilang
- "webapp.html" faylini bosing
- Qalamcha (✏️) tugmasini bosing
- Barchasini o'chirib, yangi kodni joylashtiring
- "Commit changes" bosing

### 3. JSON fayllarni yuklang
Bot ishlagandan keyin users.json, davomat.json, analitika.json fayllari
bot.py bilan bir papkada paydo bo'ladi.
Shu fayllarni ham GitHub repoga yuklang:
- "Add file" → "Upload files" → fayllarni tortib tashlang → Commit

### 4. GitHub Actions (avtomatik yangilash - ixtiyoriy)
Bot JSON fayllarni o'zgartirgan sayin GitHub ga push qilish uchun
bot.py ga git push qo'shish mumkin (keyinroq).

## Muhim!
webapp.html va JSON fayllar BIR repoda bo'lishi kerak:
https://bilol1ddin.github.io/webapp/webapp.html  ← Dashboard
https://bilol1ddin.github.io/webapp/users.json   ← Ma'lumot
https://bilol1ddin.github.io/webapp/davomat.json ← Ma'lumot

## bot.py da URL to'g'ri:
WEBAPP_URL = "https://bilol1ddin.github.io/webapp/"
