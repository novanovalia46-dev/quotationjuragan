# Database Barang & Quotation (Frontend Vercel)

Tampilan di Vercel. Data tetap di Google Sheets lewat Apps Script.

## File di root repo GitHub

```
index.html
api/gas.js
vercel.json
.gitignore
README.md
.env.example
```

Jangan upload `Code.gs` ke GitHub.

## Supaya kotak kuning tidak muncul (semua akun)

1. Vercel → Project → **Settings** → **Environment Variables**
2. **Key:** `GAS_WEB_APP_URL`
3. **Value:** URL Web App Apps Script yang berakhiran `/exec`
   (contoh `https://script.google.com/macros/s/AKfycb..../exec`)
4. Environment: Production, Preview, Development → Save
5. **Deployments** → ⋮ pada deployment terbaru → **Redeploy**
   (Save env saja tidak cukup; wajib Redeploy)

Setelah itu semua orang buka situs tanpa menempel URL.
