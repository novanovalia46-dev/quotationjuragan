# Database Barang & Quotation (Frontend Vercel)

Tampilan di Vercel. Data tetap di Google Sheets lewat Apps Script.

## File yang harus ada di root repo GitHub

```
index.html
api/gas.js
vercel.json
.gitignore
README.md
.env.example
```

Jangan upload `Code.gs` ke GitHub. Itu untuk Apps Script.

## Cara pakai

1. Upload **semua file di atas** ke root repo (bukan di dalam subfolder).
2. Tunggu Vercel selesai deploy.
3. Buka situs → tempel URL Web App yang berakhiran `/exec` di kotak kuning → Simpan.

URL `/exec` dari Apps Script → Deploy → Manage deployments → Copy URL.

Deploy Apps Script: **Execute as Me**, **Who has access = Anyone**.
