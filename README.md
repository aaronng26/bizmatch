# Business Matchmaking Dashboard (Streamlit)

## Isi folder
- `app.py` — aplikasi Streamlit
- `requirements.txt` — library Python yang dibutuhkan
- `business_matchmaking_template.xlsx` — template data nasabah (sheet `guide`, `data`, `lists`)

## Cara menjalankan
1. Pastikan Python 3.9+ terpasang.
2. Buka terminal di folder ini, lalu jalankan:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Browser akan terbuka di http://localhost:8501. Upload file Excel/CSV data nasabah di sidebar.

## Format data yang diterima
Aplikasi membaca sheet bernama `data`, atau sheet pertama kalau tidak ada (misalnya `Sheet1`). Nama kolom dikenali otomatis (huruf besar-kecil, spasi, `-`, dan `_` diabaikan):

| Dipakai untuk | Nama kolom yang dikenali |
|---|---|
| ID nasabah (wajib) | `customerid` |
| Sektor utama (wajib) | `Sector`; kalau tidak ada, `economicsector_segment` (kode di depan seperti `1600000000-` dibuang) |
| Punya PT (wajib) | `has_pt_ocbc` (Yes/No atau Y/N) |
| Chain (wajib) | `Upstream (sectors)` / `upstream`, `Mid-stream (sectors)` / `midstream`, `Downstream (sectors)` / `downstream`, `Enabling` |
| Opsional | `customername`, `customercompany`, `subsegment`, `region`, `area`, `officename`, `value_chain_supp(lier)`, `value_chain_buyer`, `sto`, `sto_annum`, `aum_indiv`, `avrg_trx_3m`, `total_PT`, `economicsector_subsubsegment2` (produk) |

Isi kolom chain:
- Sektor boleh dipisah `,` atau `;`.
- `0`, `#N/A`, `-`, dan sel kosong dianggap kosong.
- Ejaan berbeda untuk sektor yang sama disatukan otomatis, misalnya `FOOD & BEVERAGES` → `FOOD & BEVERAGE`, `FINANCIAL INSTITUTUION` → `FINANCIAL INSTITUTION`, `LOGISTICS & TRANSPORTATION` → `LOGISTIC & TRANSPORTATION`.
- Nama yang ditampilkan adalah ejaan yang paling sering dipakai di kolom sektor utama.

`customercompany` boleh dikosongkan; nasabah tetap dihitung punya PT kalau `has_pt_ocbc` = Yes.

`business_matchmaking_template.xlsx` adalah template format lama, dan tetap bisa dipakai.

## Mengubah aturan matchmaking
Bagian `Settings` di atas `app.py`:
- `TOP_PER_RELATION` — jumlah kandidat per hubungan di Match results (default 5)
- `TOP_IN_DETAIL` — jumlah kandidat di Customer detail (default 20)
- `SAME_REGION_BOOST` — bobot untuk partner di region yang sama (default 1.5)
- `LIST_POSITION_PENALTY` — penurunan skor untuk sektor yang disebut belakangan (default 6% per urutan)
- `KNOWN_SECTORS` — daftar sektor yang dikenali
- `SECTOR_MATCH_CUTOFF` — seberapa mirip ejaan harus sama untuk dianggap sektor yang sama (default 0.86)
- `COLUMN_ALIASES` — nama kolom yang dikenali, kalau data aslimu memakai nama lain
