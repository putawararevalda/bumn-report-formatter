# Brand & Design Guidelines — BUMN Task Force / Perum DAMRI

Panduan identitas visual dan standar desain presentasi untuk deck Taskforce BUMN / BP BUMN (Badan Pengelola BUMN) dan Perum DAMRI.

---

## 1. Identitas Visual & Prinsip Desain

Deck BUMN dirancang untuk laporan eksekutif, pembacaan hasil temuan (readout), rencana transformasi digital, dan roadmap implementasi AI/ML. Karakter visualnya:

- **Authoritative & Corporate**: Dominasi warna BUMN Navy (`#002060`, `#003C79`) yang profesional, terpercaya, dan kredibel.
- **Data-Dense yet Clear**: Menampung informasi analitis yang padat (matriks 2x2, kartu arsitektur, kartu modul, timeline multi-fase) dengan penataan hirarki yang rapi dan ruang bernapas yang cukup.
- **Consistent Top Header**: Setiap slide konten memiliki *section label* (`#174EA6`) dan *headline* (`#003C79`) di koordinat yang presisi.
- **Universal Typography**: Menggunakan font **Arial** yang tersedia secara native di seluruh sistem operasi (macOS, Windows, Linux) sehingga rendering konsisten tanpa risiko font missing atau text clipping.

---

## 2. Palet Warna (Color Palette)

### Warna Utama (Primary Brand Colors)
| Nama Warna | Hex | RGB | Penggunaan |
|---|---|---|---|
| **BUMN Navy** | `#002060` | `0, 32, 96` | Header tabel, badge prioritas, pill aksen, garis aksen |
| **Headline Blue** | `#003C79` | `0, 60, 121` | Judul utama slide (Headline text) |
| **Section Accent Blue** | `#174EA6` | `23, 78, 166` | Label seksi atas (Section label), link, pilar aktif |
| **Slate Navy** | `#123C79` | `18, 60, 121` | Sub-header modul, judul tahapan timeline |
| **Deep Slate** | `#0E2841` | `14, 40, 65` | Teks judul kartu, aktivitas timeline |

### Warna Netral & Kartu (Neutrals & Fills)
| Nama Warna | Hex | RGB | Penggunaan |
|---|---|---|---|
| **Card Active Tint** | `#E6EDF6` | `230, 237, 246` | Latar kartu prioritas tinggi (Quick wins), takeaway bar |
| **Card Neutral Fill** | `#F3F5F8` | `243, 245, 248` | Latar kartu standar, sidebar insights |
| **Phase Band Blue 1** | `#DCE6F2` | `220, 230, 242` | Pita fase genap pada timeline Gantt |
| **Phase Band Blue 2** | `#EEF2F7` | `238, 242, 247` | Pita fase ganjil pada timeline Gantt |
| **Pure White** | `#FFFFFF` | `255, 255, 255` | Latar belakang slide, teks pada kartu gelap |
| **Border / Rule Gray** | `#B9C7DE` | `185, 199, 222` | Garis pemisah, border kartu aktif |
| **Hairline Light** | `#E7E6E6` | `231, 230, 230` | Garis tabel, divider halus |

### Tipografi & Teks
| Nama Warna | Hex | RGB | Penggunaan |
|---|---|---|---|
| **Body Dark** | `#1F2A37` | `31, 42, 55` | Teks isi utama (body copy), daftar poin |
| **Text Muted** | `#5B6675` | `91, 102, 117` | Keterangan sub-judul, data label, nomor fase |
| **Text Secondary** | `#595959` | `89, 89, 89` | Catatan kaki, sumber data |

### Warna Aksen & Indikator (Accents & Highlights)
| Nama Warna | Hex | RGB | Penggunaan |
|---|---|---|---|
| **Milestone / Alert Red** | `#C00000` | `192, 0, 0` | Indikator milestone "We are here", status badge WIP |
| **Success / Green** | `#1E7E34` | `30, 126, 52` | KPI trend positif, status selesai |
| **Warning / Amber** | `#D97706` | `217, 119, 6` | KPI trend waspada, status on-track |

---

## 3. Tipografi (Typography)

| Elemen | Font Family | Size | Style | Warna |
|---|---|---|---|---|
| **Judul Sampul (Cover Title)** | Arial | 24–28 pt | Bold | `#FFFFFF` |
| **Subjudul Sampul** | Arial | 14–16 pt | Regular | `#E6EDF6` |
| **Section Label** | Arial | 12–14 pt | Bold | `#174EA6` |
| **Headline Slide** | Arial | 18–20 pt | Bold | `#003C79` |
| **Judul Kartu / Modul** | Arial | 11–13 pt | Bold | `#002060` |
| **Body Copy** | Arial | 9.5–11 pt | Regular | `#1F2A37` |
| **Sub-keterangan / Data Label** | Arial | 8.5–9.5 pt | Bold/Caps | `#5B6675` |
| **KPI Big Number** | Arial | 24–36 pt | Bold | `#002060` |
| **Nomor Halaman** | Arial | 9–10 pt | Regular | `#5B6675` |

---

## 4. Layout & Grid Spacing

Dimensi slide menggunakan standar **16:9 Widescreen**:
- **Lebar Slide**: `13.333 inci` (`960 pt` / `12,192,000 EMU`)
- **Tinggi Slide**: `7.500 inci` (`540 pt` / `6,858,000 EMU`)

### Zona Tata Letak Standar (Header & Content Area)
```text
(0, 0) ───────────────────────────────────────────────────────────── (13.33", 0)
│                                                                    │
│  [Section Label] (x=0.40", y=0.25", w=8.33", h=0.22")   [Badge/WIP]│
│  [Slide Headline] (x=0.40", y=0.44", w=12.10", h=0.47")            │
│                                                                    │
│  ─── Content Area (y=1.00" sampai y=6.40" / y=6.90") ────────────  │
│  x=0.40"                                              w=12.53"     │
│                                                                    │
│  [Optional Takeaway Bar] (x=0.40", y=6.45", w=12.53", h=0.45")     │
│                                                   [Page Number]    │
(0, 7.50") ───────────────────────────────────────────────────────── (13.33", 7.50")
```

---

## 5. Logo & Lockup Partners

- **Logo Utama**:
  - `bumn-primary.png`: Logo resmi BUMN dengan Garuda & "BHINNEKA TUNGGAL IKA"
  - `damri-primary.png`: Logo resmi Perum DAMRI (biru & oranye)
  - `danantara-primary.png`: Logo Danantara Indonesia
  - `telkom-indonesia-logo.png`: Logo Telkom Indonesia
- **Versi Putih**:
  - `bumn-white.png`: Untuk diletakkan di atas background gelap/foto
  - `damri-white.png`: Untuk diletakkan di atas banner gelap
- **Posisi Header Sampul & Penutup**:
  - Logo ditempatkan di kanan atas slide cover:
    - BUMN: `x=9.99"`, `y=0.26"`, `w=1.10"`, `h=0.57"`
    - DAMRI: `x=11.30"`, `y=0.39"`, `w=1.83"`, `h=0.30"`
