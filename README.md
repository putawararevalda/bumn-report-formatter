# BUMN Report Formatter

AI Agent Skill untuk memformat laporan transformasi, rencana kerja, audit, dan update proyek menjadi presentasi PowerPoint (`.pptx`) resmi berstandar **Taskforce BP BUMN (Badan Pengelola BUMN)** dan **Perum DAMRI**.

Terinspirasi dari arsitektur [aix-report-formatter](https://github.com/putawararevalda/aix-report-formatter), skill ini kompatibel langsung dengan **Antigravity**, **VoltAgent Workspace**, **Claude**, dan **Multica**.

---

## Struktur Folder

```text
bumn-report-formatter/
├── SKILL.md                          ← Definisi skill & instruksi agent (Antigravity & VoltAgent entrypoint)
├── README.md                         ← Panduan dokumentasi lengkap ini
├── assets/
│   ├── template.pptx                 ← Master template PowerPoint BUMN (terkompresi & layout bersih)
│   ├── backgrounds/
│   │   ├── cover-damri.jpg           ← Latar foto armada bus DAMRI untuk cover
│   │   └── closing-damri.jpg         ← Latar foto armada bus DAMRI untuk closing
│   └── logos/
│       ├── bumn-primary.png          ← Logo resmi BUMN (Garuda / Bhinneka Tunggal Ika)
│       ├── bumn-white.png            ← Logo BUMN versi putih untuk latar gelap
│       ├── damri-primary.png         ← Logo resmi Perum DAMRI
│       ├── damri-white.png           ← Logo DAMRI versi putih untuk latar gelap
│       ├── danantara-primary.png     ← Logo Danantara Indonesia
│       └── telkom-indonesia-logo.png ← Logo Telkom Indonesia
├── references/
│   └── brand.md                      ← Panduan warna BUMN, tipografi Arial, dan layout grid
├── scripts/
│   ├── bumn_deck.py                  ← Engine builder PowerPoint (python-pptx)
│   ├── example_deck.py               ← Generator deck contoh lengkap (13 jenis slide)
│   └── rebuild_template.py           ← Rebuilder template jika ada pembaruan master .pptx
├── fonts/
│   └── README.md                     ← Panduan font (Arial universal)
└── examples/
    └── BUMN_Sample_Transformation_Deck.pptx ← File deck PowerPoint hasil generate lengkap
```

---

## Cara Instalasi / Konsumsi Skill

### 1. Di Google Antigravity (Otomatis & Global)
Salin atau tautkan folder skill ini ke direktori konfigurasi global Antigravity:
```bash
# Salin ke folder konfigurasi global
cp -r bumn-report-formatter ~/.gemini/config/skills/

# Atau buat symlink agar perubahan di repo langsung sinkron
ln -s "$(pwd)/bumn-report-formatter" ~/.gemini/config/skills/bumn-report-formatter
```
Antigravity akan otomatis mendeteksi skill `bumn-report-formatter` dan memuatnya saat Anda meminta pembuatan deck atau laporan bertema BUMN/DAMRI.

### 2. Di VoltAgent Workspace
Tambahkan folder skill ini ke direktori root skills workspace VoltAgent Anda:
```ts
import { Agent, Workspace } from "@voltagent/core";

const workspace = new Workspace({
  skills: {
    rootPaths: ["/skills"],
  },
});
```

### 3. Di Multica (Import via Git URL)
1. Buka workspace Multica Anda.
2. Masuk ke menu **Skills** > pilih **New skill** > **Import from URL**.
3. Masukkan URL repository GitHub ini.

### 4. Di Claude (Desktop / Web)
Arsipkan folder ini menjadi file `.skill` (zip format) dan unggah melalui menu **Settings → Capabilities → Skills** pada Claude.

---

## Katalog Tipe Slide & Builder API

| Metode | Parameter Utama | Kegunaan |
|---|---|---|
| `cover(...)` | `title, subtitle, meta, partner_logo` | Halaman sampul berlatar armada bus DAMRI & logo BUMN/DAMRI |
| `toc(...)` | `items, title, section` | Daftar isi atau agenda dalam 1 atau 2 kolom |
| `section(...)` | `number, title, subtitle, section` | Halaman pemisah bab utama berlatar navy corporate |
| `matrix_2x2(...)` | `quadrants, insights, x_label, y_label` | **Matriks Prioritas Impact vs Effort** + sidebar Key Insights |
| `module_cards(...)` | `cards, preview_image, release_info` | **Kartu Modul Analitis 2x2 Grid** + bar milestone rilis & frame dashboard |
| `architecture_flow(...)`| `stages, title, section` | **Pipeline Arsitektur & Alur Data** (Sources → Model → Delivery) |
| `timeline_gantt(...)` | `phases, tasks, current_point, note` | **Timeline Delivery Gantt** + pita fase & penanda "We are here" |
| `kpis(...)` | `items, title, section` | 2–4 kartu KPI angka besar dengan tren & deskripsi |
| `table(...)` | `headers, rows, col_widths` | Tabel data dengan navy header dan baris selang-seling |
| `chart(...)` | `kind, categories, series` | Grafik native (bar, barh, stacked, line, doughnut) |
| `two_column(...)` | `title, left, right` | Perbandingan dua kondisi (*As-Is vs To-Be*) |
| `key_finding(...)` | `headline, statement, body, number` | Penekanan satu temuan strategis berdampak besar |
| `closing(...)` | `title, subtitle, contact` | Halaman penutup ("Terima kasih") dengan kontak resmi |

---

## Contoh Penggunaan Cepat (Quick Start)

```python
import sys
sys.path.insert(0, "bumn-report-formatter/scripts")
from bumn_deck import BUMNDeck

deck = BUMNDeck(lang="id", footer="Taskforce BP BUMN · Perum DAMRI")

# 1. Sampul
deck.cover(
    title="Taskforce BP BUMN - Perum Damri",
    subtitle="AI for Public Services: DAMRI's Transformation Roadmap",
    meta="Desember 2026 · AI & Analytics Transformation"
)

# 2. Matriks Prioritas Impact vs Effort
deck.matrix_2x2(
    title="Prioritization Framework (Impact vs Effort)",
    section="Transformation Plan - Dynamic Pricing",
    badge="WIP",
    quadrants={
        "A": {
            "title": "Quick wins",
            "subtitle": "High impact · low effort",
            "count": 5,
            "items": [
                "**Load Factor Prediction**",
                "**Booking Curve Analytics**",
                "**Competitor Price Intelligence**",
                "Promotion Optimization",
                "Monitoring Dashboard",
            ],
        },
        "B": {
            "title": "Major projects",
            "subtitle": "High impact · high effort",
            "count": 5,
            "items": ["Demand Forecasting", "Optimal Fare Engine"],
        },
        "C": {"title": "Low-hanging fruit", "subtitle": "Low impact · low effort", "count": 2, "items": ["Segmentation"]},
        "D": {"title": "Hygiene factors", "subtitle": "Low impact · high effort", "count": 3, "items": ["Simulator"]},
    },
    takeaway="Quick Wins Fase 1 dapat dijalankan langsung dengan data reservasi historis."
)

# 3. Penutup
deck.closing(
    title="Terima kasih",
    subtitle="Satgas Transformasi Digital BP BUMN & Perum DAMRI",
    contact="info@damri.co.id"
)

deck.save("output_laporan_damri.pptx")
```

Jalankan generator deck contoh:
```bash
python3 bumn-report-formatter/scripts/example_deck.py
```
Output tersimpan di `bumn-report-formatter/examples/BUMN_Sample_Transformation_Deck.pptx`.
