# BUMN Report Formatter

AI Agent Skill untuk memformat rencana transformasi, kerangka analitik, dan roadmap implementasi menjadi presentasi PowerPoint (`.pptx`) resmi berstandar **Taskforce BP BUMN (Badan Pengelola BUMN)** dan **Perum DAMRI**.

Format dan tata letak dirancang khusus mereplikasi secara presisi **7 slide template** dari master presentation `Framework Task Force BUMN.pptx`.

---

## 7 Slide Templates BUMN

| Slide | Metode Builder / Alias | Komponen & Fitur |
|---|---|---|
| **Slide 1** | `cover(...)` / `slide_1_cover(...)` | Background foto asli armada bus DAMRI, logo asli BUMN dan DAMRI di sudut kanan atas, serta banner bawah dengan judul dan subjudul |
| **Slide 2** | `prioritization_matrix(...)` / `slide_2_matrix(...)` | Matriks Prioritas Impact vs Effort 2x2 (4 kuadran: *Quick wins*, *Major projects*, *Low-hanging fruit*, *Hygiene factors*) + panel sidebar *Key insights* |
| **Slide 3** | `quick_wins_modules(...)` / `slide_3_modules(...)` | 4 Kartu Modul Analitis 2x2 (*DATA NEEDED*, *OUTCOME*) + bar milestone rilis (*Beta* & *Prototype*) + frame preview dashboard |
| **Slide 4** | `implementation_pipeline(...)` / `slide_4_architecture(...)` | Pipeline Implementasi 3 tahap (*Data Sourcing & Pipeline* → *Modelling & Optimization* → *Recommendation*) dengan panah penghubung |
| **Slide 5** | `flow_journey(...)` / `slide_5_journey(...)` | Alur perjalanan operasional bertahap (*Customer Intent* → *Dynamic Fare Engine* → *Seat Allocation* → *Dispatch* → *Feedback*) |
| **Slide 6** | `timeline_delivery(...)` / `slide_6_timeline(...)` | Timeline Delivery Gantt: pita fase (*Discover*, *Analyze*, *Build*, *Release*), daftar tugas, schedule bar, penanda *"We are here"* merah |
| **Slide 7** | `closing(...)` / `slide_7_closing(...)` | Halaman penutup resmi berlatar foto asli armada bus DAMRI, logo asli BUMN dan DAMRI, banner "Terima kasih", dan kontak resmi |

Setiap slide konten (Slide 2 s.d. Slide 6) memuat header resmi, dan mewarisi secara native dari master layout BUMN:
- Logo BUMN di kanan atas
- Garis pemisah bawah
- Nomor slide otomatis
- **Footer konfidensial tunggal (tanpa tumpuk)**:
  > `CONFIDENTIAL AND PROPRIETARY | Any use of this material without specific permission of State-Owned Enterprises Regulatory Agency is strictly prohibited`

---

## Struktur Folder

```text
bumn-report-formatter/
├── SKILL.md                          ← Definisi instruksi skill untuk AI Agent
├── README.md                         ← Dokumentasi lengkap ini
├── assets/
│   ├── template.pptx                 ← Master template BUMN (666 KB)
│   ├── backgrounds/
│   │   ├── cover-damri.png           ← Foto asli armada DAMRI untuk sampul (1672x941 px)
│   │   └── closing-damri.png         ← Foto asli armada DAMRI untuk penutup (1672x941 px)
│   └── logos/
│       ├── bumn-logo.png             ← Logo asli BUMN (Picture 6 dari master)
│       └── damri-logo.png            ← Logo asli DAMRI (Picture 7 dari master)
├── references/
│   └── brand.md                      ← Panduan identitas visual BUMN & DAMRI
├── scripts/
│   ├── bumn_deck.py                  ← Core PowerPoint builder engine (7 slide templates)
│   ├── example_deck.py               ← Generator deck contoh lengkap (7 slide Lorem Ipsum)
│   └── rebuild_template.py           ← Rebuilder template
├── fonts/
│   └── README.md                     ← Panduan tipografi (Arial universal)
└── examples/
    └── BUMN_Sample_Transformation_Deck.pptx ← File deck hasil generate (7 slide lengkap)
```

---

## Contoh Penggunaan (Quick Start)

```python
import sys
sys.path.insert(0, "bumn-report-formatter/scripts")
from bumn_deck import BUMNDeck

deck = BUMNDeck()

# 1. Slide 1: Sampul
deck.cover(
    title="Taskforce BP BUMN - Perum Damri",
    subtitle="AI for Public Services: DAMRI's Transformation Roadmap for Public Transportation"
)

# 2. Slide 2: Matriks Prioritas Impact vs Effort
deck.prioritization_matrix(
    title="Prioritization Framework (Impact vs Effort)",
    section_label="Transformation Plan - Dynamic Pricing",
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
                "**Promotion Optimization**",
                "**Monitoring & Alerting Dashboard**",
            ],
        },
        "B": {
            "title": "Major projects",
            "subtitle": "High impact · high effort",
            "count": 5,
            "items": ["Demand Forecasting Engine", "Price Elasticity Engine"],
        },
        "C": {"title": "Low-hanging fruit", "subtitle": "Low impact · low effort", "count": 2, "items": ["Customer Segmentation"]},
        "D": {"title": "Hygiene factors", "subtitle": "Low impact · high effort", "count": 3, "items": ["Scenario Simulator"]},
    }
)

# 3. Slide 7: Penutup
deck.closing(
    title="Terima kasih",
    subtitle="Taskforce BP BUMN - Perum Damri",
    contact="info@damri.co.id"
)

deck.save("output_laporan_damri.pptx")
```

Jalankan generator deck contoh:
```bash
python3 bumn-report-formatter/scripts/example_deck.py
```
Hasil presentasi tersimpan di `bumn-report-formatter/examples/BUMN_Sample_Transformation_Deck.pptx`.
