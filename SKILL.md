---
name: bumn-report-formatter
description: "Format transformation roadmaps, analytical frameworks, and project updates into PowerPoint (.pptx) decks using the 7 official Taskforce BP BUMN & Perum DAMRI slide templates — BUMN Navy #002060, headline blue #003C79, section blue #174EA6, with exact original BUMN & DAMRI slide logos, confidential footer, 2x2 prioritization matrix, analytical module cards, implementation pipeline, flow journey, timeline delivery Gantt, and closing slides."
version: "2.1.0"
tags:
  - bumn
  - damri
  - report
  - pptx
  - presentation
  - transformation
references:
  - references/brand.md
scripts:
  - scripts/bumn_deck.py
  - scripts/example_deck.py
  - scripts/rebuild_template.py
assets:
  - assets/template.pptx
  - assets/logos/bumn-logo.png
  - assets/logos/damri-logo.png
  - assets/backgrounds/cover-damri.png
  - assets/backgrounds/closing-damri.png
---

# BUMN & DAMRI Report Formatter

AI Agent Skill untuk memformat rencana transformasi, kerangka analitik, dan roadmap implementasi ke dalam deck PowerPoint (`.pptx`) resmi berstandar **Taskforce BP BUMN** dan **Perum DAMRI**.

Engine generator dibangun oleh `scripts/bumn_deck.py` di atas basis template resmi (`assets/template.pptx`), mereplikasi secara presisi **7 slide template** dari master presentation `Framework Task Force BUMN.pptx`.

---

## 7 Slide Templates BUMN

| Template | Metode Builder | Mengacu pada Master PPT | Deskripsi & Komponen |
|---|---|---|---|
| **Slide 1** | `cover(title, subtitle)` | Slide 1 Master | **Halaman Sampul**: Menggunakan foto asli armada bus DAMRI, logo asli BUMN dan DAMRI di sudut kanan atas, serta banner gelap di bagian bawah dengan judul dan subjudul. |
| **Slide 2** | `prioritization_matrix(quadrants, insights, ...)` | Slide 2 Master | **Matriks Prioritas Impact vs Effort**: Matriks 2x2 (4 kuadran: *Quick wins*, *Major projects*, *Low-hanging fruit*, *Hygiene factors*) dilengkapi panel kanan *Key insights* dan footer konfidensial master. |
| **Slide 3** | `quick_wins_modules(cards, preview_title, ...)` | Slide 3 Master | **Modul Cepat & Rilis 1**: 4 kartu modul analitis (2x2 grid) berisi *DATA NEEDED* dan *OUTCOME*, bar milestone rilis (*Beta* & *Prototype*), serta frame preview dashboard di sebelah kanan. |
| **Slide 4** | `implementation_pipeline(stages, ...)` | Slide 4 Master | **Arsitektur Implementasi & Alur Data**: Pipeline 3 tahap (*Data Sourcing & Pipeline* → *Modelling & Optimization* → *Recommendation*) dengan konektor panah. |
| **Slide 5** | `flow_journey(steps, ...)` | Slide 5 Master | **Flow Journey**: Alur perjalanan operasional bertahap (*Customer Intent* → *Dynamic Fare Engine* → *Seat Allocation* → *Dispatch* → *Feedback*). |
| **Slide 6** | `timeline_delivery(phases, tasks, ...)` | Slide 6 Master | **Timeline Delivery Gantt**: Pita fase (*Discover*, *Analyze*, *Build*, *Release*), daftar aktivitas, schedule bar, penanda milestone *"We are here"* merah, dan catatan kaki. |
| **Slide 7** | `closing(title, subtitle, contact)` | Slide 7 Master | **Halaman Penutup**: Foto asli armada bus DAMRI, logo asli BUMN dan DAMRI di sudut kanan atas, banner penutup bertuliskan "Terima kasih" dan kontak resmi. |

Setiap slide konten (Slide 2 s.d. Slide 6) mewarisi langsung dari master layout BUMN:
- **Logo BUMN**: Sudut kanan atas pada koordinat `(12.05", 0.22")`.
- **Garis Pembatas Bawah**: Di atas footer (`y=7.09"`).
- **Footer Konfidensial Resmi Tunggal** (tanpa duplikasi/tumpuk):
  > `CONFIDENTIAL AND PROPRIETARY | Any use of this material without specific permission of State-Owned Enterprises Regulatory Agency is strictly prohibited`
- **Nomor Halaman**: Sudut kanan bawah (`y=7.18"`).

---

## Contoh Penggunaan (Quick Start)

```python
import sys
sys.path.insert(0, "<skill-dir>/scripts")
from bumn_deck import BUMNDeck

deck = BUMNDeck()

# Slide 1: Sampul
deck.cover(
    title="Taskforce BP BUMN - Perum Damri",
    subtitle="AI for Public Services: DAMRI's Transformation Roadmap for Public Transportation"
)

# Slide 2: Matriks Prioritas Impact vs Effort
deck.prioritization_matrix(
    title="Prioritization Framework (Impact vs Effort)",
    section_label="Transformation Plan - Dynamic Pricing",
    badge="WIP",
    quadrants={
        "A": {
            "title": "Quick wins",
            "subtitle": "High impact · low effort",
            "count": 5,
            "items": ["**Load Factor Prediction**", "**Booking Curve Analytics**", "**Competitor Price Intelligence**"],
        },
        "B": {"title": "Major projects", "subtitle": "High impact · high effort", "count": 5, "items": ["Demand Forecasting Engine"]},
        "C": {"title": "Low-hanging fruit", "subtitle": "Low impact · low effort", "count": 2, "items": ["Customer Segmentation"]},
        "D": {"title": "Hygiene factors", "subtitle": "Low impact · high effort", "count": 3, "items": ["Scenario Simulator"]},
    }
)

# Slide 7: Penutup
deck.closing(title="Terima kasih", subtitle="Taskforce BP BUMN - Perum Damri", contact="info@damri.co.id")

deck.save("output_deck.pptx")
```
