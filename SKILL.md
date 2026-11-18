---
name: bumn-report-formatter
description: "Format research reports, transformation plans, AI roadmaps, audit findings, and project updates into PowerPoint (.pptx) decks in the official Taskforce BP BUMN & Perum DAMRI branding — BUMN Navy #002060, headline blue #003C79, accent blue #174EA6, with official BUMN Garuda and DAMRI logos. Use whenever the user wants a report, roadmap presentation, use case prioritization, or update deck for BUMN, BP BUMN, Danantara, or Perum DAMRI initiatives. Supports 2x2 prioritization matrix, analytical module cards, architecture data pipelines, multi-phase delivery Gantt roadmaps, KPI tiles, data tables, native charts, two-column comparisons, key findings, and closing slides. Supports Bahasa Indonesia (default) and English."
version: "1.0.0"
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
  - assets/logos/bumn-primary.png
  - assets/logos/bumn-white.png
  - assets/logos/damri-primary.png
  - assets/logos/damri-white.png
  - assets/logos/danantara-primary.png
  - assets/logos/telkom-indonesia-logo.png
---

# BUMN & DAMRI Report Formatter

AI Agent Skill untuk memformat laporan transformasi digital, pemetaan use case, rencana kerja, dan readout hasil temuan ke dalam deck PowerPoint (`.pptx`) berstandar resmi **Taskforce BP BUMN (Badan Pengelola BUMN)** dan **Perum DAMRI**.

Engine generator dibangun oleh `scripts/bumn_deck.py` di atas basis template resmi (`assets/template.pptx`), menyajikan layout native PowerPoint (shapes, tables, charts) yang 100% editable tanpa bergantung pada image export statis.

---

## Alur Kerja (Workflow)

1. **Analisis Bahan Sumber**:
   Pelajari materi sumber (notulensi rapat, dokumen framework transformasi, tabel use case, roadmap). Susun alur narasi presentasi:
   `Cover → Agenda / TOC → Section Divider → Evidence / Analysis Slides (Matrix, Cards, Architecture, Gantt) → KPIs / Data Table → Closing`.
2. **Pilih Tipe Slide yang Tepat**:
   Gunakan katalog slide di bawah. Hindari slide dengan teks monoton panjang; prioritaskan visualisasi matriks, kartu analitis, pipeline arsitektur, atau timeline Gantt.
3. **Tulis Script Generator Python**:
   ```python
   import sys
   sys.path.insert(0, "<skill-dir>/scripts")
   from bumn_deck import BUMNDeck

   deck = BUMNDeck(lang="id", footer="Taskforce BP BUMN · Perum DAMRI")
   deck.cover(
       title="Taskforce BP BUMN - Perum Damri",
       subtitle="AI for Public Services: DAMRI's Transformation Roadmap",
       meta="Q4 2026 · AI & Analytics Transformation"
   )
   ...
   deck.save("output_deck.pptx")
   ```
4. **Verifikasi & Deliver**:
   Jalankan script generator dan serahkan file `.pptx`.

---

## Katalog Tipe Slide (Slide Catalogue)

| Metode | Berdasarkan Template | Kegunaan Utama |
|---|---|---|
| `cover(title, subtitle, meta, partner_logo)` | Slide 1 Master | Halaman sampul berlatar armada bus DAMRI, logo BUMN & DAMRI, dan navy banner bawah. |
| `toc(items, title, section)` | Agenda Standar | Daftar isi / agenda presentasi 4–8 modul dalam 1 atau 2 kolom. |
| `section(number, title, subtitle, section)` | Divider Eksekutif | Pemisah bagian utama presentasi dengan nomor bab besar dan latar navy corporate. |
| `matrix_2x2(quadrants, insights, ...)` | Slide 2 Master | **Matriks Prioritas Impact vs Effort** (4 kuadran: Quick wins, Major projects, Low-hanging fruit, Hygiene factors) dilengkapi panel sidebar Key Insights di sebelah kanan. |
| `module_cards(cards, preview_image, ...)` | Slide 3 Master | **Kartu Modul Analitis 2x2 Grid** dengan spesifikasi data, outcome, bottom Release milestone bar (Beta & Prototype), dan frame preview dashboard. |
| `architecture_flow(stages, ...)` | Slide 4 Master | **Pipeline Arsitektur & Alur Data** multi-tahap (Data Ingestion → AI/ML Engine → Business Action/Deployment). |
| `timeline_gantt(phases, tasks, ...)` | Slide 6 Master | **Timeline & Delivery Roadmap Gantt** dengan pita fase (Discover, Analyze, Build, Release), bar jadwal, dan penanda milestone *"We are here"*. |
| `kpis(items, ...)` | Executive Metric Tiles | 2–4 kartu metrik angka besar (Nilai, Satuan, Label, Pill Tren Naik/Turun, Deskripsi dampak). |
| `table(headers, rows, ...)` | Corporate Data Table | Rekapitulasi tabel data dengan navy header, baris bergantian, dan penyesuaian lebar kolom. |
| `chart(kind, categories, series, ...)` | Native Charts | Grafik native PowerPoint: kolom (`bar`), horizontal (`barh`), stacked (`stacked`), garis (`line`), atau donat (`doughnut`). |
| `two_column(title, left, right, ...)` | Perbandingan Dual | Perbandingan dua pilar atau dua kondisi (*Before vs After*, *As-Is vs To-Be*). |
| `key_finding(headline, statement, body, ...)` | Strategic Callout | Penekanan satu temuan strategis dengan statement berukuran besar dan narasi pendukung. |
| `closing(title, subtitle, contact, ...)` | Slide 7 Master | Halaman penutup resmi ("Terima kasih") dengan latar armada bus DAMRI dan kontak resmi. |

Semua metode konten mendukung parameter:
- `section="Label Seksi"`: Teks kecil di atas headline judul (`#174EA6`).
- `badge="WIP"` atau `"RELEASE 1"`: Pill status di sudut kanan atas slide (`#C00000`).
- `takeaway="Kalimat insight utama..."`: Banner aksen di bagian bawah slide bertuliskan *"Insight Utama / Key Takeaway"*.
- Format teks `**teks tebal**` secara otomatis diubah menjadi tulisan tebal (bold run).

---

## Panduan Visual & Branding (Design Rules)

Rincian lengkap terdapat di `references/brand.md`. Pedoman esensial:

1. **Warna Utama**:
   - **BUMN Navy** `#002060`: Warna utama institusional untuk header tabel, badge prioritas, dan pill penomoran.
   - **Headline Blue** `#003C79`: Warna judul slide utama (18–20pt).
   - **Section Accent Blue** `#174EA6`: Warna label seksi di atas judul utama (13–14pt).
   - **Card Active** `#E6EDF6`: Latar kartu prioritas tinggi / Quick Wins.
   - **Card Neutral** `#F3F5F8`: Latar kartu standar dan panel sidebar.
   - **Milestone Red** `#C00000`: Digunakan secara hemat untuk penanda kritis (*We are here*, badge *WIP*).
2. **Tipografi**:
   - Seluruh teks menggunakan font universal **Arial** untuk memastikan presentasi terbuka sempurna di komputer stakeholder tanpa isu font hilang atau text wrap rusak.
3. **Hirarki Informasi**:
   - Satu pesan utama per slide.
   - Selalu berikan konteks pada slide analitis menggunakan parameter `takeaway=...` di bagian bawah.
