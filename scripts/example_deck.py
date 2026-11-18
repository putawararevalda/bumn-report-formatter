"""
Example Deck Generator for BUMN Taskforce & DAMRI Presentation.

Builds a comprehensive 12-slide demonstration deck showcasing every
slide archetype in `bumn_deck.py`.

Usage:
    python3 example_deck.py [output.pptx]
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from bumn_deck import BUMNDeck


def generate_sample_deck(output_path: str = None) -> str:
    if not output_path:
        output_path = os.path.join(HERE, "..", "examples", "BUMN_Sample_Transformation_Deck.pptx")

    deck = BUMNDeck(lang="id", footer="Taskforce BP BUMN · Perum DAMRI")

    # 1. Cover
    deck.cover(
        title="Taskforce BP BUMN - Perum Damri",
        subtitle="AI for Public Services: DAMRI's Transformation Roadmap for Public Transportation",
        meta="Desember 2026 · AI & Analytics Transformation",
    )

    # 2. Table of Contents
    deck.toc(
        title="Agenda Transformasi Digital DAMRI",
        section="Daftar Isi & Roadmap",
        items=[
            {"num": "01", "title": "Prioritization Framework", "desc": "Matriks Impact vs Effort 15 use cases AI"},
            {"num": "02", "title": "Quick Wins & Output Release 1", "desc": "Detail modul analitis & target rilis Q4 2026"},
            {"num": "03", "title": "Arsitektur & Data Pipeline", "desc": "Alur integrasi data operasional ke AI model"},
            {"num": "04", "title": "Timeline & Delivery Roadmap", "desc": "Jadwal implementasi menuju peluncuran Dec W2"},
            {"num": "05", "title": "Target Kinerja & KPI Utama", "desc": "Metrik dampak pendapatan dan okupansi armada"},
            {"num": "06", "title": "Rekomendasi & Langkah Selanjutnya", "desc": "Keputusan strategis dan governance implementasi"},
        ],
    )

    # 3. Section Divider
    deck.section(
        number="01",
        title="Prioritas Program & Modul Cepat (Quick Wins)",
        subtitle="Pemetaan 15 use case analitik berbasis dampak pendapatan dan kemudahan eksekusi",
        section="Bagian 1",
    )

    # 4. Slide 2 Archetype: Impact vs Effort 2x2 Matrix
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
                    "**Promotion Optimization**",
                    "**Monitoring & Alerting Dashboard**",
                ],
            },
            "B": {
                "title": "Major projects",
                "subtitle": "High impact · high effort",
                "count": 5,
                "items": [
                    "Demand Forecasting Engine",
                    "Price Elasticity Model",
                    "Optimal Fare Recommender",
                    "Inventory & Seat Allocation",
                    "Real-Time Revenue Management",
                ],
            },
            "C": {
                "title": "Low-hanging fruit",
                "subtitle": "Low impact · low effort",
                "count": 2,
                "items": [
                    "Customer Segmentation",
                    "Pricing Governance & Rules",
                ],
            },
            "D": {
                "title": "Hygiene factors",
                "subtitle": "Low impact · high effort",
                "count": 3,
                "items": [
                    "Scenario Simulator",
                    "Ancillary Revenue Optimization",
                    "Cancellation & No-Show Prediction",
                ],
            },
        },
        insights=[
            {"badge": "A", "title": "Quick wins (Release 1)", "text": "Meningkatkan okupansi bus dan visibilitas harga kompetitor menggunakan data historis tanpa perlu integrasi sistem kompleks."},
            {"badge": "B", "title": "Major projects (Release 2)", "text": "Mesin penetapan tarif dinamis otomatis yang membutuhkan integrasi ke sistem ticketing dan otorisasi tarif."},
            {"badge": "C", "title": "Low-hanging fruit", "text": "Segmentasi pelanggan dan panduan tata kelola harga untuk mendukung tim komersial cabang."},
            {"badge": "D", "title": "Hygiene factors", "text": "Simulator skenario pendapatan dan mitigasi pembatalan tiket untuk stabilitas operasional jangka panjang."},
        ],
        takeaway="**15 use cases terbagi dalam 4 pool nilai**; Quick Wins Fase 1 dapat dijalankan langsung dengan data reservasi eksisting.",
    )

    # 5. Slide 3 Archetype: 2x2 Module Cards with Release Bar
    deck.module_cards(
        title="Quick Wins & Output Release 1",
        section="Transformation Plan - Dynamic Pricing",
        preview_title="05   Monitoring & alerting dashboard",
        preview_caption="Dashboard monitoring terpusat untuk memantau load factor dan pergerakan tarif rute perintis & komersial DAMRI",
        release_info={
            "release": "RELEASE 1",
            "beta": "Beta   W2 Nov 26",
            "prototype": "Prototype   W2 Dec 26",
        },
        cards=[
            {
                "num": "01",
                "title": "Load Factor Prediction",
                "objective": "Memprediksi tingkat okupansi per rute dan jadwal keberangkatan sejak dini.",
                "data": "Data reservasi 2–3 tahun terakhir, kapasitas terjadwal vs terealisasi, master rute, dan kalender libur nasional.",
                "outcome": "Kapasitas bus disesuaikan dengan pola permintaan aktual; kenaikan **+2–4 poin load factor** tanpa menambah armada.",
            },
            {
                "num": "02",
                "title": "Booking Curve Analytics",
                "objective": "Menetapkan kurva pemesanan normal per koridor dan mendeteksi anomali ritel.",
                "data": "Timestamp pemesanan (tanggal transaksi vs tanggal berangkat), data pembatalan tiket, dan kelas tarif.",
                "outcome": "Intervensi harga dapat dilakukan **30–60 hari sebelum keberangkatan**, saat elastisitas penumpang masih tinggi.",
            },
            {
                "num": "03",
                "title": "Competitor Price Intelligence",
                "objective": "Monitoring terstruktur pergerakan tarif operator bus kompetitor di platform OTA.",
                "data": "Tarif publik kompetitor via OTA scraping harian, tabel pencocokan koridor, dan histori tarif DAMRI.",
                "outcome": "Kesenjangan harga terdeteksi dalam hitungan hari; mencegah hilangnya potensi pendapatan di rute gemuk.",
            },
            {
                "num": "04",
                "title": "Promotion Optimization",
                "objective": "Mengarahkan alokasi diskon hanya ke jadwal dan rute dengan permintaan rendah.",
                "data": "Histori promo (diskon, kanal, periode), pemesanan harian, dan baseline kontrol keberangkatan normal.",
                "outcome": "Efisiensi biaya promosi dengan menghentikan subsidi bagi penumpang yang bersedia membayar tarif normal.",
            },
        ],
        takeaway="Empat modul analitis ini membentuk pondasi intelijen operasional DAMRI sebelum otomatisasi penetapan harga rilis berikutnya.",
    )

    # 6. Slide 4 Archetype: Architecture & Data Flow
    deck.architecture_flow(
        title="Implementation Architecture & Data Pipeline",
        section="Transformation Plan - Dynamic Pricing",
        stages=[
            {
                "num": "1",
                "title": "Data Sources",
                "category": "Operational Ingestion",
                "items": [
                    {"heading": "SIMA Operasi & LMB", "desc": "Data rit armada, keterisian kursi per rute, dan catatan kru bus."},
                    {"heading": "DAMRI Apps & Ticketing", "desc": "Log transaksi reservasi, waktu booking, dan metode pembayaran."},
                    {"heading": "OTA Market Intelligence", "desc": "Scraping berkala tarif kompetitor di rute antarkota antarprovinsi."},
                ],
            },
            {
                "num": "2",
                "title": "AI & ML Engine",
                "category": "Predictive Modeling",
                "items": [
                    {"heading": "Time-Series Forecasting", "desc": "Model estimasi permintaan penumpang per koridor dan musim."},
                    {"heading": "Booking Curve Anomaly", "desc": "Deteksi perlambatan penjualan tiket lebih dini dari pola historis."},
                    {"heading": "Dynamic Fare Optimizer", "desc": "Algoritma rekomendasi batas tarif atas dan promo berkala."},
                ],
            },
            {
                "num": "3",
                "title": "Business Delivery",
                "category": "Decision Support",
                "items": [
                    {"heading": "Executive Dashboard", "desc": "Visualisasi performa load factor dan revenue per cabang harian."},
                    {"heading": "Commercial Alerting", "desc": "Notifikasi otomatis ke manajer komersial untuk penyesuaian kuota."},
                    {"heading": "API Gateway Reservasi", "desc": "Penyediaan tarif optimal ke kanal penjualan tiket DAMRI."},
                ],
            },
        ],
        takeaway="Pemanfaatan data operasional internal DAMRI yang diperkaya intelijen pasar eksternal menghasilkan rekomendasi tarif presisi.",
    )

    # 7. Slide 6 Archetype: Delivery Timeline Gantt
    deck.timeline_gantt(
        title="Timeline delivery: Prototype launch in Dec W2",
        section="Transformation Plan - Dynamic Pricing",
        phases=[
            {"name": "Discover", "span": 2},
            {"name": "Analyze & design", "span": 2},
            {"name": "Build & test", "span": 2},
            {"name": "Release", "span": 2},
        ],
        tasks=[
            {"name": "Gather requirements & bus operational scope", "start": 0.05, "dur": 0.18},
            {"name": "Collect data (SIMA, DAMRI Apps, OTA scrape)", "start": 0.15, "dur": 0.22},
            {"name": "Data pattern analysis & corridor cleaning", "start": 0.32, "dur": 0.20},
            {"name": "Infrastructure & machine learning design", "start": 0.45, "dur": 0.20},
            {"name": "Model development & dashboard visualization", "start": 0.58, "dur": 0.22},
            {"name": "Rigorous backtesting & revenue simulation", "start": 0.70, "dur": 0.15},
            {"name": "Beta version release & stakeholder review", "start": 0.78, "dur": 0.12},
            {"name": "Launch prototype delivery in Dec W2", "start": 0.88, "dur": 0.10},
        ],
        current_point={"frac": 0.80, "label": "We are here"},
        note="Beta testing sedang berlangsung menjelang peluncuran prototipe rilis pertama pada Desember Minggu ke-2.",
        takeaway="Tahap pengembangan modul analitis berjalan sesuai target menuju peluncuran prototipe tepat waktu.",
    )

    # 8. Key Performance Indicators (KPIs)
    deck.kpis(
        title="Target Dampak Transformasi Terhadap Kinerja DAMRI",
        section="Target & Indikator Kinerja",
        items=[
            {
                "value": "+3.8%",
                "unit": "pts",
                "label": "Kenaikan Load Factor",
                "trend": "+3.8% MoM",
                "trend_type": "up",
                "desc": "Optimalisasi kapasitas pada koridor uji coba tanpa penambahan unit bus operasional.",
            },
            {
                "value": "12.4%",
                "unit": "Milyar",
                "label": "Potensi Kenaikan Revenue",
                "trend": "Proyeksi Tahunan",
                "trend_type": "up",
                "desc": "Tambahan pendapatan dari monetisasi kursi kosong pada periode libur Nataru dan Lebaran.",
            },
            {
                "value": "30",
                "unit": "Hari",
                "label": "Lead Time Intervensi",
                "trend": "Dini vs H-1",
                "trend_type": "up",
                "desc": "Penyesuaian promosi dan jadwal kini dapat diputuskan 30 hari sebelum keberangkatan.",
            },
            {
                "value": "94%",
                "unit": "Akurasi",
                "label": "Ketepatan Prediksi Okupansi",
                "trend": "Mean Absolute %",
                "trend_type": "up",
                "desc": "Model machine learning memprediksi keterisian armada per koridor dengan deviasi rendah.",
            },
        ],
        takeaway="Implementasi analitik memberikan dampak terukur langsung terhadap pendapatan dan efisiensi utilisasi armada bus DAMRI.",
    )

    # 9. Data Table
    deck.table(
        title="Ringkasan Use Case & Kesiapan Data DAMRI",
        section="Rencana Implementasi",
        headers=["No", "Nama Use Case", "Sumber Data Utama", "Tingkat Kompleksitas", "Estimasi Waktu", "Target Rilis"],
        rows=[
            ["01", "Load Factor Prediction", "SIMA Operasi, DRMS", "Sedang", "4 Minggu", "Release 1"],
            ["02", "Booking Curve Analytics", "DAMRI Apps, KAI Access", "Sedang", "4 Minggu", "Release 1"],
            ["03", "Competitor Price Intelligence", "OTA Scraping, Rute Master", "Rendah", "3 Minggu", "Release 1"],
            ["04", "Promotion Optimization", "Histori Promo, Tiket", "Sedang", "5 Minggu", "Release 1"],
            ["05", "Dynamic Fare Recommendation", "Model Elastisitas, OTA", "Tinggi", "8 Minggu", "Release 2"],
            ["06", "Predictive Fleet Maintenance", "DFMS IoT, Sensor Armada", "Tinggi", "10 Minggu", "Release 2"],
        ],
        col_widths=[0.6, 2.8, 2.5, 1.8, 1.6, 1.5],
        takeaway="Use case prioritas Release 1 memanfaatkan sumber data yang telah tersedia di sistem operasional DAMRI saat ini.",
    )

    # 10. Native Chart
    deck.chart(
        kind="bar",
        title="Simulasi Peningkatan Keterisian Kursi (Load Factor) per Koridor",
        section="Simulasi & Hasil Uji Coba",
        categories=["Bandung - Jakarta", "Surabaya - Malang", "Yogya - Semarang", "Lampung - Jakarta", "Bandara Soetta"],
        series=[
            {"name": "Baseline Historis (%)", "values": [68.2, 72.1, 64.5, 76.0, 70.4]},
            {"name": "Target dengan Dynamic Pricing (%)", "values": [74.5, 78.8, 71.0, 82.5, 77.2]},
        ],
        takeaway="Simulasi menunjukkan peningkatan konsisten okupansi 5–7% pada koridor antarkota dan angkutan bandara.",
    )

    # 11. Two-Column Comparison
    deck.two_column(
        title="Perbandingan Operasional: Eksisting vs Transformasi Digital",
        section="Manfaat Transformasi",
        left={
            "heading": "Kondisi Eksisting (As-Is)",
            "sub": "Karakteristik operasional konvensional",
            "items": [
                "Penetapan tarif bersifat **statis dan kaku**, jarang disesuaikan dengan dinamika pasar dan musim.",
                "Intervensi keterisian bus baru dilakukan saat **H-1 atau hari keberangkatan**, saat peluang promosi sudah lewat.",
                "Visibilitas harga kompetitor hanya diketahui melalui survei manual atau laporan tertulis lapangan yang lambat.",
                "Alokasi diskon dan promosi sering kali disamaratakan ke seluruh jadwal tanpa memperhitungkan kurva okupansi.",
            ],
        },
        right={
            "heading": "Target Masa Depan (To-Be)",
            "sub": "Operasional berbasis data & machine learning",
            "items": [
                "Penerapan **rekomendasi tarif dinamis** berbasis elastisitas harga dan lonjakan permintaan aktual penumpang.",
                "Peringatan dini keterisian bus otomatis aktif sejak **30–60 hari sebelum jadwal keberangkatan**.",
                "Pemantauan tarif kompetitor secara otomatis dan harian melalui integrasi data online travel agent (OTA).",
                "Promosi cerdas berbasis segmen pelanggan dan jadwal sepi peminat untuk memaksimalkan margin pendapatan.",
            ],
        },
        takeaway="Transformasi digital menggeser DAMRI dari pendekatan reaktif menjadi proaktif berbasis analitik cerdas.",
    )

    # 12. Strategic Key Finding
    deck.key_finding(
        number="01",
        headline="Temuan Kunci: Peluang Monetisasi Kursi Kosong Musiman",
        statement="Kapasitas armada bus DAMRI dapat ditingkatkan utilitasnya hingga +4.2% tanpa belanja modal (Capex) tambahan unit.",
        body="Analisis data pemesanan menunjukkan bahwa pada rute-rute utama, penurunan load factor terjadi bukan karena ketiadaan penumpang, melainkan karena keterlambatan penyesuaian kuota dan tarif terhadap pergerakan pasar. Dengan menerapkan modul **Booking Curve Analytics** dan **Competitor Intelligence**, DAMRI dapat menangkap segmen penumpang awal dengan harga terkelola, sekaligus mengamankan margin optimal pada periode puncak (peak season).",
        section="Kesimpulan Eksekutif",
        takeaway="Optimalisasi berbasis data analitik memberikan jalan tercepat bagi DAMRI dalam memulihkan profitabilitas rute komersial.",
    )

    # 13. Closing Slide
    deck.closing(
        title="Terima kasih",
        subtitle="Satgas Transformasi Digital BP BUMN & Perum DAMRI",
        contact="info@damri.co.id  ·  www.damri.co.id",
    )

    deck.save(output_path)
    return output_path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else None
    res = generate_sample_deck(out)
    print(f"Sample deck built: {res}")
