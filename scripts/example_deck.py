"""
BUMN Task Force & DAMRI Presentation Sample Generator (Lorem Ipsum).

Generates a complete 7-slide demonstration presentation faithfully implementing
each of the 7 official BUMN slide templates with realistic Lorem Ipsum content.

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

    deck = BUMNDeck()

    # =========================================================================
    # SLIDE 1: Cover Slide
    # =========================================================================
    deck.cover(
        title="Taskforce BP BUMN - Perum Damri",
        subtitle="AI for Public Services: DAMRI's Transformation Roadmap for Public Transportation",
    )

    # =========================================================================
    # SLIDE 2: Prioritization Framework (Impact vs Effort 2x2 Matrix)
    # =========================================================================
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
                "items": [
                    "Demand Forecasting Engine",
                    "Price Elasticity Engine",
                    "Optimal Fare Engine",
                    "Inventory & Seat Control",
                    "Revenue Optimization Model",
                ],
            },
            "C": {
                "title": "Low-hanging fruit",
                "subtitle": "Low impact · low effort",
                "count": 2,
                "items": [
                    "Customer Segmentation",
                    "Pricing Governance Rules",
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
            {
                "badge": "A",
                "title": "Quick wins",
                "text": "Improve occupancy and competitor visibility with existing data; no system integration required.",
            },
            {
                "badge": "B",
                "title": "Major projects",
                "text": "The core dynamic fare engines; high value but require automated pricing approval workflows.",
            },
            {
                "badge": "C",
                "title": "Low-hanging fruit",
                "text": "Customer segmentation & governance supporting the pricing rollouts across corridors.",
            },
            {
                "badge": "D",
                "title": "Hygiene factors",
                "text": "Scenario simulator & ancillary optimization for long-term route expansion and sustainability.",
            },
        ],
    )

    # =========================================================================
    # SLIDE 3: Quick Wins & Output Release 1 (Analytical Modules)
    # =========================================================================
    deck.quick_wins_modules(
        title="Quick Wins & Output Release 1",
        section_label="Transformation Plan - Dynamic Pricing",
        preview_title="05   Monitoring & alerting dashboard",
        preview_caption="This dashboard covered analytical modules (use case 1-4)",
        release_info={
            "tag": "RELEASE 1",
            "beta": "Beta   W2 Nov 26",
            "prototype": "Prototype   W2 Dec 26",
        },
        cards=[
            {
                "num": "01",
                "title": "Load Factor Prediction",
                "objective": "Forecast departure occupancy early enough to act and optimize capacity.",
                "data": "2–3 years of reservation bookings, scheduled vs. operated capacity, route master and holiday calendar.",
                "outcome": "Capacity matched to demand rather than timetable habit; +2–4 pts load factor on managed corridors with no added fleet.",
            },
            {
                "num": "02",
                "title": "Booking Curve Analytics",
                "objective": "Set the normal booking pace per route and flag departures that fall off pace.",
                "data": "Timestamped bookings (booking date, not only travel date), cancellations and fare class at booking.",
                "outcome": "Intervention moves from departure day to 30–60 days out, while price can still shift passenger behavior.",
            },
            {
                "num": "03",
                "title": "Competitor Price Intelligence",
                "objective": "Build a structured view of competitor fares by route, date and booking horizon.",
                "data": "Competitor published fares on a set cadence, a route-matching table and our own fare history.",
                "outcome": "Price gaps surface within days, not after revenue is lost; provides market constraints for fare engines.",
            },
            {
                "num": "04",
                "title": "Promotion Optimization",
                "objective": "Direct discounts to where they generate incremental demand and margin.",
                "data": "Promotion history (date, depth, channel, scope), bookings and revenue at matching grain, control departures.",
                "outcome": "Promotional spend stops subsidizing passengers who would have paid full fare; fastest margin recovery.",
            },
        ],
    )

    # =========================================================================
    # SLIDE 4: Implementation (Architecture & Pipeline)
    # =========================================================================
    deck.implementation_pipeline(
        title="Implementation",
        section_label="Transformation Plan - Dynamic Pricing",
        badge="WIP",
        stages=[
            {
                "num": "1",
                "title": "Data Sourcing & Pipeline",
                "cards": [
                    {
                        "name": "Load Factor Prediction",
                        "desc": "Last 2 years route & capacity\nNumber of bus, operated & fulfillment\nDeparture schedule + load factor\nMaster route",
                    },
                    {
                        "name": "Booking Curve Analysis",
                        "desc": "Historical order\nBooking date vs departure date\nCancellation rate",
                    },
                    {
                        "name": "Competitor Price Intelligence",
                        "desc": "OTA Data Scraping vs competitor\nRoute fare matching\nCadence monitoring",
                    },
                    {
                        "name": "Promotion Optimization",
                        "desc": "Pricing & tiering\nOTA Flash sale\nRevenue per branches",
                    },
                ],
            },
            {
                "num": "2",
                "title": "Modelling & Optimization",
                "cards": [
                    {
                        "name": "Machine Learning",
                        "desc": "AI / ML Model\nTime-series demand forecasting\nSeat elasticity optimization\nOccupancy prediction",
                    },
                    {
                        "name": "Gen AI",
                        "desc": "Natural language query engine\nAutomated commercial narrative\nExecutive anomaly alerts",
                    },
                ],
            },
            {
                "num": "3",
                "title": "Recommendation",
                "cards": [
                    {
                        "name": "Deploy Model",
                        "desc": "Outcome per use case\nCommercial quota adjustment\nDynamic ticket pricing output\nRevenue growth simulation",
                    },
                ],
            },
        ],
    )

    # =========================================================================
    # SLIDE 5: Flow Journey
    # =========================================================================
    deck.flow_journey(
        title="Flow Journey",
        section_label="Transformation Plan - Dynamic Pricing",
        badge="WIP\nMas Valda",
        steps=[
            {
                "num": "01",
                "stage": "Pencarian & Permintaan",
                "title": "Customer Intent",
                "desc": "Calon penumpang mencari rute dan jadwal melalui DAMRI Apps atau mitra tiket; sinyal lonjakan permintaan terdeteksi secara otomatis.",
            },
            {
                "num": "02",
                "stage": "Prediksi & Optimasi",
                "title": "Dynamic Fare Engine",
                "desc": "Mesin analitik menghitung estimasi load factor dan membandingkan tarif kompetitor OTA untuk menghasilkan rekomendasi harga optimal.",
            },
            {
                "num": "03",
                "stage": "Alokasi Kuota Kursi",
                "title": "Seat & Tier Allocation",
                "desc": "Pemberian kuota tiket berjenjang berdasarkan kurva booking keberangkatan bus (30–60 hari sebelum jadwal rit).",
            },
            {
                "num": "04",
                "stage": "Operasional Armada",
                "title": "Dispatch & LMB Online",
                "desc": "Validasi manifest penumpang dan realisasi muatan bus per rute antarkota oleh kru operasional via SIMA Operasi.",
            },
            {
                "num": "05",
                "stage": "Evaluasi & Feedback",
                "title": "Model Recalibration",
                "desc": "Pencatatan pendapatan aktual per koridor sebagai data latih untuk meningkatkan akurasi rekomendasi periode berikutnya.",
            },
        ],
    )

    # =========================================================================
    # SLIDE 6: Timeline Delivery (Gantt)
    # =========================================================================
    deck.timeline_delivery(
        title="Timeline delivery: Prototype launch in Dec W2",
        section_label="Transformation Plan - Dynamic Pricing",
        badge="WIP",
        phases=[
            {"name": "Discover", "span": 2},
            {"name": "Analyze & design", "span": 2},
            {"name": "Build & test", "span": 2},
            {"name": "Release", "span": 2},
        ],
        tasks=[
            {"name": "Gather requirements", "start": 0.05, "dur": 0.16},
            {"name": "Collect data", "start": 0.14, "dur": 0.22},
            {"name": "Data pattern analysis", "start": 0.32, "dur": 0.20},
            {"name": "Infrastructure & tech design", "start": 0.46, "dur": 0.18},
            {"name": "Model development & visualization", "start": 0.58, "dur": 0.22},
            {"name": "Testing", "start": 0.70, "dur": 0.15},
            {"name": "Beta version + review", "start": 0.78, "dur": 0.12},
            {"name": "Launch prototype", "start": 0.88, "dur": 0.10},
        ],
        current_point={"frac": 0.80, "label": "We are here"},
        note="Beta is followed by a review before the Dec W2 prototype launch",
    )

    # =========================================================================
    # SLIDE 7: Closing Slide
    # =========================================================================
    deck.closing(
        title="Terima kasih",
        subtitle="Taskforce BP BUMN - Perum Damri",
        contact="info@damri.co.id",
    )

    deck.save(output_path)
    return output_path


if __name__ == "__main__":
    out_file = sys.argv[1] if len(sys.argv) > 1 else None
    res = generate_sample_deck(out_file)
    print(f"Sample deck successfully generated: {res}")
