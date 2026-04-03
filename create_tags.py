"""
=============================================================
 TAHAP 1 — Buat PI Tags di PI Data Archive
=============================================================
 Jalankan PERTAMA sebelum script lain.
 
 CARA PAKAI:
   py -3.11 create_tags.py

 HASIL: 9 PI Tags terbuat di PISRVCISDEMO
=============================================================
"""

import PIconnect as PI
import time

PI_SERVER = "PISRVCISDEMO"

# ─────────────────────────────────────────────
#  Definisi semua tags yang mau dibuat
# ─────────────────────────────────────────────

TAGS = [
    # Compressor-01
    {
        "name"       : "DEMO.COMP01.VIB.RMS",
        "descriptor" : "Compressor 01 Vibration RMS",
        "eng_units"  : "mm/s",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 50,
    },
    {
        "name"       : "DEMO.COMP01.TEMP",
        "descriptor" : "Compressor 01 Temperature",
        "eng_units"  : "degC",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 200,
    },
    {
        "name"       : "DEMO.COMP01.PRESS",
        "descriptor" : "Compressor 01 Suction Pressure",
        "eng_units"  : "bar",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 20,
    },
    # Pump-01
    {
        "name"       : "DEMO.PUMP01.VIB.RMS",
        "descriptor" : "Pump 01 Vibration RMS",
        "eng_units"  : "mm/s",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 50,
    },
    {
        "name"       : "DEMO.PUMP01.TEMP",
        "descriptor" : "Pump 01 Bearing Temperature",
        "eng_units"  : "degC",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 150,
    },
    {
        "name"       : "DEMO.PUMP01.PRESS",
        "descriptor" : "Pump 01 Discharge Pressure",
        "eng_units"  : "bar",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 20,
    },
    # Motor-01
    {
        "name"       : "DEMO.MTR01.VIB.RMS",
        "descriptor" : "Motor 01 Vibration RMS",
        "eng_units"  : "mm/s",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 50,
    },
    {
        "name"       : "DEMO.MTR01.TEMP",
        "descriptor" : "Motor 01 Winding Temperature",
        "eng_units"  : "degC",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 200,
    },
    {
        "name"       : "DEMO.MTR01.PRESS",
        "descriptor" : "Motor 01 Current",
        "eng_units"  : "A",
        "point_type" : "Float32",
        "zero"       : 0,
        "span"       : 100,
    },
]

def main():
    print("=" * 60)
    print("  Membuat PI Tags di PISRVCISDEMO")
    print("=" * 60)

    server = PI.PIServer(server=PI_SERVER)
    print(f"  Connected: {server}\n")

    created = 0
    skipped = 0
    failed  = 0

    for tag_def in TAGS:
        name = tag_def["name"]
        try:
            # Cek apakah tag sudah ada
            existing = list(server.search(name))
            if existing:
                print(f"  [SKIP] {name} — sudah ada")
                skipped += 1
                continue

            # Buat tag baru via PI SDK
            point = server.create_point(
                name,
                point_type     = tag_def["point_type"],
                descriptor     = tag_def["descriptor"],
                eng_units      = tag_def["eng_units"],
                zero           = tag_def["zero"],
                span           = tag_def["span"],
            )
            print(f"  [OK]   {name} — {tag_def['descriptor']} ({tag_def['eng_units']})")
            created += 1
            time.sleep(0.3)  # Jeda kecil biar tidak overwhelm server

        except Exception as e:
            print(f"  [ERR]  {name} — {e}")
            failed += 1

    print(f"\n  {'─'*55}")
    print(f"  Selesai! Created: {created} | Skipped: {skipped} | Failed: {failed}")
    print(f"\n  Lanjut ke: py -3.11 create_af.py")
    print(f"  {'─'*55}\n")


if __name__ == "__main__":
    main()
