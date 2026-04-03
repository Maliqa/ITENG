"""
=============================================================
 TAHAP 1 — Buat PI Tags via AF SDK (pythonnet)
=============================================================
 CARA PAKAI:
   py -3.11 C:\PI_Demo\create_tags.py
=============================================================
"""

import clr
import sys

# Load AF SDK
AF_SDK_PATH = r"C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0"
sys.path.append(AF_SDK_PATH)
clr.AddReference("OSIsoft.AFSDK")

from OSIsoft.AF.PI import PIServers, PIPoint, PIPointClass
from OSIsoft.AF import PISystems

PI_SERVER_NAME = "PISRVCISDEMO"

TAGS = [
    {"name": "DEMO.COMP01.VIB.RMS",  "descriptor": "Compressor 01 Vibration RMS",      "eng_units": "mm/s", "zero": 0, "span": 50},
    {"name": "DEMO.COMP01.TEMP",     "descriptor": "Compressor 01 Temperature",         "eng_units": "degC", "zero": 0, "span": 200},
    {"name": "DEMO.COMP01.PRESS",    "descriptor": "Compressor 01 Suction Pressure",    "eng_units": "bar",  "zero": 0, "span": 20},
    {"name": "DEMO.PUMP01.VIB.RMS",  "descriptor": "Pump 01 Vibration RMS",             "eng_units": "mm/s", "zero": 0, "span": 50},
    {"name": "DEMO.PUMP01.TEMP",     "descriptor": "Pump 01 Bearing Temperature",       "eng_units": "degC", "zero": 0, "span": 150},
    {"name": "DEMO.PUMP01.PRESS",    "descriptor": "Pump 01 Discharge Pressure",        "eng_units": "bar",  "zero": 0, "span": 20},
    {"name": "DEMO.MTR01.VIB.RMS",   "descriptor": "Motor 01 Vibration RMS",            "eng_units": "mm/s", "zero": 0, "span": 50},
    {"name": "DEMO.MTR01.TEMP",      "descriptor": "Motor 01 Winding Temperature",      "eng_units": "degC", "zero": 0, "span": 200},
    {"name": "DEMO.MTR01.CURR",      "descriptor": "Motor 01 Current",                  "eng_units": "A",    "zero": 0, "span": 100},
]

def main():
    print("=" * 60)
    print("  Membuat PI Tags di PISRVCISDEMO (via AF SDK)")
    print("=" * 60)

    # Connect ke PI Data Archive via PIServers
    try:
        pi_servers = PIServers()
        server = pi_servers[PI_SERVER_NAME]
        server.Connect()
        print(f"  Connected: {server.Name}\n")
    except Exception as e:
        print(f"  [ERROR] Gagal connect: {e}")
        return

    created = skipped = failed = 0

    for tag_def in TAGS:
        name = tag_def["name"]
        try:
            # Cek apakah tag sudah ada
            existing = PIPoint.FindPIPoint(server, name)
            print(f"  [SKIP] {name} — sudah ada")
            skipped += 1
            continue
        except:
            pass  # Tag belum ada, lanjut buat

        try:
            # Definisi attribute tag
            point_attributes = {
                "descriptor" : tag_def["descriptor"],
                "engunits"   : tag_def["eng_units"],
                "pointtype"  : "Float32",
                "zero"       : tag_def["zero"],
                "span"       : tag_def["span"],
                "compressing": 1,
                "excmin"     : 0,
                "excmax"     : 100,
                "excdev"     : 0.5,
                "excdevpercent": 1,
            }

            # Buat tag baru
            new_point = server.CreatePIPoint(name, point_attributes)
            print(f"  [OK]   {name} — {tag_def['descriptor']} ({tag_def['eng_units']})")
            created += 1

        except Exception as e:
            print(f"  [ERR]  {name} — {e}")
            failed += 1

    print(f"\n  {'─'*55}")
    print(f"  Selesai! Created: {created} | Skipped: {skipped} | Failed: {failed}")
    print(f"\n  Lanjut ke: py -3.11 C:\\PI_Demo\\create_af.py")
    print(f"  {'─'*55}\n")


if __name__ == "__main__":
    main()
