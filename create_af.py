"""
=============================================================
 TAHAP 2 — Buat PI AF Hierarchy via AF SDK
=============================================================
 Jalankan SETELAH create_tags.py berhasil.

 CARA PAKAI:
   py -3.11 create_af.py

 HASIL:
   Database : DEMO_PLANT
   Hierarchy:
     DEMO_PLANT
     └── Rotating Equipment
         ├── Compressor-01  (VIB, TEMP, PRESS)
         ├── Pump-01        (VIB, TEMP, PRESS)
         └── Motor-01       (VIB, TEMP, PRESS)
=============================================================
"""

import clr
import sys

# Load AF SDK DLL
AF_SDK_PATH = r"C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0"
sys.path.append(AF_SDK_PATH)

clr.AddReference("OSIsoft.AFSDK")

from OSIsoft.AF import PISystems, PISystem
from OSIsoft.AF.Asset import AFDatabase, AFElement, AFAttribute
from OSIsoft.AF.PI import PIServer, PIPoint
from OSIsoft.AF.Data import AFDataReference

PI_SERVER_NAME = "PISRVCISDEMO"
AF_DB_NAME     = "DEMO_PLANT"

# ─────────────────────────────────────────────
#  Definisi Aset & Atribut
# ─────────────────────────────────────────────

ASSETS = [
    {
        "name"      : "Compressor-01",
        "description": "Centrifugal Compressor Unit 01",
        "attributes": [
            {"name": "Vibration RMS",      "tag": "DEMO.COMP01.VIB.RMS", "uom": "mm/s"},
            {"name": "Temperature",        "tag": "DEMO.COMP01.TEMP",    "uom": "degC"},
            {"name": "Suction Pressure",   "tag": "DEMO.COMP01.PRESS",   "uom": "bar"},
        ]
    },
    {
        "name"      : "Pump-01",
        "description": "Centrifugal Pump Unit 01",
        "attributes": [
            {"name": "Vibration RMS",      "tag": "DEMO.PUMP01.VIB.RMS", "uom": "mm/s"},
            {"name": "Bearing Temperature","tag": "DEMO.PUMP01.TEMP",    "uom": "degC"},
            {"name": "Discharge Pressure", "tag": "DEMO.PUMP01.PRESS",   "uom": "bar"},
        ]
    },
    {
        "name"      : "Motor-01",
        "description": "Induction Motor Unit 01",
        "attributes": [
            {"name": "Vibration RMS",      "tag": "DEMO.MTR01.VIB.RMS", "uom": "mm/s"},
            {"name": "Winding Temperature","tag": "DEMO.MTR01.TEMP",    "uom": "degC"},
            {"name": "Current",            "tag": "DEMO.MTR01.PRESS",   "uom": "A"},
        ]
    },
]

def get_or_create_element(parent, name, description=""):
    """Ambil element yang ada atau buat baru"""
    if name in parent.Elements:
        print(f"    [SKIP] Element '{name}' sudah ada")
        return parent.Elements[name]
    el = parent.Elements.Add(name)
    el.Description = description
    print(f"    [OK]   Element '{name}' dibuat")
    return el

def link_attribute_to_tag(element, attr_name, tag_name, pi_server, uom=""):
    """Buat attribute di element dan link ke PI Tag"""
    if attr_name in element.Attributes:
        print(f"      [SKIP] Attribute '{attr_name}' sudah ada")
        return

    attr = element.Attributes.Add(attr_name)
    attr.Description = f"PI Tag: {tag_name}"

    # Set data reference ke PI Point
    attr.DataReferencePlugIn = element.PISystem.DataReferencePlugIns["PI Point"]
    attr.ConfigString = f"\\\\{PI_SERVER_NAME}\\{tag_name};ReadOnly=False"

    print(f"      [OK]   Attribute '{attr_name}' → {tag_name}")

def main():
    print("=" * 60)
    print("  Membuat PI AF Hierarchy")
    print("=" * 60)

    # Connect ke PI System (AF Server — biasanya sama host dengan PI DA)
    try:
        pi_systems = PISystems()
        pi_system  = pi_systems.DefaultPISystem
        if pi_system is None:
            pi_system = pi_systems[PI_SERVER_NAME]
        pi_system.Connect()
        print(f"  Connected ke AF Server: {pi_system.Name}\n")
    except Exception as e:
        print(f"  [ERROR] Gagal connect ke AF Server: {e}")
        print(f"  Pastikan PI AF Server service berjalan.")
        return

    # Buat atau ambil database
    if AF_DB_NAME in pi_system.Databases:
        db = pi_system.Databases[AF_DB_NAME]
        print(f"  [SKIP] Database '{AF_DB_NAME}' sudah ada")
    else:
        db = pi_system.Databases.Add(AF_DB_NAME)
        print(f"  [OK]   Database '{AF_DB_NAME}' dibuat")

    # Buat root element "Rotating Equipment"
    print(f"\n  Membuat hierarki...")
    root = get_or_create_element(db, "Rotating Equipment", "Rotating Equipment Assets")

    # Buat child element per aset
    for asset in ASSETS:
        print(f"\n  → {asset['name']}")
        el = get_or_create_element(root, asset["name"], asset["description"])

        for attr_def in asset["attributes"]:
            link_attribute_to_tag(
                el,
                attr_def["name"],
                attr_def["tag"],
                PI_SERVER_NAME,
                attr_def["uom"]
            )

    # Simpan semua perubahan ke AF Server
    db.CheckIn()
    print(f"\n  {'─'*55}")
    print(f"  PI AF Hierarchy berhasil dibuat & di-check in!")
    print(f"  Buka PI System Explorer untuk verifikasi.")
    print(f"\n  Lanjut ke: py -3.11 -m streamlit run streamlit_app.py")
    print(f"  {'─'*55}\n")


if __name__ == "__main__":
    main()
