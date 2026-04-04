import clr
import sys
import time
import random
from datetime import datetime

# Load PI AF SDK
sys.path.append(r"C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0")
clr.AddReference("OSIsoft.AFSDK")

from OSIsoft.AF import *
from OSIsoft.AF.PI import *
from OSIsoft.AF.Time import *
from OSIsoft.AF.Asset import *

# === KONFIGURASI ===
PI_SERVER_NAME = "PISRVCISDEMO"
PI_TAG_NAME    = "TEST_OPC"
SCAN_INTERVAL  = 5

def main():
    print("=== PI AFSDK -> PI Data Archive ===")
    print("Tag: " + PI_TAG_NAME)
    print("Server: " + PI_SERVER_NAME)
    print("Interval: " + str(SCAN_INTERVAL) + " detik")
    print("===================================")

    # Step 1: Connect ke PI Server
    print("\n[1] Connecting ke PI Server...")
    try:
        pi_servers = PIServers()
        pi_server = pi_servers[PI_SERVER_NAME]
        pi_server.Connect()
        print("[OK] Connected ke: " + str(pi_server.Name))
    except Exception as e:
        print("[ERROR] Gagal connect ke PI Server: " + str(e))
        return

    # Step 2: Ambil PI Point
    print("\n[2] Mencari PI tag: " + PI_TAG_NAME)
    try:
        pi_point = PIPoint.FindPIPoint(pi_server, PI_TAG_NAME)
        print("[OK] Tag ditemukan: " + str(pi_point.Name))
    except Exception as e:
        print("[ERROR] Tag tidak ditemukan: " + str(e))
        pi_server.Disconnect()
        return

    # Step 3: Loop tulis data
    print("\n[3] Mulai kirim data... (Ctrl+C untuk stop)\n")
    count = 0

    while True:
        try:
            value = round(random.uniform(0, 32767), 4)
            ts = AFTime.Now
            af_value = AFValue(value, ts)
            pi_point.UpdateValue(af_value, AFUpdateOption.Replace)
            count += 1
            print("[" + str(datetime.now().strftime("%H:%M:%S")) + "] Value: " + str(value) + " -> OK (#" + str(count) + ")")
            time.sleep(SCAN_INTERVAL)

        except KeyboardInterrupt:
            print("\nDihentikan. Total kiriman: " + str(count))
            break
        except Exception as e:
            print("[ERROR] " + str(e))
            time.sleep(SCAN_INTERVAL)

    pi_server.Disconnect()
    print("Disconnected.")

main()
