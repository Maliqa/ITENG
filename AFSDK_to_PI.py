import clr
import sys
import time
import random
from datetime import datetime

# Load PI AF SDK
sdk_path = r"C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0"
sys.path.append(sdk_path)
clr.AddReference("OSIsoft.AFSDK")

from OSIsoft.AF import AFSystem
from OSIsoft.AF.PI import PIServers, PIPoint
from OSIsoft.AF.Asset import AFValue, AFValues
from OSIsoft.AF.Time import AFTime
from OSIsoft.AF.Data import AFUpdateOption, AFBufferOption

# === KONFIGURASI ===
PI_SERVER_NAME = "PISRVCISDEMO"
PI_TAG_NAME    = "TEST_OPC"
SCAN_INTERVAL  = 5

def main():
    print("=== PI AFSDK -> PI Data Archive ===")
    print("Tag    : " + PI_TAG_NAME)
    print("Server : " + PI_SERVER_NAME)
    print("Interval: " + str(SCAN_INTERVAL) + " detik")
    print("===================================")

    # Step 1: Connect ke PI Server
    print("\n[1] Connecting ke PI Server...")
    try:
        pi_servers = PIServers()
        pi_server = pi_servers[PI_SERVER_NAME]
        pi_server.Connect()
        print("[OK] Connected: " + str(pi_server.Name))
    except Exception as e:
        print("[ERROR] Connect gagal: " + str(e))
        return

    # Step 2: Cari PI Tag
    print("\n[2] Mencari tag: " + PI_TAG_NAME)
    try:
        pi_point = PIPoint.FindPIPoint(pi_server, PI_TAG_NAME)
        print("[OK] Tag ditemukan: " + str(pi_point.Name))
    except Exception as e:
        print("[ERROR] Tag tidak ditemukan: " + str(e))
        pi_server.Disconnect()
        return

    # Step 3: Loop kirim data
    print("\n[3] Kirim data tiap " + str(SCAN_INTERVAL) + " detik...")
    print("     Ctrl+C untuk stop\n")
    count = 0

    while True:
        try:
            # Generate nilai random simulasi
            value = round(random.uniform(0, 32767), 4)

            # Buat AFValue
            af_val = AFValue()
            af_val.Value = System.Double(value)
            af_val.Timestamp = AFTime.Now

            # Tulis ke PI Archive
            errors = pi_point.UpdateValue(
                af_val,
                AFUpdateOption.Replace,
                AFBufferOption.BufferIfPossible
            )

            count += 1
            now = datetime.now().strftime("%H:%M:%S")
            print("[" + now + "] Value: " + str(value) + " -> OK (#" + str(count) + ")")
            time.sleep(SCAN_INTERVAL)

        except KeyboardInterrupt:
            print("\nDihentikan. Total: " + str(count) + " nilai terkirim.")
            break
        except Exception as e:
            print("[ERROR] " + str(e))
            time.sleep(SCAN_INTERVAL)

    pi_server.Disconnect()
    print("Disconnected dari PI Server.")

main()
