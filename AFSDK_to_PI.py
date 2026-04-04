import clr
import sys
import time
import win32com.client
from datetime import datetime

# Load PI AF SDK
sys.path.append(r"C:\Program Files (x86)\PIPC\AF\PublicAssemblies\4.0")
clr.AddReference("OSIsoft.AFSDK")

from OSIsoft.AF import *
from OSIsoft.AF.PI import PIServers, PIPoint
from OSIsoft.AF.Asset import AFValue
from OSIsoft.AF.Time import AFTime
from OSIsoft.AF.Data import AFUpdateOption, AFBufferOption

# === KONFIGURASI ===
PI_SERVER_NAME = "PISRVCISDEMO"
PI_TAG_NAME    = "TEST_OPC"
OPC_SERVER     = "Matrikon.OPC.Simulation.1"
OPC_ITEM_ID    = "Random.Real8"
SCAN_INTERVAL  = 5

# === FUNGSI BACA OPC ===
def connect_opc():
    print("[OPC] Connecting ke " + OPC_SERVER + "...")
    try:
        opc = win32com.client.Dispatch(OPC_SERVER)
        groups = opc.OPCGroups
        group = groups.Add("PIBridgeGroup")
        group.IsActive = True
        group.UpdateRate = 1000
        items = group.OPCItems
        item = items.AddItem(OPC_ITEM_ID, 1)
        print("[OPC] Connected! Item: " + OPC_ITEM_ID)
        return opc, item
    except Exception as e:
        print("[OPC ERROR] " + str(e))
        return None, None

def read_opc(item):
    try:
        value = float(item.Value)
        quality = item.Quality
        return value, quality
    except Exception as e:
        print("[OPC READ ERROR] " + str(e))
        return None, None

# === FUNGSI TULIS PI ===
def connect_pi():
    print("[PI] Connecting ke " + PI_SERVER_NAME + "...")
    try:
        pi_servers = PIServers()
        pi_server = pi_servers[PI_SERVER_NAME]
        pi_server.Connect()
        print("[PI] Connected: " + str(pi_server.Name))
        return pi_server
    except Exception as e:
        print("[PI ERROR] Connect gagal: " + str(e))
        return None

def get_pi_point(pi_server):
    print("[PI] Mencari tag: " + PI_TAG_NAME)
    try:
        pi_point = PIPoint.FindPIPoint(pi_server, PI_TAG_NAME)
        print("[PI] Tag ditemukan: " + str(pi_point.Name))
        return pi_point
    except Exception as e:
        print("[PI ERROR] Tag tidak ditemukan: " + str(e))
        return None

def write_pi(pi_point, value):
    try:
        af_val = AFValue()
        af_val.Value = float(value)
        af_val.Timestamp = AFTime.Now
        pi_point.UpdateValue(
            af_val,
            AFUpdateOption.Replace,
            AFBufferOption.BufferIfPossible
        )
        return True
    except Exception as e:
        print("[PI WRITE ERROR] " + str(e))
        return False

# === MAIN ===
def main():
    print("=" * 45)
    print("  Matrikon OPC -> PI Data Archive Bridge")
    print("=" * 45)
    print("OPC Server : " + OPC_SERVER)
    print("OPC Item   : " + OPC_ITEM_ID)
    print("PI Tag     : " + PI_TAG_NAME)
    print("Interval   : " + str(SCAN_INTERVAL) + " detik")
    print("=" * 45)

    # Connect OPC
    opc, opc_item = connect_opc()
    if opc_item is None:
        print("[FATAL] Tidak bisa connect ke OPC Server!")
        return

    # Connect PI
    pi_server = connect_pi()
    if pi_server is None:
        print("[FATAL] Tidak bisa connect ke PI Server!")
        return

    # Get PI Point
    pi_point = get_pi_point(pi_server)
    if pi_point is None:
        print("[FATAL] PI tag tidak ditemukan!")
        pi_server.Disconnect()
        return

    print("\nMulai transfer data... (Ctrl+C untuk stop)\n")
    count = 0
    fail = 0

    while True:
        try:
            # Baca dari Matrikon OPC
            value, quality = read_opc(opc_item)

            if value is not None:
                # Tulis ke PI Archive
                ok = write_pi(pi_point, value)
                count += 1
                now = datetime.now().strftime("%H:%M:%S")
                status = "OK" if ok else "GAGAL"
                print("[" + now + "] OPC Value: " + str(round(value, 4)) + " | Quality: " + str(quality) + " -> PI: " + status + " (#" + str(count) + ")")
            else:
                fail += 1
                print("[WARN] OPC read gagal (#" + str(fail) + ")")

            time.sleep(SCAN_INTERVAL)

        except KeyboardInterrupt:
            print("\nDihentikan.")
            print("Total sukses : " + str(count))
            print("Total gagal  : " + str(fail))
            break
        except Exception as e:
            print("[ERROR] " + str(e))
            time.sleep(SCAN_INTERVAL)

    pi_server.Disconnect()
    print("Selesai.")

main()
