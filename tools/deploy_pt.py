import os
import time
import shutil

src = r"c:\Users\VICTUS\Downloads\CN PROJECT\tools\CiscoPacketTracer_820_setup.exe"
dst = r"C:\Users\VICTUS\Desktop\CiscoPacketTracer_Setup.exe"

print("Monitoring Packet Tracer download...")
# Expected size ~234,196,488 bytes
target_size = 230000000

while True:
    if os.path.exists(src):
        sz = os.path.getsize(src)
        print(f"Current size: {sz / (1024*1024):.1f} MB")
        if sz >= target_size:
            print("Download completed! Copying to Desktop...")
            try:
                shutil.copy2(src, dst)
                print(f"[SUCCESS] Copied installer to {dst}")
                break
            except Exception as e:
                print(f"Error copying: {e}")
                time.sleep(2)
    time.sleep(10)
