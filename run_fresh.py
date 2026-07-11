import sys
import os
import shutil

sys.path.append(r"D:\tammachart\src")

target_project = "ป8401_สทล01_ลป2_03 EAR"
layout_dir = os.path.join(r"D:\tammachart\Data_layout", target_project)

print("="*50)
if os.path.exists(layout_dir):
    print(f"กำลังลบโฟลเดอร์ผลลัพธ์เก่า: {layout_dir}")
    try:
        shutil.rmtree(layout_dir)
        print("ลบโฟลเดอร์สำเร็จ!")
    except Exception as e:
        print(f"ไม่สามารถลบโฟลเดอร์ได้ (ตรวจสอบว่าไม่ได้เปิด Excel หรือ QGIS ค้างไว้): {e}")
        sys.exit(1)
else:
    print(f"ไม่พบโฟลเดอร์เก่า ({layout_dir}) ข้ามขั้นตอนการลบ")
print("="*50)

print("กำลังเริ่มรัน Pipeline เพื่อประมวลผลใหม่ทั้งหมด...")
# เรียกใช้ pipeline
os.system(r"python D:\tammachart\src\pipeline.py")
print("="*50)
print("รันเสร็จสิ้น! สามารถเข้าไปดูผลลัพธ์ในโฟลเดอร์ Data_layout ได้เลยครับ")
