import os
import requests
import math
from config import GOOGLE_MAPS_API_KEY, DATA_LAYOUT_DIR

class StreetViewFetcher:
    def __init__(self, project_dir):
        self.project_dir = project_dir
        self.project_name = os.path.basename(project_dir)
        self.output_dir = os.path.join(DATA_LAYOUT_DIR, self.project_name, "StreetView_Images")
        
    def check_api_key(self):
        """ตรวจสอบว่ามี API Key หรือไม่"""
        if not GOOGLE_MAPS_API_KEY or GOOGLE_MAPS_API_KEY.strip() == "":
            print("\n[WARNING] ข้ามกระบวนการโหลดภาพ Street View เนื่องจากไม่มี API Key")
            print("กรุณาตั้งค่า GOOGLE_MAPS_API_KEY ในไฟล์ config.py หากต้องการใช้งาน")
            
            # บันทึก log แจ้งเตือนไว้ในโฟลเดอร์เพื่อให้ผู้ใช้ทราบ
            os.makedirs(self.output_dir, exist_ok=True)
            log_path = os.path.join(self.output_dir, "WARNING_NO_API_KEY.txt")
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write("กระบวนการดึงรูปภาพ Street View ถูกข้ามไปเนื่องจากไม่ได้ระบุ API Key\n")
                f.write("ประมวลผล GIS และ Excel เสร็จสิ้นตามปกติ\n")
            return False
        return True
        
    def fetch_image(self, lat, lng, heading=0, fov=90, pitch=0, filename="image.jpg"):
        """ดาวน์โหลดภาพ Street View ตามพิกัด"""
        if not self.check_api_key():
            return False
            
        os.makedirs(self.output_dir, exist_ok=True)
        filepath = os.path.join(self.output_dir, filename)
        
        # ถ้ารูปมีอยู่แล้วไม่ต้องโหลดซ้ำ
        if os.path.exists(filepath):
            print(f"มีภาพ {filename} อยู่แล้ว ข้ามการดาวน์โหลด")
            return True
            
        url = "https://maps.googleapis.com/maps/api/streetview"
        params = {
            "size": "640x640",
            "location": f"{lat},{lng}",
            "heading": heading,
            "fov": fov,
            "pitch": pitch,
            "key": GOOGLE_MAPS_API_KEY
        }
        
        print(f"กำลังดาวน์โหลดภาพ Street View ที่พิกัด {lat},{lng} ...")
        try:
            response = requests.get(url, params=params)
            if response.status_code == 200:
                with open(filepath, 'wb') as f:
                    f.write(response.content)
                print(f"บันทึกภาพสำเร็จ: {filepath}")
                return True
            else:
                print(f"เกิดข้อผิดพลาดในการโหลดรูป HTTP {response.status_code}")
                return False
        except Exception as e:
            print(f"Network error: {e}")
            return False
            
    def run_batch_from_gis_cache(self):
        """วิ่งโหลดรูปภาพจากจุดตัดที่คำนวณไว้ใน Cache"""
        if not self.check_api_key():
            return
            
        json_path = os.path.join(self.project_dir, "cache", "master_gis_results.json")
        if not os.path.exists(json_path):
            print("ไม่พบข้อมูล Cache สำหรับดาวน์โหลด Street View")
            return
            
        import json
        from pyproj import Transformer
        
        with open(json_path, 'r', encoding='utf-8') as f:
            master_records = json.load(f)
            
        # ตัวแปลงพิกัดจาก UTM 47N (EPSG:32647) เป็น Lat/Lng (EPSG:4326)
        transformer = Transformer.from_crs("EPSG:32647", "EPSG:4326", always_xy=True)
        
        print("\n--- เริ่มกระบวนการดาวน์โหลดภาพ Street View ---")
        downloaded_count = 0
        
        for sheet_name, records in master_records.items():
            for i, r in enumerate(records):
                # ตรวจสอบว่ามีพิกัดหรือไม่
                if 'center_x' not in r or 'center_y' not in r:
                    continue
                    
                x, y = r['center_x'], r['center_y']
                lng, lat = transformer.transform(x, y)
                
                # ตั้งชื่อไฟล์ให้สื่อความหมาย เช่น "1.ป่า_KM_5002+153.jpg"
                safe_sheet = sheet_name.replace("/", "_").replace("\\", "_")
                filename = f"{safe_sheet}_{r['KM In']}.jpg"
                
                success = self.fetch_image(lat, lng, filename=filename)
                if success:
                    downloaded_count += 1
                    
        print(f"ดาวน์โหลดภาพ Street View สำเร็จทั้งหมด {downloaded_count} ภาพ")

if __name__ == "__main__":
    fetcher = StreetViewFetcher(r"D:\tammachart\ป.84-1_สทล.1_มส._1 EAR")
    fetcher.run_batch_from_gis_cache()
