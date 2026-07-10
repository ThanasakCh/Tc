import os

# --- PATH CONFIGURATION ---
BASE_DIR = r"D:\tammachart"
SHP_DIR = os.path.join(BASE_DIR, "ข้อมูลShp")
DATA_LAYOUT_DIR = os.path.join(BASE_DIR, "Data_layout")

# --- API CONFIGURATION ---
# ใส่ API Key ของ Google Maps ได้ที่นี่
# หากปล่อยว่างไว้ ระบบจะข้ามการดึงรูป Street View แต่จะประมวลผลส่วนอื่นต่อจนจบ
GOOGLE_MAPS_API_KEY = ""

# --- GIS CONFIGURATION ---
# พิกัดอ้างอิงหลักที่ใช้ในการคำนวณ (UTM Zone 47N)
TARGET_CRS = "EPSG:32647"

def get_project_settings(project_folder_name):
    """
    ตรวจจับว่าเป็นงาน EC หรือ EAR จากชื่อโฟลเดอร์ 
    และคืนค่าระยะ Buffer ที่ถูกต้อง (เมตร)
    """
    name_upper = project_folder_name.upper()
    if "EAR" in name_upper:
        return {"type": "EAR", "buffer_m": 1000}
    elif "EC" in name_upper:
        return {"type": "EC", "buffer_m": 500}
    else:
        # ค่าเริ่มต้น หากระบุไม่ได้
        return {"type": "UNKNOWN", "buffer_m": 500}
