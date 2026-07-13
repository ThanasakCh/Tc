import os
import glob
import json
import numpy as np

from config import BASE_DIR, get_project_settings, DATA_LAYOUT_DIR
from gis_engine import GISEngine

def sanitize_json(obj):
    """ทำความสะอาดข้อมูล NaN/Inf ก่อนเซฟลง JSON ตาม Step 5"""
    if isinstance(obj, dict):
        return {k: sanitize_json(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [sanitize_json(v) for v in obj]
    elif isinstance(obj, float):
        if np.isnan(obj) or np.isinf(obj):
            return None
    return obj

def run_pipeline():
    from datetime import datetime
    start_time = datetime.now()
    summary_stats = {"total_projects": 0, "successful_projects": 0, "failed_projects": 0}
    
    print("=== เริ่มต้นกระบวนการอัตโนมัติ EC & EAR ===")
    
    MASTER_REQUESTS_DIR = os.path.join(BASE_DIR, "คำขอทั้งหมด")
    if not os.path.exists(MASTER_REQUESTS_DIR):
        print(f"ไม่พบโฟลเดอร์ {MASTER_REQUESTS_DIR} กรุณาสร้างและนำคำขอไปใส่ไว้")
        return

    print("กรุณาเลือกโหมดการทำงาน:")
    print("1. รันทั้งหมด (สแกนหาโฟลเดอร์ EC/EAR ใน 'คำขอทั้งหมด')")
    print("2. เลือกเฉพาะคำขอ (ป้อนชื่อเพื่อค้นหาและยืนยัน Y/N)")
    choice = input("เลือกโหมด (1/2): ").strip()
    
    project_dirs = []
    all_ec_ear = []
    for root, dirs, files in os.walk(MASTER_REQUESTS_DIR):
        for d in dirs:
            if "EC" in d.upper() or "EAR" in d.upper():
                all_ec_ear.append(os.path.join(root, d))
                
    if choice == "1":
        project_dirs = all_ec_ear
    elif choice == "2":
        keywords_input = input("\nกรุณาใส่ชื่อคำขอที่ต้องการค้นหา (คั่นด้วยคอมม่าถ้ามีหลายอัน เช่น ลป2, 8401) \n[หรือกด Enter เพื่อดึงรายชื่อทั้งหมดมาให้เลือกทีละอัน]: ")
        keywords = [k.strip() for k in keywords_input.split(",") if k.strip()]
        
        import re
        def is_fuzzy_match(keyword, text):
            # ลบจุดและขีดออกให้หมดเพื่อการเทียบที่ง่ายขึ้น
            kw_clean = re.sub(r'[\.\-_\s]', '', keyword).upper()
            txt_clean = re.sub(r'[\.\-_\s]', '', text).upper()
            
            # ถ้าเป็น 84-1 -> 841, แต่โฟลเดอร์เป็น 8401
            # ป้องกัน PatternError ด้วย re.escape
            pattern = '.*'.join(re.escape(c) for c in list(kw_clean))
            return re.search(pattern, txt_clean) is not None

        if not keywords:
            # ถ้าไม่ใส่อะไรเลย ให้ถือว่าค้นหาทั้งหมด
            keywords = [""]
            
        for kw in keywords:
            matched = False
            for p in all_ec_ear:
                if not kw or is_fuzzy_match(kw, os.path.basename(p)):
                    matched = True
                    ans = input(f"-> พบโฟลเดอร์: [{os.path.relpath(p, MASTER_REQUESTS_DIR)}] ต้องการรันหรือไม่? (y/n): ")
                    if ans.lower() == 'y':
                        if p not in project_dirs:
                            project_dirs.append(p)
                            print(f"   เพิ่ม {os.path.basename(p)} เข้าสู่คิวแล้ว")
            if not matched:
                print(f"-> ไม่พบโฟลเดอร์ใดที่ตรงกับคำค้นหา: '{kw}'")
    else:
        print("ตัวเลือกไม่ถูกต้อง ยกเลิกการทำงาน")
        return
        
    if not project_dirs:
        print("\nไม่มีคำขอในคิว ยกเลิกการทำงาน")
        return
        
    summary_stats["total_projects"] = len(project_dirs)
    
    engine = GISEngine(os.path.join(BASE_DIR, "ข้อมูลShp"))
    
    for i, project_path in enumerate(project_dirs, 1):
        proj_name = os.path.basename(project_path)
        print(f"\n--- [{i}/{len(project_dirs)}] กำลังประมวลผลคำขอ: {proj_name} ---")
        
        # 2. เตรียม Buffer (500m หรือ 1000m)
        settings = get_project_settings(proj_name)
        buffer_m = settings["buffer_m"]
        print(f"[{settings['type']} Mode] เซ็ตระยะ Buffer ที่ {buffer_m} เมตร")
        
        # 3. โหลดเส้นกลางถนนและใช้ Fuzzy Match (STEP 1)
        engine.load_road_centerline(project_path)
        
        # --- [NEW] Export Start and End Points from KM ---
        try:
            import geopandas as gpd
            import pandas as pd
            from shapely.geometry import Point
            
            # Find KM shapefile
            km_shp = None
            for f in os.listdir(project_path):
                if f.upper().endswith("_KM.SHP"):
                    km_shp = os.path.join(project_path, f)
                    break
                    
            if km_shp and os.path.exists(km_shp):
                km_gdf = gpd.read_file(km_shp)
                
                km_col = None
                for col in km_gdf.columns:
                    if col.lower() == 'km':
                        km_col = col
                        break
                        
                if km_col:
                    km_gdf['_km_num'] = pd.to_numeric(km_gdf[km_col].astype(str).str.replace('+', ''), errors='coerce')
                    km_gdf = km_gdf.dropna(subset=['_km_num'])
                    
                    if not km_gdf.empty:
                        min_km_row = km_gdf.loc[km_gdf['_km_num'].idxmin()]
                        max_km_row = km_gdf.loc[km_gdf['_km_num'].idxmax()]
                        
                        pts = []
                        labels = []
                        
                        start_geom = min_km_row.geometry
                        end_geom = max_km_row.geometry
                        
                        if isinstance(start_geom, Point):
                            pts.append(start_geom)
                            labels.append(f"Start (KM {min_km_row[km_col]})")
                        if isinstance(end_geom, Point) and end_geom != start_geom:
                            pts.append(end_geom)
                            labels.append(f"End (KM {max_km_row[km_col]})")
                            
                        if pts:
                            pt_gdf = gpd.GeoDataFrame({'PointType': labels}, geometry=pts, crs=km_gdf.crs)
                            if not engine.road_gdf.empty:
                                pt_gdf = pt_gdf.to_crs(engine.road_gdf.crs)
                                
                            layout_shp_dir = os.path.join(DATA_LAYOUT_DIR, proj_name, "Shp")
                            os.makedirs(layout_shp_dir, exist_ok=True)
                            pt_out_path = os.path.join(layout_shp_dir, "Start_End_Point.shp")
                            pt_gdf.to_file(pt_out_path)
                            print(f"[{proj_name}] สร้างไฟล์จุดเริ่มต้น-สิ้นสุดแล้ว (จาก KM): Start_End_Point.shp")
            else:
                print(f"[{proj_name}] ไม่พบไฟล์ _KM.shp สำหรับสร้างจุดหัวท้าย")
        except Exception as e:
            print(f"[{proj_name}] Error generating Start_End_Point: {e}")
            
        # --- [NEW] Export Clipped Tambon ---
        try:
            import geopandas as gpd
            admin_path = os.path.join(BASE_DIR, "ข้อมูลShp", "ขอบเขตการปกครอง", "ขอบเขตตำบล.shp")
            if os.path.exists(admin_path):
                print(f"[{proj_name}] กำลังตัดแผนที่ขอบเขตตำบล...")
                admin_gdf = gpd.read_file(admin_path).to_crs(engine.road_gdf.crs)
                
                if hasattr(engine.road_gdf.geometry, 'union_all'):
                    road_union = engine.road_gdf.geometry.buffer(buffer_m).union_all()
                else:
                    import shapely.ops
                    road_union = shapely.ops.unary_union(engine.road_gdf.geometry.buffer(buffer_m))
                    
                mask_gdf = gpd.GeoDataFrame(geometry=[road_union], crs=engine.road_gdf.crs)
                clipped_admin = gpd.overlay(admin_gdf, mask_gdf, how='intersection')
                
                if not clipped_admin.empty:
                    layout_shp_dir = os.path.join(DATA_LAYOUT_DIR, proj_name, "Shp")
                    os.makedirs(layout_shp_dir, exist_ok=True)
                    tambon_out_path = os.path.join(layout_shp_dir, "Clipped_Tambon.shp")
                    clipped_admin.to_file(tambon_out_path)
                    print(f"[{proj_name}] สร้างไฟล์ขอบเขตตำบล (Clipped) แล้ว: Clipped_Tambon.shp")
        except Exception as e:
            print(f"[{proj_name}] Error generating Clipped_Tambon: {e}")
        # ----------------------------------------

        # 3.1 ค้นหาตัวย่อจังหวัดจากชื่อโปรเจกต์
        import re
        province_abbr = ""
        PROVINCE_MAP = {
            "มส": "แม่ฮ่องสอน",
            "ลป": "ลำปาง",
            "ชม": "เชียงใหม่",
            "ชร": "เชียงราย",
            "พย": "พะเยา",
            "นน": "น่าน",
            "พร": "แพร่",
            "ตาก": "ตาก"
        }
        
        for part in proj_name.split('_'):
            clean_part = re.sub(r'[\d\.]', '', part).strip()
            if clean_part in PROVINCE_MAP:
                province_abbr = clean_part
                break
                
        full_prov_name = PROVINCE_MAP.get(province_abbr, "")
        is_ear = (settings["type"] == "EAR")
        
        # 4. ทดสอบ GIS Overlay แบบวนลูปทุก Shapefile (STEP 2)
        # รายชื่อ Shapefile พื้นฐานที่ไม่ได้แยกจังหวัด
        shapefiles_to_process = [
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "01_ป่า", "ป่าสงวนแห่งชาติ.shp"), "name_col": "fr_name", "do_clip": True, "sheet": {"EAR": "1.ป่า_สงวน", "EC": "1.ป่า_สงวน"}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "01_ป่า", "01_ป่าถาวร.shp"), "name_col": "name_th", "do_clip": True, "sheet": {"EAR": "1.ป่า", "EC": "1.ป่า"}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "Forest Area 2565", "forestarea2565_wgs1984.shp"), "name_col": "f_code", "do_clip": True, "sheet": {"EAR": "4.พื้นที่คงสภาพป่า", "EC": None}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "Forest Area 2565", "forestarea2565_wgs1984.shp"), "name_col": "f_code", "do_clip": False, "buffer_m_override": 30, "sheet": {"EAR": "4.พื้นที่คงสภาพป่า_เขตทาง", "EC": None}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "ชั้นคุณภาพลุ่มน้ำ", "WSC_Edit_3112025.shp"), "name_col": "wsc_ver", "do_clip": True, "excel_name": "ชั้นคุณภาพลุ่มน้ำ", "sheet": {"EAR": "10.ชั้นคุณภาพลุ่มน้ำ", "EC": "12.ชั้นคุณภาพลุ่มน้ำ"}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "แม่น้ำ", "Stream.shp"), "name_col": "HY_LNAME", "do_clip": True, "sheet": {"EAR": "5.จุดตัดแหล่งน้ำ", "EC": "9.จุดตัดแหล่งน้ำ"}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "หมู่บ้าน", "ตำแหน่งหมู่บ้าน.shp"), "name_col": "MUBAN", "do_clip": True, "excel_name": "บ้าน", "sheet": {"EAR": "2.พื้นที่หมู่บ้าน", "EC": "2.พื้นที่หมู่บ้าน"}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "04_ตำแหน่งโบราณสถาน", "โบราณสถานประเทศไทย", "historic_site_export.shp"), "name_col": "HISNAME", "do_clip": True, "excel_name": "โบราณสถาน", "sheet": {"EAR": "9.แหล่งโบราณสถาน", "EC": "5.แหล่งโบราณสถาน"}},
        ]
        
        # เพิ่มพื้นที่อ่อนไหว: EAR ใช้สถานศึกษา+วัด+โรงพยาบาล, EC ใช้สถานศึกษา+ดินถล่ม
        shapefiles_to_process.append(
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "02_พื้นที่อ่อนไหว", "ตำแหน่งสถานศึกษา_Point.shp"), "name_col": "Name", "do_clip": True, "excel_name": "โรงเรียน", "sheet": {"EAR": "พื้นที่อ่อนไหว_สถานศึกษา", "EC": "3.ตรวจสอบสถานศึกษา"}}
        )
        
        shapefiles_to_process.extend([
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "02_พื้นที่อ่อนไหว", "ตำแหน่งวัด_Point.shp"), "name_col": "Name", "do_clip": True, "excel_name": "วัด", "sheet": {"EAR": "พื้นที่อ่อนไหว_ศาสนสถาน", "EC": "3.ตรวจสอบสถานศึกษา_วัด"}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "02_พื้นที่อ่อนไหว", "ตำแหน่งโรงพยาบาล_Point.shp"), "name_col": "Name", "do_clip": True, "excel_name": "รพ", "sheet": {"EAR": "พื้นที่อ่อนไหว_สถานพยาบาล", "EC": "3.ตรวจสอบสถานศึกษา_รพ"}},
            {"path": os.path.join(BASE_DIR, "ข้อมูลShp", "02_พื้นที่อ่อนไหว", "ตำแหน่งรพสต_Point.shp"), "name_col": "Name", "do_clip": True, "excel_name": "รพสต", "sheet": {"EAR": "พื้นที่อ่อนไหว_สถานพยาบาล", "EC": "3.ตรวจสอบสถานศึกษา_รพสต"}},
        ])

        if not is_ear:
            shapefiles_to_process.extend([
                {"path": "https://gisportal.dmr.go.th/arcgis/rest/services/MINERAL/MIN_AREA/MapServer/0", "name_col": "MIN_GEO", "do_clip": False, "excel_name": "แร่", "sheet": {"EAR": None, "EC": "6.แหล่งทรัพยากรทางธรณี"}},
                {"path": "https://gis.dmr.go.th/arcgis/rest/services/DMR_Hazard/WebService_DataCatalog_GeoHazard/MapServer/4", "name_col": "Level_T", "do_clip": False, "excel_name": "ดินถล่ม", "sheet": {"EAR": None, "EC": "7.เสี่ยงต่อการเกิดดินถล่ม"}},
                {"path": "https://gisportal.dmr.go.th/arcgis/rest/services/HAZARD/H_SINKHOLE/MapServer/0", "name_col": "PROV_NAM_T", "do_clip": False, "excel_name": "หลุมยุบ", "sheet": {"EAR": None, "EC": "7.เสี่ยงต่อการเกิดดินถล่ม"}},
                {"path": "https://gisportal.dmr.go.th/arcgis/rest/services/HAZARD/EQ_ZONE/MapServer/0", "name_col": "INTEN_T", "do_clip": False, "excel_name": "แผ่นดินไหว", "sheet": {"EAR": None, "EC": "8.เสี่ยงต่อการเกิดแผ่นดินไหว"}}
            ])
        
        # เพิ่มไฟล์รายจังหวัดถ้าหาเจอ
        import glob
        import geopandas as gpd
        crossed_provinces_thai = [full_prov_name] if full_prov_name else []
        try:
            admin_shp_path = os.path.join(BASE_DIR, "ข้อมูลShp", "ขอบเขตการปกครอง", "ขอบเขตตำบล.shp")
            if os.path.exists(admin_shp_path):
                admin_gdf = gpd.read_file(admin_shp_path).to_crs(engine.road_gdf.crs)
                road_union = engine.road_gdf.geometry.buffer(buffer_m).union_all()
                mask_gdf = gpd.GeoDataFrame(geometry=[road_union], crs=engine.road_gdf.crs)
                intersected = admin_gdf[admin_gdf.intersects(mask_gdf.geometry.iloc[0])]
                crossed = intersected['PROV_TH'].dropna().unique()
                for p in crossed:
                    p_str = engine.safe_decode(p)
                    p_str = p_str.replace("จังหวัด", "").strip()
                    if p_str and p_str not in crossed_provinces_thai:
                        crossed_provinces_thai.append(p_str)
        except Exception as e:
            print(f"Error detecting crossed provinces: {e}")
            
        print(f"พบว่าแนวเส้นทางพาดผ่านจังหวัด: {', '.join(crossed_provinces_thai)}")
        
        PROV_ENG_MAP_EXTENDED = {
            "กรุงเทพมหานคร": "bkk", "กระบี่": "kbi", "กาญจนบุรี": "kri", "กาฬสินธุ์": "ksn",
            "กำแพงเพชร": "kpt", "ขอนแก่น": "kkn", "จันทบุรี": "cti", "ฉะเชิงเทรา": "cco",
            "ชลบุรี": "cbi", "ชัยนาท": "cnt", "ชัยภูมิ": "cpm", "ชุมพร": "cpn",
            "เชียงราย": "cri", "เชียงใหม่": "cmi", "ตรัง": "trg", "ตราด": "trt",
            "ตาก": "tak", "นครนายก": "nyk", "นครปฐม": "npt", "นครพนม": "npn",
            "นครราชสีมา": "nma", "นครศรีธรรมราช": "nrt", "นครสวรรค์": "nsn", "นนทบุรี": "ntb",
            "นราธิวาส": "nwt", "น่าน": "nan", "บึงกาฬ": "bkn", "บุรีรัมย์": "brm",
            "ปทุมธานี": "ptm", "ประจวบคีรีขันธ์": "pkn", "ปราจีนบุรี": "pri", "ปัตตานี": "ptn",
            "พระนครศรีอยุธยา": "aya", "พะเยา": "pyo", "พังงา": "pna", "พัทลุง": "plg",
            "พิจิตร": "pct", "พิษณุโลก": "plk", "เพชรบุรี": "pbi", "เพชรบูรณ์": "pbn",
            "แพร่": "pre", "ภูเก็ต": "pkt", "มหาสารคาม": "mkm", "มุกดาหาร": "mdh",
            "แม่ฮ่องสอน": "msn", "ยโสธร": "yst", "ยะลา": "yla", "ร้อยเอ็ด": "ret",
            "ระนอง": "rng", "ระยอง": "ryg", "ราชบุรี": "rbr", "ลพบุรี": "lri",
            "ลำปาง": "lpg", "ลำพูน": "lpn", "เลย": "lei", "ศรีสะเกษ": "ssk",
            "สกลนคร": "snk", "สงขลา": "ska", "สตูล": "stn", "สมุทรปราการ": "smp",
            "สมุทรสงคราม": "skm", "สมุทรสาคร": "skn", "สระแก้ว": "sko", "สระบุรี": "sri",
            "สิงห์บุรี": "sbr", "สุโขทัย": "sti", "สุพรรณบุรี": "spb", "สุราษฎร์ธานี": "sni",
            "สุรินทร์": "srn", "หนองคาย": "nki", "หนองบัวลำภู": "nbl", "อ่างทอง": "atg",
            "อำนาจเจริญ": "anc", "อุดรธานี": "udn", "อุตรดิตถ์": "utt", "อุทัยธานี": "uti",
            "อุบลราชธานี": "ubn"
        }
        
        # รวบรวมไฟล์
        soil_files_all = []
        ero_files_all = []
        luse_files_all = []
        
        for prov_thai in crossed_provinces_thai:
            prov_eng = PROV_ENG_MAP_EXTENDED.get(prov_thai, "")
            
            # ดิน
            soil_f = glob.glob(os.path.join(BASE_DIR, "ข้อมูลShp", "00_Soil_1ต่อ25000", f"*{prov_thai}*.shp"))
            if soil_f: soil_files_all.extend(soil_f)
            
            # การชะล้างย้ายไปหาด้วยพิกัด Bounding Box แทนเพื่อความแม่นยำ 100%
            
            # การใช้ประโยชน์ที่ดิน
            luse_f = []
            if prov_eng:
                for pattern in [f"*Luse*{prov_eng}*.shp", f"*LU_{prov_eng.upper()}*.shp", f"*LU_{prov_eng.lower()}*.shp", f"*LandUse*{prov_eng}*.shp"]:
                    luse_f = glob.glob(os.path.join(BASE_DIR, "ข้อมูลShp", "**", pattern), recursive=True)
                    if luse_f: break
            if not luse_f:
                for pattern in [f"*Luse*{prov_thai}*.shp", f"*LU_{prov_thai}*.shp", f"*LandUse*{prov_thai}*.shp"]:
                    luse_f = glob.glob(os.path.join(BASE_DIR, "ข้อมูลShp", "**", pattern), recursive=True)
                    if luse_f: break
                    
            if not luse_f and prov_eng:
                zip_path = os.path.join(BASE_DIR, "Landuse_All", f"Landuse_{prov_eng.lower()}.zip")
                if os.path.exists(zip_path):
                    import zipfile
                    extract_dir = os.path.join(BASE_DIR, "ข้อมูลShp", "LandUse", f"{prov_thai}_{prov_eng}")
                    print(f"Extracting missing LandUse data for {prov_thai} ({prov_eng})...")
                    os.makedirs(extract_dir, exist_ok=True)
                    try:
                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                            zip_ref.extractall(extract_dir)
                        # Retry finding
                        for pattern in [f"*Luse*{prov_eng}*.shp", f"*LU_{prov_eng.upper()}*.shp", f"*LU_{prov_eng.lower()}*.shp", f"*LandUse*{prov_eng}*.shp"]:
                            luse_f = glob.glob(os.path.join(BASE_DIR, "ข้อมูลShp", "**", pattern), recursive=True)
                            if luse_f: break
                    except Exception as e:
                        print(f"Failed to extract {zip_path}: {e}")
                        
            if luse_f: luse_files_all.extend(luse_f)
            
        # ลบไฟล์ซ้ำ
        soil_files_all = list(set(soil_files_all))
        luse_files_all = list(set(luse_files_all))
        
        # --- [NEW] ค้นหาการชะล้างพังทลายด้วย Bounding Box ---
        print("[Spatial Search] กำลังค้นหาไฟล์ 'การชะล้างพังทลาย' จากพิกัด Bounding Box...")
        try:
            import pyogrio
            from pyproj import CRS, Transformer
            road_bounds = engine.road_gdf.geometry.buffer(buffer_m).total_bounds # (minx, miny, maxx, maxy) EPSG:32647
            all_ero_files = glob.glob(os.path.join(BASE_DIR, "ข้อมูลShp", "08_การชะล้างพังทลายของดิน_2563", "**", "*.shp"), recursive=True)
            
            road_crs = engine.road_gdf.crs
            for f in all_ero_files:
                try:
                    info = pyogrio.read_info(f)
                    shp_bounds = info['total_bounds']
                    shp_crs = info.get('crs', None)
                    
                    chk_minx, chk_miny, chk_maxx, chk_maxy = road_bounds
                    if shp_crs and CRS.from_user_input(shp_crs) != CRS.from_user_input(road_crs):
                        transformer = Transformer.from_crs(road_crs, shp_crs, always_xy=True)
                        p1x, p1y = transformer.transform(road_bounds[0], road_bounds[1])
                        p2x, p2y = transformer.transform(road_bounds[2], road_bounds[3])
                        chk_minx, chk_maxx = min(p1x, p2x), max(p1x, p2x)
                        chk_miny, chk_maxy = min(p1y, p2y), max(p1y, p2y)
                        
                    if not (chk_maxx < shp_bounds[0] or chk_minx > shp_bounds[2] or chk_maxy < shp_bounds[1] or chk_miny > shp_bounds[3]):
                        ero_files_all.append(f)
                except Exception:
                    pass
        except Exception as e:
            print(f"Error in spatial search for erosion files: {e}")
            
        ero_files_all = list(set(ero_files_all))
        
        if soil_files_all:
            shapefiles_to_process.append({"path": soil_files_all, "name_col": "SOIL_SERIE", "do_clip": True, "excel_name": "ชุดดิน", "sheet": {"EAR": "7.ชุดดิน", "EC": "7.ชุดดิน"}})
        if ero_files_all:
            shapefiles_to_process.append({"path": ero_files_all, "name_col": "SEV_CLASS", "do_clip": True, "sheet": {"EAR": "3.การชะล้างพังทลาย", "EC": "11.การชะล้างพังทลาย"}})
            shapefiles_to_process.append({"path": ero_files_all, "name_col": "SEV_CLASS", "do_clip": False, "buffer_m_override": 30, "sheet": {"EAR": "3.การชะล้างพังทลาย_เขตทาง", "EC": "11.การชะล้างพังทลาย_เขตทาง"}})
        if luse_files_all:
            shapefiles_to_process.append({"path": luse_files_all, "name_col": "LU_NAME", "do_clip": True, "excel_name": "LU", "sheet": {"EAR": "8.การใช้ประโยชน์ที่ดิน", "EC": "4.การใช้ประโยชน์ที่ดิน"}})
        
        all_gis_results = {}
        
        for shp_info in shapefiles_to_process:
            shp_path = shp_info["path"]
            is_valid_path = False
            
            if isinstance(shp_path, list) and len(shp_path) > 0:
                is_valid_path = True
                base_name = os.path.basename(shp_path[0]).replace('.shp', '')
                if len(shp_path) > 1:
                    # e.g., LU_LPG_2564 -> LU_merged
                    parts = base_name.split('_')
                    base_name = f"{parts[0]}_merged" if len(parts) > 0 else "merged_shp"
            elif isinstance(shp_path, str) and shp_path.startswith('http'):
                is_valid_path = True
                base_name = shp_info.get("excel_name", shp_path.split('/')[-2])
            elif isinstance(shp_path, str) and os.path.exists(shp_path):
                is_valid_path = True
                base_name = os.path.basename(shp_path).replace('.shp', '')
                
            if is_valid_path:
                
                # 4.1 คำนวณจุดตัดเพื่อเอาไปหยอด Excel
                custom_buffer = shp_info.get("buffer_m_override", buffer_m)
                
                # สำหรับ Point layers (เช่น สถานศึกษา, หมู่บ้าน, โบราณสถาน) ให้แนบ crossed_provinces ไปกรองด้วย
                is_point_layer = any(kw in str(shp_path) for kw in ['ตำแหน่ง', 'หมู่บ้าน', 'historic', 'โรงพยาบาล', 'วัด', 'รพสต'])
                if is_point_layer:
                    engine.crossed_provinces = crossed_provinces_thai
                else:
                    engine.crossed_provinces = []
                    
                df = engine.calculate_intersections(shp_path, shp_info["name_col"], custom_buffer)
                
                # Fallback for landslide if REST API fails or returns no data
                if df.empty and shp_info.get("excel_name") == "ดินถล่ม" and isinstance(shp_path, str) and shp_path.startswith('http'):
                    fallback_shp = os.path.join(BASE_DIR, "ข้อมูลShp", "shp_พื้นที่อ่อนไหวต่อการเกิดแผ่นดินถล่ม", "SHP_05010103 พื้นที่อ่อนไหวต่อการเกิดแผ่นดินถล่ม", "LANDSLIDE_SUSCEPTIBILITY.shp")
                    if os.path.exists(fallback_shp):
                        print(f"  [Fallback] REST API ดินถล่มไม่พบข้อมูล หรือเชื่อมต่อไม่ได้ ลองใช้ไฟล์ Local แทน: {fallback_shp}")
                        df = engine.calculate_intersections(fallback_shp, "Level_T", custom_buffer)
                
                # 4.2 ตัด Shapefile สำหรับทำ Layout (เซฟลง Data_layout)
                layout_shp_dir = os.path.join(DATA_LAYOUT_DIR, proj_name, "Shp")
                layout_excel_dir = os.path.join(DATA_LAYOUT_DIR, proj_name, "Excel")
                
                layout_shp_path = os.path.join(layout_shp_dir, f"{base_name}_clipped.shp")
                excel_path = None
                if "excel_name" in shp_info:
                    excel_path = os.path.join(layout_excel_dir, f'{shp_info["excel_name"]}.xlsx')
                    
                do_clip = shp_info.get("do_clip", True)
                if do_clip:
                    engine.clip_and_save_shapefile(shp_path, buffer_m, layout_shp_path, excel_path=excel_path, do_clip=do_clip)
                
                # 5. ระบบ Cache (STEP 5)
                if not df.empty:
                    df['KM In'] = df.apply(lambda r: "-" if r['km_in_m'] == 0 and r['km_out_m'] == 0 and r.get('length_m', 0) == 0 else engine.format_km(r['km_in_m']), axis=1)
                    df['KM Out'] = df.apply(lambda r: "-" if r['km_in_m'] == 0 and r['km_out_m'] == 0 and r.get('length_m', 0) == 0 else engine.format_km(r['km_out_m']), axis=1)
                    
                    cache_dir = os.path.join(project_path, "cache")
                    os.makedirs(cache_dir, exist_ok=True)
                    
                    records = df.to_dict(orient="records")
                    sanitized_records = sanitize_json(records)
                    
                    # รวมผลลัพธ์แยกตามชีทปลายทาง
                    sheet_map = shp_info.get("sheet", {})
                    sheet_name = sheet_map.get(settings["type"])
                    
                    # Skip if sheet is not required for this mode
                    if sheet_name is None:
                        continue
                        
                    if sheet_name not in all_gis_results:
                        all_gis_results[sheet_name] = []
                        
                    # แทร็กว่ามาจาก shapefile ไหน (เผื่อชีทนึงมีหลายข้อมูล)
                    for r in sanitized_records:
                        r['source_shp'] = base_name
                        all_gis_results[sheet_name].append(r)
                        
                    print(f"เพิ่มข้อมูลจาก {base_name} เข้าไปเตรียมหยอดชีท {sheet_name} แล้ว")
                else:
                    print(f"ไม่พบจุดตัดกับ {os.path.basename(shp_path)}")
            else:
                print(f"ไม่พบไฟล์ {shp_path}")
                
        # เซฟ Cache รวมสำหรับ Excel Reporter
        cache_dir = os.path.join(project_path, "cache")
        os.makedirs(cache_dir, exist_ok=True)
        master_json_path = os.path.join(cache_dir, "master_gis_results.json")
        with open(master_json_path, 'w', encoding='utf-8') as f:
            json.dump(all_gis_results, f, ensure_ascii=False, indent=2)
            
        # --- NEW LOGIC: Copy files to match layout structure ---
        import shutil
        layout_root = os.path.join(DATA_LAYOUT_DIR, proj_name)
        layout_shp_dir = os.path.join(layout_root, "Shp")
        layout_map_dir = os.path.join(layout_root, "Map")
        os.makedirs(layout_shp_dir, exist_ok=True)
        os.makedirs(layout_map_dir, exist_ok=True)
        
        # Save buffer shapefile (แก้บั๊กเส้นม้วน/ซ้อนทับกัน ด้วย union_all)
        buffer_radius_str = "1กิโลเมตร" if buffer_m == 1000 else f"{int(buffer_m)}เมตร"
        buffer_shp_path = os.path.join(layout_shp_dir, f"รัศมี{buffer_radius_str}.shp")
        try:
            import geopandas as gpd
            merged_geom = engine.road_gdf.copy().geometry.buffer(buffer_m).union_all()
            merged_gdf = gpd.GeoDataFrame(geometry=[merged_geom], crs=engine.road_gdf.crs)
            merged_gdf.to_file(buffer_shp_path)
            print(f"บันทึกไฟล์รัศมี Buffer (Unioned) ไว้ที่ {buffer_shp_path}")
        except Exception as e:
            print(f"Error saving buffer: {e}")
        
        # Copy original shapefiles to Shp
        for shp_ext in [".shp", ".shx", ".dbf", ".prj", ".cpg", ".sbn", ".sbx", ".shp.xml"]:
            for pattern in ["*_CL", "*_KM", "*_ROW", "จุดเริ่มต้นสิ้นสุดโครงการ*", "หลักกิโลเมตร*"]:
                for file_path in glob.glob(os.path.join(project_path, f"{pattern}{shp_ext}")):
                    shutil.copy2(file_path, os.path.join(layout_shp_dir, os.path.basename(file_path)))
                    
        # Copy KMZ and MXD to root
        for ext in ["*.kmz", "*.mxd"]:
            for file_path in glob.glob(os.path.join(project_path, ext)):
                shutil.copy2(file_path, os.path.join(layout_root, os.path.basename(file_path)))
        # --------------------------------------------------------
                
        # 6. รายงานผลลง Excel (STEP 4)
        from excel_reporter import ExcelReporter
        reporter = ExcelReporter(project_path)
        reporter.write_report()
        
        # 7. ระบบดึงภาพ Street View
        from streetview_fetcher import StreetViewFetcher
        fetcher = StreetViewFetcher(project_path)
        fetcher.run_batch_from_gis_cache()
        
        # เก็บสถิติสรุปผล
        summary_stats["successful_projects"] += 1
            
    # สร้าง Run Summary
    from datetime import datetime
    end_time = datetime.now()
    duration = end_time - start_time
    
    summary_path = os.path.join(DATA_LAYOUT_DIR, "run_summary.txt")
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("=== รายงานสรุปผลการทำงาน (Run Summary) ===\n")
        f.write(f"เวลาเริ่ม: {start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"เวลาสิ้นสุด: {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"ใช้เวลาทั้งหมด: {duration}\n")
        f.write(f"จำนวนโปรเจกต์ทั้งหมด: {summary_stats['total_projects']}\n")
        f.write(f"สำเร็จ: {summary_stats['successful_projects']}\n")
        f.write(f"ล้มเหลว: {summary_stats['failed_projects']}\n")
        f.write("==============================================\n")
        
    print(f"\n=== กระบวนการเสร็จสิ้น! บันทึกรายงานสรุปที่: {summary_path} ===")

if __name__ == "__main__":
    run_pipeline()
