import os
import glob
import shutil
import json
import re
import xlwings as xw
import geopandas as gpd
from config import DATA_LAYOUT_DIR, TARGET_CRS

# แมพตัวย่อจังหวัด -> ชื่อจังหวัดเต็ม เพื่อป้องกันข้อมูลเทมเพลตหลุดมาผิดพื้นที่
PROVINCE_MAP = {
    "มส": "แม่ฮ่องสอน", "ลป": "ลำปาง", "ชม": "เชียงใหม่",
    "ชร": "เชียงราย", "พย": "พะเยา", "นน": "น่าน",
    "พร": "แพร่", "ตาก": "ตาก", "ลพ": "ลำพูน",
    "กจ": "กาญจนบุรี", "นม": "นครราชสีมา", "ขก": "ขอนแก่น",
    "อด": "อุดรธานี", "สร": "สุรินทร์", "นพ": "นครพนม",
    "สน": "สกลนคร", "มค": "มหาสารคาม", "รอ": "ร้อยเอ็ด",
    "อบ": "อุบลราชธานี", "ศก": "ศรีสะเกษ", "บร": "บุรีรัมย์",
    "ชย": "ชัยภูมิ", "นศ": "นครศรีธรรมราช", "สฎ": "สุราษฎร์ธานี",
    "พง": "พังงา", "กบ": "กระบี่", "ตง": "ตรัง", "สข": "สงขลา",
}

class ExcelReporter:
    def __init__(self, project_dir):
        self.project_dir = project_dir
        self.project_name = os.path.basename(project_dir)
        self.cache_dir = os.path.join(project_dir, "cache")
        self.output_dir = os.path.join(DATA_LAYOUT_DIR, self.project_name)
        
        # ดึงจังหวัดจากชื่อโปรเจกต์เพื่อป้องกันข้อมูลเทมเพลตผิดพื้นที่
        self.project_province = ""
        for part in self.project_name.split('_'):
            clean_part = re.sub(r'[\d\.]', '', part).strip()
            if clean_part in PROVINCE_MAP:
                self.project_province = PROVINCE_MAP[clean_part]
                break
        if self.project_province:
            print(f"[ExcelReporter] พื้นที่โครงการ: จังหวัด{self.project_province}")

    def prepare_output_file(self):
        os.makedirs(self.output_dir, exist_ok=True)
        backup_dir = os.path.join(self.project_dir, "_backup")
        os.makedirs(backup_dir, exist_ok=True)

        excel_files = glob.glob(os.path.join(self.project_dir, "*.xlsx"))
        excel_files = [f for f in excel_files if "_backup" not in f and not os.path.basename(f).startswith("test")]

        template_path = None
        expected_name = f"{self.project_name}.xlsx"
        for f in excel_files:
            if os.path.basename(f) == expected_name:
                template_path = f
                break
        
        if not template_path and excel_files:
            template_path = excel_files[0]

        if not template_path:
            print("ไม่พบไฟล์ Excel ต้นฉบับ... ดึงจาก Master Template อัตโนมัติ!")
            from config import BASE_DIR
            master_dir = os.path.join(BASE_DIR, "Excel_Templates")
            if "EAR" in self.project_name.upper():
                master_file = os.path.join(master_dir, "Template_EAR.xlsx")
            else:
                master_file = os.path.join(master_dir, "Template_EC.xlsx")
            if os.path.exists(master_file):
                new_excel_path = os.path.join(self.project_dir, expected_name)
                shutil.copy(master_file, new_excel_path)
                template_path = new_excel_path
            else:
                return None

        base_name = os.path.basename(template_path)
        output_path = os.path.join(self.output_dir, base_name)

        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(backup_dir, f"{base_name.replace('.xlsx', '')}_{timestamp}.xlsx")
        import time
        max_retries = 3
        for attempt in range(max_retries):
            try:
                shutil.copy2(template_path, backup_path)
                shutil.copy2(template_path, output_path)
                break
            except PermissionError:
                if attempt < max_retries - 1:
                    time.sleep(1)
                else:
                    print(f"\n[WARNING] ไฟล์ผลลัพธ์เดิมถูกเปิดค้างอยู่ จะทำการบันทึกเป็นไฟล์ใหม่แทน!")
                    output_path = output_path.replace(".xlsx", f"_{timestamp}.xlsx")
                    shutil.copy2(template_path, output_path)
        return output_path

    def load_cache(self):
        json_path = os.path.join(self.cache_dir, "master_gis_results.json")
        if not os.path.exists(json_path):
            return {}
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _safe_clear(self, sht, col, start_row, end_row=50):
        """ล้างข้อมูลทีละเซลล์ เพื่อเลี่ยงการล้างฟอร์แมตขอบตาราง"""
        for r in range(start_row, end_row + 1):
            try:
                sht.range(f'{col}{r}').clear_contents()
            except Exception:
                pass

    def _set_cell_value_only(self, sht, cell_ref, value):
        """เขียนค่าลงเซลล์โดยไม่เปลี่ยนฟอนต์เพื่อรักษาฟอร์แมตของเทมเพลต"""
        try:
            sht.range(cell_ref).value = value
        except Exception as e:
            print(f"  Warning: Cannot write to {cell_ref}: {e}")

    def _find_sheet(self, wb, partial_name):
        """หาชีทที่ชื่อมี partial_name"""
        for sht in wb.sheets:
            if partial_name in sht.name:
                try:
                    # จัดแนวตั้งกึ่งกลางสำหรับทั้งชีทเพื่อความสมบูรณ์และสวยงาม
                    sht.used_range.api.VerticalAlignment = -4108
                except Exception:
                    pass
                return sht
        return None

    def _clear_template_defaults(self, wb, is_ear):
        """ล้างข้อมูลตัวอย่างจากเทมเพลตทุกชีตก่อนเขียนข้อมูลจริง
        ป้องกันข้อมูลเทมเพลต (เช่น จังหวัดลำปาง) หลุดมาผิดพื้นที่"""
        print(f"  ล้างข้อมูลตัวอย่างจากเทมเพลต (ป้องกันข้อมูลผิดพื้นที่)...")
        
        if not is_ear:
            # ชีต 3: ตรวจสอบสถานศึกษา - ล้างข้อมูลตัวอย่างทั้ง 3 ตาราง
            sht3 = self._find_sheet(wb, "3.ตรวจสอบสถานศึกษา")
            if sht3:
                for col in ['A', 'B', 'D', 'E', 'F']:
                    self._safe_clear(sht3, col, 3, 17)   # โรงเรียน
                    self._safe_clear(sht3, col, 20, 30)   # ศาสนสถาน
                    self._safe_clear(sht3, col, 33, 42)   # สถานพยาบาล
            
            # ชีต 6: แหล่งทรัพยากรทางธรณี - ล้างข้อมูลตัวอย่าง
            sht6 = self._find_sheet(wb, "6.แหล่งทรัพยากรทางธรณี")
            if sht6:
                for col in ['A', 'B', 'C', 'D']:
                    self._safe_clear(sht6, col, 2, 50)
                    
            # ชีต 8: แผ่นดินไหว - ล้างข้อมูลตัวอย่าง (ลบคำว่า "จังหวัดลำปาง" ฯลฯ)
            sht8 = self._find_sheet(wb, "8.เสี่ยงต่อการเกิดแผ่นดินไหว")
            if sht8:
                self._safe_clear(sht8, 'A', 1, 15)

    def _get_buffer_areas(self):
        """หาพื้นที่บัฟเฟอร์รวมของโครงการ (1000ม. และ 30ม.)"""
        cl_files = glob.glob(os.path.join(self.project_dir, "*_CL.shp"))
        if not cl_files:
            return 0.0, 0.0
        try:
            road_gdf = gpd.read_file(cl_files[0]).to_crs(TARGET_CRS)
            # 1000m Buffer
            buf_1000 = road_gdf.copy()
            buf_1000['geometry'] = buf_1000.geometry.buffer(1000)
            area_1000 = buf_1000.geometry.unary_union.area
            # 30m Buffer
            buf_30 = road_gdf.copy()
            buf_30['geometry'] = buf_30.geometry.buffer(30)
            area_30 = buf_30.geometry.unary_union.area
            return area_1000, area_30
        except Exception as e:
            print(f"Error calculating buffer areas: {e}")
            return 0.0, 0.0

    def _prepare_rows(self, sht, start_row, end_row, target_count):
        """
        เตรียมแถวสำหรับเขียนข้อมูล:
        - หากข้อมูลมากกว่าจำนวนแถวที่มีในเทมเพลต จะแทรกแถวเพิ่มและก๊อปปี้ฟอร์แมตลงมา
        - คืนค่าแถวสุดท้าย (new_end_row) เพื่อใช้ในการเคลียร์ข้อมูลและเขียนต่อ
        """
        styled_rows_count = end_row - start_row + 1
        if target_count > styled_rows_count:
            needed = target_count - styled_rows_count
            print(f"  แทรกแถวเพิ่มสำหรับตาราง {needed} แถว...")
            for _ in range(needed):
                # แทรกแถวใหม่ที่ถัดจาก end_row โดยใช้ฟอร์แมตของแถว end_row
                insert_row_idx = end_row + 1
                sht.range(f"{insert_row_idx}:{insert_row_idx}").insert(
                    shift='down', 
                    copy_origin='format_from_left_or_above'
                )
                # ล้างข้อมูลที่อาจก๊อปปี้มาด้วย ให้เหลือแต่ฟอร์แมต/ขอบตาราง
                sht.range(f"{insert_row_idx}:{insert_row_idx}").clear_contents()
                end_row += 1
        return end_row

    # ========================================
    # ฟังก์ชันเขียนข้อมูลลงแต่ละชีท
    # ========================================

    def _apply_thin_borders(self, sht, cell_ref):
        """ใส่เส้นขอบบาง (Thin) ให้เซลล์เพื่อให้เหมือนเทมเพลต"""
        try:
            cell_api = sht.range(cell_ref).api
            # xlEdgeLeft=7, xlEdgeTop=8, xlEdgeBottom=9, xlEdgeRight=10
            # xlInsideVertical=11, xlInsideHorizontal=12
            # xlThin=2, xlContinuous=1
            for edge in (7, 8, 9, 10):
                border = cell_api.Borders(edge)
                border.LineStyle = 1  # xlContinuous
                border.Weight = 2     # xlThin
        except Exception:
            pass

    def _write_pa_table_sequential(self, sht, records, start_col):
        """เขียนตารางป่า (ป่าสงวน หรือ ป่าถาวร) แบบเรียงแถวลงมาใหม่"""
        cols = {
            'n': start_col,
            'sqkm': chr(ord(start_col) + 1),
            'sqm': chr(ord(start_col) + 2),
            'rai': chr(ord(start_col) + 3),
            'k': chr(ord(start_col) + 4),
            'd': chr(ord(start_col) + 5)
        }
        
        # ล้างข้อมูลเดิมทั้งหมดในแถว 3-30 (โดยไม่ลบฟอร์แมตขอบตาราง)
        for col_letter in cols.values():
            self._safe_clear(sht, col_letter, 3, 30)
            
        if not records:
            return
            
        # จัดกลุ่มตาม area_name
        grouped = {}
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''):
                a_name = 'ไม่ระบุ'
                
            props = r.get('properties', {})
            nrf_zone = props.get('NRF_Zone')
            if nrf_zone and str(nrf_zone) not in ('None', 'nan', ''):
                a_name = f"{a_name} Zone {nrf_zone}"
                
            if a_name not in grouped:
                grouped[a_name] = {'kms': [], 'length_km': 0, 'area_sqm': 0}
            km_str = f"{r['KM In']} - {r['KM Out']}"
            grouped[a_name]['kms'].append(km_str)
            grouped[a_name]['length_km'] += r.get('length_m', 0) / 1000
            grouped[a_name]['area_sqm'] += r.get('intersect_area_sqm', 0)
            
        row = 3
        for a_name, data in grouped.items():
            area_sqm = data['area_sqm']
            area_sqkm = area_sqm / 1_000_000 if area_sqm else 0
            area_rai = area_sqm / 1600 if area_sqm else 0
            
            self._set_cell_value_only(sht, f"{cols['n']}{row}", a_name)
            self._set_cell_value_only(sht, f"{cols['sqkm']}{row}", round(area_sqkm, 6) if area_sqkm else '')
            self._set_cell_value_only(sht, f"{cols['sqm']}{row}", round(area_sqm, 2) if area_sqm else '')
            self._set_cell_value_only(sht, f"{cols['rai']}{row}", round(area_rai, 6) if area_rai else '')
            self._set_cell_value_only(sht, f"{cols['k']}{row}", "\n".join(data['kms']))
            
            dist_list = []
            for r2 in records:
                rn = str(r2.get('area_name', ''))
                if rn in ('None', 'nan', ''):
                    rn = 'ไม่ระบุ'
                
                props2 = r2.get('properties', {})
                nrf_zone2 = props2.get('NRF_Zone')
                if nrf_zone2 and str(nrf_zone2) not in ('None', 'nan', ''):
                    rn = f"{rn} Zone {nrf_zone2}"
                
                if rn == a_name:
                    dist_list.append(f"{r2.get('length_m', 0)/1000:.3f}")
            self._set_cell_value_only(sht, f"{cols['d']}{row}", "\n".join(dist_list))
            
            for c_letter in cols.values():
                try:
                    cell = sht.range(f"{c_letter}{row}")
                    cell.api.Font.Name = 'TH Sarabun New'
                    cell.api.Font.Size = 14
                    
                    # ตั้งค่าการจัดหน้าให้เหมือนเทมเพลต (จัดกลางแนวตั้ง)
                    cell.api.VerticalAlignment = -4108 # xlVAlignCenter
                    if c_letter == cols['n']:
                        cell.api.HorizontalAlignment = -4131 # xlHAlignLeft
                    else:
                        cell.api.HorizontalAlignment = -4108 # xlHAlignCenter
                        
                    # ตั้งค่าปัดบรรทัด (Wrap Text) สำหรับคอลัมน์กม.และระยะทางที่มีหลายบรรทัด
                    if c_letter in (cols['k'], cols['d']):
                        cell.api.WrapText = True
                        
                except Exception:
                    pass
                self._apply_thin_borders(sht, f"{c_letter}{row}")
                
            # ขยายขนาดแถวอัตโนมัติตามเนื้อหา (AutoFit เฉพาะความสูงแถว)
            try:
                sht.range(f"{row}:{row}").autofit('r')
            except Exception:
                pass
                
            row += 1

    def _write_sheet_pa(self, wb, records_nrf=None, records_permanent=None):
        """ชีท 1.ป่า"""
        sht = self._find_sheet(wb, "1.ป่า")
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")
        self._write_pa_table_sequential(sht, records_nrf, 'A')
        self._write_pa_table_sequential(sht, records_permanent, 'H')

    def _write_forest_status_table(self, sht, records, total_buffer_area_sqm, start_row):
        """เขียนตารางพื้นที่คงสภาพป่า (รัศมี 1 กม. หรือ ในเขตทาง)"""
        self._safe_clear(sht, 'B', start_row, start_row + 2)
        self._safe_clear(sht, 'C', start_row, start_row + 2)

        if total_buffer_area_sqm <= 0:
            return

        forest_area_sqm = sum(r.get('intersect_area_sqm', 0) for r in records) if records else 0.0
        forest_area_sqm = min(forest_area_sqm, total_buffer_area_sqm)
        non_forest_area_sqm = total_buffer_area_sqm - forest_area_sqm

        self._set_cell_value_only(sht, f'B{start_row}', round(forest_area_sqm / 1_000_000, 6))
        self._set_cell_value_only(sht, f'C{start_row}', round(forest_area_sqm / 1600, 2))

        self._set_cell_value_only(sht, f'B{start_row+1}', round(non_forest_area_sqm / 1_000_000, 6))
        self._set_cell_value_only(sht, f'C{start_row+1}', round(non_forest_area_sqm / 1600, 2))

        self._set_cell_value_only(sht, f'B{start_row+2}', round(total_buffer_area_sqm / 1_000_000, 6))
        self._set_cell_value_only(sht, f'C{start_row+2}', round(total_buffer_area_sqm / 1600, 2))

    def _write_sheet_forest_status(self, wb, records_1km, records_30m):
        """ชีท 4.พื้นที่คงสภาพป่า"""
        sht = self._find_sheet(wb, "4.พื้นที่คงสภาพป่า")
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        area_1000_sqm, area_30_sqm = self._get_buffer_areas()
        self._write_forest_status_table(sht, records_1km, area_1000_sqm, 2)
        self._write_forest_status_table(sht, records_30m, area_30_sqm, 9)

    def _write_sheet_watershed_lookup(self, wb, records, is_ear=True):
        """ชีท 10.ชั้นคุณภาพลุ่มน้ำ / 12.ชั้นคุณภาพลุ่มน้ำ (ใช้แบบ lookup เพราะระดับ 1A, 1B, 2... นั้นคงที่ในตาราง)"""
        sheet_name = "10.ชั้นคุณภาพลุ่มน้ำ" if is_ear else "12.ชั้นคุณภาพลุ่มน้ำ"
        sht = self._find_sheet(wb, sheet_name)
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        # หาแถวรวม (total_row) ก่อนทำการเคลียร์ข้อมูลในช่อง B, C, D
        total_row = None
        for row in range(3, 20):
            val_a = sht.range(f"A{row}").value
            val_b = sht.range(f"B{row}").value
            if val_a is None and val_b is not None:
                total_row = row
                break
            if val_a and 'รวม' in str(val_a):
                total_row = row
                break

        self._safe_clear(sht, 'B', 3, 20)
        self._safe_clear(sht, 'C', 3, 20)
        self._safe_clear(sht, 'D', 3, 20)
        self._safe_clear(sht, 'E', 3, 20)
        self._safe_clear(sht, 'F', 3, 20)

        if not records:
            return

        grouped = {}
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''):
                a_name = 'ไม่ระบุ'
            if a_name not in grouped:
                grouped[a_name] = {'kms': [], 'area_sqm': 0, 'length_km': 0.0}
                
            if is_ear:
                km_str = f"{r.get('KM In', '')}, {r.get('KM Out', '')}".strip(', ')
            else:
                km_in = r.get('KM In', '').strip()
                km_str = f"กม.ที่ {km_in}" if km_in else ""
                
            if km_str and km_str not in grouped[a_name]['kms']:
                grouped[a_name]['kms'].append(km_str)
            grouped[a_name]['area_sqm'] += r.get('intersect_area_sqm', 0)
            grouped[a_name]['length_km'] += r.get('length_m', 0) / 1000

        total_sqkm = 0
        total_sqm = 0
        total_rai = 0
        total_length_km = 0.0
        
        rows_to_delete = []

        for row in range(3, 20):
            if row == total_row:
                continue
            val_a = sht.range(f"A{row}").value
            if not val_a:
                continue
            cell_str = str(val_a).strip()

            matched_class = None
            for g_class in grouped.keys():
                if g_class.lower() in cell_str.lower():
                    matched_class = g_class
                    break

            if matched_class:
                data = grouped[matched_class]
                area_sqm = data['area_sqm']
                area_sqkm = area_sqm / 1_000_000
                area_rai = area_sqm / 1600

                self._set_cell_value_only(sht, f"B{row}", round(area_sqkm, 6))
                self._set_cell_value_only(sht, f"C{row}", round(area_sqm, 2))
                self._set_cell_value_only(sht, f"D{row}", round(area_rai, 6))
                
                if is_ear:
                    self._set_cell_value_only(sht, f"E{row}", "") # leave blank or calculate percentage
                    self._set_cell_value_only(sht, f"F{row}", "\n".join(data['kms']))
                else:
                    self._set_cell_value_only(sht, f"E{row}", "\n".join(data['kms']))

                total_sqkm += area_sqkm
                total_sqm += area_sqm
                total_rai += area_rai
                total_length_km += data['length_km']
            else:
                if is_ear:
                    rows_to_delete.append(row)
                else:
                    # In EC, don't delete. Write 'ไม่ตัดผ่านแนวเส้น'
                    self._set_cell_value_only(sht, f"B{row}", "-")
                    self._set_cell_value_only(sht, f"C{row}", "-")
                    self._set_cell_value_only(sht, f"D{row}", "-")
                    self._set_cell_value_only(sht, f"E{row}", "ไม่ตัดผ่านแนวเส้น")

        if total_row:
            self._set_cell_value_only(sht, f"B{total_row}", round(total_sqkm, 6))
            self._set_cell_value_only(sht, f"C{total_row}", round(total_sqm, 2))
            self._set_cell_value_only(sht, f"D{total_row}", round(total_rai, 6))
            
        for r in reversed(rows_to_delete):
            sht.range(f"{r}:{r}").delete(shift='up')

    def _write_sheet_villages_sequential(self, wb, records):
        """เขียนชีท 2.พื้นที่หมู่บ้าน แบบเรียงแถวลงมาใหม่ เติมข้อมูลให้เต็มทุกคอลัมน์ และแทรกแถวเพิ่มหากพื้นที่ไม่พอ"""
        sht = self._find_sheet(wb, "2.พื้นที่หมู่บ้าน")
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        if not records:
            # ถ้าไม่มีข้อมูล ให้ล้างข้อมูลเดิมในแถว 3-50
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                self._safe_clear(sht, col, 3, 50)
            return

        # จัดกลุ่มตามชื่อหมู่บ้าน
        grouped = {}
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''):
                continue
            if a_name not in grouped:
                props = r.get('properties', {})
                grouped[a_name] = {
                    'moo': props.get('VILL_NO', '01'),
                    'tambon': props.get('TAMBOL_TH', props.get('TAMBON', '')),
                    'amphoe': props.get('AMPHOE_TH', props.get('AMPHOE', '')),
                    'province': props.get('PROV_TH', props.get('PROVINCE', ''))
                }

        # แทรกแถวเพิ่มหากข้อมูลมากกว่า 48 หมู่บ้าน (แถว 3-50)
        end_row = self._prepare_rows(sht, 3, 50, len(grouped))

        # ล้างข้อมูลเดิมทั้งหมดในพื้นที่ทำงาน (แถว 3 ถึง end_row) เฉพาะคอลัมน์ A-G
        clear_end = max(100, end_row)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            self._safe_clear(sht, col, 3, clear_end)

        row = 3
        idx = 1
        for a_name, data in grouped.items():
            self._set_cell_value_only(sht, f"A{row}", idx)
            self._set_cell_value_only(sht, f"B{row}", data['moo'])
            self._set_cell_value_only(sht, f"C{row}", a_name)
            
            tb = data['tambon']
            admin_str = f"อบต.{tb}" if tb else ""
            self._set_cell_value_only(sht, f"D{row}", admin_str)
            
            self._set_cell_value_only(sht, f"E{row}", tb)
            self._set_cell_value_only(sht, f"F{row}", data['amphoe'])
            self._set_cell_value_only(sht, f"G{row}", data['province'])
            
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                try:
                    cell = sht.range(f"{col}{row}")
                    cell.api.Font.Name = 'TH Sarabun New'
                    cell.api.Font.Size = 14
                    cell.row_height = 15
                except Exception:
                    pass
                self._apply_thin_borders(sht, f"{col}{row}")
            row += 1
            idx += 1

    def _write_sheet_historic_sequential(self, wb, records, is_ear=True):
        """เขียนชีท แหล่งโบราณสถาน แบบเรียงแถวลงมาใหม่ และแทรกแถวเพิ่มหากพื้นที่ไม่พอ"""
        sheet_name = "9.แหล่งโบราณสถาน" if is_ear else "5.แหล่งโบราณสถาน"
        sht = self._find_sheet(wb, sheet_name)
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        if not records:
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                self._safe_clear(sht, col, 3, 20)
            return

        grouped = {}
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''):
                continue
            if a_name not in grouped:
                props = r.get('properties', {})
                grouped[a_name] = {
                    'type': props.get('HISGEN', ''),
                    'status': props.get('REGSTATUS', 'ขึ้นทะเบียนแล้ว'),
                    'tambon': props.get('TAMBOL_TH', props.get('TAMBON', '')),
                    'amphoe': props.get('AMPHOE_TH', props.get('AMPHOE', '')),
                    'province': props.get('PROV_TH', props.get('PROVINCE', ''))
                }

        # แทรกแถวเพิ่มหากข้อมูลมากกว่า 18 รายการ (แถว 3-20)
        end_row = self._prepare_rows(sht, 3, 20, len(grouped))

        clear_end = max(100, end_row)
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            self._safe_clear(sht, col, 3, clear_end)

        row = 3
        idx = 1
        for a_name, data in grouped.items():
            self._set_cell_value_only(sht, f"A{row}", idx)
            self._set_cell_value_only(sht, f"B{row}", a_name)
            self._set_cell_value_only(sht, f"C{row}", data['type'])
            self._set_cell_value_only(sht, f"D{row}", data['status'])
            self._set_cell_value_only(sht, f"E{row}", data['tambon'])
            self._set_cell_value_only(sht, f"F{row}", data['amphoe'])
            self._set_cell_value_only(sht, f"G{row}", data['province'])

            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                try:
                    cell = sht.range(f"{col}{row}")
                    cell.api.Font.Name = 'TH Sarabun New'
                    cell.api.Font.Size = 14
                    cell.row_height = 15
                except Exception:
                    pass
                self._apply_thin_borders(sht, f"{col}{row}")
            row += 1
            idx += 1

    def _write_sheet_mineral_ec(self, wb, records):
        """ชีท 6.แหล่งทรัพยากรทางธรณี (EC เท่านั้น)"""
        sht = self._find_sheet(wb, "6.แหล่งทรัพยากรทางธรณี")
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")
        
        # Clear existing
        for col in ['A', 'B', 'C', 'D']:
            self._safe_clear(sht, col, 2, 50)
            
        if not records:
            return
            
        row = 2
        for r in records:
            props = r.get('properties', {})
            # Name from COMNAME_T
            name = props.get('COMNAME_T', 'ไม่ระบุ')
            desc = props.get('MIN_GEO', '')
            
            self._set_cell_value_only(sht, f"A{row}", name)
            self._set_cell_value_only(sht, f"B{row}", "www.dmr.go.th")
            row += 1
            
            self._set_cell_value_only(sht, f"A{row}", desc)
            row += 1
            
    def _write_sheet_earthquake_ec(self, wb, records):
        """ชีท 8.เสี่ยงต่อการเกิดแผ่นดินไหว (EC เท่านั้น)"""
        sht = self._find_sheet(wb, "8.เสี่ยงต่อการเกิดแผ่นดินไหว")
        if not sht: return
        
        # ล้างข้อความเดิมในคอลัมน์ A (เพื่อลบคำว่า จังหวัดลำปาง ในเทมเพลต)
        self._safe_clear(sht, 'A', 1, 10)
        
        if not records: return
        
        # ดึงข้อความที่ไม่ซ้ำจาก API (area_name)
        areas = list(set(r.get('area_name') for r in records if r.get('area_name') and r.get('area_name') not in ('None', 'nan', '')))
        if areas:
            self._set_cell_value_only(sht, "A1", "พื้นที่เสี่ยงต่อการเกิดแผ่นดินไหว")
            row = 2
            for a in areas:
                self._set_cell_value_only(sht, f"A{row}", a)
                row += 1

    def _write_sheet_landslide_lookup(self, wb, records, is_ear=True):
        """ชีท 7.เสี่ยงต่อการเกิดดินถล่ม (EC เท่านั้น)"""
        # ผู้ใช้แจ้งว่าไม่ต้องการตัวเลขกิโลเมตร ต้องการแค่ตัวหนังสือแบบเทมเพลต
        # ดังนั้นข้ามการหยอดตัวเลขในโหมด EC
        pass

    def _prepare_sub_rows(self, sht, start_row, end_row, target_count):
        """
        เตรียมแถวสำหรับตารางย่อย แทรกแถวระหว่าง start_row และ end_row
        และคืนค่า end_row ใหม่หลังแทรก เพื่อเลี่ยงการทับซ้อนกับตารางถัดไปที่อยู่ข้างใต้
        """
        current_count = end_row - start_row + 1
        if target_count > current_count:
            needed = target_count - current_count
            print(f"  แทรกแถวเพิ่ม {needed} แถว ในช่วงแถว {start_row}-{end_row}...")
            for _ in range(needed):
                # แทรกแถวใหม่ที่ end_row เพื่อขยายตาราง
                sht.range(f"{end_row}:{end_row}").insert(
                    shift='down',
                    copy_origin='format_from_left_or_above'
                )
                sht.range(f"{end_row}:{end_row}").clear_contents()
                end_row += 1
        return end_row

    def _write_sheet_sensitive_sequential(self, wb, master_records):
        """เขียนชีท พื้นที่อ่อนไหว (สำหรับ EAR เท่านั้น) - แสดงรายชื่อสถานศึกษา ศาสนสถาน และสถานพยาบาล"""
        sht = self._find_sheet(wb, "พื้นที่อ่อนไหว")
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        # ดึงข้อมูลแยกตารางจาก master_records
        school_records = master_records.get('พื้นที่อ่อนไหว_สถานศึกษา', [])
        religion_records = master_records.get('พื้นที่อ่อนไหว_ศาสนสถาน', [])
        hospital_records = master_records.get('พื้นที่อ่อนไหว_สถานพยาบาล', [])

        # ฟังก์ชันจัดกลุ่มและจัดทำข้อมูลโรงเรียน/ศาสนสถาน/โรงพยาบาล
        def get_grouped_data(records):
            grouped = {}
            for r in records:
                a_name = str(r.get('area_name', 'ไม่ระบุ'))
                if a_name in ('None', 'nan', ''):
                    continue
                if a_name not in grouped:
                    props = r.get('properties', {})
                    grouped[a_name] = {
                        'tambon': props.get('TAMBOL_TH', props.get('TAMBON', '')),
                        'amphoe': props.get('AMPHOE_TH', props.get('AMPHOE', '')),
                        'province': props.get('PROV_TH', props.get('PROVINCE', ''))
                    }
            return grouped

        schools = get_grouped_data(school_records)
        religious = get_grouped_data(religion_records)
        hospitals = get_grouped_data(hospital_records)

        def find_headers():
            idx_s, idx_r, idx_h = None, None, None
            for r in range(1, 150):
                val = sht.range(f"A{r}").value
                if val:
                    val_str = str(val).strip()
                    if "สถานศึกษา" in val_str:
                        idx_s = r
                    elif "ศาสนสถาน" in val_str:
                        idx_r = r
                    elif "สถานพยาบาล" in val_str:
                        idx_h = r
            return idx_s, idx_r, idx_h

        # 1. เขียนตารางสถานศึกษา
        idx_s, idx_r, idx_h = find_headers()
        if idx_s and idx_r:
            start_s = idx_s + 2
            end_s = idx_r - 2
            end_s = self._prepare_sub_rows(sht, start_s, end_s, len(schools))
            
            # ล้างเนื้อหาเดิมในตารางสถานศึกษา
            for col in ['A', 'B', 'C', 'D', 'E']:
                self._safe_clear(sht, col, start_s, end_s)
                
            # หยอดข้อมูลสถานศึกษา
            row = start_s
            idx = 1
            for a_name, data in schools.items():
                self._set_cell_value_only(sht, f"A{row}", idx)
                self._set_cell_value_only(sht, f"B{row}", a_name)
                
                tb = data['tambon']
                if tb and not tb.startswith('ตำบล'):
                    tb = f"ตำบล{tb}"
                self._set_cell_value_only(sht, f"C{row}", tb)
                
                amp = data['amphoe']
                if amp and not (amp.startswith('อำเภอ') or amp.startswith('เมือง')):
                    if 'เมือง' in amp:
                        amp = amp
                    else:
                        amp = f"อำเภอ{amp}"
                self._set_cell_value_only(sht, f"D{row}", amp)
                
                prov = data['province']
                if prov and not prov.startswith('จังหวัด'):
                    prov = f"จังหวัด{prov}"
                self._set_cell_value_only(sht, f"E{row}", prov)
                
                for col in ['A', 'B', 'C', 'D', 'E']:
                    try:
                        sht.range(f"{col}{row}").api.Font.Name = 'TH Sarabun New'
                        sht.range(f"{col}{row}").api.Font.Size = 14
                    except Exception:
                        pass
                row += 1
                idx += 1

        # 2. เขียนตารางศาสนสถาน
        # หาตำแหน่งแถวของหัวข้อใหม่หลังจากแถวของสถานศึกษาถูกเลื่อนไปแล้ว
        _, idx_r, idx_h = find_headers()
        if idx_r and idx_h:
            start_r = idx_r + 2
            end_r = idx_h - 2
            end_r = self._prepare_sub_rows(sht, start_r, end_r, len(religious))
            
            # ล้างเนื้อหาเดิมในตารางศาสนสถาน
            for col in ['A', 'B', 'C', 'D', 'E']:
                self._safe_clear(sht, col, start_r, end_r)
                
            # หยอดข้อมูลศาสนสถาน
            row = start_r
            idx = 1
            for a_name, data in religious.items():
                self._set_cell_value_only(sht, f"A{row}", idx)
                self._set_cell_value_only(sht, f"B{row}", a_name)
                
                tb = data['tambon']
                if tb and not tb.startswith('ตำบล'):
                    tb = f"ตำบล{tb}"
                self._set_cell_value_only(sht, f"C{row}", tb)
                
                amp = data['amphoe']
                if amp and not (amp.startswith('อำเภอ') or amp.startswith('เมือง')):
                    if 'เมือง' in amp:
                        amp = amp
                    else:
                        amp = f"อำเภอ{amp}"
                self._set_cell_value_only(sht, f"D{row}", amp)
                
                prov = data['province']
                if prov and not prov.startswith('จังหวัด'):
                    prov = f"จังหวัด{prov}"
                self._set_cell_value_only(sht, f"E{row}", prov)
                
                for col in ['A', 'B', 'C', 'D', 'E']:
                    try:
                        sht.range(f"{col}{row}").api.Font.Name = 'TH Sarabun New'
                        sht.range(f"{col}{row}").api.Font.Size = 14
                    except Exception:
                        pass
                row += 1
                idx += 1

        # 3. เขียนตารางสถานพยาบาล
        _, _, idx_h = find_headers()
        if idx_h:
            start_h = idx_h + 2
            # ในเทมเพลตปกติจะมี 3 แถวให้กรอกเบื้องต้น
            end_h = start_h + 2
            end_h = self._prepare_sub_rows(sht, start_h, end_h, len(hospitals))
            
            # ล้างเนื้อหาเดิมในตารางสถานพยาบาล
            for col in ['A', 'B', 'C', 'D', 'E']:
                self._safe_clear(sht, col, start_h, end_h)
                
            # หยอดข้อมูลสถานพยาบาล
            row = start_h
            idx = 1
            for a_name, data in hospitals.items():
                self._set_cell_value_only(sht, f"A{row}", idx)
                self._set_cell_value_only(sht, f"B{row}", a_name)
                
                tb = data['tambon']
                if tb and not tb.startswith('ตำบล'):
                    tb = f"ตำบล{tb}"
                self._set_cell_value_only(sht, f"C{row}", tb)
                
                amp = data['amphoe']
                if amp and not (amp.startswith('อำเภอ') or amp.startswith('เมือง')):
                    if 'เมือง' in amp:
                        amp = amp
                    else:
                        amp = f"อำเภอ{amp}"
                self._set_cell_value_only(sht, f"D{row}", amp)
                
                prov = data['province']
                if prov and not prov.startswith('จังหวัด'):
                    prov = f"จังหวัด{prov}"
                self._set_cell_value_only(sht, f"E{row}", prov)
                
                for col in ['A', 'B', 'C', 'D', 'E']:
                    try:
                        sht.range(f"{col}{row}").api.Font.Name = 'TH Sarabun New'
                        sht.range(f"{col}{row}").api.Font.Size = 14
                    except Exception:
                        pass
                row += 1
                idx += 1

    def _write_sheet_schools_ec(self, wb, master_records):
        """เขียนชีท 3.ตรวจสอบสถานศึกษา (EC) รองรับสถานศึกษา วัด และสถานพยาบาล"""
        sht = self._find_sheet(wb, "3.ตรวจสอบสถานศึกษา")
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")
        
        # จัดเตรียมข้อมูล
        schools = master_records.get('3.ตรวจสอบสถานศึกษา', [])
        temples = master_records.get('3.ตรวจสอบสถานศึกษา_วัด', [])
        hospitals = master_records.get('3.ตรวจสอบสถานศึกษา_รพ', [])
        hospitals.extend(master_records.get('3.ตรวจสอบสถานศึกษา_รพสต', []))
        
        def format_records(records_list):
            grp = []
            for r in records_list:
                a_name = str(r.get('area_name', 'ไม่ระบุ'))
                if a_name in ('None', 'nan', ''): continue
                props = r.get('properties', {})
                grp.append({
                    'name': a_name,
                    'tambon': props.get('TAMBOL_TH', props.get('TAMBON', '')),
                    'amphoe': props.get('AMPHOE_TH', props.get('AMPHOE', '')),
                    'province': props.get('PROV_TH', props.get('PROVINCE', ''))
                })
            return grp
            
        schools_grp = format_records(schools)
        temples_grp = format_records(temples)
        hospitals_grp = format_records(hospitals)
        
        # ถ้าไม่มีข้อมูลเลย ให้ใส่ '-' เฉพาะช่องแรกของแต่ละตาราง
        if not schools_grp: schools_grp = [{'name': '-', 'tambon': '-', 'amphoe': '-', 'province': '-'}]
        if not temples_grp: temples_grp = [{'name': '-', 'tambon': '-', 'amphoe': '-', 'province': '-'}]
        if not hospitals_grp: hospitals_grp = [{'name': '-', 'tambon': '-', 'amphoe': '-', 'province': '-'}]

        # ล้างข้อมูลเดิมในช่องกรอก (แถว 3-17 สำหรับโรงเรียน, 19-30 สำหรับวัด, 32-42 สำหรับ รพ)
        # แต่เพื่อความปลอดภัย เราจะเขียนทับไปเลย
        # เริ่มเขียนโรงเรียน (เริ่มแถว 3)
        row = 3
        idx = 1
        for data in schools_grp:
            if row > 17: break
            self._set_cell_value_only(sht, f"A{row}", idx if data['name'] != '-' else '-')
            self._set_cell_value_only(sht, f"B{row}", data['name'])
            self._set_cell_value_only(sht, f"D{row}", data['tambon'])
            self._set_cell_value_only(sht, f"E{row}", data['amphoe'])
            self._set_cell_value_only(sht, f"F{row}", data['province'])
            row += 1
            idx += 1
            
        # ล้างแถวว่างที่เหลือของโรงเรียน
        for r in range(row, 18):
            self._safe_clear(sht, 'A', r, r)
            self._safe_clear(sht, 'B', r, r)
            self._safe_clear(sht, 'D', r, r)
            self._safe_clear(sht, 'E', r, r)
            self._safe_clear(sht, 'F', r, r)

        # เริ่มเขียนวัด (เริ่มแถว 20)
        row = 20
        idx = 1
        for data in temples_grp:
            if row > 30: break
            self._set_cell_value_only(sht, f"A{row}", idx if data['name'] != '-' else '-')
            self._set_cell_value_only(sht, f"B{row}", data['name'])
            self._set_cell_value_only(sht, f"D{row}", data['tambon'])
            self._set_cell_value_only(sht, f"E{row}", data['amphoe'])
            self._set_cell_value_only(sht, f"F{row}", data['province'])
            row += 1
            idx += 1
            
        # ล้างแถวว่างที่เหลือของวัด
        for r in range(row, 31):
            self._safe_clear(sht, 'A', r, r)
            self._safe_clear(sht, 'B', r, r)
            self._safe_clear(sht, 'D', r, r)
            self._safe_clear(sht, 'E', r, r)
            self._safe_clear(sht, 'F', r, r)

        # เริ่มเขียนโรงพยาบาล (เริ่มแถว 33)
        row = 33
        idx = 1
        for data in hospitals_grp:
            if row > 42: break
            self._set_cell_value_only(sht, f"A{row}", idx if data['name'] != '-' else '-')
            self._set_cell_value_only(sht, f"B{row}", data['name'])
            self._set_cell_value_only(sht, f"D{row}", data['tambon'])
            self._set_cell_value_only(sht, f"E{row}", data['amphoe'])
            self._set_cell_value_only(sht, f"F{row}", data['province'])
            row += 1
            idx += 1
            
        # ล้างแถวว่างที่เหลือของ รพ
        for r in range(row, 43):
            self._safe_clear(sht, 'A', r, r)
            self._safe_clear(sht, 'B', r, r)
            self._safe_clear(sht, 'D', r, r)
            self._safe_clear(sht, 'E', r, r)
            self._safe_clear(sht, 'F', r, r)

    def _write_sheet_stream(self, wb, records, is_ear=True):
        """ชีท 5.จุดตัดแหล่งน้ำ / 9.จุดตัดแหล่งน้ำ (เขียนเรียงแถวลงมา และแทรกแถวเพิ่มหากแถวไม่พอ)"""
        sheet_name = "5.จุดตัดแหล่งน้ำ" if is_ear else "9.จุดตัดแหล่งน้ำ"
        sht = self._find_sheet(wb, sheet_name)
        if not sht:
            return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        if not records:
            self._safe_clear(sht, 'A', 3, 100)
            self._safe_clear(sht, 'B', 3, 100)
            return

        # แทรกแถวเพิ่มหากข้อมูลมากกว่า 28 รายการ (แถว 3-30)
        end_row = self._prepare_rows(sht, 3, 30, len(records))

        # ล้างข้อมูลเดิมทั้งหมดในพื้นที่ทำงาน (เคลียร์ลงไปลึกหน่อยเพื่อกันข้อมูลเทมเพลตปน)
        clear_end = max(100, end_row)
        self._safe_clear(sht, 'A', 3, clear_end)
        self._safe_clear(sht, 'B', 3, clear_end)

        row = 3
        for r in records:
            a_name = str(r.get('area_name', ''))
            if a_name in ('None', 'nan', ''):
                a_name = 'แหล่งน้ำไม่ทราบชื่อ'
            km_val = f"{r.get('KM In', '')}" if r.get('KM In', '') else ''
            
            self._set_cell_value_only(sht, f'A{row}', a_name)
            self._set_cell_value_only(sht, f'B{row}', km_val)
            
            try:
                for col in ['A', 'B']:
                    cell = sht.range(f'{col}{row}')
                    cell.api.Font.Name = 'TH Sarabun New'
                    cell.api.Font.Size = 14
                    cell.row_height = 15
                    self._apply_thin_borders(sht, f'{col}{row}')
            except Exception:
                pass
            row += 1

    def _write_sheet_soil_sequential(self, wb, records):
        """เขียนชีท 7.ชุดดิน แบบเรียงแถวลงมาใหม่"""
        sht = self._find_sheet(wb, "7.ชุดดิน")
        if not sht: return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        if not records:
            self._safe_clear(sht, 'A', 2, 26)
            self._safe_clear(sht, 'B', 2, 26)
            self._safe_clear(sht, 'C', 2, 26)
            self._safe_clear(sht, 'D', 2, 26)
            self._safe_clear(sht, 'E', 2, 26)
            return

        # จัดกลุ่ม
        grouped = {}
        total_sqm = 0.0
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''): continue
            
            props = r.get('properties', {})
            s_name = props.get('seriesname', props.get('soilseries', a_name))
            s_group = props.get('soilgroup', '')
            
            if s_name not in grouped:
                grouped[s_name] = {
                    'group': s_group,
                    'area_sqm': 0.0
                }
            grouped[s_name]['area_sqm'] += r.get('intersect_area_sqm', 0.0)
            total_sqm += r.get('intersect_area_sqm', 0.0)

        # เตรียมแถว แถวข้อมูลในเทมเพลตคือ 2-26
        end_row = self._prepare_rows(sht, 2, 26, len(grouped))
        
        # ล้างข้อมูลแถว 2 ถึง end_row
        for col in ['A', 'B', 'C', 'D', 'E']:
            self._safe_clear(sht, col, 2, end_row)

        row = 2
        for s_name, data in grouped.items():
            area_sqm = data['area_sqm']
            area_sqkm = area_sqm / 1_000_000
            area_rai = area_sqm / 1600
            pct = (area_sqm / total_sqm * 100) if total_sqm > 0 else 0
            
            self._set_cell_value_only(sht, f"A{row}", data['group'])
            self._set_cell_value_only(sht, f"B{row}", s_name)
            self._set_cell_value_only(sht, f"C{row}", round(area_sqkm, 6))
            self._set_cell_value_only(sht, f"D{row}", round(area_rai, 2))
            self._set_cell_value_only(sht, f"E{row}", round(pct, 2))

            for col in ['A', 'B', 'C', 'D', 'E']:
                try:
                    sht.range(f"{col}{row}").api.Font.Name = 'TH Sarabun New'
                    sht.range(f"{col}{row}").api.Font.Size = 14
                except Exception:
                    pass
            row += 1

        # เขียนบรรทัดรวมที่ถัดจาก end_row
        total_row = end_row + 1
        total_sqkm = total_sqm / 1_000_000
        total_rai = total_sqm / 1600
        self._set_cell_value_only(sht, f"C{total_row}", round(total_sqkm, 6))
        self._set_cell_value_only(sht, f"D{total_row}", round(total_rai, 2))
        self._set_cell_value_only(sht, f"E{total_row}", 100.0)

    def _fill_erosion_table(self, sht, records, start_row, total_row):
        """เติมข้อมูลลงตารางการชะล้างพังทลาย (มีระดับ น้อยมาก-รุนแรงมาก คงที่)"""
        # ล้างข้อมูลเดิม
        self._safe_clear(sht, 'C', start_row, start_row + 4)
        self._safe_clear(sht, 'D', start_row, start_row + 4)
        self._safe_clear(sht, 'E', start_row, start_row + 4)
        
        self._safe_clear(sht, 'C', total_row, total_row)
        self._safe_clear(sht, 'D', total_row, total_row)
        self._safe_clear(sht, 'E', total_row, total_row)

        if not records:
            return

        # จัดกลุ่มตาม SEV_CLASS
        grouped = {}
        total_sqm = 0.0
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''): continue
            
            props = r.get('properties', {})
            sev_class = None
            for key in ['soillosscl', 'slc_code', 'gridcode', 'sev_class', 'sev_desc', 'class', 'severity', 'sev_class_t']:
                for prop_key, prop_val in props.items():
                    if prop_key.lower() == key.lower():
                        sev_class = prop_val
                        break
                if sev_class is not None:
                    break
            if sev_class is None:
                sev_class = a_name
            
            sev_class_str = str(sev_class).lower()
            std_class = None
            if 'รุนแรงมากที่สุด' in sev_class_str or 'รุนแรงที่สุด' in sev_class_str or 'very severe' in sev_class_str or '5' == sev_class_str:
                std_class = 'รุนแรงมากที่สุด'
            elif 'รุนแรงมาก' in sev_class_str or '4' == sev_class_str:
                std_class = 'รุนแรงมาก'
            elif 'รุนแรง' in sev_class_str or 'severe' in sev_class_str or '3' == sev_class_str:
                std_class = 'รุนแรง'
            elif 'ปานกลาง' in sev_class_str or 'moderate' in sev_class_str or '2' == sev_class_str:
                std_class = 'ปานกลาง'
            elif 'น้อย' in sev_class_str or 'low' in sev_class_str or 'น้อยมาก' in sev_class_str or 'very low' in sev_class_str or '1' == sev_class_str:
                std_class = 'น้อย'

            if not std_class:
                continue

            grouped[std_class] = grouped.get(std_class, 0.0) + r.get('intersect_area_sqm', 0.0)
            total_sqm += r.get('intersect_area_sqm', 0.0)

        total_sqkm = 0.0
        total_rai = 0.0
        
        for r_offset in range(5):
            row = start_row + r_offset
            cell_val = sht.range(f"A{row}").value
            if not cell_val: continue
            cell_str = str(cell_val).strip()

            matched_class = None
            if 'น้อยมาก' in cell_str:
                matched_class = 'น้อยมาก'
            elif 'รุนแรงมากที่สุด' in cell_str or 'รุนแรงที่สุด' in cell_str:
                matched_class = 'รุนแรงมากที่สุด'
            elif 'รุนแรงมาก' in cell_str:
                matched_class = 'รุนแรงมาก'
            elif 'รุนแรง' in cell_str:
                matched_class = 'รุนแรง'
            elif 'ปานกลาง' in cell_str:
                matched_class = 'ปานกลาง'
            elif 'น้อย' in cell_str:
                matched_class = 'น้อย'
            
            if matched_class and matched_class in grouped:
                area_sqm = grouped[matched_class]
                area_sqkm = area_sqm / 1_000_000
                area_rai = area_sqm / 1600
                pct = (area_sqm / total_sqm * 100) if total_sqm > 0 else 0

                self._set_cell_value_only(sht, f"C{row}", round(area_sqkm, 6))
                self._set_cell_value_only(sht, f"D{row}", round(area_rai, 2))
                self._set_cell_value_only(sht, f"E{row}", round(pct, 2))
                
                total_sqkm += area_sqkm
                total_rai += area_rai

        # รวม
        if total_sqm > 0:
            self._set_cell_value_only(sht, f"C{total_row}", round(total_sqkm, 6))
            self._set_cell_value_only(sht, f"D{total_row}", round(total_rai, 2))
            self._set_cell_value_only(sht, f"E{total_row}", 100.0)

    def _write_sheet_erosion_lookup(self, wb, records_1km, records_30m, is_ear=True):
        """ชีท การชะล้างพังทลาย"""
        sheet_name = "3.การชะล้างพังทลาย" if is_ear else "11.การชะล้างพังทลาย"
        sht = self._find_sheet(wb, sheet_name)
        if not sht: return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        # 1. ตารางบน (รัศมี 1 กม. หรือ 500ม.) -> แถวข้อมูล 3-7, รวมแถว 8
        self._fill_erosion_table(sht, records_1km, 3, 8)
        
        # 2. ตารางล่าง (ในเขตทาง 30ม.) -> แถวข้อมูล 15-19, รวมแถว 20
        self._fill_erosion_table(sht, records_30m, 15, 20)


    def _write_sheet_landuse_sequential_ec(self, wb, records):
        """ชีท 4.การใช้ประโยชน์ที่ดิน (EC แบบกางบรรทัดใหม่ทั้งหมด)"""
        sht = self._find_sheet(wb, "4.การใช้ประโยชน์ที่ดิน")
        if not sht: return
        print(f"  กำลังหยอดชีท '{sht.name}' แบบ List...")
        
        self._safe_clear(sht, 'A', 2, 74)
        self._safe_clear(sht, 'B', 2, 74)
        self._safe_clear(sht, 'C', 2, 74)
        self._safe_clear(sht, 'D', 2, 74)
        
        if not records: return
        
        grouped = {}
        total_sqm = 0.0
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''): continue
            props = r.get('properties', {})
            lu_name = None
            for key in ['LU_NAME', 'lu_name', 'LU_DES_TH', 'lu_des_th', 'LU_DES', 'lu_des', 'LU_DES_EN', 'lu_des_en']:
                for prop_key, prop_val in props.items():
                    if prop_key.lower() == key.lower():
                        lu_name = prop_val
                        break
                if lu_name is not None: break
            if lu_name is None: lu_name = a_name
            grouped[lu_name] = grouped.get(lu_name, 0.0) + r.get('intersect_area_sqm', 0.0)
            total_sqm += r.get('intersect_area_sqm', 0.0)
            
        if total_sqm <= 0: return
        
        LU_CLASSIFICATION = {
            '1. พื้นที่เกษตรกรรม': ['นา', 'ไร่', 'สวน', 'พืช', 'ทุ่ง', 'เลี้ยงสัตว์', 'เกษตร', 'a'],
            '2. พื้นที่ป่าไม้': ['ป่า', 'ไม้', 'f'],
            '3. พื้นที่แหล่งน้ำ': ['น้ำ', 'คลอง', 'อ่าง', 'บ่อ', 'ทะเลสาบ', 'หนอง', 'บึง', 'w'],
            '4. พื้นที่ชุมชนและสิ่งปลูกสร้าง': ['ชุมชน', 'หมู่บ้าน', 'เมือง', 'สิ่งปลูกสร้าง', 'ถนน', 'สนามบิน', 'โรงงาน', 'u'],
            '5. พื้นที่เบ็ดเตล็ด': ['ว่าง', 'เบ็ดเตล็ด', 'เหมือง', 'm']
        }
        
        cats = {k: {} for k in LU_CLASSIFICATION.keys()}
        for g_key, area in grouped.items():
            g_clean = str(g_key).lower()
            matched_p = '5. พื้นที่เบ็ดเตล็ด'
            for p_name, keywords in LU_CLASSIFICATION.items():
                if any(kw == g_clean or (len(kw) > 1 and kw in g_clean) for kw in keywords):
                    matched_p = p_name
                    break
            cats[matched_p][g_key] = area
            
        row = 2
        for p_name, items in cats.items():
            if not items: continue
            
            p_area = sum(items.values())
            self._set_cell_value_only(sht, f"A{row}", p_name)
            self._set_cell_value_only(sht, f"B{row}", round(p_area/1_000_000, 6))
            self._set_cell_value_only(sht, f"C{row}", round(p_area/1600, 2))
            self._set_cell_value_only(sht, f"D{row}", round(p_area/total_sqm*100, 2))
            
            try:
                tr = sht.range(f"A{row}:D{row}")
                tr.api.Font.Bold = True
            except: pass
            
            row += 1
            
            for lu_name, area in items.items():
                self._set_cell_value_only(sht, f"A{row}", f"    {lu_name}")
                self._set_cell_value_only(sht, f"B{row}", round(area/1_000_000, 6))
                self._set_cell_value_only(sht, f"C{row}", round(area/1600, 2))
                self._set_cell_value_only(sht, f"D{row}", round(area/total_sqm*100, 2))
                row += 1
                
        self._set_cell_value_only(sht, f"A{row}", "รวม")
        self._set_cell_value_only(sht, f"B{row}", round(total_sqm/1_000_000, 6))
        self._set_cell_value_only(sht, f"C{row}", round(total_sqm/1600, 2))
        self._set_cell_value_only(sht, f"D{row}", 100.0)
        try:
            tr = sht.range(f"A{row}:D{row}")
            tr.api.Font.Bold = True
            tr.color = (217, 217, 217)
            for b_id in [7, 8, 9, 10, 11, 12]:
                tr.api.Borders(b_id).LineStyle = 1
                tr.api.Borders(b_id).Weight = 2
        except: pass
        
        try:
            if row < 74:
                # Clear all formats and borders for unused rows
                sht.range(f"A{row+1}:D74").clear()
        except: pass

    def _write_sheet_landuse_lookup(self, wb, records, is_ear=True):
        """ชีท การใช้ประโยชน์ที่ดิน แบบแมปค่าตามประเภทที่มีและคำนวณยอดรวมหัวข้อใหญ่"""
        sheet_name = "8.การใช้ประโยชน์ที่ดิน" if is_ear else "4.การใช้ประโยชน์ที่ดิน"
        sht = self._find_sheet(wb, sheet_name)
        if not sht: return
        print(f"  กำลังหยอดชีท '{sht.name}'...")

        # ล้างตัวเลขในคอลัมน์ B, C, D แถว 2-74
        self._safe_clear(sht, 'B', 2, 74)
        self._safe_clear(sht, 'C', 2, 74)
        self._safe_clear(sht, 'D', 2, 74)

        if not records:
            return

        # จัดกลุ่มพื้นที่ตามชื่อประเภทสิ่งปกคลุมดิน (LU_NAME)
        grouped = {}
        total_sqm = 0.0
        for r in records:
            a_name = str(r.get('area_name', 'ไม่ระบุ'))
            if a_name in ('None', 'nan', ''): continue
            
            props = r.get('properties', {})
            # ค้นหาชื่อภาษาไทย/อังกฤษของสิ่งปกคลุมดินจากหลากหลายคอลัมน์ที่เป็นไปได้
            lu_name = None
            for key in ['LU_NAME', 'lu_name', 'LU_DES_TH', 'lu_des_th', 'LU_DES', 'lu_des', 'LU_DES_EN', 'lu_des_en']:
                for prop_key, prop_val in props.items():
                    if prop_key.lower() == key.lower():
                        lu_name = prop_val
                        break
                if lu_name is not None:
                    break
            if lu_name is None:
                lu_name = a_name
            
            grouped[lu_name] = grouped.get(lu_name, 0.0) + r.get('intersect_area_sqm', 0.0)
            total_sqm += r.get('intersect_area_sqm', 0.0)

        if total_sqm <= 0:
            return

        # 1. ค้นหาแถวของหมวดหมู่หลัก (Parent Nodes) และช่องรวม (Total Row) แบบอัตโนมัติ
        parent_rows = []
        total_row = None
        for r in range(2, 100):
            val = sht.range(f"A{r}").value
            if not val:
                continue
            val_str = str(val).strip()
            if any(val_str.startswith(prefix) for prefix in ["1.", "2.", "3.", "4.", "5."]):
                parent_rows.append(r)
            elif "รวม" in val_str:
                total_row = r
                break
                
        if not total_row:
            total_row = 74 # fallback
            
        end_leaf_row = total_row - 1

        matched_keys_set = set()
        row_matches = {}

        # Pass 1: Exact match
        for row in range(2, end_leaf_row + 1):
            if row in parent_rows:
                continue
            cell_val = sht.range(f"A{row}").value
            if not cell_val: continue
            cell_str = str(cell_val).strip()
            cell_str_clean = cell_str.replace(" ", "").lower()
            
            for g_key in grouped.keys():
                if g_key in matched_keys_set: continue
                g_key_clean = str(g_key).replace(" ", "").lower()
                if g_key_clean == cell_str_clean:
                    row_matches[row] = g_key
                    matched_keys_set.add(g_key)
                    break

        # Pass 2: Substring match
        for row in range(2, end_leaf_row + 1):
            if row in parent_rows or row in row_matches:
                continue
            cell_val = sht.range(f"A{row}").value
            if not cell_val: continue
            cell_str = str(cell_val).strip()
            cell_str_clean = cell_str.replace(" ", "").lower()

            for g_key in grouped.keys():
                if g_key in matched_keys_set: continue
                g_key_clean = str(g_key).replace(" ", "").lower()
                if cell_str_clean in g_key_clean:
                    row_matches[row] = g_key
                    matched_keys_set.add(g_key)
                    break

        # Write matched values
        for row in range(2, end_leaf_row + 1):
            if row in parent_rows:
                continue
            if row in row_matches:
                matched_key = row_matches[row]
                area_sqm = grouped[matched_key]
                area_sqkm = area_sqm / 1_000_000
                area_rai = area_sqm / 1600
                pct = (area_sqm / total_sqm * 100)
                
                self._set_cell_value_only(sht, f"B{row}", round(area_sqkm, 6))
                self._set_cell_value_only(sht, f"C{row}", round(area_rai, 2))
                self._set_cell_value_only(sht, f"D{row}", round(pct, 2))

        # Pass 3: Classify unmatched keys to parent categories
        LU_CLASSIFICATION = {
            'เกษตรกรรม': ['นา', 'ไร่', 'สวน', 'พืช', 'ทุ่ง', 'เลี้ยงสัตว์', 'เกษตร', 'a'],
            'ป่าไม้': ['ป่า', 'ไม้', 'f'],
            'แหล่งน้ำ': ['น้ำ', 'คลอง', 'อ่าง', 'บ่อ', 'ทะเลสาบ', 'หนอง', 'บึง', 'w'],
            'ชุมชนและสิ่งปลูกสร้าง': ['ชุมชน', 'หมู่บ้าน', 'เมือง', 'สิ่งปลูกสร้าง', 'ถนน', 'สนามบิน', 'โรงงาน', 'u'],
            'เบ็ดเตล็ด': ['ว่าง', 'เบ็ดเตล็ด', 'เหมือง', 'm']
        }
        
        parent_unmatched_area = {p_row: 0.0 for p_row in parent_rows}
        
        for g_key, area in grouped.items():
            if g_key not in matched_keys_set:
                g_clean = str(g_key).lower()
                matched_p = None
                for p_name, keywords in LU_CLASSIFICATION.items():
                    # For single character codes like 'a', 'f', 'w', 'u', 'm', check exact or prefix
                    if any(kw == g_clean or (len(kw) > 1 and kw in g_clean) for kw in keywords):
                        matched_p = p_name
                        break
                
                if not matched_p:
                    matched_p = 'เบ็ดเตล็ด'
                    
                for p_row in parent_rows:
                    p_val = str(sht.range(f"A{p_row}").value)
                    if matched_p in p_val:
                        parent_unmatched_area[p_row] += area
                        break

        # 2. คำนวณหัวข้อใหญ่ (Parent Nodes) จากผลรวมของแถวย่อยแบบอัตโนมัติ
        parent_ranges = {}
        for i in range(len(parent_rows)):
            p_row = parent_rows[i]
            start_child = p_row + 1
            if i < len(parent_rows) - 1:
                end_child = parent_rows[i+1] - 1
            else:
                end_child = total_row - 1
            parent_ranges[p_row] = (start_child, end_child)

        total_parent_sqkm = 0.0
        total_parent_rai = 0.0
        total_parent_pct = 0.0

        for p_row, (start_r, end_r) in parent_ranges.items():
            p_sqkm = 0.0
            p_rai = 0.0
            p_pct = 0.0
            for r in range(start_r, end_r + 1):
                val_sqkm = sht.range(f"B{r}").value
                val_rai = sht.range(f"C{r}").value
                val_pct = sht.range(f"D{r}").value
                if val_sqkm: p_sqkm += float(val_sqkm)
                if val_rai: p_rai += float(val_rai)
                if val_pct: p_pct += float(val_pct)

            if p_sqkm > 0 or parent_unmatched_area[p_row] > 0:
                p_sqkm += parent_unmatched_area[p_row] / 1_000_000
                p_rai += parent_unmatched_area[p_row] / 1600
                p_pct += (parent_unmatched_area[p_row] / total_sqm * 100)
                
                self._set_cell_value_only(sht, f"B{p_row}", round(p_sqkm, 6))
                self._set_cell_value_only(sht, f"C{p_row}", round(p_rai, 2))
                self._set_cell_value_only(sht, f"D{p_row}", round(p_pct, 2))
                total_parent_sqkm += p_sqkm
                total_parent_rai += p_rai
                total_parent_pct += p_pct

        # 3. เขียนแถวรวมทั้งหมด (Row 74)
        if total_parent_sqkm > 0:
            self._set_cell_value_only(sht, f"A{total_row}", "รวม")
            self._set_cell_value_only(sht, f"B{total_row}", round(total_parent_sqkm, 6))
            self._set_cell_value_only(sht, f"C{total_row}", round(total_parent_rai, 2))
            self._set_cell_value_only(sht, f"D{total_row}", 100.0)
            
            # ย้ำฟอร์แมตช่องรวม (เทาและตีกรอบ) เผื่อบางเทมเพลตไม่ได้ทำไว้
            try:
                tr = sht.range(f"A{total_row}:D{total_row}")
                tr.api.Font.Bold = True
                tr.color = (217, 217, 217)  # เทา
                # เส้นขอบบนซ้ายขวาล่าง และเส้นคั่นใน
                for b_id in [7, 8, 9, 10, 11, 12]:
                    tr.api.Borders(b_id).LineStyle = 1
                    tr.api.Borders(b_id).Weight = 2
            except Exception:
                pass
            
        # 4. ไม่ลบแถวที่ไม่มีข้อมูล เพื่อคงโครงสร้าง 5 หมวดหลักไว้
            
        # หาบรรทัดรวมใหม่ (เพราะแถวเลื่อน) เพื่อขีดเส้นหรือแก้ไขในอนาคตถ้าต้องการ
        # ไม่จำเป็นต้องแก้ เพราะเขียนข้อมูลไปหมดแล้ว

    def _write_sheet_slope(self, wb, is_ear=True):
        """ชีท 6.ความลาดชันของพื้นที่ (EAR) หรือ 10.ความลาดชัน (EC)"""
        sheet_name = "6.ความลาดชันของพื้นที่" if is_ear else "10.ความลาดชัน"
        sht = self._find_sheet(wb, sheet_name)
        if not sht:
            return
        print(f"  กำลังคำนวณและเขียนชีท '{sht.name}'...")

        cl_files = glob.glob(os.path.join(self.project_dir, "*_CL.shp"))
        if not cl_files:
            return

        try:
            import numpy as np
            import urllib.request
            from shapely.geometry import Point

            # 1. โหลดแนวถนนและแปลงพิกัดเป็น WGS84 (EPSG:4326) เพื่อดึง Lat/Lon
            road_gdf = gpd.read_file(cl_files[0])
            road_wgs84 = road_gdf.to_crs("EPSG:4326")
            line = road_wgs84.geometry.iloc[0]

            # 2. ทำการ Sample 30 จุดตามแนวถนนโครงการ
            num_samples = 30
            distances = np.linspace(0, line.length, num_samples)
            sampled_coords = [line.interpolate(d) for d in distances]

            elevations = []
            
            # 1. ลองดึงจาก OpenTopoData (SRTM 30m) ก่อน เพราะเสถียรกว่ามาก
            try:
                locations_str = "|".join([f"{p.y},{p.x}" for p in sampled_coords])
                url = f"https://api.opentopodata.org/v1/srtm30m?locations={locations_str}"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=15) as resp:
                    result = json.loads(resp.read().decode('utf-8'))
                    if result.get('status') == 'OK':
                        elevations = [item['elevation'] for item in result['results']]
            except Exception as e:
                print(f"  [Info] OpenTopoData ไม่ตอบสนอง กำลังสลับไปใช้ Open-Elevation... ({e})")
            
            # 2. ถ้าวิธีแรกไม่ได้ผล ให้ Fallback ไปใช้ Open-Elevation
            if not elevations or len(elevations) != num_samples:
                locations = [{'latitude': p.y, 'longitude': p.x} for p in sampled_coords]
                data = json.dumps({'locations': locations}).encode('utf-8')
                req = urllib.request.Request(
                    'https://api.open-elevation.com/api/v1/lookup', 
                    data=data, 
                    headers={'Content-Type': 'application/json', 'User-Agent': 'Mozilla/5.0'}
                )
                with urllib.request.urlopen(req, timeout=15) as resp:
                    result = json.loads(resp.read().decode('utf-8'))
                    elevations = [item['elevation'] for item in result['results']]

            if len(elevations) == num_samples:
                # 3. คำนวณความสูงต่ำสุด-สูงสุด
                min_el = round(min(elevations))
                max_el = round(max(elevations))
                el_range_str = f"{min_el} - {max_el} ม"

                # 4. คำนวณความลาดชันเฉลี่ย (Slope) 
                # แปลงจุดกลับมาเป็นพิกัดโครงการ (Metric CRS เช่น UTM) เพื่อให้คำนวณระยะทางได้ถูกต้อง
                road_metric = road_gdf.to_crs(TARGET_CRS)
                line_metric = road_metric.geometry.iloc[0]
                distances_metric = np.linspace(0, line_metric.length, num_samples)
                points_metric = [line_metric.interpolate(d) for d in distances_metric]

                slopes_signed = []
                for i in range(num_samples - 1):
                    p1 = points_metric[i]
                    p2 = points_metric[i+1]
                    dist_h = distances_metric[i+1] - distances_metric[i] # ระยะทางราบ (ตามแนวเส้น)
                    dist_v = elevations[i+1] - elevations[i] # ระยะทางดิ่งแบบมีทิศทาง
                    if dist_h > 0:
                        slopes_signed.append(dist_v / dist_h)
                
                up_slopes = [s for s in slopes_signed if s > 0]
                down_slopes = [s for s in slopes_signed if s < 0]
                avg_up = (sum(up_slopes) / len(up_slopes)) if up_slopes else 0.0
                avg_down = (sum(down_slopes) / len(down_slopes)) if down_slopes else 0.0
                
                # ความลาดชันเฉลี่ยที่มากสุดตามที่ user ระบุ (Maximum Average Slope)
                final_slope = max(avg_up, abs(avg_down))

                # 5. หยอดข้อมูลลงใน Excel
                # A2: ความลาดชันเฉลี่ย (ทศนิยม เช่น 0.045)
                # C2: ความสูงจากรทก (เช่น "146 - 432 ม")
                final_slope_rounded = round(final_slope * 100, 1) / 100.0
                self._set_cell_value_only(sht, 'A2', final_slope_rounded)
                self._set_cell_value_only(sht, 'C2', el_range_str)
                
                # ตั้งฟอนต์ขนาด 14
                for cell_ref in ['A2', 'C2']:
                    try:
                        cell = sht.range(cell_ref)
                        cell.api.Font.Name = 'TH Sarabun New'
                        cell.api.Font.Size = 14
                        cell.api.HorizontalAlignment = -4108 # จัดกลาง
                        cell.api.VerticalAlignment = -4108 # จัดกลาง
                    except:
                        pass
                print(f"  คำนวณความลาดชันสำเร็จ: Slope={round(final_slope*100, 2)}%, Elevation Range={el_range_str}")
                
                # วาดกราฟ Elevation Profile สไตล์ Google Earth (Pixel-Perfect)
                try:
                    import matplotlib.pyplot as plt
                    import matplotlib.ticker as ticker
                    import matplotlib.patches as patches
                    import matplotlib.transforms as transforms
                    import numpy as np
                    
                    dist_km = np.array([d / 1000.0 for d in distances_metric])
                    elevations_arr = np.array(elevations)
                    
                    y_min, y_max = np.min(elevations_arr), np.max(elevations_arr)
                    avg_el = np.mean(elevations_arr)
                    max_dist = dist_km[-1]
                    
                    # คำนวณสถิติเพิ่มเติม
                    diffs = np.diff(elevations_arr)
                    dist_diffs = np.diff(distances_metric)
                    safe_dist_diffs = np.where(dist_diffs == 0, 1, dist_diffs)
                    signed_slopes = np.where(dist_diffs > 0, diffs / safe_dist_diffs, 0)
                    
                    gain = np.sum(diffs[diffs > 0])
                    loss = np.sum(np.abs(diffs[diffs < 0]))
                    
                    max_up = np.max(signed_slopes) * 100 if len(signed_slopes) > 0 else 0
                    max_down = np.min(signed_slopes) * 100 if len(signed_slopes) > 0 else 0
                    
                    up_slopes = signed_slopes[signed_slopes > 0]
                    down_slopes = signed_slopes[signed_slopes < 0]
                    avg_up = np.mean(up_slopes) * 100 if len(up_slopes) > 0 else 0
                    avg_down = np.mean(down_slopes) * 100 if len(down_slopes) > 0 else 0
                    
                    plt.rcParams['font.family'] = 'Tahoma'
                    
                    bg_color = '#363636'
                    fig = plt.figure(figsize=(20, 3.8), facecolor=bg_color)
                    ax = fig.add_axes([0.06, 0.12, 0.92, 0.65])
                    ax.set_facecolor('#FFFFFF')
                    
                    fill_color = '#FFCACA'
                    ax.plot(dist_km, elevations_arr, color='#600000', linewidth=2.0)
                    ax.fill_between(dist_km, elevations_arr, y_min, color=fill_color, alpha=1.0)
                    
                    end_x, end_y = dist_km[-1], elevations_arr[-1]
                    ax.plot(end_x, end_y, marker='o', markersize=7, markerfacecolor='none', markeredgecolor='red', markeredgewidth=2.0, clip_on=False)
                    
                    y_range = y_max - y_min
                    step = 25 if y_range < 300 else 50
                    start_tick = (int(y_min) // step + 1) * step
                    regular_y_ticks = np.arange(start_tick, y_max, step)
                    min_y_dist = step * 0.4
                    filtered_y_ticks = [t for t in regular_y_ticks if (t - y_min) > min_y_dist and (y_max - t) > min_y_dist]
                    y_ticks = [y_min] + filtered_y_ticks + [y_max]
                    ax.set_yticks(y_ticks)
                    
                    y_margin = y_range * 0.05 if y_range > 0 else 10
                    ax.set_ylim(y_min, y_max + y_margin)
                    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda val, pos: f"{int(val)} ม."))
                    
                    ax.set_xlim(0, max_dist)
                    if max_dist <= 10:
                        x_step = 1.0
                    elif max_dist <= 25:
                        x_step = 2.5
                    elif max_dist <= 50:
                        x_step = 5.0
                    else:
                        x_step = 10.0
                    x_ticks = np.arange(x_step, max_dist, x_step)
                    ax.set_xticks(x_ticks)
                    ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda val, pos: f"{val:g} กม."))
                    
                    ax.tick_params(axis='y', colors=fill_color, labelsize=10, length=0, pad=8)
                    ax.tick_params(axis='x', colors='#E0E0E0', labelsize=8, length=0, pad=5)
                    
                    ax.yaxis.grid(True, color='#FFD0D0', linestyle='-', linewidth=1.2)
                    ax.xaxis.grid(True, color='#E0E0E0', linestyle='-', linewidth=1.0)
                    
                    for spine in ax.spines.values():
                        spine.set_edgecolor('black')
                        spine.set_linewidth(2.0)
                        
                    fig.text(0.06, 0.88, "กราฟ: ต่ำสุด, เฉลี่ย, สูงสุด", color='white', fontsize=10)
                    fig.text(0.147, 0.88, f" ระดับความสูง: {int(y_min)}, {int(avg_el)}, {int(y_max)} ม. ", color='white', fontsize=10, 
                             bbox=dict(facecolor='#8B0000', edgecolor='none', pad=2.0))
                             
                    fig.text(0.08, 0.81, "ช่วงทั้งหมด:", color='white', fontsize=10)
                    
                    box_style = dict(facecolor='none', edgecolor='black', linewidth=1.5, pad=3.0)
                    fig.text(0.125, 0.81, f" ระยะทาง: {max_dist:.1f} กม. ", color='white', fontsize=9.5, bbox=box_style)
                    fig.text(0.185, 0.81, f" ความสูงเพิ่ม/ลด: {int(gain)} ม.-{int(loss)} ม. ", color='white', fontsize=9.5, bbox=box_style)
                    fig.text(0.297, 0.81, f" ความลาดชันสูงสุด: {max_up:.1f}%, {max_down:.1f}% ", color='white', fontsize=9.5, bbox=box_style)
                    fig.text(0.413, 0.81, f" ความลาดชันเฉลี่ย: {avg_up:.1f}%, {avg_down:.1f}% ", color='white', fontsize=9.5, bbox=box_style)
                    
                    bbox_red_black_edge = dict(facecolor='#8B0000', edgecolor='black', linewidth=1.5, pad=2.0)
                    trans_axes = ax.transAxes
                    trans_blend = transforms.blended_transform_factory(ax.transAxes, ax.transAxes)
                    
                    end_slope = signed_slopes[-1] * 100 if len(signed_slopes) > 0 else 0.0
                        
                    ax.text(1.0, 0.0, f" {end_slope:.1f}% ", color="white", fontsize=10, ha="right", va="bottom", 
                            bbox=bbox_red_black_edge, transform=trans_axes)
                            
                    ax.text(1.0, -0.01, f" {max_dist:.1f} กม. ", color="white", fontsize=10, ha="right", va="top", 
                            bbox=bbox_red_black_edge, transform=trans_blend, clip_on=False)
                            
                    ax.text(0.965, 0.17, f" {int(end_y)} ม. ", color="white", fontsize=10, ha="right", va="bottom", 
                            bbox=bbox_red_black_edge, transform=trans_axes)
                    
                    graph_path = os.path.join(self.project_dir, 'elevation_profile.png')
                    plt.savefig(graph_path, dpi=100, facecolor=fig.get_facecolor(), edgecolor='none')
                    plt.close(fig)
                    
                    # ค้นหาภาพเดิมในชีทเพื่อดึงขนาดและตำแหน่งมาใช้ (ถ้ามี)
                    target_left = sht.range('A5').left
                    target_top = sht.range('A5').top
                    target_width = 600
                    target_height = 250
                    
                    if len(sht.pictures) > 0:
                        first_pic = sht.pictures[0]
                        target_left = first_pic.left
                        target_top = first_pic.top
                        target_width = first_pic.width
                        target_height = first_pic.height
                        
                        # ลบรูปเดิมทั้งหมดในชีทนี้ออก (ภาพ placeholder หรือกราฟที่เคยวาดไว้)
                        for pic in sht.pictures:
                            pic.delete()
                            
                    # นำรูปล่าสุดไปวางแทนที่ในขนาดและตำแหน่งเดิมเป๊ะๆ
                    sht.pictures.add(graph_path, name='ElevationProfile', 
                                     left=target_left, top=target_top,
                                     width=target_width, height=target_height)
                    print(f"  สร้างและแปะกราฟ Elevation Profile ลงในเซลล์ A5 สำเร็จ")
                except Exception as e:
                    print(f"  [Warning] ไม่สามารถสร้างกราฟ Elevation Profile ได้: {e}")
        except Exception as e:
            print(f"  Warning: ไม่สามารถดึงข้อมูลความลาดชันอัตโนมัติได้ (ข้ามไปใช้ค่าเริ่มต้นจากเทมเพลต): {e}")

    # ========================================
    # MAIN WRITE REPORT
    # ========================================

    def write_report(self):
        output_path = self.prepare_output_file()
        if not output_path:
            return

        master_records = self.load_cache()
        if not master_records:
            return

        is_ear = "EAR" in self.project_name.upper()

        print("กำลังเปิด Excel แบบ Background (xlwings)...")
        app = xw.App(visible=False)
        try:
            wb = app.books.open(output_path)

            # ล้างข้อมูลตัวอย่างจากเทมเพลตก่อนเขียนข้อมูลจริง (ป้องกันข้อมูลผิดพื้นที่)
            self._clear_template_defaults(wb, is_ear)

            # 1. ชีท 1.ป่า (ป่าสงวน & ป่าถาวร)
            self._write_sheet_pa(wb, master_records.get('1.ป่า_สงวน'), master_records.get('1.ป่า'))

            # 2. ชีท 4.พื้นที่คงสภาพป่า (EAR เท่านั้น)
            if is_ear:
                self._write_sheet_forest_status(
                    wb, 
                    master_records.get('4.พื้นที่คงสภาพป่า'), 
                    master_records.get('4.พื้นที่คงสภาพป่า_เขตทาง')
                )

            # 3. ชีท 10/12.ชั้นคุณภาพลุ่มน้ำ (ใช้แบบ Lookup)
            self._write_sheet_watershed_lookup(wb, master_records.get('10.ชั้นคุณภาพลุ่มน้ำ' if is_ear else '12.ชั้นคุณภาพลุ่มน้ำ'), is_ear)

            # 4. ชีท 2.พื้นที่หมู่บ้าน (เขียนแบบเรียงแถว เติมรายละเอียดเต็ม)
            self._write_sheet_villages_sequential(wb, master_records.get('2.พื้นที่หมู่บ้าน'))

            # 5. ชีท 9/5.แหล่งโบราณสถาน (เขียนแบบเรียงแถว เติมรายละเอียดเต็ม)
            self._write_sheet_historic_sequential(wb, master_records.get('9.แหล่งโบราณสถาน' if is_ear else '5.แหล่งโบราณสถาน'), is_ear)

            # 6. ชีท 7.เสี่ยงต่อการเกิดดินถล่ม (EC) / พื้นที่อ่อนไหว (EAR)
            if is_ear:
                self._write_sheet_sensitive_sequential(wb, master_records)
            else:
                self._write_sheet_landslide_lookup(wb, master_records.get('7.เสี่ยงต่อการเกิดดินถล่ม'), is_ear)
                self._write_sheet_schools_ec(wb, master_records)
                self._write_sheet_mineral_ec(wb, master_records.get('6.แหล่งทรัพยากรทางธรณี'))
                self._write_sheet_earthquake_ec(wb, master_records.get('8.เสี่ยงต่อการเกิดแผ่นดินไหว'))

            # 7. ชีท 5/9.จุดตัดแหล่งน้ำ (เขียนแบบเรียงแถวใหม่)
            self._write_sheet_stream(wb, master_records.get('5.จุดตัดแหล่งน้ำ' if is_ear else '9.จุดตัดแหล่งน้ำ'), is_ear)

            # 8. ชีท 7.ชุดดิน (เขียนแบบเรียงแถวใหม่ และคำนวณร้อยละ)
            self._write_sheet_soil_sequential(wb, master_records.get('7.ชุดดิน'))

            # 9. ชีท 3/11.การชะล้างพังทลาย (ใช้แบบ Lookup คัดแยก 1กม. และ เขตทาง)
            self._write_sheet_erosion_lookup(
                wb, 
                master_records.get('3.การชะล้างพังทลาย' if is_ear else '11.การชะล้างพังทลาย'), 
                master_records.get('3.การชะล้างพังทลาย_เขตทาง' if is_ear else '11.การชะล้างพังทลาย_เขตทาง'), 
                is_ear
            )

            # 10. ชีท 8/4.การใช้ประโยชน์ที่ดิน
            if is_ear:
                self._write_sheet_landuse_lookup(wb, master_records.get('8.การใช้ประโยชน์ที่ดิน'), is_ear)
            else:
                self._write_sheet_landuse_sequential_ec(wb, master_records.get('4.การใช้ประโยชน์ที่ดิน'))

            # 11. ชีท 6/10.ลาดชันของพื้นที่ (ดึงความสูงจริงและคำนวณเฉลี่ยอัตโนมัติ)
            self._write_sheet_slope(wb, is_ear)

            print("กำลังจัดเรียงชีตตามตัวเลข...")
            try:
                sheet_names = [sht.name for sht in wb.sheets]
                import re
                def get_sheet_num(s):
                    m = re.match(r'^(\d+)', s)
                    return int(m.group(1)) if m else 999
                    
                sorted_names = sorted(sheet_names, key=get_sheet_num)
                for s_name in sorted_names:
                    wb.sheets[s_name].api.Move(After=wb.sheets[-1].api)
            except Exception as e:
                print(f"  Warning: ไม่สามารถจัดเรียงชีตได้: {e}")

            print("บันทึกและปิดไฟล์ Excel...")
            wb.save()
        except Exception as e:
            import traceback
            print(f"Error during excel writing: {e}")
            traceback.print_exc()
        finally:
            if 'wb' in locals():
                wb.close()
            app.quit()
        print("เขียน Excel เสร็จสมบูรณ์!")
        
        try:
            import shutil
            dest_in_project = os.path.join(self.project_dir, os.path.basename(output_path))
            print(f"กำลังคัดลอกไฟล์รายงานที่เขียนเสร็จสมบูรณ์กลับไปยังโฟลเดอร์โครงการ: {dest_in_project}")
            shutil.copy2(output_path, dest_in_project)
        except Exception as e:
            print(f"  Warning: ไม่สามารถคัดลอกไฟล์รายงานกลับไปยังโฟลเดอร์โครงการได้: {e}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = r"D:\tammachart\ป8401_สทล01_ลป2_13 EAR"
    reporter = ExcelReporter(path)
    reporter.write_report()
