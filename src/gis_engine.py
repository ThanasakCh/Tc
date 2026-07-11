import geopandas as gpd
from shapely.geometry import Point, LineString, MultiLineString
import pandas as pd
import os
import math

from config import TARGET_CRS

class GISEngine:
    def __init__(self, shp_dir):
        self.shp_dir = shp_dir
        self.road_gdf = None
        self.crossed_provinces = []
        
    def safe_decode(self, val):
        if not isinstance(val, str):
            return val
        try:
            return val.encode('latin-1').decode('cp874')
        except UnicodeEncodeError:
            return val
        except Exception:
            return val
            
    def _load_poly_gdf(self, shp_paths, read_bbox):
        import pandas as pd
        if isinstance(shp_paths, str):
            shp_paths = [shp_paths]
        gdfs = []
        for path in shp_paths:
            try:
                gdf = gpd.read_file(path, bbox=read_bbox, engine="pyogrio")
            except Exception:
                try:
                    gdf = gpd.read_file(path, bbox=read_bbox)
                except Exception:
                    gdf = gpd.read_file(path)
            if not gdf.empty:
                gdfs.append(gdf)
        if not gdfs:
            return gpd.GeoDataFrame()
        if len(gdfs) == 1:
            return gdfs[0]
        return pd.concat(gdfs, ignore_index=True)
        
        
    def load_road_centerline(self, project_path):
        """โหลดเส้นกลางถนนจากโฟลเดอร์โปรเจกต์ ถ้าไม่มีค่อยใช้ Fuzzy Match กรองจากเส้นรวม"""
        import glob
        
        # 1. หาไฟล์เส้นกลางถนน (*_CL.shp) จากในโฟลเดอร์คำขอก่อน
        cl_files = glob.glob(os.path.join(project_path, "*_CL.shp"))
        if cl_files:
            road_path = cl_files[0]
            print(f"Loading PROJECT road centerline from {os.path.basename(road_path)}")
            gdf = gpd.read_file(road_path)
            self.road_gdf = gdf.to_crs(TARGET_CRS)
            return self.road_gdf
            
        # 2. ถ้าไม่มีไฟล์ _CL.shp ให้หยุดทำงานทันทีเพื่อป้องกันการโหลดถนนทั้งประเทศจนเครื่องค้าง
        project_name = os.path.basename(project_path)
        raise ValueError(f"\n[ERROR] ไม่พบไฟล์แนวเส้นทาง (_CL.shp) ในโฟลเดอร์ {project_name}\nกรุณาตรวจสอบและนำไฟล์ _CL.shp มาใส่ให้ครบถ้วนเพื่อป้องกันไม่ให้โปรแกรมทำงานหนักและเครื่องค้าง!")

        # แปลงพิกัดเป็น UTM 47N (EPSG:32647)
        self.road_gdf = gdf.to_crs(TARGET_CRS)
        return self.road_gdf

    def _fetch_rest_api_as_gdf(self, url, bounds_utm):
        """Fetch features from ESRI MapServer REST API within a bounding box"""
        import urllib.request
        import urllib.parse
        import json
        import ssl
        from shapely.geometry import shape

        # Create bounding box in UTM and reproject to WGS84 for the API
        from shapely.geometry import box
        bbox_poly = box(*bounds_utm)
        bbox_gdf = gpd.GeoDataFrame({'geometry': [bbox_poly]}, crs=self.road_gdf.crs)
        bbox_wgs84 = bbox_gdf.to_crs(epsg=4326).total_bounds

        geometry_json = {
            "xmin": bbox_wgs84[0],
            "ymin": bbox_wgs84[1],
            "xmax": bbox_wgs84[2],
            "ymax": bbox_wgs84[3],
            "spatialReference": {"wkid": 4326}
        }
        
        # Determine layer id from url if possible or assume 0
        layer_id = url.split('/')[-1]
        if not layer_id.isdigit():
            layer_id = "0"
            query_url = f"{url}/{layer_id}/query"
        else:
            query_url = f"{url}/query"

        params = {
            'f': 'json',
            'geometry': json.dumps(geometry_json),
            'geometryType': 'esriGeometryEnvelope',
            'spatialRel': 'esriSpatialRelIntersects',
            'outFields': '*',
            'returnGeometry': 'true',
            'outSR': '4326'
        }
        
        data = urllib.parse.urlencode(params).encode('utf-8')
        req = urllib.request.Request(query_url, data=data, headers={'User-Agent': 'Mozilla/5.0'})
        
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        try:
            with urllib.request.urlopen(req, context=ctx, timeout=30) as response:
                resp_json = json.loads(response.read().decode())
                
                features = resp_json.get('features', [])
                if not features:
                    return gpd.GeoDataFrame()
                
                # Convert ESRI features to Shapely geometries
                geoms = []
                attrs_list = []
                for f in features:
                    geom = f.get('geometry')
                    if not geom:
                        continue
                    # Simple conversion for ESRI polygon to GeoJSON-like
                    if 'rings' in geom:
                        poly_geojson = {'type': 'Polygon', 'coordinates': geom['rings']}
                        geoms.append(shape(poly_geojson))
                    elif 'paths' in geom:
                        line_geojson = {'type': 'MultiLineString', 'coordinates': geom['paths']}
                        geoms.append(shape(line_geojson))
                    elif 'x' in geom and 'y' in geom:
                        pt_geojson = {'type': 'Point', 'coordinates': [geom['x'], geom['y']]}
                        geoms.append(shape(pt_geojson))
                    else:
                        continue
                        
                    attrs = f.get('attributes', {})
                    attrs_list.append(attrs)
                
                if not geoms:
                    return gpd.GeoDataFrame()
                    
                gdf = gpd.GeoDataFrame(attrs_list, geometry=geoms, crs="EPSG:4326")
                return gdf.to_crs(self.road_gdf.crs)
        except Exception as e:
            print(f"Error fetching from REST API {url}: {e}")
            return gpd.GeoDataFrame()

    def calculate_intersections(self, polygon_shp_path, name_col, buffer_m):
        import time
        start_time = time.time()
        
        # ตรวจสอบว่ามีไฟล์ GPKG (Geopackage) ที่แปลงไว้หรือไม่ ถ้ามีให้ใช้ GPKG แทน .shp เพื่อความเร็วสูงสุด
        if isinstance(polygon_shp_path, str):
            gpkg_path = polygon_shp_path.replace(".shp", ".gpkg")
            if os.path.exists(gpkg_path):
                polygon_shp_path = gpkg_path
            
        first_path = polygon_shp_path[0] if isinstance(polygon_shp_path, list) else polygon_shp_path
        print(f"[{time.strftime('%X')}] Calculating intersections for {os.path.basename(first_path)}")
        
        # ตี Buffer รอบเส้นถนน
        road_buffer = self.road_gdf.copy()
        road_buffer['geometry'] = road_buffer.geometry.buffer(buffer_m)
        bounds = tuple(road_buffer.total_bounds)
        print(f"[{time.strftime('%X')}] Road bounding box (UTM): {bounds}", flush=True)
        
        # ยูเนียนบัฟเฟอร์ของถนนทั้งหมดและเตรียมตัวนับเพื่อกันการนับพื้นที่ซ้ำ (Double Counting)
        road_buf_union = road_buffer.unary_union
        assigned_poly_ids = set()
        
        # หาวิธีแปลงขอบเขต (bbox) ให้ตรงกับ CRS ของ Shapefile เพื่อไม่ให้โหลดทั้งประเทศ
        import pyogrio
        from pyproj import CRS
        from shapely.geometry import box
        
        shp_crs = None
        bbox_in_shp_crs = None
        
        try:
            meta = pyogrio.read_info(first_path)
            shp_crs = meta['crs']
        except Exception:
            try:
                import fiona
                with fiona.open(first_path) as src:
                    shp_crs = src.crs_wkt
            except:
                pass
                
        if shp_crs:
            try:
                crs_from = CRS.from_user_input(TARGET_CRS)
                crs_to = CRS.from_user_input(shp_crs)
                if not crs_from.equals(crs_to):
                    b = box(*bounds)
                    temp_gdf = gpd.GeoDataFrame(geometry=[b], crs=TARGET_CRS).to_crs(crs_to)
                    bbox_in_shp_crs = tuple(temp_gdf.total_bounds)
                    print(f"[{time.strftime('%X')}] Converted road bounds to shapefile CRS: {bbox_in_shp_crs}", flush=True)
                else:
                    bbox_in_shp_crs = bounds
            except Exception as e:
                print(f"[{time.strftime('%X')}] Warning converting bounds: {e}", flush=True)

        read_bbox = bbox_in_shp_crs if bbox_in_shp_crs else bounds

        is_api = isinstance(polygon_shp_path, str) and polygon_shp_path.startswith('http')

        # โหลด Polygon และแปลงพิกัด (ใช้ bbox ที่แปลง CRS แล้วช่วยกรองข้อมูล)
        if is_api:
            print(f"[{time.strftime('%X')}] Reading polygon from REST API...", flush=True)
            poly_gdf = self._fetch_rest_api_as_gdf(polygon_shp_path, bounds)
        else:
            print(f"[{time.strftime('%X')}] Reading polygon shapefile...", flush=True)
            poly_gdf = self._load_poly_gdf(polygon_shp_path, read_bbox)
            
        if poly_gdf.empty:
            is_fallback = False
            path_str = str(polygon_shp_path)
            if any(x in path_str for x in ['สถานศึกษา', 'วัด', 'โรงพยาบาล', 'รพสต']):
                is_fallback = True
            
            if not is_fallback or is_api:
                return pd.DataFrame()
            else:
                print(f"[{time.strftime('%X')}] ไม่มีข้อมูลใน Bounding Box ทำการโหลดทั้งจังหวัดเพื่อหาจุดที่ใกล้ที่สุด...", flush=True)
                poly_gdf = self._load_poly_gdf(polygon_shp_path, None)
                if poly_gdf.empty:
                    return pd.DataFrame()
            
        if not is_api:
            print(f"[{time.strftime('%X')}] Converting CRS to {TARGET_CRS}...", flush=True)
            poly_gdf = poly_gdf.to_crs(TARGET_CRS)
        
        # หั่น (Clip) geometry ของ Polygon ให้เหลือเฉพาะขอบเขตพื้นที่ศึกษา เพื่อเร่งความเร็วการคำนวณ 150 เท่า!
        if not locals().get('is_fallback', False):
            xmin, ymin, xmax, ymax = bounds
            try:
                import shapely
                poly_gdf['geometry'] = poly_gdf.geometry.apply(
                    lambda g: shapely.clip_by_rect(g, xmin, ymin, xmax, ymax) if g and not g.is_empty else g
                )
                poly_gdf = poly_gdf[~poly_gdf.geometry.is_empty]
            except Exception as e:
                print(f"[{time.strftime('%X')}] Warning clipping geometries: {e}", flush=True)
            
        print(f"[{time.strftime('%X')}] Total polygons loaded: {len(poly_gdf)}", flush=True)
        
        # เชื่อมโยงข้อมูลตำบล อำเภอ จังหวัด อัตโนมัติด้วยขอบเขตการปกครอง
        # สำหรับชั้นข้อมูลประเภทจุด (Point/MultiPoint) หรือชั้นข้อมูลหมู่บ้าน/โบราณสถาน/สถานศึกษา
        is_point_layer = any(geom_type in str(g).lower() for g in poly_gdf.geometry.geom_type.unique() for geom_type in ['point', 'multipoint'])
        if 'หมู่บ้าน' in polygon_shp_path or 'historic' in polygon_shp_path or 'สถานศึกษา' in polygon_shp_path or 'อ่อนไหว' in polygon_shp_path or is_point_layer:
            print(f"[{time.strftime('%X')}] ดึงข้อมูลการปกครอง (ตำบล/อำเภอ/จังหวัด) จาก ขอบเขตตำบล.shp...", flush=True)
            poly_gdf = self.enrich_with_admin_boundaries(poly_gdf)
        
        results = []
        
        # ค้นหาชื่อคอลัมน์จริงที่มีใน shapefile ด้วยระบบ fallback
        actual_name_col = name_col
        if name_col not in poly_gdf.columns:
            cols_lower = {c.lower(): c for c in poly_gdf.columns}
            if name_col.lower() in cols_lower:
                actual_name_col = cols_lower[name_col.lower()]
            else:
                # เช็คคำใกล้เคียงสำหรับ ดิน
                if 'soil' in name_col.lower() or 'series' in name_col.lower():
                    for fallback in ['seriesname', 'soilseries', 'soilgroup', 'soilgp', 'soil_gp']:
                        if fallback in cols_lower:
                            actual_name_col = cols_lower[fallback]
                            break
                # เช็คคำใกล้เคียงสำหรับ การใช้ประโยชน์ที่ดิน
                elif 'lu_name' in name_col.lower() or 'luse' in name_col.lower():
                    for fallback in ['lu_des_th', 'lu_des', 'lu_name_th', 'lu_name', 'lu_code', 'luname', 'lucode', 'lu_t', 'lu_code_t']:
                        if fallback in cols_lower:
                            actual_name_col = cols_lower[fallback]
                            break
                # เช็คคำใกล้เคียงสำหรับ การชะล้างพังทลายของดิน
                elif 'sev_class' in name_col.lower() or 'erosion' in name_col.lower():
                    for fallback in ['soillosscl', 'slc_code', 'gridcode', 'sev_class', 'sev_desc', 'class', 'severity', 'sev_class_t']:
                        if fallback in cols_lower:
                            actual_name_col = cols_lower[fallback]
                            break
                # เช็คคำใกล้เคียงสำหรับ ตำแหน่งหมู่บ้าน
                elif 'vill_nm_t' in name_col.lower() or 'village' in name_col.lower() or 'muban' in name_col.lower():
                    for fallback in ['muban', 'vill_name', 'vill_nm_t', 'vill_nm', 'vill_no']:
                        if fallback in cols_lower:
                            actual_name_col = cols_lower[fallback]
                            break
        
        # วนลูปตามเส้นถนน (เผื่อมีหลายเส้น)
        for idx, road in self.road_gdf.iterrows():
            road_line = road.geometry
            if road_line is None or road_line.is_empty:
                continue
                
            km_start_base = road.get('km_start', 0.0) # ดึงค่าเริ่มต้นกิโลเมตร
            try:
                km_start_base = float(km_start_base)
            except:
                km_start_base = 0.0
            if pd.isna(km_start_base):
                km_start_base = 0.0
            # ป้องกันกรณีใส่เลขสายทางลงใน km_start
            if km_start_base > 2000:
                km_start_base = 0.0
            
            # หา Polygon ที่ซ้อนทับกับ Buffer ของเส้นถนนเส้นนี้
            road_buf_geom = road_buffer.loc[idx, 'geometry']
            if road_buf_geom is None or road_buf_geom.is_empty:
                continue
                
            # กรอง polygon ที่ว่างออกก่อน
            valid_polys = poly_gdf[~poly_gdf.geometry.is_empty]
            try:
                intersecting_polys = valid_polys[valid_polys.intersects(road_buf_geom)]
            except Exception:
                # ถ้าเจอปัญหา topology พัง ให้รัน buffer(0) เฉพาะจุด
                try:
                    fixed_geoms = valid_polys.geometry.buffer(0)
                    intersecting_polys = valid_polys[fixed_geoms.intersects(road_buf_geom)]
                except Exception:
                    # ถ้ายังพังอีก ให้ fallback เป็นการกวาดเช็คทีละอันแบบปลอดภัย
                    intersecting_polys = valid_polys
            
            for p_idx, poly in intersecting_polys.iterrows():
                poly_name = self.safe_decode(poly.get(actual_name_col, "Unknown"))
                if name_col == 'fr_name' and 'NRF_Zone' in poly:
                    zone_val = str(poly.get('NRF_Zone', ''))
                    if zone_val and zone_val != 'None' and zone_val != 'nan':
                        poly_name = f"{poly_name} Zone {self.safe_decode(zone_val)}"
                    else:
                        poly_name = f"{poly_name} Zone None"
                poly_geom = poly.geometry
                if not poly_geom.is_valid:
                    try:
                        poly_geom = poly_geom.buffer(0)
                    except:
                        pass

                # ถ้าลักษณะภูมิศาสตร์เป็นจุด (เช่น ตำแหน่งหมู่บ้าน, แหล่งโบราณสถาน)
                if poly_geom.geom_type in ['Point', 'MultiPoint']:
                    from shapely.geometry import Point
                    pts = [poly_geom] if poly_geom.geom_type == 'Point' else poly_geom.geoms
                    for pt in pts:
                        dist = road_line.project(pt)
                        km = (km_start_base * 1000) + dist
                        results.append({
                            'route': road.get('route', 'Unknown'),
                            'area_name': poly_name,
                            'km_in_m': km,
                            'km_out_m': km,
                            'length_m': 0,
                            'intersect_area_sqm': 0,
                            'center_x': pt.x,
                            'center_y': pt.y,
                            'properties': {k: self.safe_decode(v) if isinstance(v, str) else str(v) for k, v in poly.items() if k != 'geometry'}
                        })
                    continue
                
                # หาเส้นถนนเฉพาะส่วนที่ตัดกับ Polygon
                intersection = road_line.intersection(poly_geom)
                
                if intersection.is_empty:
                    continue
                    
                # คำนวณระยะทางและพื้นที่ตัดจริงกับบัฟเฟอร์ถนนทั้งหมดครั้งเดียว (ถ้าเป็น Polygon)
                area_sqm = 0
                if poly_geom.geom_type in ['Polygon', 'MultiPolygon']:
                    area_sqm = poly_geom.intersection(road_buf_union).area
                
                # จัดการกรณีเป็นจุด (เช่น ถนนตัดแม่น้ำ)
                from shapely.geometry import Point, MultiPoint, LineString, MultiLineString
                if isinstance(intersection, (Point, MultiPoint)):
                    pts = [intersection] if isinstance(intersection, Point) else intersection.geoms
                    for pt in pts:
                        dist = road_line.project(pt)
                        km = (km_start_base * 1000) + dist
                        results.append({
                            'route': road.get('route', 'Unknown'),
                            'area_name': poly_name,
                            'km_in_m': km,
                            'km_out_m': km,
                            'length_m': 0,
                            'intersect_area_sqm': 0,
                            'center_x': pt.x,
                            'center_y': pt.y,
                            'properties': {k: self.safe_decode(v) if isinstance(v, str) else str(v) for k, v in poly.items() if k != 'geometry'}
                        })
                    continue
                
                # จัดการกรณีเป็นเส้น
                lines = []
                if isinstance(intersection, LineString):
                    lines = [intersection]
                elif hasattr(intersection, 'geoms'):
                    lines = [geom for geom in intersection.geoms if isinstance(geom, LineString)]
                
                for segment_idx, line_segment in enumerate(lines):
                    # คำนวณระยะทางจากจุดเริ่มต้นของถนนสายหลัก
                    start_dist = road_line.project(Point(line_segment.coords[0]))
                    end_dist = road_line.project(Point(line_segment.coords[-1]))
                    
                    # ป้องกันการสลับหน้าหลัง
                    actual_start = min(start_dist, end_dist)
                    actual_end = max(start_dist, end_dist)
                    
                    # หาจุดกึ่งกลางของเส้นตัดเพื่อถ่าย Street View
                    mid_pt = line_segment.interpolate(0.5, normalized=True)
                    
                    km_in = (km_start_base * 1000) + actual_start
                    km_out = (km_start_base * 1000) + actual_end
                    length_m = actual_end - actual_start
                    
                    # หยอดพื้นที่ทับซ้อนเฉพาะแถวแรกของโพลีกอนนี้เท่านั้นเพื่อป้องกันการบวกพื้นที่ซ้ำ
                    poly_unique_id = poly.get('ID', poly.get('objectid', poly.get('FID', p_idx)))
                    if poly_unique_id not in assigned_poly_ids:
                        seg_area = area_sqm
                        assigned_poly_ids.add(poly_unique_id)
                    else:
                        seg_area = 0
                    
                    results.append({
                        'route': road.get('route', 'Unknown'),
                        'area_name': poly_name,
                        'km_in_m': km_in,
                        'km_out_m': km_out,
                        'length_m': length_m,
                        'intersect_area_sqm': seg_area,
                        'center_x': mid_pt.x,
                        'center_y': mid_pt.y,
                        'properties': {k: self.safe_decode(v) if isinstance(v, str) else str(v) for k, v in poly.items() if k != 'geometry'}
                    })
                    
        # กฎพิเศษ: ถ้าเป็นชั้นข้อมูลหมู่บ้าน ต้องมีหมู่บ้านครอบคลุมทุกตำบลที่ถนนตัดผ่าน
        is_village_layer = 'บ้าน' in first_path or 'village' in first_path.lower()
        if is_village_layer:
            from config import BASE_DIR
            admin_shp = os.path.join(BASE_DIR, "ข้อมูลShp", "ขอบเขตการปกครอง", "ขอบเขตตำบล.shp")
            if os.path.exists(admin_shp) and not self.road_gdf.empty:
                try:
                    admin_gdf = gpd.read_file(admin_shp).to_crs(TARGET_CRS)
                    admin_gdf = admin_gdf[admin_gdf.is_valid & ~admin_gdf.is_empty]
                    road_union = self.road_gdf.geometry.unary_union
                    intersected_admin = admin_gdf[admin_gdf.intersects(road_union)]
                    crossed_tambol_names = set(intersected_admin['TAMBOL_TH'].dropna())
                    
                    covered_tambols = set()
                    for r in results:
                        t = r.get('properties', {}).get('TAMBOL_TH')
                        if t:
                            covered_tambols.add(t)
                            
                    missing_tambols = crossed_tambol_names - covered_tambols
                    if missing_tambols:
                        print(f"  [Info] ตำบลที่ถนนผ่านแต่ไม่มีหมู่บ้านในรัศมี {buffer_m} ม.: {missing_tambols}")
                        for tb in missing_tambols:
                            if 'TAMBOL_TH' not in poly_gdf.columns:
                                continue
                            tb_villages = poly_gdf[poly_gdf['TAMBOL_TH'] == tb]
                            if not tb_villages.empty:
                                distances = tb_villages.geometry.distance(road_union)
                                nearest_idx = distances.idxmin()
                                nearest_village = tb_villages.loc[nearest_idx]
                                
                                pt = nearest_village.geometry
                                min_dist_to_road = float('inf')
                                best_km = 0
                                best_route = 'Unknown'
                                
                                for r_idx, road in self.road_gdf.iterrows():
                                    rl = road.geometry
                                    if rl and not rl.is_empty:
                                        d = rl.distance(pt)
                                        if d < min_dist_to_road:
                                            min_dist_to_road = d
                                            km_base = road.get('km_start', 0.0)
                                            try: km_base = float(km_base)
                                            except: km_base = 0.0
                                            if pd.isna(km_base): km_base = 0.0
                                            if km_base > 2000: km_base = 0.0
                                            best_km = (km_base * 1000) + rl.project(pt)
                                            best_route = road.get('route', 'Unknown')
                                            
                                poly_name = self.safe_decode(nearest_village.get(actual_name_col, "Unknown"))
                                results.append({
                                    'route': best_route,
                                    'area_name': poly_name,
                                    'km_in_m': best_km,
                                    'km_out_m': best_km,
                                    'length_m': 0,
                                    'intersect_area_sqm': 0,
                                    'center_x': pt.x if hasattr(pt, 'x') else pt.centroid.x,
                                    'center_y': pt.y if hasattr(pt, 'y') else pt.centroid.y,
                                    'properties': {k: self.safe_decode(v) if isinstance(v, str) else str(v) for k, v in nearest_village.items() if k != 'geometry'}
                                })
                                print(f"  -> ดึงหมู่บ้าน {poly_name} (อ.{nearest_village.get('AMPHOE_TH','')} จ.{nearest_village.get('PROV_TH','')}) มาชดเชยให้ ต.{tb}")
                except Exception as e:
                    print(f"  [Warning] เกิดข้อผิดพลาดในการตรวจสอบหมู่บ้านตามขอบเขตตำบล: {e}")

        # --- Fallback: หากไม่พบจุดตัดในรัศมี 500ม. ให้ดึงจุดที่ใกล้ที่สุด 1 จุด (เฉพาะสถานศึกษา วัด รพ.)
        if len(results) == 0 and poly_gdf is not None and not poly_gdf.empty:
            if 'สถานศึกษา' in polygon_shp_path or 'วัด' in polygon_shp_path or 'โรงพยาบาล' in polygon_shp_path or 'รพสต' in polygon_shp_path:
                print(f"[{time.strftime('%X')}] ไม่พบสถานที่ในระยะตัดผ่าน ทำการดึงสถานที่ที่ใกล้ที่สุด 1 แห่งแทน...", flush=True)
                
                # หาจุดที่ใกล้ที่สุดจาก road_buf_union
                road_union_geom = self.road_gdf.geometry.union_all()
                distances = poly_gdf.geometry.distance(road_union_geom)
                nearest_idx = distances.idxmin()
                nearest_poly = poly_gdf.loc[nearest_idx]
                
                poly_name = self.safe_decode(nearest_poly.get(actual_name_col, "Unknown"))
                poly_geom = nearest_poly.geometry
                
                from shapely.geometry import Point
                pt = poly_geom.geoms[0] if hasattr(poly_geom, 'geoms') else poly_geom
                if not isinstance(pt, Point):
                    pt = pt.centroid
                
                # หา km อ้างอิงจากเส้นถนน
                road_line = self.road_gdf.geometry.iloc[0]
                km_start_base = self.road_gdf.iloc[0].get('km_start', 0)
                dist_on_line = road_line.project(pt)
                km = (km_start_base * 1000) + dist_on_line
                
                results.append({
                    'route': self.road_gdf.iloc[0].get('route', 'Unknown'),
                    'area_name': poly_name + " (อยู่นอกระยะ 500ม.)",
                    'km_in_m': km,
                    'km_out_m': km,
                    'length_m': 0,
                    'intersect_area_sqm': 0,
                    'center_x': pt.x,
                    'center_y': pt.y,
                    'properties': {k: self.safe_decode(v) if isinstance(v, str) else str(v) for k, v in nearest_poly.items() if k != 'geometry'}
                })

        return pd.DataFrame(results)

    def clip_and_save_shapefile(self, shp_path, buffer_m, output_path, excel_path=None, do_clip=True):
        """ตัดขอบเขต Shapefile ด้วย Buffer และเซฟไฟล์สำหรับนำไปจัด Layout พร้อมออปชั่น Export Excel"""
        from pyproj import CRS
        from shapely.geometry import box
        
        # ตรวจสอบว่ามีไฟล์ GPKG ที่แปลงไว้หรือไม่ ถ้ามีให้ใช้ GPKG แทน .shp เพื่อความเร็วสูงสุด
        if isinstance(shp_path, str):
            gpkg_path = shp_path.replace(".shp", ".gpkg")
            if os.path.exists(gpkg_path):
                shp_path = gpkg_path
            
        first_path = shp_path[0] if isinstance(shp_path, list) else shp_path
        print(f"Creating Layout Shapefile: {os.path.basename(output_path)}")
        road_buffer = self.road_gdf.copy()
        road_buffer['geometry'] = road_buffer.geometry.buffer(buffer_m)
        bounds = tuple(road_buffer.total_bounds)
        
        # ปรับแก้ CRS Bounding Box ให้ตรงกับ Shapefile ต้นฉบับ
        try:
            temp_check_gdf = gpd.read_file(first_path, rows=1)
            shp_crs = temp_check_gdf.crs
        except Exception:
            shp_crs = None
            
        read_bbox = bounds
        if shp_crs:
            try:
                crs_from = CRS.from_user_input(TARGET_CRS)
                crs_to = CRS.from_user_input(shp_crs)
                if not crs_from.equals(crs_to):
                    b = box(*bounds)
                    temp_gdf = gpd.GeoDataFrame(geometry=[b], crs=TARGET_CRS).to_crs(crs_to)
                    read_bbox = tuple(temp_gdf.total_bounds)
            except Exception as e:
                print(f"Warning: BBox CRS translation failed: {e}")
        
        poly_gdf = self._load_poly_gdf(shp_path, read_bbox)
        if poly_gdf.empty:
            print("  No intersecting features found (empty reading). Skipping.")
            return
            
        poly_gdf = poly_gdf.to_crs(TARGET_CRS)
        
        # [SPEED HACK] ถ้ารูปทรงใหญ่ระดับประเทศ การเอาไป make_valid หรือตัดกับ Buffer จะค้างนานมาก
        # ให้ตัดด้วยกรอบสี่เหลี่ยม (Bounding Box) ก่อน เพื่อหั่นรูปทรงให้เหลือแค่ชิ้นเล็กๆ ก่อนทำอย่างอื่น
        try:
            import shapely
            xmin, ymin, xmax, ymax = road_buffer.total_bounds
            poly_gdf['geometry'] = poly_gdf.geometry.apply(
                lambda g: shapely.clip_by_rect(g, xmin, ymin, xmax, ymax) if g and not g.is_empty else g
            )
            poly_gdf = poly_gdf[~poly_gdf.geometry.is_empty]
        except Exception as e:
            print(f"Warning: clip_by_rect failed: {e}")
            
        if poly_gdf.empty:
            print("  No intersecting features found after BBox clip. Skipping.")
            return

        # แก้ไขปัญหา "แหว่ง": ซ่อมแซม Geometry ที่ไม่สมบูรณ์ เฉพาะตัวที่พัง (ทำหลังจากหั่นให้เล็กแล้ว จะไวมาก)
        invalid_mask = ~poly_gdf.is_valid
        if invalid_mask.any():
            try:
                import pandas as pd
                fixed_geoms = poly_gdf.loc[invalid_mask, 'geometry'].make_valid()
                
                # ถ้า make_valid ได้ GeometryCollection ให้พยายามดึงมาเฉพาะ Polygon/MultiPolygon
                def keep_polygons(geom):
                    if geom.geom_type in ['Polygon', 'MultiPolygon']:
                        return geom
                    elif geom.geom_type == 'GeometryCollection':
                        from shapely.geometry import MultiPolygon
                        polys = [g for g in geom.geoms if g.geom_type in ['Polygon', 'MultiPolygon']]
                        return MultiPolygon(polys) if polys else geom
                    return geom
                
                poly_gdf.loc[invalid_mask, 'geometry'] = fixed_geoms.apply(keep_polygons)
            except Exception as e:
                print(f"Warning: Geometry repair failed: {e}")
            
        # กรองข้อมูลที่ว่างเปล่า 
        is_gpkg = isinstance(shp_path, str) and shp_path.endswith('.gpkg')
        if is_gpkg:
            poly_gdf = poly_gdf[~poly_gdf.geometry.is_empty]
        else:
            poly_gdf = poly_gdf[poly_gdf.is_valid & ~poly_gdf.is_empty]
            
        if poly_gdf.empty:
            print("  No intersecting features found. Skipping.")
            return

        print(f"Creating Layout Shapefile: {os.path.basename(output_path)}")
        
        # ตัดให้ขาดพอดีขอบเขต Buffer (สำหรับ ป่า, ลุ่มน้ำ, หมู่บ้าน)
        if do_clip:
            try:
                # แก้บั๊ก Buffer ม้วนตอนตัด ให้เชื่อมเป็นชิ้นเดียวกันก่อนตัด
                road_union = road_buffer.geometry.union_all()
                mask_gdf = gpd.GeoDataFrame(geometry=[road_union], crs=road_buffer.crs)
                clipped = gpd.clip(poly_gdf, mask_gdf)
            except Exception as e:
                print(f"Warning: Detailed clip failed, using bbox clipped. Error: {e}")
                clipped = poly_gdf
        else:
            # ไม่ตัดเนื้อข้างใน เอามาเฉพาะฟีเจอร์ที่ตัดผ่าน (สำหรับ ป่าเขตทาง, ชะล้างเขตทาง, เส้นแม่น้ำ)
            road_union = road_buffer.unary_union
            clipped = poly_gdf[poly_gdf.intersects(road_union)]
            
        if not clipped.empty:
            clipped = clipped.copy()
            # แก้บั๊ก: ป้องกัน Error "Coordinates with non-finite values are not allowed" ตอนเซฟไฟล์
            import numpy as np
            import shapely
            
            # บังคับให้เป็น 2D ทั้งหมด เพื่อลบค่าพิกัดแกน Z ที่อาจจะเป็น NaN ทิ้ง
            try:
                clipped['geometry'] = clipped.geometry.apply(lambda g: shapely.force_2d(g) if g else g)
            except Exception as e:
                print(f"Warning: force_2d failed: {e}")
                
            def is_valid_geom(geom):
                if geom is None or geom.is_empty: return False
                try:
                    # เช็คเฉพาะ X, Y เพราะ Z ถูก force_2d ลบทิ้งไปแล้ว (ใส่ include_z=True จะทำให้พัง)
                    coords = shapely.get_coordinates(geom)
                    if np.isnan(coords).any() or np.isinf(coords).any(): return False
                    return True
                except:
                    return False
            
            clipped = clipped[clipped.geometry.apply(is_valid_geom)]
            
            if not clipped.empty:
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                clipped.to_file(output_path)
                msg = "clipped" if do_clip else "filtered"
                print(f"  Saved {msg} shapefile to {output_path}")
            
        # Export Attribute Table เป็น Excel ถ้าระบุ path มา
        if excel_path and not clipped.empty:
            try:
                os.makedirs(os.path.dirname(excel_path), exist_ok=True)
                df_export = clipped.drop(columns=['geometry'])
                df_export.to_excel(excel_path, index=False, engine='openpyxl')
                print(f"  Saved attribute table to {excel_path}")
            except Exception as e:
                print(f"  Error saving attribute table to Excel: {e}")

    def enrich_with_admin_boundaries(self, gdf):
        """
        เชื่อมโยงข้อมูลกับขอบเขตตำบลเพื่อหา ตำบล อำเภอ จังหวัด
        """
        import os
        from config import BASE_DIR
        admin_shp = os.path.join(BASE_DIR, "ข้อมูลShp", "ขอบเขตการปกครอง", "ขอบเขตตำบล.shp")
        if not os.path.exists(admin_shp):
            print("  [Warning] ไม่พบไฟล์ ขอบเขตตำบล.shp ข้ามขั้นตอนการเชื่อมโยงขอบเขตปกครอง")
            return gdf
            
        try:
            admin_gdf = gpd.read_file(admin_shp).to_crs(TARGET_CRS)
            admin_gdf = admin_gdf[admin_gdf.is_valid & ~admin_gdf.is_empty]
            
            gdf_projected = gdf.copy()
            if 'geometry' in gdf_projected:
                # ใช้ Centroid เพื่อจุดตัดหรือโพลีกอนตัดกับขอบเขตการปกครองที่แม่นยำ
                gdf_projected['join_geom'] = gdf_projected.geometry.centroid
                gdf_projected = gdf_projected.set_geometry('join_geom')
                
                joined = gpd.sjoin(
                    gdf_projected, 
                    admin_gdf[['TAMBOL_TH', 'AMPHOE_TH', 'PROV_TH', 'geometry']], 
                    how='left', 
                    predicate='within'
                )
                
                # กู้คืน geometry เดิม
                joined = joined.set_geometry('geometry')
                joined.drop(columns=['index_right', 'join_geom'], errors='ignore', inplace=True)
                
                # ถอดรหัสอักขระไทยเพื่อให้อ่านได้ถูกต้อง
                for col in ['TAMBOL_TH', 'AMPHOE_TH', 'PROV_TH']:
                    if col in joined.columns:
                        joined[col] = joined[col].apply(lambda x: self.safe_decode(x) if x and not pd.isna(x) else "")
                
                # กรองเฉพาะข้อมูลที่อยู่ใน crossed_provinces (ถ้ามีการระบุไว้)
                if hasattr(self, 'crossed_provinces') and self.crossed_provinces and 'PROV_TH' in joined.columns:
                    original_len = len(joined)
                    def check_prov(p):
                        if not p: return False
                        for cp in self.crossed_provinces:
                            if cp in p or p in cp:
                                return True
                        return False
                    joined = joined[joined['PROV_TH'].apply(check_prov)]
                    if len(joined) < original_len:
                        print(f"  [Filter] กรองข้อมูลจุดที่อยู่นอกพื้นที่โครงการออก (เหลือ {len(joined)} จาก {original_len})")
                        
                return joined
        except Exception as e:
            print(f"  [Warning] เกิดข้อผิดพลาดในการเชื่อมโยงขอบเขตการปกครอง: {e}")
        return gdf

    def format_km(self, meters):
        """แปลงระยะเมตร เป็นรูปแบบ กม. เช่น 190350 -> 190+350"""
        km = int(meters // 1000)
        m = int(meters % 1000)
        return f"{km}+{m:03d}"

# สามารถทดสอบโค้ดได้โดยการรันไฟล์นี้ตรงๆ
if __name__ == "__main__":
    import config
    engine = GISEngine(config.SHP_DIR)
    engine.load_road_centerline()
    # ทดสอบรันกับป่าถาวร ด้วย Buffer 500m
    df = engine.calculate_intersections(os.path.join(config.SHP_DIR, "01_ป่า", "01_ป่าถาวร.shp"), "name_th", 500)
    
    # แปลง กม. ให้สวยงาม
    if not df.empty:
        df['KM In'] = df['km_in_m'].apply(engine.format_km)
        df['KM Out'] = df['km_out_m'].apply(engine.format_km)
        print("\n--- ผลลัพธ์การคำนวณป่าถาวร ---")
        print(df[['area_name', 'KM In', 'KM Out', 'length_m', 'intersect_area_sqm']])
    else:
        print("ไม่พบการตัดผ่าน")
