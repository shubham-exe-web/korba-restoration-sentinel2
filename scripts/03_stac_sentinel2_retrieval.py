"""
Script: 03_stac_sentinel2_retrieval.py
Phases 4, 5, 6, 7:
- Multi-temporal Sentinel-2 Level-2A surface reflectance retrieval across Summer, Monsoon, and Winter.
- Pixel-level SCL Cloud/Shadow Quality Control.
- Multi-scale spatial window extraction (Zone A: 10m, Zone B: 30m, Zone C: 50m, Zone D: 100m).
- Generation of Scene Inventory (04), QC table (05 & Table 3), and Phenology Time Series (08).
"""

import os
import requests
import numpy as np
import pandas as pd
import rasterio
from rasterio.windows import from_bounds
from pyproj import Transformer
from scipy.ndimage import zoom
import time

# 1. Load quadrat coordinates
coords_df = pd.read_csv("data/01_Korba_quadrat_coordinates.csv")
trans = Transformer.from_crs("EPSG:4326", "EPSG:32644", always_xy=True)
coords_df["UTM_X"], coords_df["UTM_Y"] = trans.transform(coords_df["Longitude"].values, coords_df["Latitude"].values)

# Master bounding box covering all 10 quadrats with 500m buffer
min_x = coords_df["UTM_X"].min() - 500
max_x = coords_df["UTM_X"].max() + 500
min_y = coords_df["UTM_Y"].min() - 500
max_y = coords_df["UTM_Y"].max() + 500

print(f"Extraction Bounding Box (UTM 44N): X=[{min_x:.1f}, {max_x:.1f}], Y=[{min_y:.1f}, {max_y:.1f}]")
width_m = max_x - min_x
height_m = max_y - min_y
print(f"Dimensions: {width_m/1000:.2f} km x {height_m/1000:.2f} km")

# 2. Multi-temporal observations: 2 cloud-screened dates per season (6 total)
scenes_meta = [
    # Summer (Pre-monsoon dry season)
    {"id": "S2A_44QPK_20240315_0_L2A", "date": "2024-03-15", "season": "Summer", "cloud_scene_pct": 0.00, "url_base": "https://sentinel-cogs.s3.us-west-2.amazonaws.com/sentinel-s2-l2a-cogs/44/Q/PK/2024/3/S2A_44QPK_20240315_0_L2A"},
    {"id": "S2A_44QPK_20240424_0_L2A", "date": "2024-04-24", "season": "Summer", "cloud_scene_pct": 0.00, "url_base": "https://sentinel-cogs.s3.us-west-2.amazonaws.com/sentinel-s2-l2a-cogs/44/Q/PK/2024/4/S2A_44QPK_20240424_0_L2A"},
    
    # Monsoon (Wet season herbaceous & canopy flush)
    {"id": "S2A_44QPK_20240613_0_L2A", "date": "2024-06-13", "season": "Monsoon", "cloud_scene_pct": 4.28, "url_base": "https://sentinel-cogs.s3.us-west-2.amazonaws.com/sentinel-s2-l2a-cogs/44/Q/PK/2024/6/S2A_44QPK_20240613_0_L2A"},
    {"id": "S2B_44QPK_20241006_0_L2A", "date": "2024-10-06", "season": "Monsoon", "cloud_scene_pct": 4.11, "url_base": "https://sentinel-cogs.s3.us-west-2.amazonaws.com/sentinel-s2-l2a-cogs/44/Q/PK/2024/10/S2B_44QPK_20241006_0_L2A"},
    
    # Winter (Post-monsoon dry season)
    {"id": "S2A_44QPK_20241120_0_L2A", "date": "2024-11-20", "season": "Winter", "cloud_scene_pct": 0.00, "url_base": "https://sentinel-cogs.s3.us-west-2.amazonaws.com/sentinel-s2-l2a-cogs/44/Q/PK/2024/11/S2A_44QPK_20241120_0_L2A"},
    {"id": "S2A_44QPK_20241230_0_L2A", "date": "2024-12-30", "season": "Winter", "cloud_scene_pct": 0.49, "url_base": "https://sentinel-cogs.s3.us-west-2.amazonaws.com/sentinel-s2-l2a-cogs/44/Q/PK/2024/12/S2A_44QPK_20241230_0_L2A"}
]

# Save Scene Inventory (04_Sentinel2_scene_inventory.csv)
df_scenes = pd.DataFrame(scenes_meta)
df_scenes.drop(columns=["url_base"]).to_csv("data/04_Sentinel2_scene_inventory.csv", index=False)
print("Saved data/04_Sentinel2_scene_inventory.csv")

# Bands to extract
bands_10m = ["B02", "B03", "B04", "B08"]
bands_20m = ["B05", "B06", "B07", "B8A", "B11", "B12", "SCL"]

zone_specs = {
    "Zone_A_10m": {"radius": 5, "type": "square"},
    "Zone_B_50m": {"radius": 15, "type": "square"},
    "Zone_C_50m": {"radius": 50, "type": "circle"},
    "Zone_D_100m": {"radius": 100, "type": "circle"}
}

all_timeseries_records = []
qc_records = []

# Configure GDAL environment for rapid remote COG window reads
env_config = {
    "GDAL_DISABLE_READDIR_ON_OPEN": "EMPTY_DIR",
    "CPL_VSIL_CURL_ALLOWED_EXTENSIONS": ".tif",
    "VSI_CACHE": "TRUE",
    "VSI_CACHE_SIZE": "50000000"
}

with rasterio.Env(**env_config):
    for s_idx, scene in enumerate(scenes_meta):
        t_scene_start = time.time()
        print(f"\nProcessing scene [{s_idx+1}/{len(scenes_meta)}]: {scene['id']} ({scene['date']} - {scene['season']})")
        url_base = scene["url_base"]
        
        band_arrays = {}
        master_transform = None
        
        # 1. Read 10m bands
        read_success = True
        for b in bands_10m:
            band_url = f"{url_base}/{b}.tif"
            try:
                with rasterio.open(band_url) as src:
                    win = from_bounds(min_x, min_y, max_x, max_y, src.transform)
                    band_arrays[b] = src.read(1, window=win).astype(float)
                    if master_transform is None:
                        master_transform = rasterio.windows.transform(win, src.transform)
            except Exception as e:
                print(f"Error reading {band_url}: {e}")
                read_success = False
                break
                
        if not read_success:
            print(f"Skipping scene {scene['id']} due to 10m read error.")
            continue
            
        ref_shape = band_arrays["B04"].shape
        
        # 2. Read 20m bands and resample via scipy zoom in memory
        for b in bands_20m:
            band_url = f"{url_base}/{b}.tif"
            try:
                with rasterio.open(band_url) as src:
                    win = from_bounds(min_x, min_y, max_x, max_y, src.transform)
                    raw_arr = src.read(1, window=win).astype(float)
                    zoom_y = ref_shape[0] / raw_arr.shape[0]
                    zoom_x = ref_shape[1] / raw_arr.shape[1]
                    order = 0 if b == "SCL" else 1
                    band_arrays[b] = zoom(raw_arr, (zoom_y, zoom_x), order=order)
            except Exception as e:
                print(f"Error reading {band_url}: {e}")
                read_success = False
                break
                
        if not read_success:
            continue
            
        inv_transform = ~master_transform

        for _, q in coords_df.iterrows():
            q_id = q["ID"]
            mining_type = q["Mining_type"]
            qx, qy = q["UTM_X"], q["UTM_Y"]
            
            c_col, c_row = inv_transform * (qx, qy)
            c_row, c_col = int(round(c_row)), int(round(c_col))
            
            # SCL quality check within 50m window
            r50_pix = int(round(50 / 10.0))
            r_slice = slice(max(0, c_row - r50_pix), min(ref_shape[0], c_row + r50_pix + 1))
            c_slice = slice(max(0, c_col - r50_pix), min(ref_shape[1], c_col + r50_pix + 1))
            scl_window = band_arrays["SCL"][r_slice, c_slice]
            
            # SCL: 4=veg, 5=bare soil, 6=water, 7=unclassified
            valid_mask = np.isin(scl_window, [4, 5, 6, 7])
            valid_pct = float(np.mean(valid_mask) * 100.0) if scl_window.size > 0 else 0.0
            is_usable = valid_pct >= 50.0
            
            qc_records.append({
                "Scene_ID": scene["id"],
                "Date": scene["date"],
                "Season": scene["season"],
                "Quadrat_ID": q_id,
                "Mining_type": mining_type,
                "Scene_Cloud_Pct": scene["cloud_scene_pct"],
                "Quadrat_Valid_Pixel_Pct": round(valid_pct, 2),
                "Retained": is_usable
            })
            
            if not is_usable:
                continue
                
            for z_name, z_spec in zone_specs.items():
                rad_m = z_spec["radius"]
                rad_pix = int(round(rad_m / 10.0))
                
                row_start = max(0, c_row - rad_pix)
                row_end = min(ref_shape[0], c_row + rad_pix + 1)
                col_start = max(0, c_col - rad_pix)
                col_end = min(ref_shape[1], c_col + rad_pix + 1)
                
                rows_grid, cols_grid = np.ogrid[row_start:row_end, col_start:col_end]
                if z_spec["type"] == "circle":
                    dist_sq = (rows_grid - c_row)**2 + (cols_grid - c_col)**2
                    mask = dist_sq <= (rad_m / 10.0)**2
                else:
                    mask = np.ones((row_end - row_start, col_end - col_start), dtype=bool)
                    
                if not np.any(mask):
                    mask = np.ones((1, 1), dtype=bool)
                    row_start, row_end = c_row, c_row + 1
                    col_start, col_end = c_col, c_col + 1
                    
                band_vals = {}
                for b in bands_10m + bands_20m:
                    sub_arr = band_arrays[b][row_start:row_end, col_start:col_end]
                    band_vals[b] = float(np.mean(sub_arr[mask]))
                    
                # Sentinel-2 BOA surface reflectance = DN / 10000.0
                b02 = band_vals["B02"] / 10000.0
                b03 = band_vals["B03"] / 10000.0
                b04 = band_vals["B04"] / 10000.0
                b05 = band_vals["B05"] / 10000.0
                b06 = band_vals["B06"] / 10000.0
                b07 = band_vals["B07"] / 10000.0
                b08 = band_vals["B08"] / 10000.0
                b8a = band_vals["B8A"] / 10000.0
                b11 = band_vals["B11"] / 10000.0
                b12 = band_vals["B12"] / 10000.0
                
                ndvi = (b08 - b04) / (b08 + b04 + 1e-6)
                evi = 2.5 * (b08 - b04) / (b08 + 6.0 * b04 - 7.5 * b02 + 1.0 + 1e-6)
                ndre = (b08 - b05) / (b08 + b05 + 1e-6)
                ndwi = (b03 - b08) / (b03 + b08 + 1e-6)
                ndwi_moist = (b08 - b11) / (b08 + b11 + 1e-6)
                re_slope = (b07 - b05) / (b07 + b05 + 1e-6)
                swir_ratio = b11 / (b12 + 1e-6)
                
                all_timeseries_records.append({
                    "Date": scene["date"],
                    "Season": scene["season"],
                    "Scene_ID": scene["id"],
                    "Quadrat_ID": q_id,
                    "Mining_type": mining_type,
                    "Spatial_Zone": z_name,
                    "B02": round(b02, 4),
                    "B03": round(b03, 4),
                    "B04": round(b04, 4),
                    "B05": round(b05, 4),
                    "B06": round(b06, 4),
                    "B07": round(b07, 4),
                    "B08": round(b08, 4),
                    "B8A": round(b8a, 4),
                    "B11": round(b11, 4),
                    "B12": round(b12, 4),
                    "NDVI": round(ndvi, 4),
                    "EVI": round(evi, 4),
                    "NDRE": round(ndre, 4),
                    "NDWI": round(ndwi, 4),
                    "NDWI_Moisture": round(ndwi_moist, 4),
                    "RedEdge_Slope": round(re_slope, 4),
                    "SWIR_Ratio": round(swir_ratio, 4)
                })
        print(f"Completed scene {scene['id']} in {time.time() - t_scene_start:.2f}s")

# Save outputs
df_ts = pd.DataFrame(all_timeseries_records)
df_ts.to_csv("data/08_Korba_phenology_timeseries.csv", index=False)
print(f"\nSaved data/08_Korba_phenology_timeseries.csv ({len(df_ts)} records)")

df_qc = pd.DataFrame(qc_records)
df_qc.to_csv("data/05_Sentinel2_QC.csv", index=False)
print(f"Saved data/05_Sentinel2_QC.csv ({len(df_qc)} records)")

qc_agg = df_qc.groupby(["Quadrat_ID", "Mining_type", "Season"]).agg(
    Images_Available=("Scene_ID", "count"),
    Images_Retained=("Retained", "sum"),
    Mean_Valid_Pixel_Pct=("Quadrat_Valid_Pixel_Pct", "mean")
).reset_index()

qc_agg["Mean_Valid_Pixel_Pct"] = qc_agg["Mean_Valid_Pixel_Pct"].round(1)
qc_agg.to_csv("tables/Table_3_Seasonal_Sentinel2_QC.csv", index=False)
print("Saved tables/Table_3_Seasonal_Sentinel2_QC.csv")
print("\n=== Table 3: Seasonal Sentinel-2 QC Preview ===")
print(qc_agg.head(15))
