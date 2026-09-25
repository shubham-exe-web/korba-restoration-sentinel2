"""
Calculate exact centroid-to-pixel-center offsets and distance to pixel boundaries
for all 10 quadrats in Korba, Chhattisgarh, India using Sentinel-2 UTM Zone 44N raster grid.
"""

import os
import rasterio
import pandas as pd
import numpy as np

def calculate_offsets():
    raster_path = "rasters/Korba_20240325_stack.tif"
    coords_path = "data/01_Korba_quadrat_coordinates.csv"
    
    if not os.path.exists(raster_path):
        raise FileNotFoundError(f"Raster not found: {raster_path}")
    if not os.path.exists(coords_path):
        raise FileNotFoundError(f"Coordinates not found: {coords_path}")
        
    df_coords = pd.read_csv(coords_path)
    
    with rasterio.open(raster_path) as src:
        transform = src.transform
        res_x = transform[0]
        res_y = transform[4] # usually -10.0
        origin_x = transform[2]
        origin_y = transform[5]
        crs = src.crs
        
        print(f"Raster CRS: {crs}")
        print(f"Transform: {transform}")
        print(f"Pixel resolution: dx={res_x}, dy={res_y}")
        
        results = []
        for idx, row in df_coords.iterrows():
            qid = row['ID']
            mtype = row['Mining_type']
            utm_x = float(row['UTM_X_m'])
            utm_y = float(row['UTM_Y_m'])
            
            # Find row, col (0-indexed integer pixel indices)
            # rasterio.transform.rowcol takes (x, y)
            py_row, px_col = rasterio.transform.rowcol(transform, utm_x, utm_y)
            
            # Pixel center coordinates in UTM:
            # rasterio.transform.xy takes (row, col, offset='center')
            pixel_center_x, pixel_center_y = rasterio.transform.xy(transform, py_row, px_col, offset='center')
            
            # Pixel bounding box
            # offset='ul' gives upper-left corner
            ul_x, ul_y = rasterio.transform.xy(transform, py_row, px_col, offset='ul')
            lr_x = ul_x + res_x
            lr_y = ul_y + res_y # res_y is negative, so this is lower y
            
            pix_min_x = min(ul_x, lr_x)
            pix_max_x = max(ul_x, lr_x)
            pix_min_y = min(ul_y, lr_y)
            pix_max_y = max(ul_y, lr_y)
            
            delta_x = utm_x - pixel_center_x
            delta_y = utm_y - pixel_center_y
            euclidean_dist_center = np.sqrt(delta_x**2 + delta_y**2)
            
            # Distance to 4 boundaries of the pixel
            dist_left = utm_x - pix_min_x
            dist_right = pix_max_x - utm_x
            dist_bottom = utm_y - pix_min_y
            dist_top = pix_max_y - utm_y
            dist_nearest_boundary = min(dist_left, dist_right, dist_bottom, dist_top)
            
            # Area overlap of 10m x 10m quadrat centered on centroid with the 10m x 10m pixel:
            # Quadrat bbox:
            quad_min_x = utm_x - 5.0
            quad_max_x = utm_x + 5.0
            quad_min_y = utm_y - 5.0
            quad_max_y = utm_y + 5.0
            
            overlap_x = max(0.0, min(quad_max_x, pix_max_x) - max(quad_min_x, pix_min_x))
            overlap_y = max(0.0, min(quad_max_y, pix_max_y) - max(quad_min_y, pix_min_y))
            overlap_area = overlap_x * overlap_y
            overlap_pct = (overlap_area / 100.0) * 100.0
            
            results.append({
                'Quadrat_ID': qid,
                'Mining_Type': mtype,
                'Quadrat_Centroid_UTM_X': round(utm_x, 2),
                'Quadrat_Centroid_UTM_Y': round(utm_y, 2),
                'Pixel_Col': px_col,
                'Pixel_Row': py_row,
                'Pixel_Center_UTM_X': round(pixel_center_x, 2),
                'Pixel_Center_UTM_Y': round(pixel_center_y, 2),
                'Offset_Delta_X_m': round(delta_x, 2),
                'Offset_Delta_Y_m': round(delta_y, 2),
                'Distance_to_Pixel_Center_m': round(euclidean_dist_center, 2),
                'Distance_to_Nearest_Boundary_m': round(dist_nearest_boundary, 2),
                'Quadrat_Pixel_Overlap_m2': round(overlap_area, 1),
                'Quadrat_Pixel_Overlap_Pct': round(overlap_pct, 1)
            })
            
    df_out = pd.DataFrame(results)
    out_csv = "tables/Table_S1B_Quadrat_Centroid_Pixel_Offsets.csv"
    df_out.to_csv(out_csv, index=False)
    print(f"Saved {out_csv}:")
    print(df_out.to_string())
    
    mean_dist_center = df_out['Distance_to_Pixel_Center_m'].mean()
    range_dist_center = (df_out['Distance_to_Pixel_Center_m'].min(), df_out['Distance_to_Pixel_Center_m'].max())
    mean_dist_bound = df_out['Distance_to_Nearest_Boundary_m'].mean()
    range_dist_bound = (df_out['Distance_to_Nearest_Boundary_m'].min(), df_out['Distance_to_Nearest_Boundary_m'].max())
    mean_overlap = df_out['Quadrat_Pixel_Overlap_Pct'].mean()
    
    print("\nSummary Metrics:")
    print(f"Mean distance to pixel center: {mean_dist_center:.2f} m (range: {range_dist_center[0]:.2f} - {range_dist_center[1]:.2f} m)")
    print(f"Mean distance to nearest boundary: {mean_dist_bound:.2f} m (range: {range_dist_bound[0]:.2f} - {range_dist_bound[1]:.2f} m)")
    print(f"Mean nominal quadrat-pixel geometric overlap: {mean_overlap:.1f}%")

if __name__ == "__main__":
    calculate_offsets()
