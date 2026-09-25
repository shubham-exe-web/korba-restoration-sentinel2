"""
Script: 03_query_imagery_manifest.py
Step 3: Query and preserve the complete Sentinel-2 L2A candidate archive manifest for Korba (Tile 44QPK).
- Searches all acquisitions across the declared 2024 study period (2024-01-01 to 2024-12-31).
- Catalogs item ID, UTC & IST timestamps, cloud cover %, IMD seasonal assignment, processing metadata, and asset URLs.
- Delineates candidate retention status and flags preliminary vs comprehensive observations.
- Generates data/04_Sentinel2_Candidate_Archive_Manifest.csv and data/04_Sentinel2_scene_inventory.csv.
"""

import requests
import pandas as pd
from datetime import datetime, timezone, timedelta

def query_and_save_manifest():
    print("Querying Element84 AWS Sentinel-2 L2A STAC archive for Tile 44QPK across 2024...")
    url = "https://earth-search.aws.element84.com/v1/search"
    bbox = [82.55, 22.34, 82.65, 22.44] # Covers all 10 quadrats
    
    payload = {
        "collections": ["sentinel-2-l2a"],
        "bbox": bbox,
        "datetime": "2024-01-01T00:00:00Z/2024-12-31T23:59:59Z",
        "limit": 150
    }
    
    response = requests.post(url, json=payload, timeout=30)
    if response.status_code != 200:
        raise RuntimeError(f"STAC API query failed with status code {response.status_code}")
        
    data = response.json()
    features = data.get("features", [])
    print(f"Retrieved {len(features)} total candidate scenes for Korba in 2024.")
    
    # Preliminary 6 dates selected previously
    preliminary_ids = {
        "S2A_44QPK_20240315_0_L2A",
        "S2A_44QPK_20240424_0_L2A",
        "S2A_44QPK_20240613_0_L2A",
        "S2B_44QPK_20241006_0_L2A",
        "S2A_44QPK_20241120_0_L2A",
        "S2A_44QPK_20241230_0_L2A"
    }
    
    ist_tz = timezone(timedelta(hours=5, minutes=30))
    records = []
    
    for f in features:
        props = f["properties"]
        item_id = f["id"]
        dt_utc_str = props.get("datetime")
        dt_utc = datetime.fromisoformat(dt_utc_str.replace("Z", "+00:00"))
        dt_ist = dt_utc.astimezone(ist_tz)
        
        month = dt_ist.month
        day = dt_ist.day
        
        # Season assignment following IMD & CGWB Korba District Climatology:
        # Summer (Pre-monsoon): March 1 - May 31
        # Southwest Monsoon: June 1 - September 30 (with retreat in early Oct)
        # Winter / Post-monsoon: October 1 - February 28
        if month in [3, 4, 5]:
            season_imd = "Summer (Pre-monsoon)"
            season_code = "Summer"
        elif month in [6, 7, 8, 9]:
            season_imd = "Southwest Monsoon"
            season_code = "Monsoon"
        else:
            season_imd = "Winter / Post-monsoon"
            season_code = "Winter"
            
        cloud_pct = props.get("eo:cloud_cover", 100.0)
        platform = props.get("platform", "Sentinel-2")
        grid_square = props.get("grid:code", "MGRS-44QPK")
        
        assets = f.get("assets", {})
        b04_href = assets.get("red", {}).get("href", "")
        b08_href = assets.get("nir", {}).get("href", "")
        scl_href = assets.get("scl", {}).get("href", "")
        
        # Categorize retention tier
        is_prelim = item_id in preliminary_ids
        if is_prelim:
            status = "Selected - Primary Benchmark Run"
        elif cloud_pct < 10.0:
            status = "Retained - Clear Candidate"
        elif cloud_pct < 30.0:
            status = "Retained - Moderately Clear Candidate"
        else:
            status = "Excluded - Cloud Cover >30%"
            
        records.append({
            "Item_ID": item_id,
            "Acquisition_UTC": dt_utc_str,
            "Acquisition_IST": dt_ist.strftime("%Y-%m-%d %H:%M:%S"),
            "Date_YYYY_MM_DD": dt_ist.strftime("%Y-%m-%d"),
            "Tile_MGRS": "44QPK",
            "Platform": platform,
            "Season_IMD_Climatology": season_imd,
            "Season_Code": season_code,
            "Scene_Cloud_Cover_Pct": round(cloud_pct, 2),
            "Processing_Level": "Level-2A (Surface Reflectance BOA)",
            "STAC_Collection": "sentinel-2-l2a",
            "Asset_URL_B04": b04_href,
            "Asset_URL_B08": b08_href,
            "Asset_URL_SCL": scl_href,
            "Archive_Status": status,
            "Preliminary_Run_Included": is_prelim
        })
        
    df_manifest = pd.DataFrame(records).sort_values("Date_YYYY_MM_DD").reset_index(drop=True)
    df_manifest.to_csv("data/04_Sentinel2_Candidate_Archive_Manifest.csv", index=False)
    print(" - Saved data/04_Sentinel2_Candidate_Archive_Manifest.csv")
    
    # Save the active scene inventory (04_Sentinel2_scene_inventory.csv)
    active = df_manifest[df_manifest["Preliminary_Run_Included"]].copy()
    active.to_csv("data/04_Sentinel2_scene_inventory.csv", index=False)
    print(f" - Saved data/04_Sentinel2_scene_inventory.csv ({len(active)} active observation scenes)")
    
    print("\nCandidate Manifest Summary by IMD Season:")
    print(df_manifest.groupby(["Season_Code", "Archive_Status"])["Item_ID"].count())

if __name__ == "__main__":
    query_and_save_manifest()
