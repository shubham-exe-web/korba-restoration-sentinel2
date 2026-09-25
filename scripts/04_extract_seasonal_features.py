"""
Script: 04_extract_seasonal_features.py
Phase 8: Calculate seasonal spectral features and vegetation indices across Summer, Monsoon, and Winter.
Computes:
- Band medians (B02 to B12) per season and spatial zone.
- Vegetation indices (NDVI, EVI, NDRE, NDWI, NDWI_Moisture, RedEdge_Slope, SWIR_Ratio).
- Seasonal dynamics: Delta-NDVI (Monsoon - Summer, Winter - Monsoon), Amplitude, and Persistence ratios.
Outputs:
- data/06_Korba_seasonal_spectral_features.csv
- data/07_Korba_vegetation_indices.csv
"""

import pandas as pd
import numpy as np

def main():
    print("Loading data/08_Korba_phenology_timeseries.csv...")
    df_ts = pd.read_csv("data/08_Korba_phenology_timeseries.csv")
    
    bands = ["B02", "B03", "B04", "B05", "B06", "B07", "B08", "B8A", "B11", "B12"]
    indices = ["NDVI", "EVI", "NDRE", "NDWI", "NDWI_Moisture", "RedEdge_Slope", "SWIR_Ratio"]
    
    # 1. Compute seasonal aggregation (median and standard deviation across observations)
    group_cols = ["Quadrat_ID", "Mining_type", "Spatial_Zone", "Season"]
    
    # Spectral bands
    df_bands_agg = df_ts.groupby(group_cols)[bands].agg(["median", "std"]).reset_index()
    # Flatten column multi-index
    df_bands_agg.columns = [f"{c[0]}_{c[1]}" if c[1] else c[0] for c in df_bands_agg.columns]
    df_bands_agg.to_csv("data/06_Korba_seasonal_spectral_features.csv", index=False)
    print(f"Saved data/06_Korba_seasonal_spectral_features.csv ({len(df_bands_agg)} rows)")
    
    # Vegetation indices
    df_idx_agg = df_ts.groupby(group_cols)[indices].agg("median").reset_index()
    
    # Pivot indices to wide format by Season to compute delta and persistence metrics
    pivoted = df_idx_agg.pivot(index=["Quadrat_ID", "Mining_type", "Spatial_Zone"], columns="Season")
    
    # Flatten multiindex columns: e.g. ('NDVI', 'Summer') -> 'NDVI_Summer'
    pivoted.columns = [f"{col[0]}_{col[1]}" for col in pivoted.columns]
    piv = pivoted.reset_index()
    
    # 2. Compute Seasonal Dynamic Metrics
    # Seasonal amplitude / Green-up flush (Monsoon - Summer)
    piv["Delta_NDVI_Monsoon_Summer"] = (piv["NDVI_Monsoon"] - piv["NDVI_Summer"]).round(4)
    piv["Delta_NDRE_Monsoon_Summer"] = (piv["NDRE_Monsoon"] - piv["NDRE_Summer"]).round(4)
    piv["Delta_EVI_Monsoon_Summer"] = (piv["EVI_Monsoon"] - piv["EVI_Summer"]).round(4)
    
    # Senescence / Post-monsoon decay (Winter - Monsoon)
    piv["Delta_NDVI_Winter_Monsoon"] = (piv["NDVI_Winter"] - piv["NDVI_Monsoon"]).round(4)
    piv["Delta_NDRE_Winter_Monsoon"] = (piv["NDRE_Winter"] - piv["NDRE_Monsoon"]).round(4)
    
    # Phenological amplitude
    piv["Seasonal_NDVI_Amplitude"] = piv["Delta_NDVI_Monsoon_Summer"]
    
    # Vegetation Persistence
    # Winter persistence: ratio of winter greenness to monsoon peak
    piv["Winter_Persistence_Ratio"] = (piv["NDVI_Winter"] / (piv["NDVI_Monsoon"] + 1e-6)).round(4)
    # Summer baseline persistence: ratio of summer baseline to monsoon peak
    piv["Summer_Baseline_Persistence"] = (piv["NDVI_Summer"] / (piv["NDVI_Monsoon"] + 1e-6)).round(4)
    # Red-edge persistence
    piv["Winter_NDRE_Persistence"] = (piv["NDRE_Winter"] / (piv["NDRE_Monsoon"] + 1e-6)).round(4)
    
    # Save 07_Korba_vegetation_indices.csv
    piv.to_csv("data/07_Korba_vegetation_indices.csv", index=False)
    print(f"Saved data/07_Korba_vegetation_indices.csv ({len(piv)} rows)")
    
    print("\n=== Sample Vegetation Indices (Zone A - 10m Quadrat Scale) ===")
    zone_a = piv[piv["Spatial_Zone"] == "Zone_A_10m"][
        ["Quadrat_ID", "Mining_type", "NDVI_Summer", "NDVI_Monsoon", "NDVI_Winter", "Seasonal_NDVI_Amplitude", "Winter_Persistence_Ratio"]
    ]
    print(zone_a.to_string(index=False))

if __name__ == "__main__":
    main()
