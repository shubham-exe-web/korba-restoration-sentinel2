"""
Script: 05_extract_seasonal_features.py
Step 5: Recompute seasonal metrics, signed changes, ratios, and field-satellite linkage.
- Builds seasonal medians from retained valid dates per quadrat and spatial zone.
- Accurately names:
    Delta_NDVI_Monsoon_minus_Summer (Signed seasonal change)
    Delta_NDVI_Winter_minus_Monsoon (Signed post-monsoon senescence)
    Seasonal_NDVI_Amplitude (Monsoon NDVI - Summer NDVI)
    Winter_to_Monsoon_Ratio (Winter NDVI / Monsoon NDVI)
    Summer_to_Monsoon_Ratio (Summer NDVI / Monsoon NDVI)
- Strictly avoids claiming satellite identified species; labels remotely characterized signals as:
    - Canopy/Understory Vegetation Signal: Persistent Woody / Shrub Baseline
    - Canopy/Understory Vegetation Signal: Mixed Deciduous Canopy with Understory Turnover
    - Canopy/Understory Vegetation Signal: High Seasonal Green-Up Flush
- Fills and saves:
    - data/06_Korba_seasonal_spectral_features.csv
    - data/07_Korba_vegetation_indices.csv
    - data/09_Korba_field_satellite_linkage.csv
    - data/10_Korba_analysis_master.csv
    - tables/Table_4_Master_Field_RemoteSensing_Comparison.csv
    - tables/Table_Phase21_Vegetation_Richness_Phenology.csv
"""

import pandas as pd
import numpy as np

def run_seasonal_feature_extraction():
    print("Loading data/08_Korba_phenology_timeseries.csv...")
    df_ts = pd.read_csv("data/08_Korba_phenology_timeseries.csv")
    
    bands = ["B02", "B03", "B04", "B05", "B06", "B07", "B08", "B8A", "B11", "B12"]
    indices = ["NDVI", "EVI", "NDRE", "NDWI", "NDWI_Moisture", "RedEdge_Slope", "SWIR_Ratio"]
    
    # 1. Spectral bands aggregation (data/06_Korba_seasonal_spectral_features.csv)
    group_cols = ["Quadrat_ID", "Mining_type", "Spatial_Zone", "Season"]
    df_bands_agg = df_ts.groupby(group_cols)[bands].agg(["median", "std"]).reset_index()
    df_bands_agg.columns = [f"{c[0]}_{c[1]}" if c[1] else c[0] for c in df_bands_agg.columns]
    df_bands_agg.to_csv("data/06_Korba_seasonal_spectral_features.csv", index=False)
    print(f" - Saved data/06_Korba_seasonal_spectral_features.csv ({len(df_bands_agg)} rows)")
    
    # 2. Vegetation indices seasonal aggregation & pivoting
    df_idx_agg = df_ts.groupby(group_cols)[indices].agg("median").reset_index()
    
    pivoted = df_idx_agg.pivot(index=["Quadrat_ID", "Mining_type", "Spatial_Zone"], columns="Season")
    pivoted.columns = [f"{col[0]}_{col[1]}" for col in pivoted.columns]
    piv = pivoted.reset_index()
    
    # 3. Accurately named signed changes and ratios
    piv["Delta_NDVI_Monsoon_minus_Summer"] = (piv["NDVI_Monsoon"] - piv["NDVI_Summer"]).round(4)
    piv["Delta_NDVI_Winter_minus_Monsoon"] = (piv["NDVI_Winter"] - piv["NDVI_Monsoon"]).round(4)
    piv["Delta_NDRE_Monsoon_minus_Summer"] = (piv["NDRE_Monsoon"] - piv["NDRE_Summer"]).round(4)
    piv["Delta_NDRE_Winter_minus_Monsoon"] = (piv["NDRE_Winter"] - piv["NDRE_Monsoon"]).round(4)
    
    # Phenological amplitude
    piv["Seasonal_NDVI_Amplitude"] = piv["Delta_NDVI_Monsoon_minus_Summer"]
    
    # Persistence ratios
    piv["Winter_to_Monsoon_Ratio"] = (piv["NDVI_Winter"] / (piv["NDVI_Monsoon"] + 1e-6)).round(4)
    piv["Summer_to_Monsoon_Ratio"] = (piv["NDVI_Summer"] / (piv["NDVI_Monsoon"] + 1e-6)).round(4)
    piv["Winter_to_Monsoon_NDRE_Ratio"] = (piv["NDRE_Winter"] / (piv["NDRE_Monsoon"] + 1e-6)).round(4)
    
    piv.to_csv("data/07_Korba_vegetation_indices.csv", index=False)
    print(f" - Saved data/07_Korba_vegetation_indices.csv ({len(piv)} rows)")
    
    # 4. Field evidence linkage & signal categorization (data/09 & data/10)
    t1_df = pd.read_csv("tables/Table_1_Quadrat_Characteristics.csv")
    zone_a = piv[piv["Spatial_Zone"] == "Zone_A_10m"].copy()
    zone_b = piv[piv["Spatial_Zone"] == "Zone_B_50m"].copy()
    
    link_rows = []
    p21_rows = []
    
    for _, t1 in t1_df.iterrows():
        q_id = t1["Quadrat_ID"]
        za = zone_a[zone_a["Quadrat_ID"] == q_id].iloc[0]
        zb = zone_b[zone_b["Quadrat_ID"] == q_id].iloc[0]
        
        amp = za["Seasonal_NDVI_Amplitude"]
        w_ratio = za["Winter_to_Monsoon_Ratio"]
        s_ratio = za["Summer_to_Monsoon_Ratio"]
        
        # Defensible signal classification without overclaiming species-level identification
        if amp >= 0.22 and s_ratio < 0.55:
            sig = "Canopy/Understory Signal: High Seasonal Green-Up Flush"
            mech = "High seasonal amplitude driven by annual herbaceous green-up on open ground, followed by post-monsoon senescence."
        elif w_ratio >= 0.85 and s_ratio >= 0.60:
            sig = "Canopy/Understory Signal: Persistent Woody / Shrub Baseline"
            mech = "High dry-season vegetation persistence reflecting evergreen/semi-deciduous overstory trees and perennial woody shrub layer."
        else:
            sig = "Canopy/Understory Signal: Mixed Deciduous Canopy with Understory Turnover"
            mech = "Intermediate seasonal amplitude reflecting deciduous crown leaf-drop and localized understory regeneration."
            
        link_rows.append({
            "Quadrat_ID": q_id,
            "Mining_type": t1["Mining_type"],
            "Latitude": t1["Latitude"],
            "Longitude": t1["Longitude"],
            "Elevation_m": t1["Elevation_m"],
            "Colliery_Sector": t1["Colliery_Sector"],
            "Verified_Overstory_Stems": t1["Verified_Overstory_Stems"],
            "Verified_Tree_Species_Richness": t1["Verified_Tree_Species_Richness"],
            "Dominant_Tree_Taxa": t1["Dominant_Tree_Taxa"],
            "AGB_t_ha": t1["AGB_t_ha"],
            "Carbon_Stock_tC_ha": t1["Carbon_Stock_tC_ha"],
            "In_Plot_Woody_Shrub_Status": t1["In_Plot_Woody_Shrub_Status"],
            "Field_Plate_Understory_Observation": t1["Field_Plate_Understory_Observation"],
            "ZoneA_NDVI_Summer": za["NDVI_Summer"],
            "ZoneA_NDVI_Monsoon": za["NDVI_Monsoon"],
            "ZoneA_NDVI_Winter": za["NDVI_Winter"],
            "ZoneA_Seasonal_Amplitude": amp,
            "ZoneA_Winter_to_Monsoon_Ratio": w_ratio,
            "ZoneA_Summer_to_Monsoon_Ratio": s_ratio,
            "ZoneB_NDVI_Summer": zb["NDVI_Summer"],
            "ZoneB_NDVI_Monsoon": zb["NDVI_Monsoon"],
            "ZoneB_NDVI_Winter": zb["NDVI_Winter"],
            "ZoneB_Seasonal_Amplitude": zb["Seasonal_NDVI_Amplitude"],
            "ZoneB_Winter_to_Monsoon_Ratio": zb["Winter_to_Monsoon_Ratio"],
            "Remotely_Characterized_Signal": sig,
            "Phenological_Mechanism_Rationale": mech
        })
        
        p21_rows.append({
            "Site": q_id,
            "Type": "UG" if "UG" in q_id else "OC",
            "Verified_Tree_Richness": t1["Verified_Tree_Species_Richness"],
            "Verified_Tree_Stems": t1["Verified_Overstory_Stems"],
            "Tree_Carbon_Stock_tC_ha": t1["Carbon_Stock_tC_ha"],
            "Summer_NDVI": za["NDVI_Summer"],
            "Monsoon_NDVI": za["NDVI_Monsoon"],
            "Winter_NDVI": za["NDVI_Winter"],
            "Seasonal_Amplitude": amp,
            "Winter_to_Monsoon_Ratio": w_ratio,
            "In_Plot_Woody_Shrub_Status": t1["In_Plot_Woody_Shrub_Status"],
            "Remotely_Characterized_Signal": sig
        })
        
    df_link = pd.DataFrame(link_rows)
    df_link.to_csv("data/09_Korba_field_satellite_linkage.csv", index=False)
    df_link.to_csv("data/10_Korba_analysis_master.csv", index=False)
    print(" - Saved data/09_Korba_field_satellite_linkage.csv & data/10_Korba_analysis_master.csv")
    
    # Table 4: Master Comparison
    df_table4 = df_link[[
        "Quadrat_ID", "Mining_type", "Verified_Tree_Species_Richness", "Verified_Overstory_Stems", "Carbon_Stock_tC_ha",
        "In_Plot_Woody_Shrub_Status", "ZoneA_NDVI_Summer", "ZoneA_NDVI_Monsoon", "ZoneA_NDVI_Winter",
        "ZoneA_Seasonal_Amplitude", "ZoneA_Winter_to_Monsoon_Ratio", "Remotely_Characterized_Signal"
    ]].copy()
    df_table4.to_csv("tables/Table_4_Master_Field_RemoteSensing_Comparison.csv", index=False)
    print(" - Saved tables/Table_4_Master_Field_RemoteSensing_Comparison.csv")
    
    # Repaired Phase 21 Richness & Phenology Table
    df_p21 = pd.DataFrame(p21_rows)
    df_p21.to_csv("tables/Table_Phase21_Vegetation_Richness_Phenology.csv", index=False)
    print(" - Saved tables/Table_Phase21_Vegetation_Richness_Phenology.csv")
    print("\n=== Repaired Phase 21 Table Preview ===")
    print(df_p21[["Site", "Type", "Verified_Tree_Richness", "Verified_Tree_Stems", "Summer_NDVI", "Monsoon_NDVI", "Winter_NDVI", "Seasonal_Amplitude", "Winter_to_Monsoon_Ratio"]].to_string(index=False))

if __name__ == "__main__":
    run_seasonal_feature_extraction()
