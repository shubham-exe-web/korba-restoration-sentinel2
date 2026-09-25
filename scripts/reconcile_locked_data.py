"""
Script: scripts/reconcile_locked_data.py
Description:
1. Replaces all instances of 'Zone_B_30m' with 'Zone_B_50m' in:
   - data/08_Korba_phenology_timeseries.csv
   - data/06_Korba_seasonal_spectral_features.csv
   - data/07_Korba_vegetation_indices.csv
   - data/09_Korba_field_satellite_linkage.csv
   - data/10_Korba_analysis_master.csv
   - data/Sentinel2_Audited_Observation_Provenance_Table.csv
   - tables/Table_S2_Sentinel2_Observation_Provenance.csv
2. Builds the definitive locked master dataset with dual explicit aliases:
   - data/Definitive_Locked_Site_By_Season_Indices.csv
   - tables/Table_Locked_Master_Site_By_Season_Indices.csv
3. Updates tables/Table_4_Phenological_Syndromes.csv with toned-down ecological descriptions.
4. Generates data/DATA_DICTIONARY_LOCKED_INDICES.md.
5. Generates tables/Table_S8_Syndrome_Threshold_Sensitivity_Analysis.csv and Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv.
"""

import os
import pandas as pd
import numpy as np

def reconcile():
    print("Step 1: Renaming Zone_B_30m -> Zone_B_50m across all data CSVs...")
    files_to_update_zone = [
        "data/08_Korba_phenology_timeseries.csv",
        "data/06_Korba_seasonal_spectral_features.csv",
        "data/07_Korba_vegetation_indices.csv",
        "data/09_Korba_field_satellite_linkage.csv",
        "data/10_Korba_analysis_master.csv",
        "data/Sentinel2_Audited_Observation_Provenance_Table.csv",
        "tables/Table_S2_Sentinel2_Observation_Provenance.csv"
    ]
    
    for fp in files_to_update_zone:
        if os.path.exists(fp):
            with open(fp, "r", encoding="utf-8") as f:
                content = f.read()
            updated = content.replace("Zone_B_30m", "Zone_B_50m")
            with open(fp, "w", encoding="utf-8") as f:
                f.write(updated)
            print(f" - Updated {fp}")

    print("\nStep 2: Building Definitive Locked Site-by-Season Master Table with Dual Aliases...")
    df_indices = pd.read_csv("data/07_Korba_vegetation_indices.csv")
    
    zone_dim_map = {
        "Zone_A_10m": "10 × 10 m (1 pixel, 0.01 ha nominal support)",
        "Zone_B_50m": "50 × 50 m (25 pixels, 0.25 ha)",
        "Zone_C_50m": "Radius 50 m disc (81 pixels, 0.81 ha)",
        "Zone_D_100m": "Radius 100 m disc (317 pixels, 3.17 ha)"
    }
    
    df_indices["Zone_Dimensions"] = df_indices["Spatial_Zone"].map(zone_dim_map)
    
    # Classification rules for Zone A (Heuristic descriptive classes)
    def assign_rule(row):
        if row["Spatial_Zone"] != "Zone_A_10m":
            return "N/A (Multi-scale support)"
        s = row["NDVI_Summer"]
        r = row["Winter_to_Monsoon_Ratio"]
        a = row["Seasonal_NDVI_Amplitude"]
        q = row["Quadrat_ID"]
        
        if s >= 0.30 and r >= 0.90:
            return "Persistent Greenness Signal"
        elif a >= 0.25 and s < 0.25:
            return "Pronounced Seasonal Flush"
        else:
            if q == "OC-Q2":
                return "Disturbed Transition (Drainage Feature Trajectory)"
            elif q == "OC-Q3":
                return "Disturbed Transition (Rocky Spoil)"
            elif q == "UG-Q3":
                return "Disturbed Transition (Sal Stand Phenology)"
            elif q == "UG-Q5":
                return "Disturbed Transition (Thinned Edge Forest)"
            return "Disturbed Transition"

    df_indices["Descriptive_Classification_Rule"] = df_indices.apply(assign_rule, axis=1)
    
    # Toned-down scientific audit notes (eliminating overinterpreted unmeasured mechanisms)
    def audit_note(row):
        q = row["Quadrat_ID"]
        z = row["Spatial_Zone"]
        if z != "Zone_A_10m":
            return f"Aggregated across {row['Zone_Dimensions']}"
        notes = {
            "OC-Q1": "Planted Eucalyptus stand on legacy mine spoil overburden; exhibits relatively high dry-season spectral retention.",
            "OC-Q2": "Depression adjacent to mine water drainage feature; winter composite = 0.414 (mean of Nov 20: 0.374 and Dec 30: 0.453; Dec 30 acquisition = 0.4533). The trajectory is consistent with a possible moisture-related influence near the drainage feature; however, neither soil moisture nor the abundance of wetland vegetation was quantitatively measured.",
            "OC-Q3": "Stony overburden substrate; sparse Millettia pinnata and woody climbers; lowest overall NDVI across all seasons.",
            "OC-Q4": "Revegetated spoil bench with Holarrhena pubescens thicket (5 verified stems); high persistence ratio (1.006).",
            "OC-Q5": "Open-canopy spoil bench; pronounced late-monsoon spectral peak (Oct 6 = 0.749, Amplitude = 0.285) followed by post-monsoon decline.",
            "UG-Q1": "Shorea robusta stand with Ziziphus mauritiana in understory; high dry-season spectral persistence (winter-to-green-season ratio 1.114). High persistence ratios identify relatively strong dry-season spectral retention, but do not by themselves demonstrate perennial woody cover, rooting depth, or year-round soil stabilization.",
            "UG-Q2": "Terminalia elliptica and Shorea robusta canopy; relatively persistent dry-season spectral greenness (ratio 0.901).",
            "UG-Q3": "Shorea robusta (Sal) stand; low seasonal amplitude (0.075) consistent with spring foliar renewal; winter-to-green-season ratio 0.841.",
            "UG-Q4": "Mixed woodland (Diospyros melanoxylon, Wrightia tinctoria); high multi-season spectral retention (ratio 0.921).",
            "UG-Q5": "Woodland margin; thinned overstory with Ziziphus mauritiana; ratio 0.873."
        }
        return notes.get(q, "")

    df_indices["Field_Ecological_Audit_Notes"] = df_indices.apply(audit_note, axis=1)
    
    # Add explicit dual aliases
    df_indices["NDVI_GreenSeason"] = df_indices["NDVI_Monsoon"]
    df_indices["EVI_GreenSeason"] = df_indices["EVI_Monsoon"]
    df_indices["NDRE_GreenSeason"] = df_indices["NDRE_Monsoon"]
    df_indices["NDWI_GreenSeason"] = df_indices["NDWI_Monsoon"]
    df_indices["NDWI_Moisture_GreenSeason"] = df_indices["NDWI_Moisture_Monsoon"]
    df_indices["RedEdge_Slope_GreenSeason"] = df_indices["RedEdge_Slope_Monsoon"]
    df_indices["SWIR_Ratio_GreenSeason"] = df_indices["SWIR_Ratio_Monsoon"]
    df_indices["Delta_NDVI_GreenSeason_minus_Summer"] = df_indices["Delta_NDVI_Monsoon_minus_Summer"]
    df_indices["Delta_NDVI_Winter_minus_GreenSeason"] = df_indices["Delta_NDVI_Winter_minus_Monsoon"]
    df_indices["Winter_to_GreenSeason_Ratio"] = df_indices["Winter_to_Monsoon_Ratio"]
    df_indices["Summer_to_GreenSeason_Ratio"] = df_indices["Summer_to_Monsoon_Ratio"]
    df_indices["Winter_to_GreenSeason_NDRE_Ratio"] = df_indices["Winter_to_Monsoon_NDRE_Ratio"]
    
    # Add rounded presentation columns
    df_indices["NDVI_Summer_3dec"] = df_indices["NDVI_Summer"].round(3)
    df_indices["NDVI_GreenSeason_3dec"] = df_indices["NDVI_Monsoon"].round(3)
    df_indices["NDVI_Winter_3dec"] = df_indices["NDVI_Winter"].round(3)
    df_indices["Seasonal_NDVI_Amplitude_3dec"] = df_indices["Seasonal_NDVI_Amplitude"].round(3)
    df_indices["Winter_to_GreenSeason_Ratio_3dec"] = df_indices["Winter_to_Monsoon_Ratio"].round(3)
    df_indices["Summer_to_GreenSeason_Ratio_3dec"] = df_indices["Summer_to_Monsoon_Ratio"].round(3)
    
    df_indices.to_csv("data/Definitive_Locked_Site_By_Season_Indices.csv", index=False)
    df_indices.to_csv("tables/Table_Locked_Master_Site_By_Season_Indices.csv", index=False)
    print(" - Saved data/Definitive_Locked_Site_By_Season_Indices.csv and tables/Table_Locked_Master_Site_By_Season_Indices.csv")
    
    # Step 3: Update Table 4 Phenological Syndromes
    zone_a = df_indices[df_indices["Spatial_Zone"] == "Zone_A_10m"].copy()
    table_4_rows = []
    for _, r in zone_a.iterrows():
        table_4_rows.append({
            "Quadrat_ID": r["Quadrat_ID"],
            "Mining_Type": r["Mining_type"],
            "Summer_NDVI_Exact": r["NDVI_Summer"],
            "Summer_NDVI_Reported": f"{r['NDVI_Summer']:.3f}",
            "Monsoon_Retreat_NDVI_Exact": r["NDVI_Monsoon"],
            "Monsoon_Retreat_NDVI_Reported": f"{r['NDVI_Monsoon']:.3f}",
            "Winter_NDVI_Exact": r["NDVI_Winter"],
            "Winter_NDVI_Reported": f"{r['NDVI_Winter']:.3f}",
            "Seasonal_Amplitude_ΔNDVI_Reported": f"{r['Seasonal_NDVI_Amplitude']:.3f}",
            "Winter_to_GreenSeason_Ratio_Reported": f"{r['Winter_to_Monsoon_Ratio']:.3f}",
            "Descriptive_Syndrome_Rule": r["Descriptive_Classification_Rule"],
            "Single_Date_vs_Composite_Audit_Note": r["Field_Ecological_Audit_Notes"]
        })
    df_t4 = pd.DataFrame(table_4_rows)
    df_t4.to_csv("tables/Table_4_Phenological_Syndromes.csv", index=False)
    print(" - Saved tables/Table_4_Phenological_Syndromes.csv")
    
    # Step 4: Generate Table S8 & S9 Syndrome Threshold Sensitivity Analysis
    sens_rows = []
    for s_cut in [0.25, 0.28, 0.30, 0.32, 0.35]:
        for r_cut in [0.85, 0.88, 0.90, 0.92, 0.95]:
            for a_cut in [0.20, 0.25, 0.28]:
                counts = {"Persistent Greenness": 0, "Pronounced Seasonal Flush": 0, "Disturbed Transition": 0}
                changed_quadrats = []
                for _, r in zone_a.iterrows():
                    s = r["NDVI_Summer"]
                    rat = r["Winter_to_Monsoon_Ratio"]
                    amp = r["Seasonal_NDVI_Amplitude"]
                    base_cat = r["Descriptive_Classification_Rule"]
                    if "Persistent Greenness" in base_cat:
                        base_std = "Persistent Greenness"
                    elif "Pronounced Seasonal Flush" in base_cat:
                        base_std = "Pronounced Seasonal Flush"
                    else:
                        base_std = "Disturbed Transition"
                        
                    if s >= s_cut and rat >= r_cut:
                        new_cat = "Persistent Greenness"
                    elif amp >= a_cut and s < s_cut:
                        new_cat = "Pronounced Seasonal Flush"
                    else:
                        new_cat = "Disturbed Transition"
                    counts[new_cat] += 1
                    if new_cat != base_std:
                        changed_quadrats.append(f"{r['Quadrat_ID']} ({base_std} -> {new_cat})")
                        
                sens_rows.append({
                    "Summer_NDVI_Threshold": s_cut,
                    "Persistence_Ratio_Threshold": r_cut,
                    "Seasonal_Amplitude_Threshold": a_cut,
                    "Persistent_Greenness_Count": counts["Persistent Greenness"],
                    "Seasonal_Flush_Count": counts["Pronounced Seasonal Flush"],
                    "Disturbed_Transition_Count": counts["Disturbed Transition"],
                    "Reclassified_Quadrats_Count": len(changed_quadrats),
                    "Reclassified_Quadrats_Detail": "; ".join(changed_quadrats) if changed_quadrats else "None (Exact baseline match)"
                })
    df_sens = pd.DataFrame(sens_rows)
    df_sens.to_csv("tables/Table_S8_Syndrome_Threshold_Sensitivity_Analysis.csv", index=False)
    df_sens.to_csv("tables/Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv", index=False)
    print(f" - Saved tables/Table_S8_Syndrome_Threshold_Sensitivity_Analysis.csv and Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv ({len(df_sens)} threshold configurations)")
    
    # Step 4B: Generate Compact Table S9A Sensitivity Frequency Distribution
    vc = df_sens["Reclassified_Quadrats_Count"].value_counts().sort_index()
    cum_pct = 0.0
    freq_rows = []
    for k, v in vc.items():
        pct = (v / len(df_sens)) * 100.0
        cum_pct += pct
        label = "Exact baseline concordance (10/10 quadrats preserved)" if k == 0 else f"{k} quadrat reclassified" if k == 1 else f"{k} quadrats reclassified"
        freq_rows.append({
            "Reclassification_Count": k,
            "Description": label,
            "Configurations_Count": v,
            "Percentage": f"{pct:.1f}%",
            "Cumulative_Percentage": f"{cum_pct:.1f}%"
        })
    df_freq = pd.DataFrame(freq_rows)
    df_freq.to_csv("tables/Table_S9A_Sensitivity_Frequency_Distribution.csv", index=False)
    print(" - Saved tables/Table_S9A_Sensitivity_Frequency_Distribution.csv")
    
    # Step 5: Write Data Dictionary
    data_dict_content = """# Data Dictionary: Definitive Locked Site-by-Season Vegetation Indices

**File:** `data/Definitive_Locked_Site_By_Season_Indices.csv` (and `tables/Table_Locked_Master_Site_By_Season_Indices.csv`)  
**Geographic Domain:** Korba Coalfield, Chhattisgarh, India (UTM Zone 44N, EPSG:32644)  
**Sensors:** Sentinel-2A / Sentinel-2B Multi-Spectral Instrument (Level-2A Bottom-of-Atmosphere Surface Reflectance)  
**Audit Status:** Fully audited, reproducible, and locked.

---

### Important Clarification on Nomenclature and Temporal Compositing

> [!IMPORTANT]
> **Operational Composite Definition:**  
> In the locked dataset, all fields containing `GreenSeason` (as well as legacy alias fields containing `Monsoon`) refer strictly to the **operational Monsoon & Retreat Green Season composite** formed by taking the arithmetic mean of the cloud-free observations from **June 13, 2024** (monsoon onset) and **October 6, 2024** (monsoon retreat peak).  
> 
> The terms `NDVI_GreenSeason` and `NDVI_Monsoon` are exact mathematical synonyms in this dataset. Users and downstream scripts may utilize either alias interchangeably.

---

### Variable Definitions

| Column Name | Dual / Legacy Alias | Data Type | Units | Description |
| :--- | :--- | :--- | :--- | :--- |
| `Quadrat_ID` | — | String | — | Unique quadrat identifier (UG-Q1 to UG-Q5: Underground; OC-Q1 to OC-Q5: Opencast). |
| `Mining_type` | — | String | — | Mining regime classification: `Underground` vs. `Opencast`. |
| `Spatial_Zone` | — | String | — | Multi-scale extraction support window: `Zone_A_10m` (1 pixel), `Zone_B_50m` (25 pixels), `Zone_C_50m` (81 pixels), `Zone_D_100m` (317 pixels). |
| `NDVI_Summer` | — | Float | Unitless | Pre-monsoon Summer Normalized Difference Vegetation Index (mean of March 25 and May 14, 2024). |
| `NDVI_GreenSeason` | `NDVI_Monsoon` | Float | Unitless | Monsoon & Retreat Green Season NDVI (mean of June 13 and October 6, 2024). |
| `NDVI_Winter` | — | Float | Unitless | Post-monsoon Winter NDVI (mean of November 20 and December 30, 2024). |
| `EVI_Summer` | — | Float | Unitless | Enhanced Vegetation Index for Pre-monsoon Summer. |
| `EVI_GreenSeason` | `EVI_Monsoon` | Float | Unitless | Enhanced Vegetation Index for Monsoon & Retreat Green Season. |
| `EVI_Winter` | — | Float | Unitless | Enhanced Vegetation Index for Post-monsoon Winter. |
| `NDRE_Summer` | — | Float | Unitless | Normalized Difference Red Edge Index for Pre-monsoon Summer. |
| `NDRE_GreenSeason` | `NDRE_Monsoon` | Float | Unitless | Normalized Difference Red Edge Index for Monsoon & Retreat Green Season. |
| `NDRE_Winter` | — | Float | Unitless | Normalized Difference Red Edge Index for Post-monsoon Winter. |
| `NDWI_Summer` | — | Float | Unitless | Normalized Difference Water Index (Gao 1996; Green/NIR) for Summer. |
| `NDWI_GreenSeason` | `NDWI_Monsoon` | Float | Unitless | Normalized Difference Water Index for Monsoon & Retreat Green Season. |
| `NDWI_Winter` | — | Float | Unitless | Normalized Difference Water Index for Post-monsoon Winter. |
| `NDWI_Moisture_Summer` | — | Float | Unitless | Canopy water index (NIR/SWIR1) for Summer. |
| `NDWI_Moisture_GreenSeason` | `NDWI_Moisture_Monsoon` | Float | Unitless | Canopy water index (NIR/SWIR1) for Monsoon & Retreat Green Season. |
| `NDWI_Moisture_Winter` | — | Float | Unitless | Canopy water index (NIR/SWIR1) for Post-monsoon Winter. |
| `RedEdge_Slope_Summer` | — | Float | Unitless | Red Edge Chlorophyll Slope for Summer. |
| `RedEdge_Slope_GreenSeason` | `RedEdge_Slope_Monsoon` | Float | Unitless | Red Edge Chlorophyll Slope for Monsoon & Retreat Green Season. |
| `RedEdge_Slope_Winter` | — | Float | Unitless | Red Edge Chlorophyll Slope for Post-monsoon Winter. |
| `SWIR_Ratio_Summer` | — | Float | Unitless | Shortwave Infrared Ratio (NIR / SWIR2) for Summer. |
| `SWIR_Ratio_GreenSeason` | `SWIR_Ratio_Monsoon` | Float | Unitless | Shortwave Infrared Ratio for Monsoon & Retreat Green Season. |
| `SWIR_Ratio_Winter` | — | Float | Unitless | Shortwave Infrared Ratio for Post-monsoon Winter. |
| `Delta_NDVI_GreenSeason_minus_Summer` | `Delta_NDVI_Monsoon_minus_Summer` | Float | Unitless | Difference in NDVI: `NDVI_GreenSeason - NDVI_Summer`. |
| `Delta_NDVI_Winter_minus_GreenSeason` | `Delta_NDVI_Winter_minus_Monsoon` | Float | Unitless | Difference in NDVI: `NDVI_Winter - NDVI_GreenSeason`. |
| `Delta_NDRE_Monsoon_minus_Summer` | — | Float | Unitless | Difference in NDRE between Green Season and Summer. |
| `Delta_NDRE_Winter_minus_Monsoon` | — | Float | Unitless | Difference in NDRE between Winter and Green Season. |
| `Seasonal_NDVI_Amplitude` | — | Float | Unitless | Seasonal greenness amplitude: `NDVI_GreenSeason - NDVI_Summer`. |
| `Winter_to_GreenSeason_Ratio` | `Winter_to_Monsoon_Ratio` | Float | Unitless | Persistence ratio: `NDVI_Winter / NDVI_GreenSeason`. |
| `Summer_to_GreenSeason_Ratio` | `Summer_to_Monsoon_Ratio` | Float | Unitless | Retention ratio: `NDVI_Summer / NDVI_GreenSeason`. |
| `Winter_to_GreenSeason_NDRE_Ratio` | `Winter_to_Monsoon_NDRE_Ratio` | Float | Unitless | Canopy nitrogen/chlorophyll persistence ratio: `NDRE_Winter / NDRE_GreenSeason`. |
| `Zone_Dimensions` | — | String | — | Spatial window geometry and nominal pixel count. |
| `Descriptive_Classification_Rule` | — | String | — | Heuristic phenological syndrome assigned under baseline thresholds. |
| `Field_Ecological_Audit_Notes` | — | String | — | Field verification notes, single-date context, and scientific cautions. |
| `NDVI_Summer_3dec` | — | Float | Unitless | Summer NDVI rounded to 3 decimal places for tabular reporting. |
| `NDVI_GreenSeason_3dec` | — | Float | Unitless | Green Season NDVI rounded to 3 decimal places for tabular reporting. |
| `NDVI_Winter_3dec` | — | Float | Unitless | Winter NDVI rounded to 3 decimal places for tabular reporting. |
| `Seasonal_NDVI_Amplitude_3dec` | — | Float | Unitless | Seasonal NDVI Amplitude rounded to 3 decimal places for tabular reporting. |
| `Winter_to_GreenSeason_Ratio_3dec` | — | Float | Unitless | Persistence Ratio rounded to 3 decimal places for tabular reporting. |
| `Summer_to_GreenSeason_Ratio_3dec` | — | Float | Unitless | Retention Ratio rounded to 3 decimal places for tabular reporting. |
"""
    with open("data/DATA_DICTIONARY_LOCKED_INDICES.md", "w", encoding="utf-8") as f:
        f.write(data_dict_content)
    print(" - Saved data/DATA_DICTIONARY_LOCKED_INDICES.md")

if __name__ == "__main__":
    reconcile()
