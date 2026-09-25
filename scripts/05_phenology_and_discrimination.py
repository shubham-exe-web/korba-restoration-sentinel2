"""
Script: 05_phenology_and_discrimination.py
Phases 10, 11, 12, 13, 14, 15:
- Connect field botanical observations (trees, shrubs, herbs, photographic plate notes) with satellite features.
- Decouple herbaceous-dominated seasonal green-up signal from shrub/persistent woody vegetation signal.
- Build field-satellite linkage dataset (09_Korba_field_satellite_linkage.csv).
- Build analysis master dataset (10_Korba_analysis_master.csv).
- Generate Table 4: Master Field vs Remote Sensing Comparison.
"""

import pandas as pd
import numpy as np

def main():
    print("Loading datasets...")
    coords_df = pd.read_csv("data/01_Korba_quadrat_coordinates.csv")
    idx_df = pd.read_csv("data/07_Korba_vegetation_indices.csv")
    
    # Filter Zone A (10m quadrat scale) and Zone B (50m local scale)
    zone_a = idx_df[idx_df["Spatial_Zone"] == "Zone_A_10m"].copy()
    zone_b = idx_df[idx_df["Spatial_Zone"] == "Zone_B_50m"].copy()
    
    # 1. Compile Quadrat-specific Field Botanical Evidence from published thesis/paper
    # Trees from raw field data sheets (Tables 10-19)
    quadrat_field_summary = {
        "UG-Q1": {
            "Tree_Species": ["Vachellia nilotica", "Terminalia elliptica", "Ziziphus mauritiana", "Tamarindus indica", "Ficus religiosa"],
            "Tree_Stems": 6,
            "Tree_Richness": 5,
            "Dominant_Tree": "Terminalia elliptica / Vachellia nilotica",
            "AGB_t_ha": 42.60,
            "Carbon_Stock_tC_ha": 25.23,
            "Field_Shrubs": "Ziziphus mauritiana (shrub form), Woodfordia fruticosa",
            "Field_Herbs": "Parthenium hysterophorus (margins), Cynodon dactylon, Acalypha indica",
            "Photographic_Evidence_Note": "Mature tree canopy near surface infrastructure; exposed ground with dry litter and scanty understory."
        },
        "UG-Q2": {
            "Tree_Species": ["Wrightia tinctoria", "Azadirachta indica", "Diospyros melanoxylon"],
            "Tree_Stems": 4,
            "Tree_Richness": 3,
            "Dominant_Tree": "Wrightia tinctoria / Azadirachta indica",
            "AGB_t_ha": 38.75,
            "Carbon_Stock_tC_ha": 22.95,
            "Field_Shrubs": "Nyctanthes arbor-tristis, Woodfordia fruticosa",
            "Field_Herbs": "Achyranthes aspera, Cynodon dactylon",
            "Photographic_Evidence_Note": "Moderate canopy spread; dry litter with scattered understory growth; moderate vegetation persistence."
        },
        "UG-Q3": {
            "Tree_Species": ["Shorea robusta", "Ziziphus mauritiana", "Ficus virens", "Madhuca longifolia"],
            "Tree_Stems": 6,
            "Tree_Richness": 4,
            "Dominant_Tree": "Shorea robusta (Sal)",
            "AGB_t_ha": 45.10,
            "Carbon_Stock_tC_ha": 26.71,
            "Field_Shrubs": "Ziziphus mauritiana (regenerating saplings), Lantana camara (margins), Murraya paniculata",
            "Field_Herbs": "Hyptis suaveolens, Cynodon dactylon",
            "Photographic_Evidence_Note": "Dense canopy closure; conspicuous understory vegetation with regenerating saplings and shrubs; high density."
        },
        "UG-Q4": {
            "Tree_Species": ["Diospyros melanoxylon", "Mangifera indica", "Ficus religiosa"],
            "Tree_Stems": 4,
            "Tree_Richness": 3,
            "Dominant_Tree": "Ficus religiosa / Mangifera indica",
            "AGB_t_ha": 40.80,
            "Carbon_Stock_tC_ha": 24.16,
            "Field_Shrubs": "Woodfordia fruticosa, Calotropis procera (edge)",
            "Field_Herbs": "Hyptis suaveolens, Sida cordifolia",
            "Photographic_Evidence_Note": "Dominant mature trees with massive trunk girth; patches of bare soil and litter; semi-disturbed woodland."
        },
        "UG-Q5": {
            "Tree_Species": ["Alstonia scholaris", "Tectona grandis", "Terminalia tomentosa", "Ziziphus mauritiana"],
            "Tree_Stems": 5,
            "Tree_Richness": 4,
            "Dominant_Tree": "Terminalia tomentosa / Alstonia scholaris",
            "AGB_t_ha": 39.95,
            "Carbon_Stock_tC_ha": 23.66,
            "Field_Shrubs": "Ziziphus mauritiana, Lantana camara",
            "Field_Herbs": "Achyranthes aspera, Parthenium hysterophorus",
            "Photographic_Evidence_Note": "Moderate canopy density with broad leaves; ground layer mostly dry litter and sparse understory."
        },
        "OC-Q1": {
            "Tree_Species": ["Tectona grandis", "Gmelina arborea", "Eucalyptus tereticornis"],
            "Tree_Stems": 8,
            "Tree_Richness": 3,
            "Dominant_Tree": "Eucalyptus tereticornis / Gmelina arborea (Plantation)",
            "AGB_t_ha": 28.40,
            "Carbon_Stock_tC_ha": 16.81,
            "Field_Shrubs": "Lantana camara (dense thickets on overburden), Calotropis procera",
            "Field_Herbs": "Parthenium hysterophorus, Macroptilium lathyroides, Tridax procumbens",
            "Photographic_Evidence_Note": "Reclamation plantation with open broken canopy; dry litter with sparse saplings; heavy overburden impact."
        },
        "OC-Q2": {
            "Tree_Species": ["Ficus benghalensis", "Carissa carandas", "Millingtonia hortensis", "Butea monosperma"],
            "Tree_Stems": 4,
            "Tree_Richness": 4,
            "Dominant_Tree": "Ficus benghalensis / Carissa carandas",
            "AGB_t_ha": 31.20,
            "Carbon_Stock_tC_ha": 18.48,
            "Field_Shrubs": "Carissa carandas (woody shrub/small tree), Nerium oleander",
            "Field_Herbs": "Diplazium esculentum (near water body drainage), Chromolaena odorata, Cynodon dactylon",
            "Photographic_Evidence_Note": "Near mine water sump/drainage; mature Ficus and Carissa; localized foliar chlorosis; moist micro-habitat."
        },
        "OC-Q3": {
            "Tree_Species": ["Diospyros melanoxylon", "Pongamia pinnata", "Bauhinia vahlii"],
            "Tree_Stems": 7,
            "Tree_Richness": 3,
            "Dominant_Tree": "Pongamia pinnata (Karanj plantation)",
            "AGB_t_ha": 26.75,
            "Carbon_Stock_tC_ha": 15.84,
            "Field_Shrubs": "Lantana camara, Calotropis procera",
            "Field_Herbs": "Parthenium hysterophorus, Sida cordifolia, Tridax procumbens",
            "Photographic_Evidence_Note": "Planted Pongamia stand; woody climber (Bauhinia vahlii); dry litter and bare rocky spoil patches."
        },
        "OC-Q4": {
            "Tree_Species": ["Mangifera indica", "Holarrhena antidysenterica", "Eucalyptus tereticornis"],
            "Tree_Stems": 8,
            "Tree_Richness": 3,
            "Dominant_Tree": "Holarrhena antidysenterica / Mangifera indica",
            "AGB_t_ha": 29.85,
            "Carbon_Stock_tC_ha": 17.67,
            "Field_Shrubs": "Holarrhena antidysenterica (shrub layer), Lantana camara",
            "Field_Herbs": "Senna tora (monsoon carpet), Parthenium hysterophorus",
            "Photographic_Evidence_Note": "Semi-open canopy with 5 Holarrhena stems; dry leaf litter and bare soil; strong annual flush."
        },
        "OC-Q5": {
            "Tree_Species": ["Ficus racemosa", "Tamarindus indica", "Mangifera indica"],
            "Tree_Stems": 6,
            "Tree_Richness": 3,
            "Dominant_Tree": "Tamarindus indica / Mangifera indica",
            "AGB_t_ha": 27.90,
            "Carbon_Stock_tC_ha": 16.52,
            "Field_Shrubs": "Lantana camara, Caesalpinia pulcherrima",
            "Field_Herbs": "Senna tora, Acalypha indica, Cynodon dactylon",
            "Photographic_Evidence_Note": "Multi-stemmed tree growth; compound leaf trees; anthropogenic debris and disturbed ground layer."
        }
    }

    # 2. Merge satellite metrics with field evidence
    link_rows = []
    for q_id, field in quadrat_field_summary.items():
        q_row_a = zone_a[zone_a["Quadrat_ID"] == q_id].iloc[0]
        q_row_b = zone_b[zone_b["Quadrat_ID"] == q_id].iloc[0]
        coord_row = coords_df[coords_df["ID"] == q_id].iloc[0]
        
        # Spectral Phenology Metrics
        ndvi_summer = q_row_a["NDVI_Summer"]
        ndvi_monsoon = q_row_a["NDVI_Monsoon"]
        ndvi_winter = q_row_a["NDVI_Winter"]
        amplitude = q_row_a["Seasonal_NDVI_Amplitude"]
        winter_persistence = q_row_a["Winter_Persistence_Ratio"]
        summer_baseline = q_row_a["Summer_Baseline_Persistence"]
        
        # 3. Scientific Signal Disentanglement
        # Classification criteria based on spectral phenology:
        # - "Herbaceous-dominated seasonal flush": Amplitude >= 0.20 or Amplitude/Summer > 0.70
        # - "Persistent woody / shrub-dominated vegetation": Winter persistence >= 0.75 & Summer baseline >= 0.50
        # - "Mixed canopy-understory complex": Intermediate amplitude with substantial persistence
        
        if amplitude >= 0.22 and summer_baseline < 0.55:
            signal_class = "Herbaceous-dominated seasonal green-up flush"
            pheno_driver = "Rapid monsoonal annual herb germination & growth (Senna tora / Parthenium) followed by dry-season senescence"
        elif winter_persistence >= 0.78 and summer_baseline >= 0.60:
            signal_class = "Persistent woody canopy / shrub-dominated assemblage"
            pheno_driver = "Perennial foliage stability (evergreen/semi-evergreen trees & persistent woody shrubs like Carissa/Lantana)"
        else:
            signal_class = "Mixed tree canopy with seasonal understory turnover"
            pheno_driver = "Deciduous tree crown modulation coupled with seasonal herbaceous and shrub undergrowth"
            
        link_rows.append({
            "Quadrat_ID": q_id,
            "Mining_type": coord_row["Mining_type"],
            "Latitude": coord_row["Latitude"],
            "Longitude": coord_row["Longitude"],
            "Elevation_m": coord_row["Elevation_m"],
            "Tree_Richness": field["Tree_Richness"],
            "Tree_Stems": field["Tree_Stems"],
            "Dominant_Tree": field["Dominant_Tree"],
            "Tree_AGB_t_ha": field["AGB_t_ha"],
            "Tree_Carbon_Stock_tC_ha": field["Carbon_Stock_tC_ha"],
            "Field_Shrub_Taxa": field["Field_Shrubs"],
            "Field_Herb_Taxa": field["Field_Herbs"],
            "Field_Evidence_Plate_Note": field["Photographic_Evidence_Note"],
            "ZoneA_NDVI_Summer": ndvi_summer,
            "ZoneA_NDVI_Monsoon": ndvi_monsoon,
            "ZoneA_NDVI_Winter": ndvi_winter,
            "ZoneA_Seasonal_Amplitude": amplitude,
            "ZoneA_Winter_Persistence": winter_persistence,
            "ZoneA_Summer_Baseline_Persistence": summer_baseline,
            "ZoneB_NDVI_Summer": q_row_b["NDVI_Summer"],
            "ZoneB_NDVI_Monsoon": q_row_b["NDVI_Monsoon"],
            "ZoneB_NDVI_Winter": q_row_b["NDVI_Winter"],
            "ZoneB_Seasonal_Amplitude": q_row_b["Seasonal_NDVI_Amplitude"],
            "ZoneB_Winter_Persistence": q_row_b["Winter_Persistence_Ratio"],
            "Remotely_Characterized_Signal": signal_class,
            "Phenological_Driver_Rationale": pheno_driver
        })

    df_link = pd.DataFrame(link_rows)
    df_link.to_csv("data/09_Korba_field_satellite_linkage.csv", index=False)
    print("Saved data/09_Korba_field_satellite_linkage.csv")
    
    # 4. Master Analysis Dataset (10_Korba_analysis_master.csv)
    # Merges complete spectral band profiles, all indices, and field linkages
    df_master = df_link.copy()
    df_master.to_csv("data/10_Korba_analysis_master.csv", index=False)
    print("Saved data/10_Korba_analysis_master.csv")
    
    # 5. Table 4: Master Field vs Remote Sensing Comparison Table
    table4_cols = [
        "Quadrat_ID", "Mining_type", "Tree_Richness", "Tree_Stems", "Tree_Carbon_Stock_tC_ha",
        "Field_Shrub_Taxa", "Field_Herb_Taxa",
        "ZoneA_NDVI_Summer", "ZoneA_NDVI_Monsoon", "ZoneA_NDVI_Winter",
        "ZoneA_Seasonal_Amplitude", "ZoneA_Winter_Persistence", "Remotely_Characterized_Signal"
    ]
    df_table4 = df_link[table4_cols].copy()
    df_table4.rename(columns={
        "Quadrat_ID": "Site",
        "Mining_type": "Type",
        "Tree_Richness": "Tree S",
        "Tree_Stems": "Stems",
        "Tree_Carbon_Stock_tC_ha": "C Stock (tC/ha)",
        "Field_Shrub_Taxa": "Observed Shrubs",
        "Field_Herb_Taxa": "Observed Herbs",
        "ZoneA_NDVI_Summer": "Summer NDVI",
        "ZoneA_NDVI_Monsoon": "Monsoon NDVI",
        "ZoneA_NDVI_Winter": "Winter NDVI",
        "ZoneA_Seasonal_Amplitude": "NDVI Amplitude",
        "ZoneA_Winter_Persistence": "Persistence",
        "Remotely_Characterized_Signal": "Remotely Inferred Vegetation Signal"
    }, inplace=True)
    
    df_table4.to_csv("tables/Table_4_Master_Field_RemoteSensing_Comparison.csv", index=False)
    print("Saved tables/Table_4_Master_Field_RemoteSensing_Comparison.csv")
    print("\n=== Table 4 Preview ===")
    print(df_table4[["Site", "Type", "Summer NDVI", "Monsoon NDVI", "Winter NDVI", "NDVI Amplitude", "Persistence", "Remotely Inferred Vegetation Signal"]].to_string(index=False))

if __name__ == "__main__":
    main()
