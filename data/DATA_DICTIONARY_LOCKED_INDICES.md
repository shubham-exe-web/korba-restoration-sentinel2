# Data Dictionary: Definitive Locked Site-by-Season Vegetation Indices

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
