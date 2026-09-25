# DATA PROVENANCE RECORD: AUDITED RE-RUN (V2.0-FROZEN)

**Project Title:** Seasonal Herbaceous and Shrub Vegetation Dynamics Across Underground and Opencast Coal Mining Landscapes in Korba, Chhattisgarh, India  
**Investigator / Author Baseline:** Shubham Sharma, School of Studies in Environmental Science, Pt. Ravishankar Shukla University, Raipur, India  
**Parent Publication Baseline:** *Comparative assessment of phytosociology, diversity and carbon dynamics of tree species of opencast and underground coal mining areas of Korba, Chhattisgarh, India* (Sharma, Banerjee & Deb, 2026)  
**Climatological Authority:** India Meteorological Department (IMD) Bilaspur/Korba Climatological Normals & Central Ground Water Board (CGWB, 2012/2020) Groundwater Information Booklet for Korba District  
**Status:** Audited, Reproducible, Defensible Rerun (Frozen Dataset)  
**Date of Record:** 2026-09-25  

---

## 1. Field Evidence Architecture (Three-Tier Evidence Hierarchy)

To eliminate any risk of unverified species attribution, ground evidence is strictly partitioned into three independent tiers:

### Tier 1: Verified In-Plot Woody Stems (Raw Field Sheets)
- **Primary Source:** M.Sc. Dissertation Appendix A (Field Data Sheets, Tables 10 to 19).
- **Sampling Design:** Standard stratified quadrat method ($10 \times 10\text{ m}$, $100\text{ m}^2$, $0.01\text{ ha}$).
- **Inventory File:** [`data/01_Quadrat_Verified_Tree_Woody_Inventory.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/data/01_Quadrat_Verified_Tree_Woody_Inventory.csv).
- **Content:** Exact enumeration of all **58 individual woody stems** ($\text{GBH} \ge 10\text{ cm}$) measured in the field:
  - Tree Stem No, Quadrat ID, Reported Name, Accepted Name (POWO / IPNI standard), Family, Girth at Breast Height ($\text{GBH}$ in m), Tree Height (in m), and Primary Source Sheet.
- **Shrub/Small-Tree Inclusions:** Includes measured individuals of woody taxa that function as shrubs or small trees:
  - *Carissa carandas* (OC-Q2, Stem 2: GBH $2.0\text{ m}$, Height $7.3\text{ m}$)
  - *Holarrhena pubescens* (syn. *H. antidysenterica*) (OC-Q4, 5 stems: GBH $0.4 - 1.1\text{ m}$, Height $4.3 - 8.3\text{ m}$)
  - *Ziziphus mauritiana* (UG-Q1, UG-Q3, UG-Q5: GBH $0.7 - 2.6\text{ m}$, Height $2.8 - 13.2\text{ m}$)
  - *Phanera vahlii* (syn. *Bauhinia vahlii*) (OC-Q3, Stem 4: GBH $0.32\text{ m}$, Height $3.0\text{ m}$; giant woody liana)
- **Confidence:** High (Direct in-plot botanical measurement).

### Tier 2: Verified Field Photographic Plates (Habitat & Ground Truth)
- **Primary Source:** M.Sc. Dissertation Appendix F (Photographic Plates).
- **Inventory File:** [`data/02_Field_Photographic_Evidence_Register.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/data/02_Field_Photographic_Evidence_Register.csv).
- **Content:** All **40 photographic plates** (stored as high-resolution JPEGs in [`field_plates/`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/field_plates)).
- **Documentation:** Structural condition of overstory tree boles, canopy continuity/bifurcation, foliar discoloration, leaf litter accumulation, bare compacted soil patches, and visible understory stratum.
- **Confidence:** High (Direct photographic evidence of ground reality).

### Tier 3: Regional / Off-Plot Floristic Leads Register (Unverified Leads)
- **Inventory File:** [`data/03_Regional_Floristic_Leads_Register.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/data/03_Regional_Floristic_Leads_Register.csv).
- **Status:** **Off-Plot / Regional Associate (Unverified at Quadrat Level)**.
- **Scientific Caveat:** The published study focused on trees ($\text{GBH} \ge 10\text{ cm}$) and did not tally $1 \times 1\text{ m}$ herbaceous or $5 \times 5\text{ m}$ shrub sub-quadrats inside the 10 plots. Common regional taxa (*Parthenium hysterophorus*, *Senna tora*, *Hyptis suaveolens*, *Lantana camara*, *Calotropis procera*, *Acalypha indica*, *Diplazium esculentum*, *Cynodon dactylon*) documented in regional mining surveys and trait collections are retained strictly as **candidate floristic leads** for future botanical audits, and are **NEVER** asserted as confirmed in-plot species tallies.

---

## 2. Sampling Geometry & Climatological Framing

### Sampling Geometry & Bounding Coordinates
- **Center Point Coordinates:** Table 28 / Table 1 verified coordinates in WGS 84 (EPSG:4326) and UTM Zone 44N (EPSG:32644).
- **Published Header Inversion:** Formally reconciled; numerical coordinates confirmed authentic.
- **GPS Device & Jitter Footprint:** Handheld GPS receiver with documented horizontal uncertainty of $\pm 5.0\text{ m}$.
- **Plot Dimensions:** $10.0 \times 10.0\text{ m}$ square ($100.0\text{ m}^2$, $0.01\text{ ha}$) centered on GPS coordinates.
- **Effective Spatial Support Footprint:** $20.0 \times 20.0\text{ m}$ square ($400.0\text{ m}^2$) incorporating GPS uncertainty buffer.
- File: [`data/01_Korba_quadrat_coordinates.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/data/01_Korba_quadrat_coordinates.csv).

### Cited Local Climatological Calendar (IMD & CGWB)
In accordance with IMD Bilaspur/Korba Climatological Normals and the CGWB (2012/2020) Groundwater Information Booklet for Korba District, the climatic cycle is partitioned into three formal agro-ecological seasons:
1. **☀️ Summer (Pre-monsoon Season):** March 1 to May 31. Extreme heat (mean daily maximum $>40^\circ\text{C}$ in May), intense moisture deficit, deciduous leaf-drop, and bare soil thermal exposure.
2. **🌧️ Southwest Monsoon Season:** June 1 to September 30. Regional precipitation exceeds $1250-1300\text{ mm}$ ($>85\%$ annual total). Dramatic vegetative flush, rapid germination of annual forbs and grasses, full tree leaf flush. Early October marks the post-monsoon retreat / maximum residual biomass peak.
3. **❄️ Post-Monsoon / Winter Season:** October 15 to February 28. Mild, dry, cool season (temperatures drop to $10-12^\circ\text{C}$). Senescence of monsoonal annual herbs; retention of perennial woody shrub and semi-evergreen tree canopies.

---

## 3. Sentinel-2 L2A Candidate Manifest & Observation Archive

- **STAC Source:** Element84 AWS Open Data Sentinel-2 L2A COG Archive (`https://earth-search.aws.element84.com/v1/collections/sentinel-2-l2a`).
- **Tile:** `44QPK`.
- **Candidate Search Period:** 2024-01-01 to 2024-12-31.
- **Candidate Archive Manifest:** [`data/04_Sentinel2_Candidate_Archive_Manifest.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/data/04_Sentinel2_Candidate_Archive_Manifest.csv) cataloging all **73 candidate acquisitions** with acquisition timestamps (UTC & IST), tile, platform, cloud cover %, IMD seasonal assignment, processing baseline, and direct S3 COG asset URLs.
- **Audited Observation Dates Selected:**
  - **Summer Date 1:** `2024-03-25` (`S2A_44QPK_20240325_0_L2A`, $0.00\%$ cloud cover)
  - **Summer Date 2:** `2024-05-14` (`S2A_44QPK_20240514_0_L2A`, $0.12\%$ cloud cover, peak late-summer heat)
  - **Monsoon Date 1:** `2024-06-13` (`S2A_44QPK_20240613_0_L2A`, $4.28\%$ cloud cover, onset green-up)
  - **Monsoon Date 2:** `2024-10-06` (`S2B_44QPK_20241006_0_L2A`, $4.11\%$ cloud cover, post-rain peak biomass)
  - **Winter Date 1:** `2024-11-20` (`S2A_44QPK_20241120_0_L2A`, $0.00\%$ cloud cover, early post-monsoon)
  - **Winter Date 2:** `2024-12-30` (`S2A_44QPK_20241230_0_L2A`, $0.49\%$ cloud cover, mid/late winter)

---

## 4. Geospatial Processing, Strict QC & Checksums

### Documented Grid Definition
- **Native 10 m Grid:** UTM Zone 44N (EPSG:32644), origin $(600000.0, 2500020.0)$, pixel size $10.0\text{ m} \times -10.0\text{ m}$.
- **Native 20 m Grid:** UTM Zone 44N (EPSG:32644), origin $(600000.0, 2500020.0)$, pixel size $20.0\text{ m} \times -20.0\text{ m}$.
- **Spatial Alignment:** 20 m bands resampled to the 10 m grid using bilinear interpolation; Scene Classification Layer (SCL) resampled using nearest-neighbor interpolation.

### Strict Pixelwise Terrestrial Masking & Quality Control
- **Per-Pixel Computation:** All spectral indices (NDVI, EVI, NDRE, NDWI, NDWI-Moisture, SWIR Ratio) were computed **per pixel** on the 2D array prior to window extraction.
- **Valid Terrestrial Filter:** Retains strictly:
  - $\text{SCL} == 4$ (Vegetation)
  - $\text{SCL} == 5$ (Non-vegetated bare soil)
- **Explicitly Excluded Mask:**
  - $\text{SCL} == 6$ (Water — excluded for vegetation analysis)
  - $\text{SCL} == 7$ (Unclassified — excluded for vegetation analysis)
  - $\text{SCL} \in \{0, 1, 2, 3, 8, 9, 10, 11\}$ (No-data, defective, dark area, shadows, clouds, cirrus, snow)
- **Predeclared Retention Rule:** A quadrat observation was retained if and only if the valid terrestrial pixel fraction in the 50 m buffer was $\ge 80.0\%$.
- **Validation Result:** In the audited 6-date dataset, all 10 quadrats achieved $92.0\% - 100.0\%$ valid terrestrial fractions across all dates ($0$ exclusions; $100\%$ retention; see [`tables/Table_3_Seasonal_Sentinel2_QC.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/tables/Table_3_Seasonal_Sentinel2_QC.csv)).

### Intermediate GeoTIFF Rasters & SHA-256 Checksums
Intermediate cropped multi-band GeoTIFFs (11 bands per date) were written to [`rasters/`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/rasters) and verified:
- `rasters/Korba_20240325_stack.tif` (SHA-256: `205a6053a986...`)
- `rasters/Korba_20240514_stack.tif` (SHA-256: `eb864da3d200...`)
- `rasters/Korba_20240613_stack.tif` (SHA-256: `3014a3c7bd24...`)
- `rasters/Korba_20241006_stack.tif` (SHA-256: `0516c537e57b...`)
- `rasters/Korba_20241120_stack.tif` (SHA-256: `936421aabfb6...`)
- `rasters/Korba_20241230_stack.tif` (SHA-256: `b8b394ae444d...`)
- Full manifest: [`rasters/CHECKSUMS.sha256`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/rasters/CHECKSUMS.sha256).

---

## 5. Statistical Methods & Uncertainty Reporting

- **Non-Parametric Group Contrasts:** Evaluated between Underground ($n=5$) and Opencast ($n=5$) locations in Zone A (10 m quadrat scale).
  - Exact two-sided Mann-Whitney U tests (exact permutation method for small samples).
  - Non-parametric effect size: Cliff's Delta ($\delta$).
  - Parametric small-sample corrected effect size: Hedges' $g$ with 95% bootstrap confidence intervals ($2,000$ iterations).
  - File: [`tables/Table_5_Statistical_Contrasts_UG_vs_OC.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/tables/Table_5_Statistical_Contrasts_UG_vs_OC.csv).
- **Conservative Inference Mandate:**
  - All statistical tests report explicit sample size caveats: *"Descriptive / Exploratory ($n=5$ per group); causal landscape extrapolation avoided."*
- **Reconciled Principal Component Analysis (PCA):**
  - Features standardized to zero mean, unit variance.
  - PC1: $60.49\%$ variance explained (Woody Biomass & Canopy Persistence Axis).
  - PC2: $21.13\%$ variance explained (Herbaceous Green-Up Amplitude Axis).
  - Cumulative variance: $81.62\%$ across the first two components.
  - File: [`data/PCA_scores_loadings.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/data/PCA_scores_loadings.csv).

---

## 6. Software & Pipeline Execution Environment

- **OS:** macOS (Darwin 24.3.0, Apple Silicon ARM64).
- **Python:** 3.12 (`/Users/shubhamsharma/Earth-One/.venv/bin/python`).
- **Core Libraries:** `rasterio` 1.5.1, `pyproj` 3.7.2, `scipy` 1.18.0, `scikit-learn` 1.9.0, `pandas` 2.3.3, `numpy` 2.5.2, `matplotlib` 3.11.2, `python-docx` 1.2.0, `requests` 2.34.2.
- **Pipeline Scripts:**
  1. `scripts/01_build_master_locations.py`
  2. `scripts/02_build_inventories.py`
  3. `scripts/03_query_imagery_manifest.py`
  4. `scripts/04_geospatial_processing_qc.py`
  5. `scripts/05_extract_seasonal_features.py`
  6. `scripts/06_statistical_analysis.py`
  7. `scripts/07_generate_publication_figures.py`
  8. `scripts/enact_ethical_data_governance.py`

---

## 7. Scientific Integrity & Ethical Data Governance Mandate (v3.0)

To ensure peer-review defensibility and prevent any misrepresentation of prospective research designs as completed empirical measurements, all prospective additions are governed by strict usage protocols:

| Generated Item / Dossier Artifact | Ethical Use Now | Strict Prohibition (Do Not Do) | Governed Destination |
| :--- | :--- | :--- | :--- |
| **Reference-site table (Table 9)** | Convert to a field-sampling sheet for future reference quadrats | Call it a sampled reference forest dataset | `templates/Template_Reference_Forest_Field_Sampling_Sheet.csv` |
| **Expanded N=75 design (Table 18)** | Use as a proposed hierarchical sampling design | Report it as completed replication | `templates/Template_Proposed_N75_Hierarchical_Sampling_Design.csv` |
| **Restoration-history metadata (Table 10)** | Use as a blank site-interview and mine-record template | Invent mine age, topsoil depth, seed-source distance, or restoration history | `templates/Template_Mine_Restoration_History_Interview_Audit.csv` |
| **Soil physical/biological tables (Tables 11 & 12)** | Use as laboratory and field-data templates (ASTM / Page protocols) | Report infiltration, MBC, enzymes, AMF, earthworms, etc. without measurement | `templates/Template_Soil_Physical_Hydrological_Lab_Datasheet.csv` |
| **Functional trait matrix (Table 14)** | Use if traits are verified against authoritative floras/databases and cited | Treat unverified classifications as measured results | Verified with citations: Haines (1925), Flora MP/CG, TRY Database |
| **Sentinel-2 table (Table 15)** | Replace with reproducible real extraction from specified imagery and dates | State values were "reconstructed" without a documented workflow | Audited real extractions preserved in `data/07_Korba_vegetation_indices.csv` |
| **Butterfly / pollinator table (Table 16)** | Use as a future sampling template (Pollard walk protocols) | Report transect observations not performed | `templates/Template_Pollinator_Butterfly_Pollard_Transect_Datasheet.csv` |
| **LMM model table (Table 19)** | Run models only on actual observations and appropriate design | Report model coefficients or $R^2$ from synthetic data | Retained as prospective power analysis & model specification |
| **Restoration figure (Figure 5)** | Label as a conceptual framework only, or regenerate entirely from real data | Publish it as empirical evidence | Formally labeled: *"Conceptual Restoration Trajectory and Methodological Framework"* |

