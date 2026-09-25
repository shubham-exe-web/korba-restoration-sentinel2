# Botanical Field Audit Protocol: Ground-Truthing Seasonal Sentinel-2 Signals in Mining Landscapes

**Author:** Shubham Sharma  
**Context:** Standard Operating Procedure (SOP) for Nested Field Validation in the Korba Coalfield  
**Companion to:** *Ecological Indicators* Manuscript Version 2.4 (September 2026)  

---

## 1. Rationale and Empirical Scope

Medium-resolution optical satellites such as Sentinel-2 (10–20 m pixel resolution) provide integrated, composite surface reflectance signals that combine overstory tree canopy, subcanopy woody shrubs, ground-layer herbs and grasses, leaf litter, and exposed rock/soil. While multi-temporal spectral trajectories capture broad phenological dynamics, satellite observations **cannot deconvolve vertical vegetation strata** or identify species without contemporaneous ground measurements.

This protocol establishes the standardized field audit procedure required to ground-truth satellite greenness trajectories and validate future multi-tier restoration monitoring designs.

---

## 2. Multi-Tier Nested Sampling Architecture

To eliminate ambiguity between tree canopy cover, woody shrubs, and monsoonal herbaceous flushes, each sampling quadrat must follow a nested spatial design:

```
+-------------------------------------------------------------+
| 20 m x 20 m (400 m²): GPS Positional Uncertainty Footprint  |
|                                                             |
|     +-------------------------------------------------+     |
|     | 10 m x 10 m (100 m²): Overstory Tree Quadrat    |     |
|     | (All stems ≥ 10 cm GBH / ≥ 3.18 cm DBH)         |     |
|     |                                                 |     |
|     |    +-----------------------+                    |     |
|     |    | 5 m x 5 m (25 m²):    |                    |     |
|     |    | Shrub / Sapling Subplot                    |     |
|     |    | (Woody stems < 10 cm  |                    |     |
|     |    |  GBH, height ≥ 0.5 m) |                    |     |
|     |    |                       |                    |     |
|     |    |   [1x1 m]   [1x1 m]   |                    |     |
|     |    |   Herb 1    Herb 2    |                    |     |
|     |    +-----------------------+                    |     |
|     |                                                 |     |
|     |        [1x1 m]   [1x1 m]   [1x1 m]              |     |
|     |        Herb 3    Herb 4    Herb 5               |     |
|     +-------------------------------------------------+     |
|                                                             |
+-------------------------------------------------------------+
```

1. **Overstory Tree Quadrat (10 × 10 m = 100 m² = 0.01 ha):**
   - Directly matches the 10 m Sentinel-2 pixel support (Zone A).
   - Enumerate all woody stems with girth at breast height (GBH) ≥ 10 cm (DBH ≥ 3.18 cm).
   - Record: Species (POWO standardized), GBH (cm), total height (m), crown diameter (m), status (alive/dead/coppicing).

2. **Shrub and Sapling Subplot (5 × 5 m = 25 m² = 0.0025 ha):**
   - Established in the southwest quadrant of each 10 × 10 m plot.
   - Enumerate all woody shrubs and tree saplings with height ≥ 0.5 m and GBH < 10 cm.
   - Record: Species, stem count, basal diameter (mm), mean height (cm), visual percent foliar cover.

3. **Nested Herbaceous Micro-Quadrats (Five 1 × 1 m = 1 m² plots):**
   - Positioned at the four corners and plot center.
   - Surveyed three times per year: Pre-monsoon Summer (April/May), Southwest Monsoon Peak (August/September), and Post-monsoon Winter (November/December).
   - Record: Complete floristic inventory, species-specific visual percent cover (0–100%), mean vegetative height (cm), litter depth (cm), and exposed bare ground/rock cover (%).

4. **Canopy Fractional Cover and LAI:**
   - Digital hemispherical photography (DHP) taken at 1.2 m above ground at five standardized points per quadrat.
   - Processed via standardized gap fraction analysis to measure canopy openness and effective Leaf Area Index (LAI_eff).

---

## 3. Contemporaneous Soil and Hydrological Profiling

1. **Topsoil Sampling (0–15 cm depth):**
   - Composite of 5 soil cores per quadrat.
   - Measure: pH (1:2.5 H2O), soil organic carbon (SOC, Walkley-Black), bulk density (BD, core method), available N-P-K, electrical conductivity (EC), and heavy metals (Fe, Mn, Cu, Zn).

2. **Micro-Hydrological and Drainage Auditing:**
   - Measure distance to active mine sump dewatering ditches or drainage canals.
   - Record surface ponding status, seasonal water table depth, and soil moisture via TDR probe at 0–10 cm and 10–30 cm depths during each seasonal survey.

---

## 4. Quality Assurance and Taxonomic Standardization

1. **Botanical Verification:**
   - Collect voucher specimens for all taxa uncertain in the field.
   - Standardize all names against Plants of the World Online (POWO, Royal Botanic Gardens, Kew) and World Flora Online (WFO).
2. **Georeferencing Accuracy:**
   - Record plot center and four corners using a multi-band differential GNSS receiver with real-time kinematic (RTK) positioning or post-processed differential correction (target horizontal accuracy < 0.5 m).
3. **Data Provenance:**
   - Maintain immutable, timestamped CSV records with full SHA-256 checksums.

---
