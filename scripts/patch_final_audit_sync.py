"""
Patch script for scripts/08_generate_manuscript_and_supplemental.py
Implementing all final audit synchronizations requested by the Chief Editor:
1. Reconcile threshold-sensitivity distribution: 0 changes: 24 (32.0%), 1 change: 21 (28.0%),
   2 changes: 9 (12.0%), 3 changes: 15 (20.0%), 4 changes: 6 (8.0%). Total: 75 configurations.
2. Confirm amplitude cutoff grid is strictly (0.20, 0.25, 0.28) and avoid "stable across all 75".
   Use: "In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28)
   did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications."
3. Add Python classify_quadrat function to Methods 3.5 and Supplement.
4. Replace causal "canopy dilution" wording with:
   "The scale-dependent decline is consistent with inclusion of lower-greenness surfaces
   surrounding the nominal pixel, including the documented colliery matrix; the relative
   contribution of each surface type was not independently quantified."
   And replace "remarkable spatial invariance" with:
   "The sampled opencast quadrats showed smaller changes in mean NDVI across the tested windows."
5. Remove "permits" from seasonal flush:
   "The open, low-greenness spoil setting exhibited a pronounced seasonal amplitude that is
   compatible with seasonal ground-cover development, but the ground-layer contribution was not directly measured."
6. Replace "annual maximum" with "the maximum greenness among the six audited observations",
   and "annual minimum" with "minimum greenness among the audited observations".
7. Change "60/60 pixels" to "60 of 60 nominal Zone A pixel observations".
8. Clarify species richness: 58 stems across 24 accepted species (25 reported field morphotaxa),
   using accepted-name field as the sole basis for richness (Terminalia tomentosa synonymized with T. elliptica).
9. Statistical softening: "The wide uncertainty and small clustered sample make the non-significant
   result uninformative about ecological equivalence."
10. Consistent spatial support: "enclosing 81 raster pixels with a rasterized support area of 8,100 m²"
"""

import os
import re

SCRIPT_PATH = "scripts/08_generate_manuscript_and_supplemental.py"

with open(SCRIPT_PATH, "r", encoding="utf-8") as f:
    text = f.read()

# 1. Highlights updates (docx & md)
old_h5 = '        "Demonstrates scale-dependent canopy dilution in remnant forest quadrats (summer NDVI declines from 0.381 in Zone A [10 m] to 0.271 in Zone D [100 m] as extraction encompasses colliery openings).",'
new_h5 = '        "Sampled underground forest quadrats exhibit a scale-dependent decline in mean summer NDVI (0.381 at 10 m to 0.271 at 100 m) consistent with matrix inclusion, whereas opencast quadrats showed smaller changes (0.303 to 0.287).",'
text = text.replace(old_h5, new_h5)

old_h5_md = "- **Spatial Support Sensitivity:** Remnant underground forest quadrats exhibit scale-dependent canopy dilution (mean summer NDVI declines from 0.381 in Zone A to 0.271 in Zone D as extraction encompasses colliery openings), whereas opencast plantations remain spatially invariant (0.303 to 0.287).\n"
new_h5_md = "- **Spatial Support Sensitivity:** Sampled underground forest quadrats exhibit a scale-dependent decline in mean summer NDVI (0.381 in Zone A to 0.271 in Zone D) consistent with matrix inclusion, whereas opencast quadrats showed smaller changes (0.303 to 0.287).\n"
text = text.replace(old_h5_md, new_h5_md)

# 2. Abstract updates (docx & md)
# 2a. Species richness in abstract
text = text.replace(
    'Tier 1 comprises 58 verified in-plot woody stems (GBH ≥ 10 cm) across 25 species with Plants of the World Online (POWO) standardized taxonomy;',
    'Tier 1 comprises 58 verified in-plot woody stems (GBH ≥ 10 cm) across 24 accepted species (25 reported field morphotaxa) with Plants of the World Online (POWO) standardized taxonomy;'
)

# 2b. 60/60 pixel observations in abstract
text = text.replace(
    'across all 60 site-date cases (60/60 pixels)',
    'across all 60 site-date cases (60 of 60 nominal Zone A pixel observations)'
)

# 2c. Sensitivity in abstract
old_abs_sens = (
    '        "Threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations) "\n'
    '        "indicated central classification stability (preserving 100% of site assignments in 32.0% of configurations and ≤ 1 reclassification in 60.0%), with "\n'
    '        "reclassifications driven primarily by persistence and summer-NDVI cutoffs in this dataset (amplitude variation from 0.20 to 0.28 caused zero reclassifications across all 75 configurations; Supplementary Table S9A). "\n'
    '        "Multi-scale spatial evaluation revealed canopy dilution in remnant underground forest quadrats (mean summer NDVI declining from 0.381 in "\n'
    '        "Zone A to 0.271 in Zone D) as extraction captures colliery infrastructure and clearings, whereas opencast spoil plantations remained "\n'
    '        "spatially stable (0.303 to 0.287). Principal Component Analysis accounted for 81.62% of variance across two primary axes, demonstrating "\n'
    '        "close agreement (score correlation r > 0.94) with an 8-feature sensitivity model excluding derived metrics. Non-parametric contrasts between "\n'
    '        "the sampled underground and opencast quadrats (n=5 per group) using SciPy\'s exact two-sided Mann-Whitney U procedure yielded non-significant "\n'
    '        "differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, exact p = 0.2222, Cliff\'s delta = +0.520, Hedges\' g = 0.647, approximate 95% analytical Student\'s t CI [-0.850, 2.143]), "\n'
    '        "demonstrating that non-significance reflects low statistical power in a pilot sample rather than ecological equivalence. We provide a complete machine-readable "\n'
    '        "provenance audit and outline requirements for future hierarchical, replicated landscape restoration designs."'
)

new_abs_sens = (
    '        "Threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations) "\n'
    '        "yielded 0 reclassifications in 24 configurations (32.0%), 1 in 21 (28.0%, cumulative ≤ 1: 45 [60.0%]), 2 in 9 (12.0%), 3 in 15 (20.0%), and 4 in 6 (8.0%; Supplementary Table S9A). "\n'
    '        "In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28) did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications. "\n'
    '        "Multi-scale spatial evaluation revealed a scale-dependent decline in mean summer NDVI in sampled underground forest quadrats (0.381 in Zone A to 0.271 in Zone D) "\n'
    '        "consistent with matrix inclusion, whereas sampled opencast quadrats showed smaller changes (0.303 to 0.287). Principal Component Analysis accounted "\n'
    '        "for 81.62% of variance across two primary axes, demonstrating close agreement (score correlation r > 0.94) with an 8-feature sensitivity model "\n'
    '        "excluding derived metrics. Non-parametric contrasts between the sampled underground and opencast quadrats (n=5 per group) using SciPy\'s exact two-sided "\n'
    '        "Mann-Whitney U procedure yielded non-significant differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, exact p = 0.2222, Cliff\'s delta = +0.520, "\n'
    '        "Hedges\' g = 0.647, approximate 95% analytical Student\'s t CI [-0.850, 2.143]); the wide uncertainty and small clustered sample make the non-significant result "\n'
    '        "uninformative about ecological equivalence. We provide a complete machine-readable provenance audit and outline requirements for future hierarchical, replicated landscape restoration designs."'
)

text = text.replace(old_abs_sens, new_abs_sens)

# Same for markdown abstract
old_abs_sens_md = (
    '        "Threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations) "\n'
    '        "indicated central classification stability (preserving 100% of site assignments in 32.0% of configurations and ≤ 1 reclassification in 60.0%), with "\n'
    '        "reclassifications driven primarily by persistence and summer-NDVI cutoffs in this dataset (amplitude variation from 0.20 to 0.28 caused zero reclassifications across all 75 configurations; Supplementary Table S9A). "\n'
    '        "Multi-scale spatial evaluation revealed canopy dilution in remnant underground forest quadrats (mean summer NDVI declining from 0.381 in "\n'
    '        "Zone A to 0.271 in Zone D) as extraction captures colliery infrastructure and clearings, whereas opencast spoil plantations remained "\n'
    '        "spatially stable (0.303 to 0.287). Principal Component Analysis accounted for 81.62% of variance across two primary axes, demonstrating "\n'
    '        "close agreement (score correlation r > 0.94) with an 8-feature sensitivity model excluding derived metrics. Non-parametric contrasts between "\n'
    '        "the sampled underground and opencast quadrats (n=5 per group) using SciPy\'s exact two-sided Mann-Whitney U procedure yielded non-significant "\n'
    '        "differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, exact p = 0.2222, Cliff\'s delta = +0.520, Hedges\' g = 0.647, approximate 95% analytical Student\'s t CI [-0.850, 2.143]), "\n'
    '        "demonstrating that non-significance reflects low statistical power in a pilot sample rather than ecological equivalence. We provide a complete machine-readable "\n'
    '        "provenance audit and outline requirements for future hierarchical, replicated landscape restoration designs.\\n"'
)

new_abs_sens_md = (
    '        "Threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations) "\n'
    '        "yielded 0 reclassifications in 24 configurations (32.0%), 1 in 21 (28.0%, cumulative ≤ 1: 45 [60.0%]), 2 in 9 (12.0%), 3 in 15 (20.0%), and 4 in 6 (8.0%; Supplementary Table S9A). "\n'
    '        "In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28) did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications. "\n'
    '        "Multi-scale spatial evaluation revealed a scale-dependent decline in mean summer NDVI in sampled underground forest quadrats (0.381 in Zone A to 0.271 in Zone D) "\n'
    '        "consistent with matrix inclusion, whereas sampled opencast quadrats showed smaller changes (0.303 to 0.287). Principal Component Analysis accounted "\n'
    '        "for 81.62% of variance across two primary axes, demonstrating close agreement (score correlation r > 0.94) with an 8-feature sensitivity model "\n'
    '        "excluding derived metrics. Non-parametric contrasts between the sampled underground and opencast quadrats (n=5 per group) using SciPy\'s exact two-sided "\n'
    '        "Mann-Whitney U procedure yielded non-significant differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, exact p = 0.2222, Cliff\'s delta = +0.520, "\n'
    '        "Hedges\' g = 0.647, approximate 95% analytical Student\'s t CI [-0.850, 2.143]); the wide uncertainty and small clustered sample make the non-significant result "\n'
    '        "uninformative about ecological equivalence. We provide a complete machine-readable provenance audit and outline requirements for future hierarchical, replicated landscape restoration designs.\\n"'
)

text = text.replace(old_abs_sens_md, new_abs_sens_md)

# 3. Section 3.1 Botanical Evidence (docx & md)
old_s31 = (
    '        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "\n'
    '        "across 25 species. All botanical nomenclature was audited and standardized according to Plants of the World Online (POWO; Royal Botanic "\n'
    '        "Gardens, Kew), resolving standard taxonomic synonyms (e.g., Shorea robusta Gaertn., Terminalia elliptica Willd. [syn. T. tomentosa (Roxb.) "\n'
    '        "Wight & Arn.], Holarrhena pubescens Wall. ex G.Don [syn. H. antidysenterica (L.) Wall.], Millettia pinnata (L.) Panigrahi [syn. Pongamia "\n'
    '        "pinnata (L.) Pierre]). Several subcanopy woody taxa (Ziziphus mauritiana, Carissa carandas, Holarrhena pubescens) were directly measured "\n'
    '        "within this Tier 1 census.\\n"'
)

new_s31 = (
    '        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "\n'
    '        "representing 25 reported field morphotaxa. All botanical nomenclature was audited and standardized according to Plants of the World Online "\n'
    '        "(POWO; Royal Botanic Gardens, Kew), using the accepted-name field as the sole basis for taxonomic richness. Following synonym harmonization, "\n'
    '        "the inventory comprises 24 unique accepted species (e.g., Terminalia tomentosa (Roxb.) Wight & Arn. is harmonized as a synonym of Terminalia elliptica Willd., "\n'
    '        "Holarrhena antidysenterica (L.) Wall. as a synonym of Holarrhena pubescens Wall. ex G.Don, and Bauhinia vahlii as Phanera vahlii (Wight & Arn.) Benth.). "\n'
    '        "Several subcanopy woody taxa (Ziziphus mauritiana, Carissa carandas, Holarrhena pubescens) were directly measured within this Tier 1 census.\\n"'
)

text = text.replace(old_s31, new_s31)

old_s31_md = (
    '        "- **Tier 1: Verified In-Plot Woody Stems (Empirical Ground Census).** The published field data sheets from the 10 quadrats enumerate 58 "\n'
    '        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "\n'
    '        "across 25 species. All botanical nomenclature was audited and standardized according to Plants of the World Online (POWO; Royal Botanic "\n'
    '        "Gardens, Kew), resolving standard taxonomic synonyms (e.g., *Shorea robusta* Gaertn., *Terminalia elliptica* Willd. [syn. *T. tomentosa* (Roxb.) "\n'
    '        "Wight & Arn.], *Holarrhena pubescens* Wall. ex G.Don [syn. *H. antidysenterica* (L.) Wall.], *Millettia pinnata* (L.) Panigrahi [syn. *Pongamia "\n'
    '        "pinnata* (L.) Pierre]). Several subcanopy woody taxa (*Ziziphus mauritiana*, *Carissa carandas*, *Holarrhena pubescens*) were directly measured "\n'
    '        "within this Tier 1 census.\\n"\n'
    '        "- **Tier 2: Photographic Habitat Records (Structural Reality).** An audited repository of 40 field photographic plates taken across all 10 "\n'
    '        "quadrats was systematically analyzed to document overstory canopy closure, bare ground exposure, leaf litter depth, and the presence "\n'
    '        "of ground-layer vegetation.\\n"\n'
    '        "- **Tier 3: Regional Floristic Candidate Leads (Unverified Context).** A register of 12 regional herbaceous and shrub species common to "\n'
    '        "disturbed mining areas in Korba (e.g., *Parthenium hysterophorus*, *Senna tora*, *Hyptis suaveolens*, *Lantana camara*, *Cynodon dactylon*) was "\n'
    '        "compiled from regional floristic surveys. These are treated strictly as unverified regional associates representing candidate hypotheses "\n'
    '        "for future nested micro-quadrat audits, rather than confirmed in-plot tallies.\\n\\n"'
)

new_s31_md = (
    '        "- **Tier 1: Verified In-Plot Woody Stems (Empirical Ground Census).** The published field data sheets from the 10 quadrats enumerate 58 "\n'
    '        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "\n'
    '        "representing 25 reported field morphotaxa. All botanical nomenclature was audited and standardized according to Plants of the World Online "\n'
    '        "(POWO; Royal Botanic Gardens, Kew), using the accepted-name field as the sole basis for taxonomic richness. Following synonym harmonization, "\n'
    '        "the inventory comprises 24 unique accepted species (e.g., *Terminalia tomentosa* (Roxb.) Wight & Arn. is harmonized as a synonym of *Terminalia elliptica* Willd., "\n'
    '        "*Holarrhena antidysenterica* (L.) Wall. as a synonym of *Holarrhena pubescens* Wall. ex G.Don, and *Bauhinia vahlii* as *Phanera vahlii* (Wight & Arn.) Benth.). "\n'
    '        "Several subcanopy woody taxa (*Ziziphus mauritiana*, *Carissa carandas*, *Holarrhena pubescens*) were directly measured within this Tier 1 census.\\n"\n'
    '        "- **Tier 2: Photographic Habitat Records (Structural Reality).** An audited repository of 40 field photographic plates taken across all 10 "\n'
    '        "quadrats was systematically analyzed to document overstory canopy closure, bare ground exposure, leaf litter depth, and the presence "\n'
    '        "of ground-layer vegetation.\\n"\n'
    '        "- **Tier 3: Regional Floristic Candidate Leads (Unverified Context).** A register of 12 regional herbaceous and shrub species common to "\n'
    '        "disturbed mining areas in Korba (e.g., *Parthenium hysterophorus*, *Senna tora*, *Hyptis suaveolens*, *Lantana camara*, *Cynodon dactylon*) was "\n'
    '        "compiled from regional floristic surveys. These are treated strictly as unverified regional associates representing candidate hypotheses "\n'
    '        "for future nested micro-quadrat audits, rather than confirmed in-plot tallies.\\n\\n"'
)

text = text.replace(old_s31_md, new_s31_md)

# Update Table 1 row 1 in docx and md
text = text.replace('58 stems (25 species)', '58 stems (24 accepted species; 25 reported taxa)')

# 4. Section 3.3 Spatial Support consistent wording
text = text.replace(
    'A circular disc of radius 50 m (theoretical geometric area = 7,854 m² [π × 50²]; raster support area = 8,100 m² across 81 pixels)',
    'A circular disc of radius 50 m (theoretical geometric area = 7,854 m² [π × 50²]; enclosing 81 raster pixels with a rasterized support area of 8,100 m²)'
)
text = text.replace(
    'A circular disc of radius 100 m (theoretical geometric area = 31,416 m² [π × 100²]; raster support area = 31,700 m² across 317 pixels)',
    'A circular disc of radius 100 m (theoretical geometric area = 31,416 m² [π × 100²]; enclosing 317 raster pixels with a rasterized support area of 31,700 m²)'
)

# 5. Section 3.5 Classification and Sensitivity (docx & md)
old_s35_docx = (
    '        "2. Pronounced Seasonal Flush: Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "\n'
    '        "disturbed ground where an open canopy permits intense monsoonal herbaceous and grass growth followed by rapid post-monsoon senescence.\\n"\n'
    '        "3. Disturbed Transition: Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "\n'
    '        "micro-hydrological conditions (e.g., mine water drainage features).\\n"\n'
    '        "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "\n'
    '        "analysis across 75 parameter configurations systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × "\n'
    '        "5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9 and Table S9A). "\n'
    '        "Across all 75 tested configurations, 24 (32.0%) preserved 100% of baseline site assignments, and 45 (60.0%) produced at most one reclassification (Table S9A). "\n'
    '        "Reclassifications were driven primarily by persistence and summer-NDVI cutoffs in this dataset (persistence ratio ≥ 0.92 reclassifies OC-Q1 and UG-Q2 to Disturbed Transition; "\n'
    '        "persistence ratio ≥ 0.95 reclassifies UG-Q4; summer NDVI ≥ 0.32 reclassifies UG-Q1), whereas varying the seasonal amplitude threshold between 0.20 and 0.28 caused zero "\n'
    '        "reclassifications across all 75 configurations. We emphasize that these groupings are heuristic descriptive spectral classes developed from this pilot dataset rather than universal ecological boundaries."'
)

new_s35_docx = (
    '        "2. Pronounced Seasonal Flush: Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "\n'
    '        "low-greenness spoil settings exhibiting a pronounced seasonal amplitude compatible with seasonal ground-cover development, but ground-layer contributions were not directly measured.\\n"\n'
    '        "3. Disturbed Transition: Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "\n'
    '        "micro-hydrological conditions (e.g., mine water drainage features).\\n\\n"\n'
    '        "The classification function evaluates each quadrat deterministically via the following logic:\\n"\n'
    '        "    def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\\n"\n'
    '        "        if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\\n"\n'
    '        "            return \'Persistent Greenness Signal\'\\n"\n'
    '        "        elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\\n"\n'
    '        "            return \'Pronounced Seasonal Flush\'\\n"\n'
    '        "        else:\\n"\n'
    '        "            return \'Disturbed Transition\'\\n\\n"\n'
    '        "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "\n'
    '        "analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × 5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × "\n'
    '        "3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B). Across all 75 tested configurations, 24 (32.0%) "\n'
    '        "preserved 100% of baseline site assignments, 21 (28.0%) produced 1 reclassification (cumulative ≤ 1: 45 [60.0%]), 9 (12.0%) produced 2 reclassifications, "\n'
    '        "15 (20.0%) produced 3 reclassifications, and 6 (8.0%) produced 4 reclassifications (Supplementary Table S9A). In this dataset, changing the amplitude "\n'
    '        "cutoff within the tested range (0.20 to 0.28) did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications "\n'
    '        "(e.g., persistence ratio ≥ 0.92 reclassifies OC-Q1 and UG-Q2; persistence ratio ≥ 0.95 reclassifies UG-Q4; summer NDVI ≥ 0.32 reclassifies UG-Q1). "\n'
    '        "We emphasize that these groupings are heuristic descriptive spectral classes developed from this pilot dataset rather than universal ecological boundaries."'
)

text = text.replace(old_s35_docx, new_s35_docx)

old_s35_md = (
    '        "2. **Pronounced Seasonal Flush:** Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "\n'
    '        "disturbed ground where an open canopy permits intense monsoonal herbaceous and grass growth followed by rapid post-monsoon senescence.\\n"\n'
    '        "3. **Disturbed Transition:** Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "\n'
    '        "micro-hydrological conditions (e.g., mine water drainage features).\\n\\n"\n'
    '        "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "\n'
    '        "analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations; "\n'
    '        "Supplementary Table S9A and Table S9B). Across all 75 tested configurations, 24 (32.0%) preserved 100% of baseline site assignments, and 45 (60.0%) produced "\n'
    '        "at most one reclassification. Reclassifications were driven primarily by persistence and summer-NDVI cutoffs in this dataset (e.g., persistence ≥ 0.92 reclassifies OC-Q1 and UG-Q2; "\n'
    '        "summer NDVI ≥ 0.32 reclassifies UG-Q1), whereas varying seasonal amplitude from 0.20 to 0.28 produced zero reclassifications across all 75 configurations. "\n'
    '        "We emphasize that these groupings are heuristic descriptive spectral classes developed from this pilot dataset rather than universal ecological boundaries.\\n"'
)

new_s35_md = (
    '        "2. **Pronounced Seasonal Flush:** Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "\n'
    '        "low-greenness spoil settings exhibiting a pronounced seasonal amplitude compatible with seasonal ground-cover development, but ground-layer contributions were not directly measured.\\n"\n'
    '        "3. **Disturbed Transition:** Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "\n'
    '        "micro-hydrological conditions (e.g., mine water drainage features).\\n\\n"\n'
    '        "The classification function evaluates each quadrat deterministically via the following logic:\\n\\n"\n'
    '        "```python\\n"\n'
    '        "def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\\n"\n'
    '        "    if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\\n"\n'
    '        "        return \'Persistent Greenness Signal\'\\n"\n'
    '        "    elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\\n"\n'
    '        "        return \'Pronounced Seasonal Flush\'\\n"\n'
    '        "    else:\\n"\n'
    '        "        return \'Disturbed Transition\'\\n"\n'
    '        "```\\n\\n"\n'
    '        "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "\n'
    '        "analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × 5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × "\n'
    '        "3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B). Across all 75 tested configurations, 24 (32.0%) "\n'
    '        "preserved 100% of baseline site assignments, 21 (28.0%) produced 1 reclassification (cumulative ≤ 1: 45 [60.0%]), 9 (12.0%) produced 2 reclassifications, "\n'
    '        "15 (20.0%) produced 3 reclassifications, and 6 (8.0%) produced 4 reclassifications (Supplementary Table S9A). In this dataset, changing the amplitude "\n'
    '        "cutoff within the tested range (0.20 to 0.28) did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications "\n'
    '        "(e.g., persistence ratio ≥ 0.92 reclassifies OC-Q1 and UG-Q2; persistence ratio ≥ 0.95 reclassifies UG-Q4; summer NDVI ≥ 0.32 reclassifies UG-Q1). "\n'
    '        "We emphasize that these groupings are heuristic descriptive spectral classes developed from this pilot dataset rather than universal ecological boundaries.\\n"'
)

text = text.replace(old_s35_md, new_s35_md)

# 6. Section 4.2 Annual maximum/minimum moderation
text = text.replace(
    'Date-by-date trajectories established that annual maximum greenness occurred on 2024-10-06 across almost all quadrats',
    'Date-by-date trajectories established that the maximum greenness among the six audited observations occurred on 2024-10-06 across almost all quadrats'
)
text = text.replace(
    'Conversely, annual minimum greenness (nadir) occurred during late pre-monsoon summer on 2024-05-14',
    'Conversely, the minimum greenness among the six audited observations occurred during late pre-monsoon summer on 2024-05-14'
)

# 7. Section 4.3 Canopy dilution & spatial window interpretation (docx & md)
old_s43_docx = (
    '        "In underground forest sites, expanding the spatial extraction window produced pronounced canopy dilution: mean Summer NDVI dropped sharply "\n'
    '        "from 0.381 in Zone A (10 m) to 0.305 in Zone B (50 × 50 m), 0.279 in Zone C (50 m radius), and 0.271 in Zone D (100 m radius)—a total dilution "\n'
    '        "of 0.111 NDVI units (29.0% reduction). This dilution occurs because underground quadrats were selectively positioned within remnant mature "\n'
    '        "forest patches; expanding the window to 50 m and 100 m incorporates dirt haulage tracks, colliery surface infrastructure, and cleared woodland "\n'
    '        "openings that dilute the dense canopy signal.\\n"\n'
    '        "In contrast, opencast spoil dump plantations exhibited remarkable spatial invariance: mean Summer NDVI remained at 0.303 in Zone A, "\n'
    '        "0.305 in Zone B, 0.302 in Zone C, and 0.287 in Zone D (only 0.016 units dilution, 5.2% reduction). Because opencast overburden dumps "\n'
    '        "present homogeneous spoil substrate and uniform plantation spacing over several hectares, spatial averaging across 100 m buffers "\n'
    '        "does not alter the underlying spectral signature."'
)

new_s43_docx = (
    '        "In underground forest sites, expanding the spatial extraction window produced a scale-dependent decline in mean Summer NDVI: dropping from "\n'
    '        "0.381 in Zone A (10 m) to 0.305 in Zone B (50 × 50 m), 0.279 in Zone C (50 m radius), and 0.271 in Zone D (100 m radius)—a total reduction "\n'
    '        "of 0.111 NDVI units (29.0%). The scale-dependent decline is consistent with inclusion of lower-greenness surfaces surrounding the nominal pixel, "\n'
    '        "including the documented colliery matrix; the relative contribution of each surface type was not independently quantified.\\n"\n'
    '        "In contrast, the sampled opencast quadrats showed smaller changes in mean NDVI across the tested windows: mean Summer NDVI was 0.303 in Zone A, "\n'
    '        "0.305 in Zone B, 0.302 in Zone C, and 0.287 in Zone D (a difference of 0.016 units, 5.2% reduction)."'
)

text = text.replace(old_s43_docx, new_s43_docx)

old_s43_md = (
    '        "In underground forest sites, expanding the spatial extraction window produced pronounced canopy dilution: mean Summer NDVI dropped sharply "\n'
    '        "from 0.381 in Zone A (10 m) to 0.305 in Zone B (50 × 50 m), 0.279 in Zone C (50 m radius), and 0.271 in Zone D (100 m radius)—a total dilution "\n'
    '        "of 0.111 NDVI units (29.0% reduction). This dilution occurs because underground quadrats were selectively positioned within remnant mature "\n'
    '        "forest patches; expanding the window to 50 m and 100 m incorporates dirt haulage tracks, colliery surface infrastructure, and cleared woodland "\n'
    '        "openings that dilute the dense canopy signal.\\n\\n"\n'
    '        "In contrast, opencast spoil dump plantations exhibited remarkable spatial invariance: mean Summer NDVI remained at 0.303 in Zone A, "\n'
    '        "0.305 in Zone B, 0.302 in Zone C, and 0.287 in Zone D (only 0.016 units dilution, 5.2% reduction). Because opencast overburden dumps "\n'
    '        "present homogeneous spoil substrate and uniform plantation spacing over several hectares, spatial averaging across 100 m buffers "\n'
    '        "does not alter the underlying spectral signature.\\n"'
)

new_s43_md = (
    '        "In underground forest sites, expanding the spatial extraction window produced a scale-dependent decline in mean Summer NDVI: dropping from "\n'
    '        "0.381 in Zone A (10 m) to 0.305 in Zone B (50 × 50 m), 0.279 in Zone C (50 m radius), and 0.271 in Zone D (100 m radius)—a total reduction "\n'
    '        "of 0.111 NDVI units (29.0%). The scale-dependent decline is consistent with inclusion of lower-greenness surfaces surrounding the nominal pixel, "\n'
    '        "including the documented colliery matrix; the relative contribution of each surface type was not independently quantified.\\n\\n"\n'
    '        "In contrast, the sampled opencast quadrats showed smaller changes in mean NDVI across the tested windows: mean Summer NDVI was 0.303 in Zone A, "\n'
    '        "0.305 in Zone B, 0.302 in Zone C, and 0.287 in Zone D (a difference of 0.016 units, 5.2% reduction).\\n"'
)

text = text.replace(old_s43_md, new_s43_md)

# 8. Section 4.4 Descriptive Classification (docx & md)
# Replace "permits" in seasonal flush
old_flush_docx = 'NDVI = 0.499; single-date peak 0.749 on Oct 06) and Summer NDVI < 0.25. The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush followed by rapid post-monsoon senescence.'
new_flush_docx = 'NDVI = 0.499; single-date peak 0.749 on Oct 06) and Summer NDVI < 0.25. The open, low-greenness spoil setting exhibited a pronounced seasonal amplitude that is compatible with seasonal ground-cover development, but the ground-layer contribution was not directly measured.'
text = text.replace(old_flush_docx, new_flush_docx)

old_flush_md = 'NDVI = 0.499; single-date peak 0.749 on Oct 06) and Summer NDVI < 0.25. The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush \n        followed by rapid post-monsoon senescence.'
# check markdown version
text = text.replace('The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush \n        followed by rapid post-monsoon senescence.',
                    'The open, low-greenness spoil setting exhibited a pronounced seasonal amplitude that is compatible with seasonal ground-cover development, but the ground-layer contribution was not directly measured.')
text = text.replace('The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush followed by rapid post-monsoon senescence.',
                    'The open, low-greenness spoil setting exhibited a pronounced seasonal amplitude that is compatible with seasonal ground-cover development, but the ground-layer contribution was not directly measured.')

# Reconcile Section 4.4 sensitivity text in docx
old_s44_docx_sens = (
    '        "The explicit threshold sensitivity analysis across 75 parameter configurations systematically varying three parameters (Summer NDVI, Persistence Ratio, "\n'
    '        "and Seasonal Amplitude; Supplementary Table S9 and Table S9A) demonstrated that the baseline classification was stable across a central range of thresholds "\n'
    '        "(preserving 100% of site assignments in 32.0% of tested configurations and ≤ 1 reclassification in 60.0%), but changed for more stringent persistence "\n'
    '        "or summer-NDVI thresholds, while variation in the seasonal amplitude cutoff (0.20 to 0.28) caused zero reclassifications across all 75 configurations. "\n'
    '        "Specifically, when the persistence cutoff was elevated to ≥ 0.92, OC-Q1 (ratio 0.908) and UG-Q2 (ratio 0.901) were reclassified to Disturbed Transition; "\n'
    '        "tightening to ≥ 0.95 reclassified UG-Q4 (ratio 0.921); and elevating the Summer NDVI cutoff to ≥ 0.32 reclassified UG-Q1 (summer 0.317). This sensitivity "\n'
    '        "underscores that these syndromic groupings are operational descriptive tools rather than rigid ecological states."'
)

new_s44_docx_sens = (
    '        "The explicit threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × "\n'
    '        "5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B) "\n'
    '        "yielded the following distribution of reclassifications: 0 changes in 24 configurations (32.0%, cumulative 32.0%), 1 change in 21 configurations "\n'
    '        "(28.0%, cumulative 60.0%), 2 changes in 9 configurations (12.0%, cumulative 72.0%), 3 changes in 15 configurations (20.0%, cumulative 92.0%), "\n'
    '        "and 4 changes in 6 configurations (8.0%, cumulative 100.0%). In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28) "\n'
    '        "did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications. Specifically, when the persistence cutoff was "\n'
    '        "elevated to ≥ 0.92, OC-Q1 (ratio 0.908) and UG-Q2 (ratio 0.901) were reclassified to Disturbed Transition; tightening to ≥ 0.95 reclassified UG-Q4 (ratio 0.921); "\n'
    '        "and elevating the Summer NDVI cutoff to ≥ 0.32 reclassified UG-Q1 (summer 0.317). This sensitivity underscores that these syndromic groupings are operational "\n'
    '        "descriptive spectral classes rather than rigid ecological states."'
)

text = text.replace(old_s44_docx_sens, new_s44_docx_sens)

# Reconcile Section 4.4 sensitivity text in md
old_s44_md_sens = (
    '        "The explicit threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations; "\n'
    '        "Supplementary Table S9A and Table S9B) demonstrated that the baseline classification was stable across a central range of thresholds (preserving 100% of site assignments "\n'
    '        "in 32.0% of tested configurations and ≤ 1 reclassification in 60.0%), but changed for more stringent persistence or summer-NDVI thresholds. Reclassifications were driven "\n'
    '        "primarily by persistence and summer-NDVI cutoffs (e.g., persistence ≥ 0.92 reclassified OC-Q1 and UG-Q2; summer NDVI ≥ 0.32 reclassified UG-Q1), whereas varying seasonal "\n'
    '        "amplitude from 0.20 to 0.28 produced zero reclassifications across all 75 configurations. This sensitivity underscores that these syndromic groupings are operational "\n'
    '        "descriptive spectral classes rather than rigid ecological states.\\n"'
)

new_s44_md_sens = (
    '        "The explicit threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × "\n'
    '        "5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B) "\n'
    '        "yielded the following distribution of reclassifications: 0 changes in 24 configurations (32.0%, cumulative 32.0%), 1 change in 21 configurations "\n'
    '        "(28.0%, cumulative 60.0%), 2 changes in 9 configurations (12.0%, cumulative 72.0%), 3 changes in 15 configurations (20.0%, cumulative 92.0%), "\n'
    '        "and 4 changes in 6 configurations (8.0%, cumulative 100.0%). In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28) "\n'
    '        "did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications. Specifically, when the persistence cutoff was "\n'
    '        "elevated to ≥ 0.92, OC-Q1 (ratio 0.908) and UG-Q2 (ratio 0.901) were reclassified to Disturbed Transition; tightening to ≥ 0.95 reclassified UG-Q4 (ratio 0.921); "\n'
    '        "and elevating the Summer NDVI cutoff to ≥ 0.32 reclassified UG-Q1 (summer 0.317). This sensitivity underscores that these syndromic groupings are operational "\n'
    '        "descriptive spectral classes rather than rigid ecological states.\\n"'
)

text = text.replace(old_s44_md_sens, new_s44_md_sens)

# 9. Statistical softening in Section 4.5
old_s45_soft = 'Non-significant p-values reflect low statistical power and substantial within-group variance (e.g., OC-Q1 plantation vs. OC-Q3 bare spoil) rather than evidence of ecological equivalence.'
new_s45_soft = 'The wide uncertainty and small clustered sample make the non-significant result uninformative about ecological equivalence, alongside substantial within-group variance (e.g., OC-Q1 plantation vs. OC-Q3 bare spoil).'
text = text.replace(old_s45_soft, new_s45_soft)

# 10. Supplementary material updates:
# In Table S3 caption:
text = text.replace(
    'Supplementary Table S3. Verified in-plot woody tree inventory (GBH ≥ 10 cm) with standardized POWO taxonomy',
    'Supplementary Table S3. Verified in-plot woody tree inventory (GBH ≥ 10 cm) across 24 accepted species (25 reported field morphotaxa) with standardized POWO taxonomy'
)

# In Table S9A caption:
text = text.replace(
    'Supplementary Table S9A. Compact frequency distribution of quadrat reclassifications across all 75 tested heuristic threshold configurations (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude).',
    'Supplementary Table S9A. Compact frequency distribution of quadrat reclassifications across all 75 tested heuristic threshold configurations (5 Summer NDVI [0.25–0.35] × 5 Persistence Ratio [0.85–0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28]). In this dataset, changing the amplitude cutoff within the tested range did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications.'
)

# Add classify_quadrat logic to Computational and Software Environment in Supplement
old_supp_env_docx = (
    '        "• python-docx: version 1.2.0\\n"\n'
    '        "All raster stacks and derived tables are archived with SHA-256 checksums in rasters/CHECKSUMS.sha256. Persistent repository link: "\n'
    '        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2."'
)

new_supp_env_docx = (
    '        "• python-docx: version 1.2.0\\n"\n'
    '        "All raster stacks and derived tables are archived with SHA-256 checksums in rasters/CHECKSUMS.sha256. Persistent repository link: "\n'
    '        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2.\\n\\n"\n'
    '        "Algorithm for Heuristic Phenological Descriptive Classification:\\n"\n'
    '        "def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\\n"\n'
    '        "    if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\\n"\n'
    '        "        return \'Persistent Greenness Signal\'\\n"\n'
    '        "    elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\\n"\n'
    '        "        return \'Pronounced Seasonal Flush\'\\n"\n'
    '        "    else:\\n"\n'
    '        "        return \'Disturbed Transition\'\\n"\n'
    '        "Threshold sensitivity analysis iteratively evaluated this function across 5 summer cutoffs (0.25, 0.28, 0.30, 0.32, 0.35), "\n'
    '        "5 persistence ratio cutoffs (0.85, 0.88, 0.90, 0.92, 0.95), and 3 seasonal amplitude cutoffs (0.20, 0.25, 0.28) for a total of 75 configurations."'
)

text = text.replace(old_supp_env_docx, new_supp_env_docx)

old_supp_env_md = (
    '- `python-docx`: version 1.2.0\\n\\n"\n'
    '        "All raster stacks and derived tables are archived with SHA-256 checksums in `rasters/CHECKSUMS.sha256`. Persistent repository link: "\n'
    '        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2.\\n\\n"'
)

new_supp_env_md = (
    '- `python-docx`: version 1.2.0\\n\\n"\n'
    '        "All raster stacks and derived tables are archived with SHA-256 checksums in `rasters/CHECKSUMS.sha256`. Persistent repository link: "\n'
    '        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2.\\n\\n"\n'
    '        "### Algorithm for Heuristic Phenological Descriptive Classification\\n\\n"\n'
    '        "```python\\n"\n'
    '        "def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\\n"\n'
    '        "    if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\\n"\n'
    '        "        return \'Persistent Greenness Signal\'\\n"\n'
    '        "    elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\\n"\n'
    '        "        return \'Pronounced Seasonal Flush\'\\n"\n'
    '        "    else:\\n"\n'
    '        "        return \'Disturbed Transition\'\\n"\n'
    '        "```\\n\\n"\n'
    '        "The threshold sensitivity analysis evaluated this function across 5 summer cutoffs (0.25, 0.28, 0.30, 0.32, 0.35), "\n'
    '        "5 persistence ratio cutoffs (0.85, 0.88, 0.90, 0.92, 0.95), and 3 seasonal amplitude cutoffs (0.20, 0.25, 0.28) for a total of 75 configurations.\\n\\n"'
)

text = text.replace(old_supp_env_md, new_supp_env_md)

with open(SCRIPT_PATH, "w", encoding="utf-8") as f:
    f.write(text)

print("Saved updated scripts/08_generate_manuscript_and_supplemental.py")
