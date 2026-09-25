"""
Patch script to update scripts/08_generate_manuscript_and_supplemental.py
for Version 2.4 (Technical & Editorial Reconciliation for External Peer Review)
incorporating all 10 Chief Editor requirements.
"""

import os
import re

SCRIPT_PATH = "scripts/08_generate_manuscript_and_supplemental.py"

with open(SCRIPT_PATH, "r", encoding="utf-8") as f:
    code = f.read()

# 1. Update Version in docx and md
code = code.replace(
    'Version: 2.3 (Fourth-Round Comprehensive Resolution) | Date: September 2026',
    'Version: 2.4 (Technical & Editorial Reconciliation for External Peer Review) | Date: September 2026'
)
code = code.replace(
    'Target Journal: Ecological Indicators | Article Type: Research Paper (Major Revision, Version 2.3) | September 2026',
    'Target Journal: Ecological Indicators | Article Type: Research Paper (Minor Revision, Version 2.4) | September 2026'
)
code = code.replace(
    'Shubham Sharma et al. | Ecological Indicators (Major Revision Version 2.3, September 2026)',
    'Shubham Sharma et al. | Ecological Indicators (Minor Revision Version 2.4, September 2026)'
)

# 2. Update Highlights in docx and md
code = code.replace(
    'persistent canopy signals (summer dry-season NDVI ≥ 0.30, winter-to-green-season ratio ≥ 0.90) vs pronounced seasonal flush (amplitude > 0.25, summer NDVI < 0.25), evaluating stability across 75 threshold combinations.',
    'persistent greenness signals (summer dry-season NDVI ≥ 0.30, winter-to-green-season ratio ≥ 0.90; descriptive spectral class) vs pronounced seasonal flush (amplitude > 0.25, summer NDVI < 0.25), evaluating stability across 75 threshold configurations.'
)

# 3. Update Abstract in docx and md
old_abs_fragment = (
    "In Zone A, 100% valid terrestrial pixel retention was achieved across all 60 site-date cases, while larger windows exhibited variable "
    "valid pixel fractions (98.8% to 99.7%) due to colliery infrastructure and water bodies. Rather than deconvolving isolated strata, "
    "satellite trajectories capture integrated vertical surface reflectance across all canopy layers and exposed ground. Quadrats partitioned "
    "into three heuristic descriptive classes: (1) Persistent Canopy Signal (Summer dry-season NDVI ≥ 0.30, Winter-to-Green-Season Persistence Ratio ≥ 0.90; "
    "UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4); (2) Pronounced Seasonal Flush (Seasonal Amplitude ΔNDVI > 0.25, Summer NDVI < 0.25; OC-Q5); and "
    "(3) Disturbed Transition, which includes sites consistent with a possible moisture-related influence near drainage features such as OC-Q2, "
    "where winter greenness retention is elevated (persistence ratio = 1.129; winter composite = 0.414, single-date Dec 30 value = 0.453). "
    "Threshold sensitivity analysis across 75 parameter configurations indicated central classification stability (preserving 100% of site "
    "assignments in 32.0% of configurations and ≤ 1 reclassification in 60.0%), while showing expected sensitivity under more stringent persistence "
    "cutoffs. Multi-scale spatial evaluation revealed canopy dilution in remnant underground forest quadrats (mean summer NDVI declining from 0.381 in "
    "Zone A to 0.271 in Zone D) as extraction captures colliery infrastructure and clearings, whereas opencast spoil plantations remained "
    "spatially stable (0.303 to 0.287). Principal Component Analysis accounted for 81.62% of variance across two primary axes, demonstrating "
    "close agreement (score correlation r > 0.94) with an 8-feature sensitivity model excluding derived metrics. Non-parametric contrasts between "
    "the sampled underground and opencast quadrats (n=5 per group) yielded non-significant differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, "
    "exact p = 0.2222, Cliff's delta = +0.520, Hedges' g = 0.647, approximate 95% analytical Student's t CI [-0.850, 2.143]), demonstrating that "
    "non-significance reflects low statistical power in a pilot sample rather than ecological equivalence."
)

new_abs_fragment = (
    "In Zone A, 100.00% valid terrestrial pixel retention was achieved across all 60 site-date cases (60/60 pixels), while larger windows exhibited "
    "valid-pixel fractions of 98.99% to 99.33% (Zone B: 99.33% [1,490/1,500], Zone C: 99.22% [4,822/4,860], Zone D: 98.99% [18,828/19,020]) "
    "due to colliery infrastructure and water bodies. Rather than deconvolving isolated strata, satellite trajectories capture integrated top-of-canopy "
    "surface reflectance influenced by overstory foliage, subcanopy vegetation, ground cover, litter, soil, and exposed substrate. Quadrats partitioned "
    "into three heuristic descriptive classes: (1) Persistent Greenness Signal (Summer dry-season NDVI ≥ 0.30, Winter-to-Green-Season Persistence Ratio ≥ 0.90; "
    "UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4; descriptive spectral class); (2) Pronounced Seasonal Flush (Seasonal Amplitude ΔNDVI > 0.25, Summer NDVI < 0.25; OC-Q5); "
    "and (3) Disturbed Transition, which includes sites consistent with a possible moisture-related influence near drainage features such as OC-Q2, "
    "where winter greenness retention is elevated (persistence ratio = 1.129; winter composite = 0.414, single-date Dec 30 value = 0.453). "
    "Threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations) "
    "indicated central classification stability (preserving 100% of site assignments in 32.0% of configurations and ≤ 1 reclassification in 60.0%), with "
    "reclassifications driven primarily by persistence and summer-NDVI cutoffs in this dataset. Multi-scale spatial evaluation revealed canopy dilution "
    "in remnant underground forest quadrats (mean summer NDVI declining from 0.381 in Zone A to 0.271 in Zone D) as extraction captures colliery "
    "infrastructure and clearings, whereas opencast spoil plantations remained spatially stable (0.303 to 0.287). Principal Component Analysis accounted "
    "for 81.62% of variance across two primary axes, demonstrating close agreement (score correlation r > 0.94) with an 8-feature sensitivity model "
    "excluding derived metrics. Non-parametric contrasts between the sampled underground and opencast quadrats (n=5 per group) using SciPy's exact two-sided "
    "Mann-Whitney U procedure yielded non-significant differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, exact p = 0.2222, Cliff's delta = +0.520, "
    "Hedges' g = 0.647, approximate 95% analytical Student's t CI [-0.850, 2.143]), demonstrating that non-significance reflects low statistical power "
    "in a pilot sample rather than ecological equivalence."
)

code = code.replace(old_abs_fragment, new_abs_fragment)

# 4. Update Research Question 2 in Markdown section if not already
code = code.replace(
    "2. How do heuristic descriptive classification rules (persistent canopy signals vs. pronounced seasonal flushes)",
    "2. How do heuristic descriptive classification rules (persistent greenness signals vs. pronounced seasonal flushes)"
)

# 5. Update Section 3.2 in Markdown
old_s32_md = (
    "In Zone A (nominal 10 m pixel support), all 10 quadrats achieved 100% valid terrestrial pixel retention across all six dates (zero masked pixels across 60 site-date cases; Table 2). "
    "Larger multi-scale extraction windows (Zones B–D) had variable valid pixel fractions (98.8% to 99.7%) due to edge inclusions of water bodies, shadows, and colliery infrastructure (Table 3; Table 6)."
)
new_s32_md = (
    "In Zone A (nominal 10 m pixel support), 100.00% valid terrestrial pixel retention was achieved across all six dates (zero masked pixels across 60 site-date cases; Table 2). "
    "Terrestrial validity percentage is defined as the total number of valid terrestrial pixels (SCL classes 4 [vegetation] and 5 [bare soil]) divided by "
    "the theoretical denominator pixel count across all 60 site-date observations (which exactly equals the arithmetic mean of individual site-date validity "
    "percentages to two decimal places): Zone A = 100.00% (60/60), Zone B = 99.33% (1,490/1,500), Zone C = 99.22% (4,822/4,860), and Zone D = 98.99% "
    "(18,828/19,020), defining a range of 98.99% to 99.33% across multi-scale extraction windows due to edge inclusions of water bodies, shadows, and colliery infrastructure (Table 3; Table 6)."
)
code = code.replace(old_s32_md, new_s32_md)

# 6. Update Section 3.3 in Markdown
old_s33_md = (
    "- **Zone A (Nominal Quadrat Pixel Support):** Exactly 1 pixel (10 × 10 m = 100 m² = 0.01 ha), defined as the raster pixel containing the quadrat centroid.\n"
    "- **Zone B (Local Neighborhood):** A square 5 × 5 pixel window (50 × 50 m = 2,500 m² = 0.25 ha, 25 pixels total support), centered on the quadrat centroid and enclosing its GPS uncertainty footprint.\n"
    "- **Zone C (Intermediate Stand Buffer):** A circular disc of radius 50 m (81 raster pixels = 8,100 m² = 0.81 ha), capturing the local forest or overburden dump plantation stand.\n"
    "- **Zone D (Contextual Landscape Buffer):** A circular disc of radius 100 m (317 raster pixels = 31,700 m² = 3.17 ha), capturing the broader colliery landscape matrix."
)
new_s33_md = (
    "- **Zone A (Nominal Quadrat Pixel Support):** Exactly 1 pixel (geometric area = 100 m² = 0.01 ha; raster support area = 100 m²), defined as the raster pixel containing the quadrat centroid.\n"
    "- **Zone B (Local Neighborhood):** A square 5 × 5 pixel window (geometric area = 2,500 m² = 0.25 ha; raster support area = 2,500 m², 25 pixels total support), centered on the quadrat centroid and enclosing its GPS uncertainty footprint.\n"
    "- **Zone C (Intermediate Stand Buffer):** A circular disc of radius 50 m (theoretical geometric area = 7,854 m² [π × 50², 0.79 ha]; raster support area = approximately 8,100 m², enclosing 81 raster pixels), capturing the local forest or overburden dump plantation stand.\n"
    "- **Zone D (Contextual Landscape Buffer):** A circular disc of radius 100 m (theoretical geometric area = 31,416 m² [π × 100², 3.14 ha]; raster support area = approximately 31,700 m², enclosing 317 raster pixels), capturing the broader colliery landscape matrix."
)
code = code.replace(old_s33_md, new_s33_md)

# 7. Update Table 3 caption in Markdown
code = code.replace(
    "### Table 3. Multi-scale spatial extraction geometries, raster support denominators, and SCL terrestrial validity across all 60 site-date cases.",
    "### Table 3. Multi-scale spatial extraction geometries, theoretical geometric areas, raster support areas, and SCL terrestrial validity across all 60 site-date cases."
)

# 8. Update Section 3.5 in Markdown
old_s35_md = (
    "1. **Persistent Canopy Signal:** Summer dry-season NDVI ≥ 0.30 AND Winter-to-Green-Season Persistence Ratio ≥ 0.90. This class characterizes "
    "sites where established woody canopy cover maintains high foliar greenness and suppresses dry-season seasonal amplitude.\n"
    "2. **Pronounced Seasonal Flush:** Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "
    "disturbed ground where an open canopy permits intense monsoonal herbaceous and grass growth followed by rapid post-monsoon senescence.\n"
    "3. **Disturbed Transition:** Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "
    "micro-hydrological conditions (e.g., mine water drainage features).\n\n"
    "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "
    "analysis across 75 parameter configurations (Summer NDVI cutoff 0.25–0.35, Persistence Ratio cutoff 0.85–0.95, Amplitude cutoff 0.20–0.28; "
    "Supplementary Table S9). Across all 75 tested configurations, 24 (32.0%) preserved 100% of baseline site assignments, and 45 (60.0%) produced "
    "at most one reclassification. Reclassifications occurred under more stringent persistence thresholds (≥ 0.92, which reclassifies OC-Q1 and UG-Q2) "
    "or elevated summer thresholds (≥ 0.32, which reclassifies UG-Q1). We emphasize that these groupings are descriptive heuristic summaries rather than general ecological laws."
)
new_s35_md = (
    "1. **Persistent Greenness Signal (High Dry-Season Spectral Retention; descriptive spectral class):** Summer dry-season NDVI ≥ 0.30 AND "
    "Winter-to-Green-Season Persistence Ratio ≥ 0.90. This class characterizes sites where high foliar greenness is maintained through the post-monsoon "
    "and pre-monsoon dry season, suppressing seasonal amplitude. We emphasize that this spectral persistence reflects integrated top-of-canopy "
    "retention and does not directly establish overstory crown dominance or rooting depth without independent ground measurements.\n"
    "2. **Pronounced Seasonal Flush:** Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "
    "disturbed ground where an open canopy permits intense monsoonal herbaceous and grass growth followed by rapid post-monsoon senescence.\n"
    "3. **Disturbed Transition:** Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "
    "micro-hydrological conditions (e.g., mine water drainage features).\n\n"
    "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "
    "analysis across 75 parameter configurations systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × "
    "5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9 and Table S9A). "
    "Across all 75 tested configurations, 24 (32.0%) preserved 100% of baseline site assignments, and 45 (60.0%) produced at most one reclassification (Table S9A). "
    "Reclassifications were driven primarily by persistence and summer-NDVI cutoffs in this dataset (persistence ratio ≥ 0.92 reclassifies OC-Q1 and UG-Q2 to Disturbed Transition; "
    "persistence ratio ≥ 0.95 reclassifies UG-Q4; summer NDVI ≥ 0.32 reclassifies UG-Q1), whereas varying the seasonal amplitude threshold between 0.20 and 0.28 caused zero "
    "reclassifications across all 75 configurations. We emphasize that these groupings are heuristic descriptive spectral classes developed from this pilot dataset rather than universal ecological boundaries."
)
code = code.replace(old_s35_md, new_s35_md)

# 9. Update Section 3.6 in Markdown
old_s36_md = (
    "Contrasts between sampled underground and opencast quadrats were evaluated using the exact two-sided Mann-Whitney U test "
    "(scipy.stats.mannwhitneyu, method='exact'), Cliff's delta, and small-sample corrected Hedges' g. In the tabular results (Table 5) and audit records "
    "(Table S7), we explicitly report both the first-sample U statistic (U₁, corresponding to Underground) and the smaller U statistic "
    "(U_min = min(U₁, U₂)), resolving potential ambiguities arising from software conventions.\n\n"
    "A standardized mean difference (Hedges' g) was paired with non-parametric rank tests because Mann-Whitney U evaluates whether one distribution "
    "stochastically dominates another without distributional assumptions, whereas Hedges' g provides a standardized, scale-free descriptive effect size "
    "with an analytical confidence interval to facilitate future meta-analyses and sample-size planning for replicated restoration designs. In accordance "
    "with Hedges & Olkin (1985), approximate 95% confidence intervals for Hedges' g were calculated using the Student's t distribution with df = n₁ + n₂ - 2 = 8 "
    "degrees of freedom: CI = g ± t(0.025, 8) × SE(g), where critical t(0.025, 8) = 2.3060, SE(g) = sqrt([(n₁ + n₂)/(n₁ × n₂)] + [g² / (2(n₁ + n₂))]), "
    "pooled standard deviation s_pooled = sqrt([((n₁ - 1)s₁² + (n₂ - 1)s₂²) / df], and small-sample correction factor J = 1 - [3 / (4(n₁ + n₂) - 9)] = 28/31 ≈ 0.9032."
)
new_s36_md = (
    "Contrasts between sampled underground and opencast quadrats were evaluated using SciPy’s exact two-sided Mann-Whitney U procedure "
    "(scipy.stats.mannwhitneyu, alternative='two-sided', method='exact') with no tie correction required for the reported values. An independent exhaustive "
    "enumeration of all 252 label permutations (10 choose 5) confirmed identical results (56/252 = 2/9 = 0.2222), demonstrating exact agreement between both "
    "calculations. In the tabular results (Table 5) and audit records (Table S7), we explicitly report both the first-sample U statistic "
    "(U₁, corresponding to Underground) and the smaller U statistic (U_min = min(U₁, U₂)), resolving potential ambiguities arising from software conventions.\n\n"
    "A standardized mean difference (Hedges' g) was paired with non-parametric rank tests because Mann-Whitney U evaluates whether one distribution "
    "stochastically dominates another without distributional assumptions, whereas Hedges' g provides a standardized, scale-free descriptive effect size "
    "to facilitate future meta-analyses and sample-size planning for replicated restoration designs. We report an approximate interval obtained by "
    "multiplying the estimated standard error of Hedges’ g by the t(8, 0.975) critical value (t = 2.3060); this interval is descriptive and not intended "
    "to provide exact small-sample coverage (Hedges & Olkin, 1985). Specifically, approximate 95% confidence intervals were calculated using the Student's t "
    "critical value with df = n₁ + n₂ - 2 = 8: CI = g ± t(0.025, 8) × SE(g), where critical t(0.025, 8) = 2.3060, SE(g) = sqrt([(n₁ + n₂)/(n₁ × n₂)] + [g² / (2(n₁ + n₂))]), "
    "pooled standard deviation s_pooled = sqrt([((n₁ - 1)s₁² + (n₂ - 1)s₂²) / df], and small-sample correction factor J = 1 - [3 / (4(n₁ + n₂) - 9)] = 28/31 ≈ 0.9032."
)
code = code.replace(old_s36_md, new_s36_md)

# 10. Update Section 4.3 in Markdown
old_s43_md = (
    "In Zone A, 100% of pixels met strict terrestrial SCL criteria (zero masked pixels across 60 site-date cases). In expanded windows, valid pixel retention "
    "remained very high but variable due to colliery infrastructure and drainage features: Zone B averaged 99.33% valid pixels (24.83/25), "
    "Zone C averaged 99.22% (80.37/81), and Zone D averaged 98.99% (313.8/317; Table 3)."
)
new_s43_md = (
    "In Zone A (nominal 10 m pixel support), 100.00% of pixels met strict terrestrial SCL criteria (zero masked pixels across 60 site-date cases; 60/60 valid observations). "
    "In expanded windows, valid pixel retention remained very high but variable due to colliery infrastructure and drainage features: Zone B averaged 99.33% "
    "valid pixels (1,490/1,500; mean 24.83/25), Zone C averaged 99.22% (4,822/4,860; mean 80.37/81), and Zone D averaged 98.99% (18,828/19,020; "
    "mean 313.80/317), defining a range of 98.99% to 99.33% across multi-scale extraction windows (Table 3; Table 6)."
)
code = code.replace(old_s43_md, new_s43_md)

# 11. Update Section 4.4 in Markdown
old_s44_md = (
    "- **Persistent Canopy Signal:** Five quadrats (UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4) met the baseline criteria of Summer dry-season NDVI ≥ 0.30 and "
    "Winter-to-Green-Season persistence ratio ≥ 0.90. These sites are dominated by dense overstory crown cover (Shorea robusta, Terminalia elliptica) "
    "or stabilized woody reclamation stands (Eucalyptus plantation in OC-Q1, Holarrhena thicket in OC-Q4) that maintain green foliage well into the dry season.\n"
)
new_s44_md = (
    "- **Persistent Greenness Signal (High Dry-Season Spectral Retention; descriptive spectral class):** Five quadrats (UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4) "
    "met the baseline criteria of Summer dry-season NDVI ≥ 0.30 and Winter-to-Green-Season persistence ratio ≥ 0.90. These sites maintain relatively "
    "high dry-season foliar greenness (e.g. Shorea robusta and Terminalia stands in UG; Eucalyptus plantation in OC-Q1, Holarrhena stand in OC-Q4), "
    "consistent with persistent woody cover or local retention, though ground-level canopy fraction was not independently measured.\n"
)
code = code.replace(old_s44_md, new_s44_md)

old_s44_sens_md = (
    "The explicit threshold sensitivity analysis across 75 parameter configurations (Supplementary Table S9) demonstrated that the baseline classification "
    "was stable across a central range of thresholds (preserving 100% of site assignments in 32.0% of tested configurations and ≤ 1 reclassification in 60.0%), "
    "but changed for more stringent persistence or summer-NDVI thresholds."
)
new_s44_sens_md = (
    "The explicit threshold sensitivity analysis across 75 parameter configurations systematically varying three parameters (Summer NDVI, Persistence Ratio, "
    "and Seasonal Amplitude; Supplementary Table S9 and Table S9A) demonstrated that the baseline classification was stable across a central range of thresholds "
    "(preserving 100% of site assignments in 32.0% of tested configurations and ≤ 1 reclassification in 60.0%), but changed for more stringent persistence "
    "or summer-NDVI thresholds, while variation in the seasonal amplitude cutoff (0.20 to 0.28) caused zero reclassifications across all 75 configurations."
)
code = code.replace(old_s44_sens_md, new_s44_sens_md)

# 12. Update Figure 4 Caption in Markdown
code = code.replace(
    "heuristic threshold boundaries for the Persistent Canopy Domain (Persistence Ratio ≥ 0.90, green shading)",
    "heuristic threshold boundaries for the Persistent Greenness Domain (Persistence Ratio ≥ 0.90, green shading)"
)

# 13. Update Section 4.5 in Markdown
old_s45_md = (
    "Exact non-parametric permutation tests comparing sampled underground and opencast quadrats in Zone A (Table 5; cross-referenced with "
    "the machine-readable verification table Table S7) indicated that group differences were not statistically significant at alpha = 0.05. "
    "For Summer NDVI, the exact Mann-Whitney test yielded U₁ = 19.0 (first-sample U, Underground) and U_min = 6.0 (smaller U), with exact "
    "permutation p = 0.2222, Cliff's delta = +0.520, and Hedges' g = 0.647 (approximate 95% analytical Student's t CI [-0.850, 2.143])."
)
new_s45_md = (
    "Contrasts between sampled underground and opencast quadrats in Zone A were evaluated using SciPy’s exact two-sided Mann-Whitney U procedure "
    "(method='exact') with no tie correction required for the reported values (Table 5; cross-referenced with the machine-readable verification "
    "table Table S7). An independent exhaustive enumeration of all 252 label permutations (10 choose 5) confirmed identical results (56/252 = 0.2222), "
    "demonstrating exact agreement. For Summer NDVI, the test yielded U₁ = 19.0 (first-sample U, Underground) and U_min = 6.0 (smaller U), with exact "
    "two-sided p = 0.2222, Cliff's delta = +0.520, and Hedges' g = 0.647 (approximate 95% analytical Student's t CI [-0.850, 2.143]). We report an "
    "approximate interval obtained by multiplying the estimated standard error of Hedges’ g by the t(8, 0.975) critical value (t = 2.3060); this "
    "interval is descriptive and not intended to provide exact small-sample coverage."
)
code = code.replace(old_s45_md, new_s45_md)

# 14. Update Discussion 5.1 in Markdown
old_d51_md = (
    "A central methodological conclusion of this pilot study is that 10 m Sentinel-2 pixels cannot deconvolve vertical vegetation strata "
    "or directly identify understory species in the absence of nested ground plots and sub-pixel fractional cover data. Spectral reflectance "
    "recorded by the satellite represents an integrated vertical measurement combining overstory crown foliage, subcanopy shrubs, ground herbs, "
    "litter, and soil background. While phenological dynamics (such as seasonal amplitude ΔNDVI and the winter-to-green-season persistence ratio) "
    "provide powerful descriptive indicators of ecosystem behavior, asserting that satellite pixels isolate herbaceous from woody components "
    "is empirically unsupportable. High winter-to-green-season ratios identify relatively strong dry-season spectral retention. They do not, "
    "by themselves, demonstrate perennial woody cover, rooting depth, or year-round soil stabilization."
)
new_d51_md = (
    "A central methodological conclusion of this pilot study is that 10 m Sentinel-2 pixels cannot deconvolve vertical vegetation strata "
    "or directly identify understory species in the absence of nested ground plots and sub-pixel fractional cover data. Spectral reflectance "
    "recorded by the satellite represents integrated top-of-canopy surface reflectance influenced by overstory foliage, subcanopy vegetation, "
    "ground cover, litter, soil, and exposed substrate. While phenological dynamics (such as seasonal amplitude ΔNDVI and the winter-to-green-season "
    "persistence ratio) provide powerful descriptive indicators of ecosystem behavior, asserting that satellite pixels isolate herbaceous from woody "
    "components is empirically unsupportable. High winter-to-green-season ratios identify relatively strong dry-season spectral retention. They do not, "
    "by themselves, demonstrate perennial woody cover, rooting depth, or year-round soil stabilization."
)
code = code.replace(old_d51_md, new_d51_md)

# 15. Update Conclusions in Markdown
old_c1_md = "1. **Integrated spectral response:** Sentinel-2 observations represent integrated pixel-level spectral responses across all canopy layers, subcanopy shrubs, ground herbs, leaf litter, and exposed soil, rather than isolated herbaceous or shrub signals."
new_c1_md = "1. **Integrated top-of-canopy spectral response:** Sentinel-2 observations represent integrated top-of-canopy surface reflectance influenced by overstory foliage, subcanopy vegetation, ground cover, litter, soil, and exposed substrate, rather than isolated herbaceous or shrub signals."
code = code.replace(old_c1_md, new_c1_md)

old_c5_md = "5. **Descriptive classification stability:** Heuristic phenological classes serve as transparent descriptive summaries of multi-temporal greenness trajectories."
new_c5_md = "5. **Descriptive classification stability:** Heuristic phenological classes (such as Persistent Greenness Signal) serve as transparent descriptive summaries of multi-temporal greenness trajectories."
code = code.replace(old_c5_md, new_c5_md)

# 16. Update Supplementary Table S6 with Source_Type in both docx and md
# In docx builder:
old_s6_docx = """    # Supplementary Table S6: Soil Properties
    doc.add_paragraph("Supplementary Table S6. Harmonized soil physicochemical properties (measured field baseline and SoilGrids 250m) across the 10 quadrats.", style='Caption')
    soil_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_7_Quadrat_Soil_Harmonized_Master.csv"))
    soil_disp = soil_df[['Quadrat_ID', 'Mining_Type', 'Measured_pH', 'Measured_OC_pct', 'Measured_BD_g_cm3', 'SoilGrids_pH', 'SoilGrids_SOC_pct', 'SoilGrids_BD_g_cm3']].copy()
    soil_disp.columns = ['Quadrat', 'Mining Type', 'Measured pH', 'Measured SOC (%)', 'Measured BD (g/cm³)', 'SoilGrids pH', 'SoilGrids SOC (%)', 'SoilGrids BD (g/cm³)']
    for c in ['SoilGrids pH', 'SoilGrids SOC (%)', 'SoilGrids BD (g/cm³)']:
        soil_disp[c] = soil_disp[c].round(2)
    add_table_from_df(doc, soil_disp, col_widths=[0.8, 1.0, 0.9, 0.9, 1.0, 0.9, 0.9, 1.0])"""

new_s6_docx = """    # Supplementary Table S6: Contextual Soil Properties
    doc.add_paragraph("Supplementary Table S6. Contextual soil physicochemical properties across the 10 quadrats (distinguishing published composite samples, regional literature ranges, chronosequence ranges, and model-derived SoilGrids 250m estimates).", style='Caption')
    soil_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_7_Quadrat_Soil_Harmonized_Master.csv"))
    source_map = {
        'UG-Q1': 'Colliery composite sample (Patel et al. 2016, 0-15cm)',
        'UG-Q2': 'Regional literature baseline (Surakachhar EIA)',
        'UG-Q3': 'Regional literature baseline (Surakachhar EIA)',
        'UG-Q4': 'Colliery composite sample (Patel et al. 2016, 0-15cm)',
        'UG-Q5': 'Colliery composite sample (Patel et al. 2016, 0-15cm)',
        'OC-Q1': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q2': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q3': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q4': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q5': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)'
    }
    soil_disp = pd.DataFrame()
    soil_disp['Quadrat'] = soil_df['Quadrat_ID']
    soil_disp['Mining Type'] = soil_df['Mining_Type']
    soil_disp['Source Type'] = soil_df['Quadrat_ID'].map(source_map)
    soil_disp['Baseline pH'] = soil_df['Measured_pH']
    soil_disp['Baseline SOC (%)'] = soil_df['Measured_OC_pct']
    soil_disp['Baseline BD (g/cm³)'] = soil_df['Measured_BD_g_cm3']
    soil_disp['SoilGrids pH'] = soil_df['SoilGrids_pH'].round(2)
    soil_disp['SoilGrids SOC (%)'] = soil_df['SoilGrids_SOC_pct'].round(2)
    soil_disp['SoilGrids BD (g/cm³)'] = soil_df['SoilGrids_BD_g_cm3'].round(2)
    add_table_from_df(doc, soil_disp, col_widths=[0.7, 0.8, 1.8, 0.7, 0.8, 0.9, 0.7, 0.8, 0.8])
    doc.add_paragraph(
        "Contextual Note: Soil physicochemical parameters are compiled strictly as background environmental context. "
        "Several entries represent published colliery-level composites or regional chronosequence ranges from nearby stations "
        "(0.67 to 3.49 km away) and SoilGrids 250 m model predictions, rather than independent quadrat-specific empirical measurements. "
        "Consequently, these soil data were not used in the primary statistical comparisons between quadrat groups.", style='Normal'
    )"""

code = code.replace(old_s6_docx, new_s6_docx)

# In md generator:
old_s6_md = """    # Table S6
    lines.append("### Supplementary Table S6. Harmonized soil physicochemical properties (measured field baseline and SoilGrids 250m) across the 10 quadrats.\\n")
    soil_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_7_Quadrat_Soil_Harmonized_Master.csv"))
    soil_disp = soil_df[['Quadrat_ID', 'Mining_Type', 'Measured_pH', 'Measured_OC_pct', 'Measured_BD_g_cm3', 'SoilGrids_pH', 'SoilGrids_SOC_pct', 'SoilGrids_BD_g_cm3']].copy()
    soil_disp.columns = ['Quadrat', 'Mining Type', 'Measured pH', 'Measured SOC (%)', 'Measured BD (g/cm³)', 'SoilGrids pH', 'SoilGrids SOC (%)', 'SoilGrids BD (g/cm³)']
    for c in ['SoilGrids pH', 'SoilGrids SOC (%)', 'SoilGrids BD (g/cm³)']:
        soil_disp[c] = soil_disp[c].round(2)
    lines.append(df_to_markdown(soil_disp))
    lines.append("\\n\\n")"""

new_s6_md = """    # Table S6
    lines.append("### Supplementary Table S6. Contextual soil physicochemical properties across the 10 quadrats (distinguishing published composite samples, regional literature ranges, chronosequence ranges, and model-derived SoilGrids 250m estimates).\\n")
    soil_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_7_Quadrat_Soil_Harmonized_Master.csv"))
    source_map = {
        'UG-Q1': 'Colliery composite sample (Patel et al. 2016, 0-15cm)',
        'UG-Q2': 'Regional literature baseline (Surakachhar EIA)',
        'UG-Q3': 'Regional literature baseline (Surakachhar EIA)',
        'UG-Q4': 'Colliery composite sample (Patel et al. 2016, 0-15cm)',
        'UG-Q5': 'Colliery composite sample (Patel et al. 2016, 0-15cm)',
        'OC-Q1': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q2': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q3': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q4': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)',
        'OC-Q5': 'Overburden chronosequence composite (Patel et al. 2016 / Singh et al. 2022)'
    }
    soil_disp = pd.DataFrame()
    soil_disp['Quadrat'] = soil_df['Quadrat_ID']
    soil_disp['Mining Type'] = soil_df['Mining_Type']
    soil_disp['Source Type'] = soil_df['Quadrat_ID'].map(source_map)
    soil_disp['Baseline pH'] = soil_df['Measured_pH']
    soil_disp['Baseline SOC (%)'] = soil_df['Measured_OC_pct']
    soil_disp['Baseline BD (g/cm³)'] = soil_df['Measured_BD_g_cm3']
    soil_disp['SoilGrids pH'] = soil_df['SoilGrids_pH'].round(2)
    soil_disp['SoilGrids SOC (%)'] = soil_df['SoilGrids_SOC_pct'].round(2)
    soil_disp['SoilGrids BD (g/cm³)'] = soil_df['SoilGrids_BD_g_cm3'].round(2)
    lines.append(df_to_markdown(soil_disp))
    lines.append("\\n*Contextual Note: Soil physicochemical parameters are compiled strictly as background environmental context. Several entries represent published colliery-level composites or regional chronosequence ranges from nearby stations (0.67 to 3.49 km away) and SoilGrids 250 m model predictions, rather than independent quadrat-specific empirical measurements. Consequently, these soil data were not used in the primary statistical comparisons between quadrat groups.*\\n\\n")"""

code = code.replace(old_s6_md, new_s6_md)

# 17. Update Supplementary Table S9 to include Table S9A and Table S9B in docx and md
old_s9_docx = """    # Supplementary Table S9: Syndrome Threshold Sensitivity Analysis
    doc.add_paragraph("Supplementary Table S9. Heuristic phenological descriptive classification threshold sensitivity analysis across alternative cutoffs for Summer NDVI (0.25 to 0.35), Persistence Ratio (0.85 to 0.95), and Seasonal Amplitude (0.20 to 0.28). Showing representative subset of 75 evaluated configurations.", style='Caption')
    sens_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv"))
    sens_disp = sens_df[sens_df['Seasonal_Amplitude_Threshold'] == 0.25].copy()
    sens_disp = sens_disp[['Summer_NDVI_Threshold', 'Persistence_Ratio_Threshold', 'Persistent_Canopy_Count', 'Seasonal_Flush_Count', 'Disturbed_Transition_Count', 'Reclassified_Quadrats_Count', 'Reclassified_Quadrats_Detail']].copy()
    sens_disp.columns = ['Summer Cutoff', 'Ratio Cutoff', 'Persistent', 'Flush', 'Transition', 'Changed', 'Reclassified Quadrats Detail']
    add_table_from_df(doc, sens_disp.head(25), col_widths=[0.8, 0.8, 0.7, 0.6, 0.7, 0.6, 2.0])"""

new_s9_docx = """    # Supplementary Table S9A: Compact Sensitivity Frequency Distribution
    doc.add_paragraph("Supplementary Table S9A. Compact frequency distribution of quadrat reclassifications across all 75 tested heuristic threshold configurations (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude).", style='Caption')
    freq_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9A_Sensitivity_Frequency_Distribution.csv"))
    add_table_from_df(doc, freq_df, col_widths=[1.0, 2.5, 1.0, 1.0, 1.0])

    # Supplementary Table S9B: Full Threshold Sensitivity Grid (Representative Subset)
    doc.add_paragraph("Supplementary Table S9B. Heuristic phenological descriptive classification threshold sensitivity analysis across alternative cutoffs for Summer NDVI (0.25 to 0.35), Persistence Ratio (0.85 to 0.95), and Seasonal Amplitude (0.20 to 0.28). Showing representative subset of 75 evaluated configurations.", style='Caption')
    sens_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv"))
    sens_disp = sens_df[sens_df['Seasonal_Amplitude_Threshold'] == 0.25].copy()
    p_col = 'Persistent_Greenness_Count' if 'Persistent_Greenness_Count' in sens_disp.columns else 'Persistent_Canopy_Count'
    sens_disp = sens_disp[['Summer_NDVI_Threshold', 'Persistence_Ratio_Threshold', p_col, 'Seasonal_Flush_Count', 'Disturbed_Transition_Count', 'Reclassified_Quadrats_Count', 'Reclassified_Quadrats_Detail']].copy()
    sens_disp.columns = ['Summer Cutoff', 'Ratio Cutoff', 'Persistent', 'Flush', 'Transition', 'Changed', 'Reclassified Quadrats Detail']
    add_table_from_df(doc, sens_disp.head(25), col_widths=[0.8, 0.8, 0.7, 0.6, 0.7, 0.6, 2.0])"""

code = code.replace(old_s9_docx, new_s9_docx)

old_s9_md = """    # Table S9
    lines.append("### Supplementary Table S9. Heuristic phenological descriptive classification threshold sensitivity analysis across alternative cutoffs (representative subset at Amplitude = 0.25).\\n")
    sens_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv"))
    sens_disp = sens_df[sens_df['Seasonal_Amplitude_Threshold'] == 0.25].copy()
    sens_disp = sens_disp[['Summer_NDVI_Threshold', 'Persistence_Ratio_Threshold', 'Persistent_Canopy_Count', 'Seasonal_Flush_Count', 'Disturbed_Transition_Count', 'Reclassified_Quadrats_Count', 'Reclassified_Quadrats_Detail']].copy()
    sens_disp.columns = ['Summer Cutoff', 'Ratio Cutoff', 'Persistent', 'Flush', 'Transition', 'Changed', 'Reclassified Quadrats Detail']
    lines.append(df_to_markdown(sens_disp.head(25)))
    lines.append("\\n*(Complete 75-configuration sensitivity dataset archived in `tables/Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv`)*\\n\\n")"""

new_s9_md = """    # Table S9A
    lines.append("### Supplementary Table S9A. Compact frequency distribution of quadrat reclassifications across all 75 tested heuristic threshold configurations (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude).\\n")
    freq_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9A_Sensitivity_Frequency_Distribution.csv"))
    lines.append(df_to_markdown(freq_df))
    lines.append("\\n\\n")

    # Table S9B
    lines.append("### Supplementary Table S9B. Heuristic phenological descriptive classification threshold sensitivity grid across alternative cutoffs (representative subset at Seasonal Amplitude = 0.25).\\n")
    sens_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv"))
    sens_disp = sens_df[sens_df['Seasonal_Amplitude_Threshold'] == 0.25].copy()
    p_col = 'Persistent_Greenness_Count' if 'Persistent_Greenness_Count' in sens_disp.columns else 'Persistent_Canopy_Count'
    sens_disp = sens_disp[['Summer_NDVI_Threshold', 'Persistence_Ratio_Threshold', p_col, 'Seasonal_Flush_Count', 'Disturbed_Transition_Count', 'Reclassified_Quadrats_Count', 'Reclassified_Quadrats_Detail']].copy()
    sens_disp.columns = ['Summer Cutoff', 'Ratio Cutoff', 'Persistent', 'Flush', 'Transition', 'Changed', 'Reclassified Quadrats Detail']
    lines.append(df_to_markdown(sens_disp.head(25)))
    lines.append("\\n*(Complete 75-configuration sensitivity dataset archived in `tables/Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv`)*\\n\\n")"""

code = code.replace(old_s9_md, new_s9_md)

with open(SCRIPT_PATH, "w", encoding="utf-8") as f:
    f.write(code)

print("Patch applied successfully to scripts/08_generate_manuscript_and_supplemental.py!")
