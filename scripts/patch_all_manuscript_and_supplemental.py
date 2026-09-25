"""
Comprehensive patch script to update scripts/08_generate_manuscript_and_supplemental.py
for Version 2.4 (Technical & Editorial Reconciliation for External Peer Review).
"""

import os
import re

SCRIPT_PATH = "scripts/08_generate_manuscript_and_supplemental.py"

with open(SCRIPT_PATH, "r", encoding="utf-8") as f:
    text = f.read()

# 1. Update Version in docx and md
text = text.replace(
    'Version: 2.3 (Fourth-Round Comprehensive Resolution) | Date: September 2026',
    'Version: 2.4 (Technical & Editorial Reconciliation for External Peer Review) | Date: September 2026'
)
text = text.replace(
    'Target Journal: Ecological Indicators | Article Type: Research Paper (Major Revision, Version 2.3) | September 2026',
    'Target Journal: Ecological Indicators | Article Type: Research Paper (Minor Revision, Version 2.4) | September 2026'
)
text = text.replace(
    'Shubham Sharma et al. | Ecological Indicators (Major Revision Version 2.3, September 2026)',
    'Shubham Sharma et al. | Ecological Indicators (Minor Revision Version 2.4, September 2026)'
)
text = text.replace(
    '**Journal:** *Ecological Indicators* (Version 2.3 Major Revision) | Date: September 2026',
    '**Journal:** *Ecological Indicators* (Version 2.4 Minor Revision) | Date: September 2026'
)
text = text.replace(
    '**Companion to:** *Ecological Indicators* Manuscript Version 2.3 (September 2026)',
    '**Companion to:** *Ecological Indicators* Manuscript Version 2.4 (September 2026)'
)
text = text.replace(
    'Generating Audited Manuscript and Supplemental Deliverables (Version 2.3)...',
    'Generating Audited Manuscript and Supplemental Deliverables (Version 2.4)...'
)

# 2. Update Abstract in Markdown
old_abs_md = (
    '        "In Zone A, 100% valid terrestrial pixel retention was achieved across all 60 site-date cases, while larger windows exhibited variable "\n'
    '        "valid pixel fractions (98.8% to 99.7%) due to colliery infrastructure and water bodies. Rather than deconvolving isolated strata, "\n'
    '        "satellite trajectories capture integrated vertical surface reflectance across all canopy layers and exposed ground. Quadrats partitioned "\n'
    '        "into three heuristic descriptive classes: (1) Persistent Canopy Signal (Summer dry-season NDVI ≥ 0.30, Winter-to-Green-Season Persistence Ratio ≥ 0.90; "\n'
    '        "UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4); (2) Pronounced Seasonal Flush (Seasonal Amplitude ΔNDVI > 0.25, Summer NDVI < 0.25; OC-Q5); and "\n'
    '        "(3) Disturbed Transition, which includes sites consistent with a possible moisture-related influence near drainage features such as OC-Q2, "\n'
    '        "where winter greenness retention is elevated (persistence ratio = 1.129; winter composite = 0.414, single-date Dec 30 value = 0.453). "\n'
    '        "Threshold sensitivity analysis across 75 parameter configurations indicated central classification stability (preserving 100% of site "\n'
    '        "assignments in 32.0% of configurations and ≤ 1 reclassification in 60.0%), while showing expected sensitivity under more stringent persistence "\n'
    '        "cutoffs. Multi-scale spatial evaluation revealed canopy dilution in remnant underground forest quadrats (mean summer NDVI declining from 0.381 in "\n'
    '        "Zone A to 0.271 in Zone D) as extraction captures colliery infrastructure and clearings, whereas opencast spoil plantations remained "\n'
    '        "spatially stable (0.303 to 0.287). Principal Component Analysis accounted for 81.62% of variance across two primary axes, demonstrating "\n'
    '        "close agreement (score correlation r > 0.94) with an 8-feature sensitivity model excluding derived metrics. Non-parametric contrasts between "\n'
    '        "the sampled underground and opencast quadrats (n=5 per group) yielded non-significant differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, "\n'
    '        "exact p = 0.2222, Cliff\'s delta = +0.520, Hedges\' g = 0.647, approximate 95% analytical Student\'s t CI [-0.850, 2.143]), demonstrating that "\n'
    '        "non-significance reflects low statistical power in a pilot sample rather than ecological equivalence. We provide a complete machine-readable "\n'
    '        "provenance audit and outline requirements for future hierarchical, replicated landscape restoration designs.\\n"'
)

new_abs_md = (
    '        "In Zone A, 100.00% valid terrestrial pixel retention was achieved across all 60 site-date cases (60/60 pixels), while larger windows exhibited "\n'
    '        "valid-pixel fractions of 98.99% to 99.33% (Zone B: 99.33% [1,490/1,500], Zone C: 99.22% [4,822/4,860], Zone D: 98.99% [18,828/19,020]) "\n'
    '        "due to colliery infrastructure and water bodies. Rather than deconvolving isolated strata, satellite trajectories capture integrated top-of-canopy "\n'
    '        "surface reflectance influenced by overstory foliage, subcanopy vegetation, ground cover, litter, soil, and exposed substrate. Quadrats partitioned "\n'
    '        "into three heuristic descriptive classes: (1) Persistent Greenness Signal (Summer dry-season NDVI ≥ 0.30, Winter-to-Green-Season Persistence Ratio ≥ 0.90; "\n'
    '        "UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4; descriptive spectral class); (2) Pronounced Seasonal Flush (Seasonal Amplitude ΔNDVI > 0.25, Summer NDVI < 0.25; OC-Q5); "\n'
    '        "and (3) Disturbed Transition, which includes sites consistent with a possible moisture-related influence near drainage features such as OC-Q2, "\n'
    '        "where winter greenness retention is elevated (persistence ratio = 1.129; winter composite = 0.414, single-date Dec 30 value = 0.453). "\n'
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

if old_abs_md in text:
    text = text.replace(old_abs_md, new_abs_md)
    print("Replaced Abstract in Markdown.")
else:
    print("WARNING: old_abs_md not found in text.")

# 3. Update Section 3.2 in Markdown
old_s32 = (
    '        "In Zone A (nominal 10 m pixel support), all 10 quadrats achieved 100% valid terrestrial pixel retention across all six dates (zero masked pixels across 60 site-date cases; Table 2). "\n'
    '        "Larger multi-scale extraction windows (Zones B–D) had variable valid pixel fractions (98.8% to 99.7%) due to edge inclusions of water bodies, shadows, and colliery infrastructure (Table 3; Table 6).\\n"'
)

new_s32 = (
    '        "In Zone A (nominal 10 m pixel support), 100.00% valid terrestrial pixel retention was achieved across all six dates (zero masked pixels across 60 site-date cases; Table 2). "\n'
    '        "Terrestrial validity percentage is defined as the total number of valid terrestrial pixels (SCL classes 4 [vegetation] and 5 [bare soil]) divided by "\n'
    '        "the theoretical denominator pixel count across all 60 site-date observations (which exactly equals the arithmetic mean of individual site-date validity "\n'
    '        "percentages to two decimal places): Zone A = 100.00% (60/60), Zone B = 99.33% (1,490/1,500), Zone C = 99.22% (4,822/4,860), and Zone D = 98.99% "\n'
    '        "(18,828/19,020), defining a range of 98.99% to 99.33% across multi-scale extraction windows due to edge inclusions of water bodies, shadows, and colliery infrastructure (Table 3; Table 6).\\n"'
)

if old_s32 in text:
    text = text.replace(old_s32, new_s32)
    print("Replaced Section 3.2 in Markdown.")
else:
    print("WARNING: old_s32 not found in text.")

# 4. Update Section 3.3 circular buffer areas in Markdown
old_s33 = (
    '        "- **Zone A (Nominal Quadrat Pixel Support):** Exactly 1 pixel (10 × 10 m = 100 m² = 0.01 ha), defined as the raster pixel containing the quadrat centroid.\\n"\n'
    '        "- **Zone B (Local Neighborhood):** A square 5 × 5 pixel window (50 × 50 m = 2,500 m² = 0.25 ha, 25 pixels total support), centered on the quadrat centroid and enclosing its GPS uncertainty footprint.\\n"\n'
    '        "- **Zone C (Intermediate Stand Buffer):** A circular disc of radius 50 m (81 raster pixels = 8,100 m² = 0.81 ha), capturing the local forest or overburden dump plantation stand.\\n"\n'
    '        "- **Zone D (Contextual Landscape Buffer):** A circular disc of radius 100 m (317 raster pixels = 31,700 m² = 3.17 ha), capturing the broader colliery landscape matrix.\\n"'
)

new_s33 = (
    '        "- **Zone A (Nominal Quadrat Pixel Support):** Exactly 1 pixel (10 × 10 m = 100 m² = 0.01 ha), defined as the raster pixel containing the quadrat centroid; field-plot overlap is variable.\\n"\n'
    '        "- **Zone B (Local Neighborhood):** A square 5 × 5 pixel window (50 × 50 m = 2,500 m² = 0.25 ha, 25 pixels total support), centered on the quadrat centroid and enclosing its GPS uncertainty footprint.\\n"\n'
    '        "- **Zone C (Intermediate Stand Buffer):** A circular disc of radius 50 m (theoretical geometric area = 7,854 m² [π × 50²]; raster support area = 8,100 m² across 81 pixels), capturing the local forest or overburden dump plantation stand.\\n"\n'
    '        "- **Zone D (Contextual Landscape Buffer):** A circular disc of radius 100 m (theoretical geometric area = 31,416 m² [π × 100²]; raster support area = 31,700 m² across 317 pixels), capturing the broader colliery landscape matrix.\\n"'
)

if old_s33 in text:
    text = text.replace(old_s33, new_s33)
    print("Replaced Section 3.3 in Markdown.")
else:
    print("WARNING: old_s33 not found in text.")

# 5. Update Section 3.5 in Markdown
old_s35 = (
    '        "1. **Persistent Canopy Signal:** Summer dry-season NDVI ≥ 0.30 AND Winter-to-Green-Season Persistence Ratio ≥ 0.90. This class characterizes "\n'
    '        "sites where established woody canopy cover maintains high foliar greenness and suppresses dry-season seasonal amplitude.\\n"\n'
    '        "2. **Pronounced Seasonal Flush:** Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "\n'
    '        "disturbed ground where an open canopy permits intense monsoonal herbaceous and grass growth followed by rapid post-monsoon senescence.\\n"\n'
    '        "3. **Disturbed Transition:** Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "\n'
    '        "micro-hydrological conditions (e.g., mine water drainage features).\\n\\n"\n'
    '        "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "\n'
    '        "analysis across 75 parameter configurations (Summer NDVI cutoff 0.25–0.35, Persistence Ratio cutoff 0.85–0.95, Amplitude cutoff 0.20–0.28; "\n'
    '        "Supplementary Table S9). Across all 75 tested configurations, 24 (32.0%) preserved 100% of baseline site assignments, and 45 (60.0%) produced "\n'
    '        "at most one reclassification. Reclassifications occurred under more stringent persistence thresholds (≥ 0.92, which reclassifies OC-Q1 and UG-Q2) "\n'
    '        "or elevated summer thresholds (≥ 0.32, which reclassifies UG-Q1). We emphasize that these groupings are descriptive heuristic summaries rather than general ecological laws.\\n"'
)

new_s35 = (
    '        "1. **Persistent Greenness Signal:** Summer dry-season NDVI ≥ 0.30 AND Winter-to-Green-Season Persistence Ratio ≥ 0.90 (descriptive spectral class; "\n'
    '        "also termed High Dry-Season Spectral Retention). This class characterizes sites where established woody cover maintains high foliar greenness and suppresses "\n'
    '        "dry-season seasonal amplitude, without establishing canopy dominance in the absence of independent understory measurements.\\n"\n'
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

if old_s35 in text:
    text = text.replace(old_s35, new_s35)
    print("Replaced Section 3.5 in Markdown.")
else:
    print("WARNING: old_s35 not found in text.")

# 6. Update Section 3.6 in Markdown
old_s36 = (
    '        "Given the small sample size of the published study (n = 10; 5 underground vs. 5 opencast), statistical analyses were conducted using "\n'
    '        "exact non-parametric permutation tests alongside standardized effect sizes. Contrasts between sampled underground and opencast quadrats "\n'
    '        "were evaluated using the exact two-sided Mann-Whitney U test (`scipy.stats.mannwhitneyu`, `method=\'exact\'`), Cliff\'s delta, and small-sample "\n'
    '        "corrected Hedges\' g. In the tabular results (Table 5) and audit records (Table S7), we explicitly report both the first-sample U statistic "\n'
    '        "(U₁, corresponding to Underground) and the smaller U statistic ($U_{\\\\text{min}} = \\\\min(U_1, U_2)$), resolving potential ambiguities arising from software conventions.\\n\\n"\n'
    '        "A standardized mean difference (Hedges\' g) was paired with non-parametric rank tests because Mann-Whitney U evaluates whether one distribution "\n'
    '        "stochastically dominates another without distributional assumptions, whereas Hedges\' g provides a standardized, scale-free descriptive effect size "\n'
    '        "with an analytical confidence interval to facilitate future meta-analyses and sample-size planning for replicated restoration designs. In accordance "\n'
    '        "with Hedges & Olkin (1985), approximate 95% confidence intervals for Hedges\' g were calculated using the Student\'s t distribution with $\\\\text{df} = n_1 + n_2 - 2 = 8$ "\n'
    '        "degrees of freedom: $\\\\text{CI} = g \\\\pm t(0.025, 8) \\\\times \\\\text{SE}(g)$, where critical $t(0.025, 8) = 2.3060$, $\\\\text{SE}(g) = \\\\sqrt{[(n_1 + n_2)/(n_1 \\\\times n_2)] + [g^2 / (2(n_1 + n_2))]}$, "\n'
    '        "pooled standard deviation $s_{\\\\text{pooled}} = \\\\sqrt{[((n_1 - 1)s_1^2 + (n_2 - 1)s_2^2) / \\\\text{df}]}$, and small-sample correction factor $J = 1 - [3 / (4(n_1 + n_2) - 9)] = 28/31 \\\\approx 0.9032$. "\n'
    '        "Contrast direction is defined as Underground minus Opencast (UG - OC). Group medians and interquartile ranges (IQR) are reported alongside means and standard deviations.\\n\\n"'
)

new_s36 = (
    '        "Given the small sample size of the published study (n = 10; 5 underground vs. 5 opencast), statistical analyses were conducted using "\n'
    '        "exact non-parametric permutation tests alongside standardized effect sizes. Contrasts between sampled underground and opencast quadrats "\n'
    '        "were evaluated using SciPy’s exact two-sided Mann-Whitney U procedure (`scipy.stats.mannwhitneyu`, `alternative=\'two-sided\'`, `method=\'exact\'`) "\n'
    '        "with no tie correction required for the reported values. An independent exhaustive enumeration of all 252 label permutations (10 choose 5) confirmed identical "\n'
    '        "results (56/252 = 2/9 = 0.2222), demonstrating exact agreement between both calculations. In the tabular results (Table 5) and audit records (Table S7), "\n'
    '        "we explicitly report both the first-sample U statistic (U₁, corresponding to Underground) and the smaller U statistic ($U_{\\\\text{min}} = \\\\min(U_1, U_2)$), "\n'
    '        "resolving potential ambiguities arising from software conventions.\\n\\n"\n'
    '        "A standardized mean difference (Hedges\' g) was paired with non-parametric rank tests because Mann-Whitney U evaluates whether one distribution "\n'
    '        "stochastically dominates another without distributional assumptions, whereas Hedges\' g provides a standardized, scale-free descriptive effect size "\n'
    '        "to facilitate future meta-analyses and sample-size planning for replicated restoration designs. We report an approximate interval obtained by "\n'
    '        "multiplying the estimated standard error of Hedges’ g by the $t(8, 0.975)$ critical value ($t = 2.3060$); this interval is descriptive and not intended "\n'
    '        "to provide exact small-sample coverage (Hedges & Olkin, 1985). Specifically, approximate 95% confidence intervals were calculated using the Student\'s t "\n'
    '        "critical value with $\\\\text{df} = n_1 + n_2 - 2 = 8$: $\\\\text{CI} = g \\\\pm t(0.025, 8) \\\\times \\\\text{SE}(g)$, where critical $t(0.025, 8) = 2.3060$, "\n'
    '        "$\\\\text{SE}(g) = \\\\sqrt{[(n_1 + n_2)/(n_1 \\\\times n_2)] + [g^2 / (2(n_1 + n_2))]}$, pooled standard deviation $s_{\\\\text{pooled}} = \\\\sqrt{[((n_1 - 1)s_1^2 + (n_2 - 1)s_2^2) / \\\\text{df}]}$, "\n'
    '        "and small-sample correction factor $J = 1 - [3 / (4(n_1 + n_2) - 9)] = 28/31 \\\\approx 0.9032$. Contrast direction is defined as Underground minus Opencast (UG - OC). "\n'
    '        "Group medians and interquartile ranges (IQR) are reported alongside means and standard deviations.\\n\\n"'
)

if old_s36 in text:
    text = text.replace(old_s36, new_s36)
    print("Replaced Section 3.6 in Markdown.")
else:
    print("WARNING: old_s36 not found in text.")

# 7. Update Section 4.4 in Markdown
old_s44 = (
    '        "- **Persistent Canopy Signal:** Five quadrats (UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4) met the baseline criteria of Summer dry-season NDVI ≥ 0.30 and "\n'
    '        "Winter-to-Green-Season persistence ratio ≥ 0.90. These sites are dominated by dense overstory crown cover (*Shorea robusta*, *Terminalia elliptica*) "\n'
    '        "or stabilized woody reclamation stands (*Eucalyptus* plantation in OC-Q1, *Holarrhena* thicket in OC-Q4) that maintain green foliage well into the dry season.\\n"\n'
    '        "- **Pronounced Seasonal Flush:** OC-Q5 displayed the highest seasonal amplitude in the study (ΔNDVI = 0.285; Summer NDVI = 0.214 -> Green Season "\n'
    '        "NDVI = 0.499; single-date peak 0.749 on Oct 06) and Summer NDVI < 0.25. The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush "\n'
    '        "followed by rapid post-monsoon senescence.\\n"\n'
    '        "- **Disturbed Transition:** Four quadrats occupied intermediate positions: UG-Q3 exhibited high greenness but very low seasonal amplitude "\n'
    '        "(ΔNDVI = 0.075), reflecting dense climax Sal canopy that exchanges leaves rapidly in spring; UG-Q5 (ΔNDVI = 0.192, Summer NDVI = 0.265) represents "\n'
    '        "a thinned edge forest; OC-Q3 (ΔNDVI = 0.158, Summer NDVI = 0.189, Winter = 0.287) represents bare rocky overburden with sparse planted boles; and OC-Q2 "\n'
    '        "(persistence ratio = 1.129; Winter composite = 0.414, Green Season composite = 0.366) represents a site where localized winter greenness retention "\n'
    '        "(single-date Dec 30 NDVI reaching 0.453) is consistent with a possible moisture-related influence near an adjacent drainage feature.\\n\\n"\n'
    '        "The explicit threshold sensitivity analysis across 75 parameter configurations (Supplementary Table S9) demonstrated that the baseline classification "\n'
    '        "was stable across a central range of thresholds (preserving 100% of site assignments in 32.0% of tested configurations and ≤ 1 reclassification in 60.0%), "\n'
    '        "but changed for more stringent persistence or summer-NDVI thresholds. Specifically, when the persistence cutoff was elevated to ≥ 0.92, OC-Q1 (ratio 0.908) "\n'
    '        "and UG-Q2 (ratio 0.901) were reclassified to Disturbed Transition; tightening to ≥ 0.95 reclassified UG-Q4 (ratio 0.921); and elevating the Summer NDVI cutoff "\n'
    '        "to ≥ 0.32 reclassified UG-Q1 (summer 0.317). This sensitivity underscores that these syndromic groupings are operational descriptive tools rather than rigid ecological states.\\n"'
)

new_s44 = (
    '        "- **Persistent Greenness Signal:** Five quadrats (UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4) met the baseline criteria of Summer dry-season NDVI ≥ 0.30 and "\n'
    '        "Winter-to-Green-Season persistence ratio ≥ 0.90 (descriptive spectral class). These sites are characterized by established woody cover (*Shorea robusta*, "\n'
    '        "*Terminalia elliptica*, *Eucalyptus tereticornis*, or *Holarrhena pubescens*) that maintains high foliar reflectance into the dry season without establishing "\n'
    '        "canopy dominance in the absence of independent understory measurements.\\n"\n'
    '        "- **Pronounced Seasonal Flush:** OC-Q5 displayed the highest seasonal amplitude in the study (ΔNDVI = 0.285; Summer NDVI = 0.214 -> Green Season "\n'
    '        "NDVI = 0.499; single-date peak 0.749 on Oct 06) and Summer NDVI < 0.25. The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush "\n'
    '        "followed by rapid post-monsoon senescence.\\n"\n'
    '        "- **Disturbed Transition:** Four quadrats occupied intermediate positions: UG-Q3 exhibited high greenness but very low seasonal amplitude "\n'
    '        "(ΔNDVI = 0.075), reflecting dense climax Sal canopy that exchanges leaves rapidly in spring; UG-Q5 (ΔNDVI = 0.192, Summer NDVI = 0.265) represents "\n'
    '        "a thinned edge forest; OC-Q3 (ΔNDVI = 0.158, Summer NDVI = 0.189, Winter = 0.287) represents bare rocky overburden with sparse planted boles; and OC-Q2 "\n'
    '        "(persistence ratio = 1.129; Winter composite = 0.414, Green Season composite = 0.366) represents a site where localized winter greenness retention "\n'
    '        "(single-date Dec 30 NDVI reaching 0.453) is consistent with a possible moisture-related influence near an adjacent drainage feature.\\n\\n"\n'
    '        "The explicit threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations; "\n'
    '        "Supplementary Table S9A and Table S9B) demonstrated that the baseline classification was stable across a central range of thresholds (preserving 100% of site assignments "\n'
    '        "in 32.0% of tested configurations and ≤ 1 reclassification in 60.0%), but changed for more stringent persistence or summer-NDVI thresholds. Reclassifications were driven "\n'
    '        "primarily by persistence and summer-NDVI cutoffs (e.g., persistence ≥ 0.92 reclassified OC-Q1 and UG-Q2; summer NDVI ≥ 0.32 reclassified UG-Q1), whereas varying seasonal "\n'
    '        "amplitude from 0.20 to 0.28 produced zero reclassifications across all 75 configurations. This sensitivity underscores that these syndromic groupings are operational "\n'
    '        "descriptive spectral classes rather than rigid ecological states.\\n"'
)

if old_s44 in text:
    text = text.replace(old_s44, new_s44)
    print("Replaced Section 4.4 in Markdown.")
else:
    print("WARNING: old_s44 not found in text.")

with open(SCRIPT_PATH, "w", encoding="utf-8") as f:
    f.write(text)

print("Saved updated scripts/08_generate_manuscript_and_supplemental.py")
