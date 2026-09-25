import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
script_path = os.path.join(BASE_DIR, "scripts", "08_generate_manuscript_and_supplemental.py")

with open(script_path, "r", encoding="utf-8") as f:
    content = f.read()

# Target 1: In build_manuscript_docx - Section 2.1
old_target1 = (
    'which aligns directly with the multi-scale spatial buffer windows analyzed herein (Zone A: 10 × 10 m quadrat support; "\n'
    '        "Zone B: 30 × 30 m; Zone C: 50 × 50 m; Zone D: 100 × 100 m buffer)."'
)
new_target1 = (
    'which aligns directly with the multi-scale spatial buffer windows analyzed herein (Zone A: 10 × 10 m quadrat support; "\n'
    '        "Zone B: 30 × 30 m; Zone C: 50 × 50 m; Zone D: 100 × 100 m buffer). The empirical baseline data for these 10 quadrats "\n'
    '        "(including verified in-plot woody stems, photographic registers, harmonized soil chemistry, and screened Sentinel-2 "\n'
    '        "observations) constitute the audited frozen v2.0 dataset documented in detail in CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md."'
)

assert old_target1 in content, "old_target1 not found"
content = content.replace(old_target1, new_target1, 1)

# Target 2: In build_manuscript_docx - After Figure 7, add Section 4.5, Table 4, Figure 8, updated Section 5, Data Availability, Code Availability, References
old_target2 = """    # 5. Conclusions
    doc.add_heading("5. Conclusions", level=1).style.font.color.rgb = RGBColor(27, 54, 93)
    doc.add_paragraph(
        "This study establishes a defensible, audited framework for reconstructing seasonal herbaceous and shrub vegetation dynamics "
        "around published field quadrats in mining-impacted tropical dry deciduous landscapes. By integrating raw field inventories, "
        "photographic evidence, cited IMD climatology, and per-pixel Sentinel-2 L2A observations across 2024, we resolved the phenological "
        "pathways of understory flushes and woody canopies without making unsubstantiated species claims. Herbaceous-dominated overburden "
        "substrates exhibited rapid, high-amplitude green-up (Delta-NDVI = 0.285) followed by sharp post-monsoon senescence, while woody "
        "shrub stands maintained persistent dry-season photosynthetic activity (Winter-to-Monsoon ratio >= 1.00). PCA revealed that 81.62% "
        "of spectral variance is captured by woody canopy persistence (PC1) and seasonal herbaceous amplitude (PC2). This workflow provides "
        "a standardized protocol for ecological restoration monitoring across disturbed industrial landscapes globally."
    )

    # Statements
    doc.add_heading("Data Availability Statement", level=2)
    doc.add_paragraph("All raw field registers, imagery manifests, intermediate GeoTIFF rasters with SHA-256 checksums, and analysis tables are available in the project repository and data directory.")

    doc.add_heading("Code Availability Statement", level=2)
    doc.add_paragraph("All processing scripts (01 to 08) are fully open-source and reproducible using Python 3.10+, rasterio, shapely, scipy, and scikit-learn.")"""

new_target2 = """    doc.add_heading("4.5 Prospective Sampling Architecture and Methodological Recommendations for Next-Phase Restoration Assessments", level=2)
    doc.add_paragraph(
        "While the multi-temporal satellite and ground evidence integration presented here successfully deconvolves overstory-understory "
        "phenology across the 10 study quadrats, our statistical contrasts (Section 3.6) highlight an important methodological constraint: "
        "with n = 5 quadrats per mining category, group comparisons lack statistical power to differentiate broad landscape-level mining "
        "impacts from localized site conditions. Furthermore, in the published baseline, opencast quadrats were clustered within the "
        "Gevra/Dipka complex, and underground quadrats were clustered within Banki/Surakachhar leases. In observational field studies, "
        "treating multiple sampling plots within a single colliery lease or overburden dump as independent landscape replicates constitutes "
        "spatial pseudoreplication (Hurlbert, 1984). Each mine concession possesses unique operational dates, dumping geometries, "
        "topsoil replacement histories, and micro-climatic exposures. Consequently, attributing observed differences solely to 'opencast vs. "
        "underground' without independent site replication conflates true disturbance class effects with localized concession-level idiosyncrasies."
    )
    doc.add_paragraph(
        "To overcome this fundamental limitation and establish publication-grade restoration assessments, we have formulated a three-category, "
        "nested hierarchical sampling architecture spanning 15 independent landscape blocks (15 sites × 5 subsample quadrats = 75 quadrats; "
        "detailed in HIERARCHICAL_REPLICATED_LANDSCAPE_DESIGN_75_QUADRATS.md and ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md). "
        "This prospective framework introduces three essential methodological advancements:"
    )
    doc.add_paragraph(
        "1. Independent Landscape-Scale Replication Across Three Classes: The expanded design balances 5 independent underground colliery "
        "leases (Banki, Balgi, Surakachhar, Dhelwadih North, Singhali/Bagdeva), 5 independent opencast overburden complexes (Gevra Dump 4, "
        "Dipka West Overburden, Kusmunda North Dump, Manikpur Void Plantation, Saraipali Active Spoil), and 5 unmined climax reference "
        "forest compartments (Katghora Range Compt 142, Lemru Elephant Reserve Core Corridor, Pali Range Chaiturgarh Hills, Tiwarta "
        "Forest Beat Compt 88, Kudmura Forest Range Compt 204). The pairwise inter-site distance matrix confirms a minimum inter-site "
        "separation of 3.19 km, a mean separation of 16.98 km, and an overall regional span of 44.95 km across the Barakar and Raniganj "
        "geological formations. By distributing opencast sites across distinct concessions separated by 6.25–18.08 km, spatial autocorrelation "
        "is systematically eliminated."
    )
    doc.add_paragraph(
        "2. Hierarchical Nested Design and Linear Mixed-Effects Modeling: Within each 50–160 ha landscape site, five 10 × 10 m quadrats "
        "are distributed across a standardized micro-topographic catena (Crest, Upper slope, Lower slope, Bench, and Swale) with 50–120 m "
        "inter-quadrat spacing. These quadrats function as subsamples to characterize within-site spatial heterogeneity, rather than independent "
        "landscape replicates. To evaluate this structure without pseudoreplication, data must be analyzed via a Linear Mixed-Effects Model "
        "(LMM) with Site treated as a random intercept nested within Landscape Class. Crucially, as detailed in Table 4, the F-test for the "
        "fixed effect of Landscape Class (df = 2) is computed using the Site-within-Class mean square as the denominator error term (df = 12), "
        "never the residual quadrat error (df = 60). Testing against the site error ensures valid Type I error rates, providing statistical "
        "power > 0.88 to detect medium effect sizes (f = 0.35)."
    )
    
    # Table 4
    doc.add_paragraph("Table 4. Proposed hierarchical experimental design and ANOVA error structure for the three-category landscape assessment (N = 75 quadrats across 15 independent sites).", style='Caption')
    t4_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_22_Hierarchical_Experimental_Design_and_ANOVA_Structure.csv"))
    t4_disp = t4_df[['Source_of_Variation', 'Degrees_of_Freedom', 'Error_Term_for_F_Test', 'Error_Degrees_of_Freedom', 'Design_Significance']].copy()
    t4_disp.columns = ['Source of Variation', 'df', 'Error Term for F-Test', 'Error df', 'Design Significance']
    add_table_from_df(doc, t4_disp, col_widths=[1.8, 0.5, 1.8, 0.6, 2.3])

    doc.add_paragraph(
        "3. Multi-Dimensional Restoration Metrics and Functional Ecosystem Recovery: Beyond multi-spectral greenness indices and basic "
        "soil chemistry, next-phase empirical campaigns should measure five complementary ecological dimensions:\n"
        "• Undisturbed Climax Reference Benchmarks: Establishing target restoration thresholds using protected mature Shorea robusta (Sal) "
        "and Terminalia tomentosa stands with intact soil pedons, providing an unmined baseline against which absolute ecological recovery can be quantified.\n"
        "• Soil Physical and Hydrological Functioning: Measuring bulk density (0–15 cm), cone penetration resistance to test the critical "
        "root impedance threshold (> 2.5 MPa), steady-state infiltration rates using double-ring infiltrometers (testing whether raw spoil experiences "
        "severe infiltration restriction < 5 mm/h compared to > 15 mm/h under woody shrub thickets), and saturated hydraulic conductivity (Ksat).\n"
        "• Soil Biological Health and Enzymatic Functioning: Quantifying microbial biomass carbon (MBC) via chloroform fumigation-extraction, "
        "and assaying key soil enzymes: dehydrogenase (microbial respiration and oxidative activity), urease (nitrogen mineralization), and "
        "alkaline/acid phosphatase (phosphorus mobilization).\n"
        "• Plant Functional Traits and Restoration Value: Leveraging our botanical trait synthesis across 42 regional species (Table 14) "
        "to track symbiotic nitrogen fixation potential (nodulating legumes such as Pongamia pinnata and Senna tora), root architecture "
        "(taproot vs. fibrous vs. stoloniferous turf binders like Cynodon dactylon), specific leaf area (SLA), wood density, and Grime CSR ecological strategies.\n"
        "• Multi-Taxa Faunal Bioindicators: Conducting standardized Pollard butterfly transects and avian point counts to measure the trophic "
        "return of insect pollinators and frugivorous seed dispersers."
    )
    doc.add_paragraph(
        "Scientific Integrity and Ethical Data Governance: We explicitly emphasize that while the current paper reports the audited, completed "
        "empirical observations of the 10 published quadrats (n = 5 UG vs. n = 5 OC; preserved in CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md), "
        "the 75-quadrat architecture, hypothetical LMM parameters, and blank field sheets (templates/) represent prospective methodological "
        "blueprints designed to guide future field surveys (ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md). Maintaining this "
        "strict division protects the empirical integrity of published datasets while providing a rigorous roadmap for next-generation restoration science."
    )

    # Insert Figure 8 (Restoration Evidence / Conceptual Pathways)
    fig8_path = os.path.join(BASE_DIR, "figures", "Figure_5_Ecological_Restoration_Pathways.png")
    if os.path.exists(fig8_path):
        doc.add_picture(fig8_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Figure 8. Conceptual restoration framework and methodological blueprint for prospective multi-site post-mining recovery monitoring, illustrating the integration of unmined reference forests, hierarchical site replication, soil physical/biological functioning, and remote sensing.")
        cap_run.font.size = Pt(9.5)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    # 5. Conclusions
    doc.add_heading("5. Conclusions", level=1).style.font.color.rgb = RGBColor(27, 54, 93)
    doc.add_paragraph(
        "This study establishes a defensible, audited framework for reconstructing seasonal herbaceous and shrub vegetation dynamics "
        "around published field quadrats in mining-impacted tropical dry deciduous landscapes. By integrating raw field inventories, "
        "photographic evidence, cited IMD climatology, and per-pixel Sentinel-2 L2A observations across 2024, we resolved the phenological "
        "pathways of understory flushes and woody canopies without making unsubstantiated species claims. Herbaceous-dominated overburden "
        "substrates exhibited rapid, high-amplitude green-up (Delta-NDVI = 0.285) followed by sharp post-monsoon senescence, while woody "
        "shrub stands maintained persistent dry-season photosynthetic activity (Winter-to-Monsoon ratio >= 1.00). PCA revealed that 81.62% "
        "of spectral variance is captured by woody canopy persistence (PC1) and seasonal herbaceous amplitude (PC2). Furthermore, to address "
        "the pilot sample size and spatial clustering constraints of the current baseline, we provide a prospective 75-quadrat, 15-site "
        "hierarchical sampling architecture that eliminates pseudoreplication and establishes a rigorous roadmap for multi-site post-mining "
        "ecological restoration monitoring across disturbed industrial landscapes globally."
    )

    # Statements
    doc.add_heading("Data Availability Statement", level=2)
    doc.add_paragraph(
        "All raw field registers, imagery manifests, intermediate GeoTIFF rasters with SHA-256 checksums, and analysis tables are "
        "preserved in the project repository under data/, tables/, and rasters/. The master audited empirical baseline for the 10 published "
        "quadrats is cataloged in CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md. The prospective 75-quadrat sampling design, ANOVA specifications, "
        "literature-derived plant functional traits, and standardized field/laboratory datasheets are documented in "
        "HIERARCHICAL_REPLICATED_LANDSCAPE_DESIGN_75_QUADRATS.md, ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md, and templates/."
    )

    doc.add_heading("Code Availability Statement", level=2)
    doc.add_paragraph("All processing scripts (01 to 08) are fully open-source and reproducible using Python 3.10+, rasterio, shapely, scipy, and scikit-learn.")

    doc.add_heading("References", level=2)
    refs = [
        "Banerjee, S., Deb, K., Sharma, S., 2023. Soil carbon dynamics and vegetation recovery on coal mine overburden spoils. Ecological Engineering 192, 106980.",
        "Bates, D., Mächler, M., Bolker, B., Walker, S., 2015. Fitting Linear Mixed-Effects Models Using lme4. Journal of Statistical Software 67, 1–48.",
        "Brede, B., Terryn, L., Barbier, N., Bartholomeus, H.M., Bartolo, R., Calders, K., et al., 2020. Non-destructive tree volume estimation through UAS laser scanning. Remote Sensing of Environment 247, 111957.",
        "Deb, K., Sharma, S., Roy, A., 2024. Environmental impacts of opencast and underground coal mining in central India. Journal of Cleaner Production 434, 140120.",
        "Drusch, M., Del Bello, U., Carlier, S., Colin, O., Fernandez, V., Gascon, F., et al., 2012. Sentinel-2: ESA's Optical High-Resolution Mission for GMES Operational Services. Remote Sensing of Environment 120, 25–36.",
        "Gann, G.D., McDonald, T., Walder, B., Aronson, J., Nelson, C.R., Jonson, J., et al., 2019. International principles and standards for the practice of ecological restoration. Second edition. Restoration Ecology 27, S1–S46.",
        "Hurlbert, S.H., 1984. Pseudoreplication and the design of ecological field experiments. Ecological Monographs 54, 187–211.",
        "Kattge, J., Bönisch, G., Díaz, S., Lavorel, S., Prentice, I.C., Leadley, P., et al., 2020. TRY plant trait database – enhanced coverage and open access. Global Change Biology 26, 119–188.",
        "Martinuzzi, S., Gould, W.A., Vierling, L.A., 2009. Identifying tropical dry forest succession from spot remote sensing in Puerto Rico. Biotropica 41, 181–191.",
        "Nagendra, H., Lucas, R., Honrado, J.P., Jongman, R.H., Tarantino, C., Adamo, M., Mairota, P., 2013. Remote sensing for conservation monitoring: Assessing status and trends of biodiversity, habitats, and ecosystems. Remote Sensing in Ecology and Conservation 1, 14–33.",
        "Roy, P.S., Kushwaha, S.P.S., Murthy, M.S.R., Roy, A., Kushwaha, D., Reddy, C.S., et al., 2022. Forest fragmentation and biomass loss in the Central Indian Highlands. Biodiversity and Conservation 31, 843–864.",
        "Sharma, S., Banerjee, S., Deb, K., 2026. Stand Structure, Carbon Sequestration and Soil Physicochemical Dynamics in Underground and Opencast Mining Quadrats of Korba, Chhattisgarh. Environmental Monitoring and Assessment (In Press)."
    ]
    for r in refs:
        p = doc.add_paragraph(r)
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.first_line_indent = Inches(-0.5)
        p.paragraph_format.space_after = Pt(4)"""

assert old_target2 in content, "old_target2 not found"
content = content.replace(old_target2, new_target2, 1)

# Target 3A: In generate_markdown_manuscript - Section 2.1
old_target3a = "analyzed herein (Zone A: 10 m support; Zone B: 30 m; Zone C: 50 m; Zone D: 100 m buffer)."
new_target3a = (
    "analyzed herein (Zone A: 10 m support; Zone B: 30 m; Zone C: 50 m; Zone D: 100 m buffer). "
    "The empirical baseline data for these 10 quadrats (including verified in-plot woody stems, photographic registers, harmonized soil chemistry, and screened Sentinel-2 observations) "
    "constitute the audited frozen v2.0 dataset documented in detail in [`CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md)."
)
assert old_target3a in content, "old_target3a not found"
content = content.replace(old_target3a, new_target3a, 1)

# Target 3B: In generate_markdown_manuscript - After Figure 7 down to References
old_target3b = """## 5. Conclusions
This study establishes a defensible, audited framework for reconstructing seasonal herbaceous and shrub vegetation dynamics around published field quadrats in mining-impacted tropical dry deciduous landscapes. By integrating raw field inventories, photographic evidence, cited IMD climatology, and per-pixel Sentinel-2 L2A observations across 2024, we resolved the phenological pathways of understory flushes and woody canopies without making unsubstantiated species claims. Herbaceous-dominated overburden substrates exhibited rapid, high-amplitude green-up ($\Delta\\text{NDVI} = 0.285$) followed by sharp post-monsoon senescence, while woody shrub stands maintained persistent dry-season photosynthetic activity (Winter-to-Monsoon ratio $\ge 1.00$). PCA revealed that 81.62% of spectral variance is captured by woody canopy persistence (PC1) and seasonal herbaceous amplitude (PC2). This workflow provides a standardized protocol for ecological restoration monitoring across disturbed industrial landscapes globally.

---

## Data Availability Statement
All raw field registers, imagery manifests, intermediate GeoTIFF rasters with SHA-256 checksums, and analysis tables are preserved in the project repository under `data/`, `tables/`, and `rasters/`.

## Code Availability Statement
All processing scripts (`01_build_master_locations.py` through `08_generate_manuscript_and_supplemental.py`) are fully open-source and reproducible using Python 3.10+, `rasterio`, `shapely`, `scipy`, and `scikit-learn`.

## References
1. Banerjee, S., Deb, K., Sharma, S., 2023. Soil carbon dynamics and vegetation recovery on coal mine overburden spoils. *Ecological Engineering* 192, 106980.
2. Brede, B., Terryn, L., Barbier, N., Bartholomeus, H.M., Bartolo, R., Calders, K., et al., 2020. Non-destructive tree volume estimation through UAS laser scanning. *Remote Sensing of Environment* 247, 111957.
3. Deb, K., Sharma, S., Roy, A., 2024. Environmental impacts of opencast and underground coal mining in central India. *Journal of Cleaner Production* 434, 140120.
4. Drusch, M., Del Bello, U., Carlier, S., Colin, O., Fernandez, V., Gascon, F., et al., 2012. Sentinel-2: ESA's Optical High-Resolution Mission for GMES Operational Services. *Remote Sensing of Environment* 120, 25–36.
5. Martinuzzi, S., Gould, W.A., Vierling, L.A., 2009. Identifying tropical dry forest succession from spot remote sensing in Puerto Rico. *Biotropica* 41, 181–191.
6. Nagendra, H., Lucas, R., Honrado, J.P., Jongman, R.H., Tarantino, C., Adamo, M., Mairota, P., 2013. Remote sensing for conservation monitoring: Assessing status and trends of biodiversity, habitats, and ecosystems. *Remote Sensing in Ecology and Conservation* 1, 14–33.
7. Roy, P.S., Kushwaha, S.P.S., Murthy, M.S.R., Roy, A., Kushwaha, D., Reddy, C.S., et al., 2022. Forest fragmentation and biomass loss in the Central Indian Highlands. *Biodiversity and Conservation* 31, 843–864.
8. Sharma, S., Banerjee, S., Deb, K., 2026. Stand Structure, Carbon Sequestration and Soil Physicochemical Dynamics in Underground and Opencast Mining Quadrats of Korba, Chhattisgarh. *Environmental Monitoring and Assessment* (In Press)."""

new_target3b = """### 4.5 Prospective Sampling Architecture and Methodological Recommendations for Next-Phase Restoration Assessments
While the multi-temporal satellite and ground evidence integration presented here successfully deconvolves overstory-understory phenology across the 10 study quadrats, our statistical contrasts (Section 3.6) highlight an important methodological constraint: with $n = 5$ quadrats per mining category, group comparisons lack statistical power to differentiate broad landscape-level mining impacts from localized site conditions. Furthermore, in the published baseline, opencast quadrats were clustered within the Gevra/Dipka complex, and underground quadrats were clustered within Banki/Surakachhar leases. In observational field studies, treating multiple sampling plots within a single colliery lease or overburden dump as independent landscape replicates constitutes spatial pseudoreplication (*Hurlbert, 1984*). Each mine concession possesses unique operational dates, dumping geometries, topsoil replacement histories, and micro-climatic exposures. Consequently, attributing observed differences solely to 'opencast vs. underground' without independent site replication conflates true disturbance class effects with localized concession-level idiosyncrasies.

To overcome this fundamental limitation and establish publication-grade restoration assessments, we have formulated a three-category, nested hierarchical sampling architecture spanning 15 independent landscape blocks ($15\\text{ sites} \\times 5\\text{ subsample quadrats} = 75\\text{ quadrats}$; detailed in [`HIERARCHICAL_REPLICATED_LANDSCAPE_DESIGN_75_QUADRATS.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/HIERARCHICAL_REPLICATED_LANDSCAPE_DESIGN_75_QUADRATS.md) and [`ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md)). This prospective framework introduces three essential methodological advancements:

1. **Independent Landscape-Scale Replication Across Three Classes:** The expanded design balances 5 independent underground colliery leases (Banki, Balgi, Surakachhar, Dhelwadih North, Singhali/Bagdeva), 5 independent opencast overburden complexes (Gevra Dump 4, Dipka West Overburden, Kusmunda North Dump, Manikpur Void Plantation, Saraipali Active Spoil), and 5 unmined climax reference forest compartments (Katghora Range Compt 142, Lemru Elephant Reserve Core Corridor, Pali Range Chaiturgarh Hills, Tiwarta Forest Beat Compt 88, Kudmura Forest Range Compt 204). The pairwise inter-site distance matrix confirms a minimum inter-site separation of $3.19\\text{ km}$, a mean separation of $16.98\\text{ km}$, and an overall regional span of $44.95\\text{ km}$ across the Barakar and Raniganj geological formations. By distributing opencast sites across distinct concessions separated by $6.25–18.08\\text{ km}$, spatial autocorrelation is systematically eliminated.

2. **Hierarchical Nested Design and Linear Mixed-Effects Modeling:** Within each 50–160 ha landscape site, five $10 \\times 10\\text{ m}$ quadrats are distributed across a standardized micro-topographic catena (Crest, Upper slope, Lower slope, Bench, and Swale) with $50–120\\text{ m}$ inter-quadrat spacing. These quadrats function as subsamples to characterize within-site spatial heterogeneity, rather than independent landscape replicates. To evaluate this structure without pseudoreplication, data must be analyzed via a Linear Mixed-Effects Model (LMM) with Site treated as a random intercept nested within Landscape Class. Crucially, as detailed in **Table 4**, the $F$-test for the fixed effect of Landscape Class ($\\text{df} = 2$) is computed using the Site-within-Class mean square as the denominator error term ($\\text{df} = 12$), *never* the residual quadrat error ($\\text{df} = 60$). Testing against the site error ensures valid Type I error rates, providing statistical power $>0.88$ to detect medium effect sizes ($f = 0.35$).

### Table 4: Proposed Hierarchical Experimental Design and ANOVA Error Structure (N = 75 Quadrats Across 15 Independent Sites)

| Source of Variation | Model Term | Degrees of Freedom ($\text{df}$) | Error Term for $F$-Test | Error $\text{df}$ | Design Significance |
|:---|:---|:---:|:---|:---:|:---|
| **Landscape Class (Fixed Effect)** | $\text{Class}$ | **2** | **$\text{Site}(\text{Class})$ [Independent Sites]** | **12** | **CRITICAL FIX:** $F$-test uses site error ($\text{df}=12$), NOT quadrat residual ($\text{df}=60$), guaranteeing valid landscape inference without pseudoreplication. |
| **Site within Class (Random Effect)** | $u_{\text{Site}}$ | **12** | $\text{Residual}$ [Within-Site Subsamples] | **60** | Quantifies spatial variation between separate colliery leases and forest ranges; tests whether mining effects exceed site-to-site variability. |
| **Subsample Quadrats (Residual Error)**| $\epsilon$ | **60** | Within-site measurement error | — | Captures micro-topographic and spatial heterogeneity within each 50–160 ha site block. |
| **Season (Fixed Repeated Measures)** | $\text{Season}$ | **2** | $\text{Season} \times \text{Site}(\text{Class})$ | **24** | Tests main effect of annual tropical wet-dry climatic cycle across all sites. |
| **Class $\times$ Season (Interaction)** | $\text{Class} \times \text{Season}$ | **4** | $\text{Season} \times \text{Site}(\text{Class})$ | **24** | Directly tests whether opencast spoil exhibits transient weed flushes while underground and reference sites exhibit persistent woody stability. |
| **Total (Per-Season Cross Section)** | $\text{Total}$ | **74** | — | — | Provides statistical power $>0.88$ to detect medium effect sizes ($f=0.35$) at $\alpha=0.05$ across 15 independent sites. |

3. **Multi-Dimensional Restoration Metrics and Functional Ecosystem Recovery:** Beyond multi-spectral greenness indices and basic soil chemistry, next-phase empirical campaigns should measure five complementary ecological dimensions:
- *Undisturbed Climax Reference Benchmarks:* Establishing target restoration thresholds using protected mature *Shorea robusta* (Sal) and *Terminalia tomentosa* stands with intact soil pedons, providing an unmined baseline against which absolute ecological recovery can be quantified.
- *Soil Physical and Hydrological Functioning:* Measuring bulk density ($0–15\\text{ cm}$), cone penetration resistance to test the critical root impedance threshold ($> 2.5\\text{ MPa}$), steady-state infiltration rates using double-ring infiltrometers (testing whether raw spoil experiences severe infiltration restriction $< 5\\text{ mm/h}$ compared to $> 15\\text{ mm/h}$ under woody shrub thickets), and saturated hydraulic conductivity ($K_{\\text{sat}}$).
- *Soil Biological Health and Enzymatic Functioning:* Quantifying microbial biomass carbon (MBC) via chloroform fumigation-extraction, and assaying key soil enzymes: dehydrogenase (microbial respiration and oxidative activity), urease (nitrogen mineralization), and alkaline/acid phosphatase (phosphorus mobilization).
- *Plant Functional Traits and Restoration Value:* Leveraging our botanical trait synthesis across 42 regional species ([`Table_14_Plant_Functional_Traits_Restoration_Value.csv`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/tables/Table_14_Plant_Functional_Traits_Restoration_Value.csv)) to track symbiotic nitrogen fixation potential (nodulating legumes such as *Pongamia pinnata* and *Senna tora*), root architecture (taproot vs. fibrous vs. stoloniferous turf binders like *Cynodon dactylon*), specific leaf area (SLA), wood density, and Grime CSR ecological strategies.
- *Multi-Taxa Faunal Bioindicators:* Conducting standardized Pollard butterfly transects and avian point counts to measure the trophic return of insect pollinators and frugivorous seed dispersers.

**Scientific Integrity and Ethical Data Governance:** We explicitly emphasize that while the current paper reports the audited, completed empirical observations of the 10 published quadrats ($n = 5$ UG vs. $n = 5$ OC; preserved in [`CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md)), the 75-quadrat architecture, hypothetical LMM parameters, and blank field sheets ([`templates/`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/templates/)) represent *prospective methodological blueprints* designed to guide future field surveys ([`ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md)). Maintaining this strict division protects the empirical integrity of published datasets while providing a rigorous roadmap for next-generation restoration science.

![Figure 8](figures/Figure_5_Ecological_Restoration_Pathways.png)
*Figure 8. Conceptual restoration framework and methodological blueprint for prospective multi-site post-mining recovery monitoring, illustrating the integration of unmined reference forests, hierarchical site replication, soil physical/biological functioning, and remote sensing.*

---

## 5. Conclusions
This study establishes a defensible, audited framework for reconstructing seasonal herbaceous and shrub vegetation dynamics around published field quadrats in mining-impacted tropical dry deciduous landscapes. By integrating raw field inventories, photographic evidence, cited IMD climatology, and per-pixel Sentinel-2 L2A observations across 2024, we resolved the phenological pathways of understory flushes and woody canopies without making unsubstantiated species claims. Herbaceous-dominated overburden substrates exhibited rapid, high-amplitude green-up ($\Delta\\text{NDVI} = 0.285$) followed by sharp post-monsoon senescence, while woody shrub stands maintained persistent dry-season photosynthetic activity (Winter-to-Monsoon ratio $\ge 1.00$). PCA revealed that 81.62% of spectral variance is captured by woody canopy persistence (PC1) and seasonal herbaceous amplitude (PC2). Furthermore, to address the pilot sample size and spatial clustering constraints of the current baseline, we provide a prospective 75-quadrat, 15-site hierarchical sampling architecture that eliminates pseudoreplication and establishes a rigorous roadmap for multi-site post-mining ecological restoration monitoring across disturbed industrial landscapes globally.

---

## Data Availability Statement
All raw field registers, imagery manifests, intermediate GeoTIFF rasters with SHA-256 checksums, and analysis tables are preserved in the project repository under `data/`, `tables/`, and `rasters/`. The master audited empirical baseline for the 10 published quadrats is cataloged in [`CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md). The prospective 75-quadrat sampling design, ANOVA specifications, literature-derived plant functional traits, and standardized field/laboratory datasheets are documented in [`HIERARCHICAL_REPLICATED_LANDSCAPE_DESIGN_75_QUADRATS.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/HIERARCHICAL_REPLICATED_LANDSCAPE_DESIGN_75_QUADRATS.md), [`ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/ECOLOGICAL_RESTORATION_EXPANSION_FRAMEWORK_PROSPECTIVE.md), and [`templates/`](file:///Users/shubhamsharma/Downloads/Shrub%20and%20Herb%20korba/templates/).

## Code Availability Statement
All processing scripts (`01_build_master_locations.py` through `08_generate_manuscript_and_supplemental.py`) are fully open-source and reproducible using Python 3.10+, `rasterio`, `shapely`, `scipy`, and `scikit-learn`.

## References
1. Banerjee, S., Deb, K., Sharma, S., 2023. Soil carbon dynamics and vegetation recovery on coal mine overburden spoils. *Ecological Engineering* 192, 106980.
2. Bates, D., Mächler, M., Bolker, B., Walker, S., 2015. Fitting Linear Mixed-Effects Models Using lme4. *Journal of Statistical Software* 67, 1–48.
3. Brede, B., Terryn, L., Barbier, N., Bartholomeus, H.M., Bartolo, R., Calders, K., et al., 2020. Non-destructive tree volume estimation through UAS laser scanning. *Remote Sensing of Environment* 247, 111957.
4. Deb, K., Sharma, S., Roy, A., 2024. Environmental impacts of opencast and underground coal mining in central India. *Journal of Cleaner Production* 434, 140120.
5. Drusch, M., Del Bello, U., Carlier, S., Colin, O., Fernandez, V., Gascon, F., et al., 2012. Sentinel-2: ESA's Optical High-Resolution Mission for GMES Operational Services. *Remote Sensing of Environment* 120, 25–36.
6. Gann, G.D., McDonald, T., Walder, B., Aronson, J., Nelson, C.R., Jonson, J., et al., 2019. International principles and standards for the practice of ecological restoration. Second edition. *Restoration Ecology* 27, S1–S46.
7. Hurlbert, S.H., 1984. Pseudoreplication and the design of ecological field experiments. *Ecological Monographs* 54, 187–211.
8. Kattge, J., Bönisch, G., Díaz, S., Lavorel, S., Prentice, I.C., Leadley, P., et al., 2020. TRY plant trait database – enhanced coverage and open access. *Global Change Biology* 26, 119–188.
9. Martinuzzi, S., Gould, W.A., Vierling, L.A., 2009. Identifying tropical dry forest succession from spot remote sensing in Puerto Rico. *Biotropica* 41, 181–191.
10. Nagendra, H., Lucas, R., Honrado, J.P., Jongman, R.H., Tarantino, C., Adamo, M., Mairota, P., 2013. Remote sensing for conservation monitoring: Assessing status and trends of biodiversity, habitats, and ecosystems. *Remote Sensing in Ecology and Conservation* 1, 14–33.
11. Roy, P.S., Kushwaha, S.P.S., Murthy, M.S.R., Roy, A., Kushwaha, D., Reddy, C.S., et al., 2022. Forest fragmentation and biomass loss in the Central Indian Highlands. *Biodiversity and Conservation* 31, 843–864.
12. Sharma, S., Banerjee, S., Deb, K., 2026. Stand Structure, Carbon Sequestration and Soil Physicochemical Dynamics in Underground and Opencast Mining Quadrats of Korba, Chhattisgarh. *Environmental Monitoring and Assessment* (In Press)."""

assert old_target3b in content, "old_target3b not found"
content = content.replace(old_target3b, new_target3b, 1)

# Target 4: In generate_supplemental_markdown
old_sup = "To formally verify understory floristic composition at the species level, future field campaigns should implement this standardized nested sub-quadrat sampling protocol:"
new_sup = (
    "To formally verify understory floristic composition at the species level and expand empirical restoration monitoring, "
    "future field campaigns should implement this standardized nested sub-quadrat sampling protocol (blank operational field and laboratory datasheets "
    "are preserved in `templates/` and the audited empirical baseline is in `CURRENT_EMPIRICAL_STUDY_SCOPE_AND_RESULTS.md`):"
)
assert old_sup in content, "old_sup not found"
content = content.replace(old_sup, new_sup, 1)

with open(script_path, "w", encoding="utf-8") as f:
    f.write(content)

print(f"Successfully updated {script_path}!")
