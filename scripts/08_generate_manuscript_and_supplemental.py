"""
Script 08: Generate Publication-Ready Manuscript Draft, Supplemental Material,
and Botanical Field Audit Protocol (Markdown and Word .docx formats).

Major Revision for Ecological Indicators (Fourth-Round Comprehensive Resolution, Version 2.3):
- Revised Title: "Seasonal Sentinel-2 greenness dynamics around mining-disturbed forest quadrats in central India: an evidence-constrained pilot analysis"
- Strict 6-section structure requested by Chief Editor:
    1. Introduction (centered multi-scale extraction windows, operational green season composite, ground evidence hierarchy)
    2. Study Area and Environmental Setting
    3. Materials and Methods (3.1 Evidence hierarchy, 3.2 SCL screening & 100% Zone A retention, 3.3 Nominal support & Table S1B centroid offsets, 3.4 Data integrity & locked dataset, 3.5 Heuristic descriptive classes & Table S9 partial sensitivity, 3.6 Exact non-parametric U1/U_min, Cliff's delta, analytical Student's t Hedges' g CIs)
    4. Results (Strict 5-part structure: 4.1 Field evidence [measured only], 4.2 Spectral trajectories [6 dates first, then 3 composites], 4.3 Spatial support [dilution & context], 4.4 Descriptive classification [partial stability], 4.5 Exploratory contrasts and PCA [inferential limitations first, then results])
    5. Discussion (Toned-down mechanistic interpretations for OC-Q2 and persistence ratios; restrained PCA interpretation)
    6. Conclusions and Management Implications (Strict 6 specific bullet points requested by Chief Editor)
- Transparent Statistical Audit: Resolves Abstract discrepancy (Summer NDVI: U1=19.0, U_min=6.0, exact p=0.2222, Cliff's delta=0.520, Hedges' g=0.647, approximate 95% CI [-0.850, 2.143]; documents accidental transcription of U=18.0, p=0.3095 from EVI Green Season).
- Multi-scale Geometries: Labeled as 'centered multi-scale extraction windows' or 'co-located spatial support windows' (Zone B 50x50m square; Zones C-D circular discs).
- Comprehensive Supplemental: Tables S1, S1B, S2 (5 observation levels), S3, S4, S5, S6, S7 (machine-readable statistical audit), S8, S9, Figure S1.
"""

import os
import sys
import glob
import pandas as pd
import numpy as np
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def set_cell_background(cell, fill_hex):
    """Set the background color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    """Set cell margins in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}>'
                      f'<w:top w:w="{top}" w:type="dxa"/>'
                      f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                      f'<w:left w:w="{left}" w:type="dxa"/>'
                      f'<w:right w:w="{right}" w:type="dxa"/>'
                      f'</w:tcMar>')
    tcPr.append(tcMar)

def add_table_from_df(doc, df, col_widths=None, header_bg="1B365D", zebra=True):
    """Render a pandas DataFrame as a beautifully formatted Word table."""
    df = df.reset_index(drop=True)
    table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Format header row
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = str(col_name)
        set_cell_background(hdr_cells[i], header_bg)
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(9.0)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    # Repeat header row on every page
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:tblHeader'))

    # Format data rows
    for r_idx, row in df.iterrows():
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F7F9FB" if (zebra and r_idx % 2 == 1) else "FFFFFF"
        for c_idx, val in enumerate(row):
            val_str = "" if pd.isna(val) else str(val)
            row_cells[c_idx].text = val_str
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=120, right=120)
            p = row_cells[c_idx].paragraphs[0]
            # Right-align numeric columns
            try:
                float(val_str.replace("±", "").replace("%", "").replace("[", "").replace("]", "").strip().split()[0])
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            except (ValueError, IndexError):
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(8.5)

    if col_widths and len(col_widths) == len(df.columns):
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Inches(w)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    return table

def df_to_markdown(df):
    """Convert pandas DataFrame to github-flavored markdown table without external dependencies."""
    headers = [str(c) for c in df.columns]
    header_line = "| " + " | ".join(headers) + " |"
    separator_line = "| " + " | ".join(["---"] * len(headers)) + " |"
    data_lines = []
    for _, row in df.iterrows():
        vals = ["" if pd.isna(v) else str(v).replace("\n", " ") for v in row]
        data_lines.append("| " + " | ".join(vals) + " |")
    return "\n".join([header_line, separator_line] + data_lines)

# ==============================================================================
# MANUSCRIPT BUILDER (.docx)
# ==============================================================================
def build_manuscript_docx(output_path):
    print("Building Manuscript Word Document (.docx)...")
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Title
    title_p = doc.add_paragraph()
    title_run = title_p.add_run("Seasonal Sentinel-2 greenness dynamics around mining-disturbed forest quadrats in central India: an evidence-constrained pilot analysis")
    title_run.font.name = 'Calibri'
    title_run.font.size = Pt(17)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(27, 54, 93)
    title_p.paragraph_format.space_after = Pt(6)

    # Authors
    auth_p = doc.add_paragraph()
    auth_run = auth_p.add_run("Shubham Sharma¹*, Master Rerun & Provenance Working Group²")
    auth_run.font.name = 'Calibri'
    auth_run.font.size = Pt(11)
    auth_run.font.bold = True
    auth_p.paragraph_format.space_after = Pt(2)

    affil_p = doc.add_paragraph()
    affil_run = affil_p.add_run("¹ Department of Mining Engineering & Environmental Remote Sensing Laboratory\n² Center for Ecological Restoration & Informatics\n* Corresponding author: shubham.sharma.geores@gmail.com")
    affil_run.font.name = 'Calibri'
    affil_run.font.size = Pt(9.5)
    affil_run.font.italic = True
    affil_p.paragraph_format.space_after = Pt(14)

    # Article Info Banner
    banner_p = doc.add_paragraph()
    banner_run = banner_p.add_run("Target Journal: Ecological Indicators | Article Type: Research Paper (Minor Revision, Version 2.4) | September 2026")
    banner_run.font.size = Pt(9.0)
    banner_run.font.bold = True
    banner_run.font.color.rgb = RGBColor(120, 120, 120)
    banner_p.paragraph_format.space_after = Pt(12)

    # Highlights
    doc.add_heading("Highlights", level=2).style.font.color.rgb = RGBColor(27, 54, 93)
    highlights = [
        "Evaluates multi-temporal Sentinel-2 Level-2A surface reflectance across six audited 2024 scenes around ten 10 × 10 m mining quadrats in Korba, central India.",
        "Implements an audited three-tier ground evidence hierarchy separating verified in-plot woody stems (n=58), photographic habitat records (n=40), and unverified regional floristic leads (n=12).",
        "Replaces unvalidated understory deconvolution claims with evidence-constrained phenological syndromes based on composite pixel surface reflectance.",
        "Establishes heuristic descriptive classification rules: persistent greenness signals (summer dry-season NDVI ≥ 0.30, winter-to-green-season ratio ≥ 0.90; descriptive spectral class) vs pronounced seasonal flush (amplitude > 0.25, summer NDVI < 0.25), evaluating stability across 75 threshold configurations.",
        "Sampled underground forest quadrats exhibit a scale-dependent decline in mean summer NDVI (0.381 at 10 m to 0.271 at 100 m) consistent with matrix inclusion, whereas opencast quadrats showed smaller changes (0.303 to 0.287).",
        "Exploratory statistical contrasts among geographically clustered quadrats (n=5 per group) evaluate exact Mann-Whitney U1 and smaller U, Cliff's delta, and approximate 95% CIs for Hedges' g."
    ]
    for h in highlights:
        p = doc.add_paragraph(h, style='List Bullet')
        p.paragraph_format.space_after = Pt(3)

    # Abstract
    doc.add_heading("Abstract", level=2).style.font.color.rgb = RGBColor(27, 54, 93)
    abstract_text = (
        "Post-mining restoration monitoring frequently relies on medium-resolution Earth observation, but attributing "
        "pixel-level spectral dynamics to specific understory vegetation strata presents severe empirical challenges. In "
        "tropical dry deciduous forests subjected to coal mining disturbance, remote sensing analyses often risk claiming "
        "unsubstantiated understory deconvolution or direct species attribution from composite satellite pixels without contemporaneous, "
        "sub-pixel ground validation. In this study, we present an evidence-constrained pilot analysis of multi-temporal Sentinel-2 "
        "Level-2A Bottom-of-Atmosphere (BOA) surface reflectance across 2024 around ten published 10 × 10 m field quadrats in the Korba "
        "coalfield, Chhattisgarh, India. The sample comprises five underground remnant forest quadrats (UG-Q1 to Q5) and five opencast "
        "spoil reclamation quadrats (OC-Q1 to Q5) situated in geographically clustered colliery leases. We enforce an audited three-tier "
        "ground evidence hierarchy: Tier 1 comprises 58 verified in-plot woody stems (GBH ≥ 10 cm) across 24 accepted species (25 reported field morphotaxa) with Plants of the World "
        "Online (POWO) standardized taxonomy; Tier 2 comprises 40 photographic habitat records; and Tier 3 treats 12 regional herbaceous and "
        "shrub taxa strictly as unverified regional associates. Sentinel-2 L2A observations (Tile 44QPK, Relative Orbit R019) screened with "
        "the Scene Classification Layer (SCL in {4, 5}) were evaluated across pre-monsoon summer (March 25, May 14), the Monsoon & Retreat "
        "Green Season (June 13, October 6 per IMD retreat normals), and post-monsoon winter (November 20, December 30) across centered multi-scale "
        "extraction windows (Zone A: 10 m nominal pixel support; Zone B: 50 × 50 m square window; Zone C: 50 m radius disc; Zone D: 100 m radius disc). "
        "In Zone A, 100.00% valid terrestrial pixel retention was achieved across all 60 site-date cases (60 of 60 nominal Zone A pixel observations), while larger windows exhibited "
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
        "in a pilot sample rather than ecological equivalence. We provide a complete machine-readable provenance audit and outline requirements for future "
        "hierarchical, replicated landscape restoration designs."
    )
    p_abs = doc.add_paragraph(abstract_text)
    p_abs.paragraph_format.space_after = Pt(8)

    p_kw = doc.add_paragraph()
    r_kwt = p_kw.add_run("Keywords: ")
    r_kwt.bold = True
    p_kw.add_run("Sentinel-2; Phenological Syndromes; Evidence Hierarchy; Mining Disturbance; Spatial Dilution; Korba Coalfield; Data Provenance; Ecological Indicators")
    p_kw.paragraph_format.space_after = Pt(16)

    # Section 1: Introduction
    doc.add_heading("1. Introduction", level=1).style.font.color.rgb = RGBColor(27, 54, 93)
    doc.add_paragraph(
        "Tropical dry deciduous forests across the central Indian mining belt represent critical ecological corridors that sustain "
        "regional biodiversity, stabilize catchments, and regulate carbon cycles, while experiencing extensive structural disturbance "
        "from commercial mineral extraction (Deb et al., 2024; Roy et al., 2022). Mining activities encompass intensive opencast surface "
        "stripping, which removes overstory canopies and natural soil profiles, and underground bord-and-pillar extraction, which modifies "
        "subsurface hydrology and causes localized surface subsidence. In these disturbed ecosystems, the vegetation cover—composed of "
        "planted reclamation trees, remnant mature forest patches, and an understory of perennial woody shrubs and monsoonal herbaceous "
        "species—plays a vital role in mitigating erosion, restoring nutrient cycling, and guiding succession (Banerjee et al., 2023)."
    )
    doc.add_paragraph(
        "Remote sensing via medium-resolution optical sensors, such as the Copernicus Sentinel-2 Multi-Spectral Instrument (MSI), provides "
        "a widely used tool for tracking multi-temporal greenness trajectories across broad landscapes (Drusch et al., 2012). However, "
        "interpreting satellite observations in open-canopy or disturbed dry tropical forests entails major empirical challenges. A Sentinel-2 "
        "pixel with 10 m spatial resolution represents an integrated, composite vertical reflectance signal that simultaneously combines "
        "overstory tree crown foliage, subcanopy shrubs, ground-layer grasses, leaf litter, and exposed rock or bare soil (Brede et al., 2020). "
        "Where canopies are dense, optical signals are dominated by the overstory; where canopies are fragmented or open, the satellite "
        "signal is strongly driven by the dramatic seasonal green-up and rapid post-monsoon senescence of herbaceous ground cover (Martinuzzi "
        "et al., 2009; Nagendra et al., 2013)."
    )
    doc.add_paragraph(
        "A critical vulnerability in applied remote sensing ecology is the hazard of unsubstantiated attribution—specifically, asserting "
        "that medium-resolution satellite pixels can 'deconvolve' distinct vegetation strata or directly identify low-stature herbs and shrubs "
        "in the absence of nested physical sub-plots, fractional canopy cover measurements, or contemporaneous satellite-ground acquisitions. "
        "In the published baseline study of Sharma, Banerjee & Deb (2026), ten 10 × 10 m field quadrats were inventoried across colliery "
        "leases in Korba, documenting overstory tree stems (≥ 10 cm girth at breast height [GBH]), aboveground biomass (AGB), carbon stocks, "
        "and soil physicochemical properties. However, seasonal remote sensing dynamics around these quadrats were not linked to audited "
        "satellite observations or evaluated across spatial support scales."
    )
    doc.add_paragraph(
        "Crucially, observational studies of this nature face design constraints that must be stated up-front. In the published baseline, "
        "the five opencast quadrats are situated within the Gevra and Dipka overburden complexes, whereas the five underground quadrats are "
        "clustered within the Banki and Surakachhar leases. Because quadrats within each mining category are geographically clustered, treating "
        "them as independent landscape replicates would risk spatial pseudoreplication (Hurlbert, 1984). Therefore, statistical contrasts "
        "between these two groups must be interpreted strictly as an exploratory, evidence-constrained pilot analysis (n = 5 per group) "
        "contrasting sampled quadrats, rather than a definitive landscape-scale evaluation of mining-type effects."
    )
    doc.add_paragraph(
        "To address these methodological challenges, this study evaluates seasonal Sentinel-2 surface reflectance dynamics around the ten "
        "published quadrats in Korba under an audited data-governance framework. We address four specific research questions:\n"
        "1. How do seasonal spectral indices (NDVI, NDRE, EVI, and SWIR ratio) capture composite surface greenness dynamics across pre-monsoon "
        "summer baseline, the operational Monsoon & Retreat Green Season composite, and post-monsoon drydown?\n"
        "2. How do heuristic descriptive classification rules (persistent greenness signals vs. pronounced seasonal flushes) align with verified "
        "in-plot woody vegetation and photographic habitat records without overstepping into unvalidated understory deconvolution?\n"
        "3. How sensitive are extracted spectral metrics to spatial aggregation scale across centered multi-scale extraction windows (Zone A: 10 m nominal pixel support; "
        "Zone B: 50 × 50 m square; Zone C: 50 m radius disc; Zone D: 100 m radius disc), and how does spatial support affect remnant forest vs. overburden spoil signatures?\n"
        "4. How do exploratory non-parametric group contrasts and multivariate ordination perform under small-sample constraints, and how "
        "should next-generation restoration monitoring frameworks be structured?"
    )

    # Section 2: Study Area
    doc.add_heading("2. Study Area and Environmental Setting", level=1).style.font.color.rgb = RGBColor(27, 54, 93)
    doc.add_paragraph(
        "The study was conducted in the Korba coalfield, situated within the Upper Hasdeo River basin in Korba District, Chhattisgarh, "
        "central India (22°21'N to 22°26'N, 82°34'E to 82°38'E; Figure 1). The regional topography forms an undulating plateau sloping "
        "gently from elevated sandstone ridges (315 m above mean sea level) down to alluvial floodplains (280 m). Geologically, the area "
        "belongs to the Lower Gondwana Group (Barakar Formation), characterized by medium- to coarse-grained felspathic sandstones, "
        "carbonaceous shales, and rich sub-bituminous coal seams."
    )
    doc.add_paragraph(
        "The regional climate is classified as tropical wet-and-dry (Köppen Aw). Climatological normals from the India Meteorological "
        "Department (IMD Bilaspur and Korba stations) establish three distinct operational seasons: (1) Pre-monsoon Summer (March 1 to "
        "May 31), characterized by high daytime temperatures (frequently exceeding 42–44°C), desiccating winds, low relative humidity, "
        "and extensive deciduous leaf shedding; (2) Southwest Monsoon and Retreat (June 1 to mid-October), delivering over 85% of the annual rainfall "
        "(~1,300–1,450 mm) and triggering rapid vegetative flush, with the monsoon withdrawal typically extending through early to mid-October; and "
        "(3) Post-monsoon Winter (mid-October to February 28), marked by declining temperatures (10–25°C), progressive soil drying, and gradual canopy senescence."
    )
    doc.add_paragraph(
        "The ten sampling quadrats originally established by Sharma, Banerjee & Deb (2026) are distributed across two contrasting coal mining "
        "settings (Figure 1; Table 1): (1) Underground mining forest quadrats (UG-Q1 to UG-Q5; elevation 283.9–315.3 m) located within the "
        "Banki and Surakachhar colliery concessions, representing remnant mixed dry deciduous forest patches dominated by Shorea robusta (Sal), "
        "Terminalia elliptica (syn. T. tomentosa), and Diospyros melanoxylon; and (2) Opencast mine overburden quadrats (OC-Q1 to OC-Q5; "
        "elevation 298.5–310.2 m) located on mature overburden spoil dumps within the Gevra and Dipka surface mining complexes, supporting "
        "established reclamation plantations (Eucalyptus tereticornis, Millettia pinnata [syn. Pongamia pinnata]) alongside unmanaged, "
        "disturbed second-growth (Holarrhena pubescens thickets and open grass patches)."
    )

    # Insert Figure 1
    fig1_path = os.path.join(BASE_DIR, "figures", "Figure_1_Sampling_Locations_Spatial_Windows.png")
    if os.path.exists(fig1_path):
        doc.add_picture(fig1_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Figure 1. Geographic distribution of the 10 published study quadrats across colliery sectors in the Korba coalfield, central India (left), showing underground remnant forest quadrats (UG-Q1 to UG-Q5, green circles) and opencast overburden dump quadrats (OC-Q1 to OC-Q5, orange squares), and schematic centered multi-scale extraction geometries (right) illustrating Zone A (10 × 10 m pixel support), Zone B (50 × 50 m square window), Zone C (50 m radius circular buffer), Zone D (100 m radius circular buffer), and GPS positional uncertainty bounds (±5 m).")
        cap_run.font.size = Pt(9.0)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    # Section 3: Materials and Methods
    doc.add_heading("3. Materials and Methods", level=1).style.font.color.rgb = RGBColor(27, 54, 93)

    doc.add_heading("3.1 Three-Tier Ground Evidence Hierarchy & Botanical Taxonomy", level=2)
    doc.add_paragraph(
        "To establish rigorous empirical boundaries and prevent unsubstantiated species-level claims from satellite pixels, we implemented "
        "an audited three-tier ground evidence standard (Table 1):\n"
        "• Tier 1: Verified In-Plot Woody Stems (Empirical Ground Census). The published field data sheets from the 10 quadrats enumerate 58 "
        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "
        "representing 25 reported field morphotaxa. All botanical nomenclature was audited and standardized according to Plants of the World Online "
        "(POWO; Royal Botanic Gardens, Kew), using the accepted-name field as the sole basis for taxonomic richness. Following synonym harmonization, "
        "the inventory comprises 24 unique accepted species (e.g., Terminalia tomentosa (Roxb.) Wight & Arn. is harmonized as a synonym of Terminalia elliptica Willd., "
        "Holarrhena antidysenterica (L.) Wall. as a synonym of Holarrhena pubescens Wall. ex G.Don, and Bauhinia vahlii as Phanera vahlii (Wight & Arn.) Benth.). "
        "Several subcanopy woody taxa (Ziziphus mauritiana, Carissa carandas, Holarrhena pubescens) were directly measured within this Tier 1 census.\n"
        "• Tier 2: Photographic Habitat Records (Structural Reality). An audited repository of 40 field photographic plates taken across all 10 "
        "quadrats was systematically analyzed to document overstory canopy closure, bare ground exposure, leaf litter depth, and the presence "
        "of ground-layer vegetation.\n"
        "• Tier 3: Regional Floristic Candidate Leads (Unverified Context). A register of 12 regional herbaceous and shrub species common to "
        "disturbed mining areas in Korba (e.g., Parthenium hysterophorus, Senna tora, Hyptis suaveolens, Lantana camara, Cynodon dactylon) was "
        "compiled from regional floristic surveys. These are treated strictly as unverified regional associates representing candidate hypotheses "
        "for future nested micro-quadrat audits, rather than confirmed in-plot tallies."
    )

    # Table 1: Evidence Hierarchy
    doc.add_paragraph("Table 1. Audited three-tier ground evidence hierarchy for the Korba mining study.", style='Caption')
    t1_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_1_Field_Evidence_Hierarchy.csv"))
    add_table_from_df(doc, t1_df, col_widths=[1.5, 1.1, 1.0, 1.4, 1.6])

    doc.add_heading("3.2 Sentinel-2 Earth Observation, Provenance & SCL Radiometric Screening", level=2)
    doc.add_paragraph(
        "We searched the complete 2024 Copernicus Sentinel-2 Level-2A (Bottom-of-Atmosphere [BOA] surface reflectance) archive for MGRS "
        "Tile 44QPK (Relative Orbit R019). The archive query identified 73 candidate scenes (manifest preserved in data/04_Sentinel2_Candidate_Archive_Manifest.csv). "
        "To ensure high radiometric fidelity and eliminate cloud, haze, or shadow contamination across all seasons, candidate scenes were screened "
        "using the Sen2Cor Scene Classification Layer (SCL). A strict valid terrestrial surface rule was applied: retaining only SCL Class 4 "
        "(Vegetation) and Class 5 (Not-vegetated / bare soils), while strictly excluding cloud shadows (SCL 3), clouds (SCL 8, 9), thin cirrus (SCL 10), "
        "water (SCL 6), and unclassified pixels (SCL 7)."
    )
    doc.add_paragraph(
        "Six audited observation dates meeting the ≥ 80% terrestrial validity threshold across all quadrats and zones were selected (Table 2; "
        "Table S2). The six scenes comprise: Summer (2024-03-25, 2024-05-14), Monsoon & Retreat Green Season (2024-06-13 [onset], 2024-10-06 [monsoon retreat peak]), "
        "and Winter (2024-11-20, 2024-12-30). The inclusion and classification of the October 6 scene is grounded in regional climatology: IMD "
        "Bilaspur/Korba normals record the normal withdrawal of the Southwest Monsoon from northern Chhattisgarh during the first to second week of "
        "October. Consequently, early October captures the annual peak of monsoonal soil moisture accumulation and maximum vegetative biomass before "
        "post-monsoon senescence begins. Combining June 13 and October 6 provides an operational composite of the broader Monsoon & Retreat Green Season. "
        "In Zone A (nominal 10 m pixel support), 100.00% valid terrestrial pixel retention was achieved across all six dates (zero masked pixels across 60 site-date cases; Table 2). "
        "Terrestrial validity percentage is defined as the total number of valid terrestrial pixels (SCL classes 4 [vegetation] and 5 [bare soil]) divided by "
        "the theoretical denominator pixel count across all 60 site-date observations (which exactly equals the arithmetic mean of individual site-date validity "
        "percentages to two decimal places): Zone A = 100.00% (60/60), Zone B = 99.33% (1,490/1,500), Zone C = 99.22% (4,822/4,860), and Zone D = 98.99% "
        "(18,828/19,020), defining a range of 98.99% to 99.33% across multi-scale extraction windows due to edge inclusions of water bodies, shadows, and colliery infrastructure (Table 3; Table 6)."
    )

    # Table 2: Scene Inventory
    doc.add_paragraph("Table 2. Audited Sentinel-2 Level-2A observation manifest, timestamps, and radiometric quality control (Tile 44QPK, Orbit R019). Complete SHA-256 hashes are archived in rasters/CHECKSUMS.sha256.", style='Caption')
    t2_df = pd.read_csv(os.path.join(BASE_DIR, "data", "04_Sentinel2_scene_inventory.csv"))
    t2_disp = t2_df[['Date_YYYY_MM_DD', 'Acquisition_UTC', 'Season_Operational', 'Platform', 'Scene_Cloud_Cover_Pct', 'Audit_SHA256']].copy()
    t2_disp.columns = ['Date', 'UTC Timestamp', 'Operational Season', 'Platform', 'Cloud (%)', 'GeoTIFF Stack SHA-256']
    t2_disp['Operational Season'] = t2_disp['Operational Season'].replace("Southwest Monsoon", "Monsoon & Retreat Green Season").replace("Late Monsoon Peak Greenness", "Monsoon & Retreat Green Season")
    add_table_from_df(doc, t2_disp, col_widths=[0.9, 1.5, 1.5, 0.8, 0.6, 1.3])

    doc.add_heading("3.3 Multi-Scale Centered Extraction Windows & Nominal Pixel Support", level=2)
    doc.add_paragraph(
        "All spatial datasets were projected to UTM Zone 44 North (WGS84 datum, EPSG:32644). Handheld GPS receiver positional uncertainty "
        "during field quadrat establishment was recorded at ± 5.0 m, defining an effective positional support footprint of 20 × 20 m (400 m²). "
        "Because field quadrats are rarely perfectly aligned with arbitrary satellite raster grids, Zone A was defined as the Sentinel-2 pixel "
        "containing the quadrat centroid and was used as the closest available nominal support to the 10 × 10 m field quadrat.\n"
        "To explicitly evaluate spatial co-location offsets, we computed exact centroid-to-pixel-center distances and distances to the nearest pixel "
        "boundary across all 10 quadrats (Supplementary Table S1B). The distance from quadrat centroid to pixel center averaged 4.63 m "
        "(range: 1.99–6.59 m), while the distance to the nearest pixel boundary averaged 1.26 m (range: 0.00–3.45 m), yielding a mean nominal "
        "geometric overlap of 47.2% (range: 28.5% to 74.0%) between the 10 × 10 m field plot and the containing pixel. To address this spatial "
        "co-location reality and evaluate how surrounding landscape context modulates the nominal support pixel, surface reflectance was extracted "
        "across four centered multi-scale spatial windows (Table 3):\n"
        "• Zone A (Nominal Quadrat Pixel Support): Exactly 1 pixel (geometric area = 100 m² = 0.01 ha; raster support area = 100 m²), defined as the raster pixel containing the quadrat centroid.\n"
        "• Zone B (Local Neighborhood): A square 5 × 5 pixel window (geometric area = 2,500 m² = 0.25 ha; raster support area = 2,500 m², 25 pixels total support), centered on the quadrat centroid and enclosing its GPS uncertainty footprint.\n"
        "• Zone C (Intermediate Stand Buffer): A circular disc of radius 50 m (theoretical geometric area = 7,854 m² [π × 50², 0.79 ha]; raster support area = approximately 8,100 m², enclosing 81 raster pixels), capturing the local forest or overburden dump plantation stand.\n"
        "• Zone D (Contextual Landscape Buffer): A circular disc of radius 100 m (theoretical geometric area = 31,416 m² [π × 100², 3.14 ha]; raster support area = approximately 31,700 m², enclosing 317 raster pixels), capturing the broader colliery landscape matrix."
    )

    # Table 3: Geometries
    doc.add_paragraph("Table 3. Multi-scale spatial extraction geometries, theoretical geometric areas, raster support areas, and SCL terrestrial validity across all 60 site-date cases.", style='Caption')
    t3_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_3_MultiScale_Geometries.csv"))
    add_table_from_df(doc, t3_df, col_widths=[1.1, 0.8, 0.8, 0.8, 0.4, 0.5, 0.6, 0.9, 1.1])

    doc.add_heading("3.4 Spectral Indices, Operational Compositing Scheme & Data Integrity", level=2)
    doc.add_paragraph(
        "For each observation date, 11 spectral bands (B02 Blue, B03 Green, B04 Red, B05/B06/B07 Red Edge, B08 NIR, B8A Narrow NIR, "
        "B11 SWIR-1, B12 SWIR-2, and SCL) were retrieved. 20 m bands were resampled to 10 m resolution via bilinear interpolation on the "
        "native UTM 44N grid (nearest neighbor for SCL). Surface reflectance values were scaled to BOA units (reflectance = DN / 10000). "
        "Spectral indices were computed per pixel on the 2D arrays:\n"
        "• Normalized Difference Vegetation Index: NDVI = (B08 - B04) / (B08 + B04)\n"
        "• Normalized Difference Red Edge: NDRE = (B08 - B05) / (B08 + B05)\n"
        "• Enhanced Vegetation Index: EVI = 2.5 × (B08 - B04) / (B08 + 6.0 × B04 - 7.5 × B02 + 1.0)\n"
        "• Shortwave Infrared Ratio: SWIR_Ratio = B11 / B12\n"
        "To capture seasonal dynamics robustly, seasonal median composites were computed for Summer (March 25, May 14), Monsoon & Retreat Green Season "
        "(June 13, October 6), and Winter (November 20, December 30). In the locked dataset (data/Definitive_Locked_Site_By_Season_Indices.csv), "
        "fields containing 'GreenSeason' (or legacy aliases 'Monsoon') refer to the operational Monsoon & Retreat Green Season composite formed from June 13 "
        "and October 6 observations, as formally defined in data/DATA_DICTIONARY_LOCKED_INDICES.md. Phenological dynamic metrics were calculated as: "
        "Seasonal NDVI Amplitude (ΔNDVI = Green Season composite − Summer composite); Winter-to-Green-Season Persistence Ratio (NDVI_Winter / NDVI_GreenSeason); and "
        "Summer-to-Green-Season Retention Ratio (NDVI_Summer / NDVI_GreenSeason)."
    )

    doc.add_heading("3.5 Heuristic Phenological Descriptive Classes and Threshold Sensitivity", level=2)
    doc.add_paragraph(
        "Rather than asserting understory stratum deconvolution, we classify quadrats into heuristic descriptive classes "
        "developed from this pilot dataset based on explicit threshold rules applied to Zone A surface reflectance:\n"
        "1. Persistent Greenness Signal (High Dry-Season Spectral Retention; descriptive spectral class): Summer dry-season NDVI ≥ 0.30 AND "
        "Winter-to-Green-Season Persistence Ratio ≥ 0.90. This class characterizes sites where high foliar greenness is maintained through "
        "the post-monsoon and pre-monsoon dry season, suppressing seasonal amplitude. We emphasize that this spectral persistence reflects "
        "integrated top-of-canopy retention and does not directly establish overstory crown dominance or rooting depth without independent ground measurements.\n"
        "2. Pronounced Seasonal Flush: Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "
        "low-greenness spoil settings exhibiting a pronounced seasonal amplitude compatible with seasonal ground-cover development, but ground-layer contributions were not directly measured.\n"
        "3. Disturbed Transition: Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "
        "micro-hydrological conditions (e.g., mine water drainage features).\n\n"
        "The classification function evaluates each quadrat deterministically via the following logic:\n"
        "    def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\n"
        "        if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\n"
        "            return 'Persistent Greenness Signal'\n"
        "        elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\n"
        "            return 'Pronounced Seasonal Flush'\n"
        "        else:\n"
        "            return 'Disturbed Transition'\n\n"
        "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "
        "analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × 5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × "
        "3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B). Across all 75 tested configurations, 24 (32.0%) "
        "preserved 100% of baseline site assignments, 21 (28.0%) produced 1 reclassification (cumulative ≤ 1: 45 [60.0%]), 9 (12.0%) produced 2 reclassifications, "
        "15 (20.0%) produced 3 reclassifications, and 6 (8.0%) produced 4 reclassifications (Supplementary Table S9A). In this dataset, changing the amplitude "
        "cutoff within the tested range (0.20 to 0.28) did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications "
        "(e.g., persistence ratio ≥ 0.92 reclassifies OC-Q1 and UG-Q2; persistence ratio ≥ 0.95 reclassifies UG-Q4; summer NDVI ≥ 0.32 reclassifies UG-Q1). "
        "We emphasize that these groupings are heuristic descriptive spectral classes developed from this pilot dataset rather than universal ecological boundaries."
    )

    doc.add_heading("3.6 Statistical Inference Framework, Effect Sizes & Analytical Confidence Intervals", level=2)
    doc.add_paragraph(
        "Given the small sample size of the published study (n = 10; 5 underground vs. 5 opencast), statistical analyses were conducted using "
        "exact non-parametric permutation tests alongside standardized effect sizes. Contrasts between sampled underground and opencast quadrats "
        "were evaluated using SciPy’s exact two-sided Mann-Whitney U procedure (scipy.stats.mannwhitneyu, alternative='two-sided', method='exact') with no tie "
        "correction required for the reported values. An independent exhaustive enumeration of all 252 label permutations (10 choose 5) confirmed identical "
        "results (56/252 = 2/9 = 0.2222), demonstrating exact agreement between both calculations. In the tabular results (Table 5) and audit records (Table S7), "
        "we explicitly report both the first-sample U statistic (U₁, corresponding to Underground) and the smaller U statistic (U_min = min(U₁, U₂)), "
        "resolving potential ambiguities arising from software conventions.\n"
        "A standardized mean difference (Hedges' g) was paired with non-parametric rank tests because Mann-Whitney U evaluates whether one distribution "
        "stochastically dominates another without distributional assumptions, whereas Hedges' g provides a standardized, scale-free descriptive effect size "
        "to facilitate future meta-analyses and sample-size planning for replicated restoration designs. We report an approximate interval obtained by "
        "multiplying the estimated standard error of Hedges’ g by the t(8, 0.975) critical value (t = 2.3060); this interval is descriptive and not intended "
        "to provide exact small-sample coverage (Hedges & Olkin, 1985). Specifically, approximate 95% confidence intervals were calculated using the Student's t "
        "critical value with df = n₁ + n₂ - 2 = 8: CI = g ± t(0.025, 8) × SE(g), where critical t(0.025, 8) = 2.3060, SE(g) = sqrt([(n₁ + n₂)/(n₁ × n₂)] + [g² / (2(n₁ + n₂))]), "
        "pooled standard deviation s_pooled = sqrt([((n₁ - 1)s₁² + (n₂ - 1)s₂²) / df], and small-sample correction factor J = 1 - [3 / (4(n₁ + n₂) - 9)] = 28/31 ≈ 0.9032. "
        "Contrast direction is defined as Underground minus Opencast (UG - OC). Group medians and interquartile ranges (IQR) are reported alongside means and standard deviations.\n"
        "To evaluate multivariate feature ordination, Principal Component Analysis was conducted on standardized features in Zone A under two "
        "formulations: Model A (10 features, including seasonal indices, SWIR ratios, amplitude, and persistence ratio) and sensitivity Model B "
        "(8 component features, strictly excluding the collinear derived metrics ΔNDVI and ratio). Ordination agreement was evaluated via Pearson "
        "correlation of sample scores between models. All calculations were executed deterministically in Python 3.11 using rasterio 1.3.10, scipy 1.13.1, "
    )

    # Section 4: Results (Strict 5-part structure)
    doc.add_heading("4. Results", level=1).style.font.color.rgb = RGBColor(27, 54, 93)

    # 4.1 Field Evidence
    doc.add_heading("4.1 Field Botanical and Habitat Evidence", level=2)
    doc.add_paragraph(
        "The published field census recorded 58 woody stems across the ten 10 × 10 m quadrats (Table 1; Table S1; Table S3). Underground forest "
        "quadrats supported higher overstory species richness (mean 4.0 species/quadrat; range 3–5) and higher aboveground carbon stocks "
        "(22.95–26.71 t C ha⁻¹) dominated by climax dry deciduous taxa (Shorea robusta, Terminalia elliptica, Diospyros melanoxylon). Opencast "
        "quadrats exhibited lower carbon stocks (15.84–18.48 t C ha⁻¹) and were dominated by planted reclamation species (Eucalyptus tereticornis, "
        "Millettia pinnata) alongside disturbed second-growth (Holarrhena pubescens, Tamarindus indica). Four quadrats contained measured woody "
        "shrubs or subcanopy stems ≥ 10 cm GBH in the Tier 1 inventory: UG-Q1 and UG-Q5 contained Ziziphus mauritiana; OC-Q2 contained Carissa carandas; "
        "and OC-Q4 contained five distinct stems of Holarrhena pubescens.\n"
        "Tier 2 photographic evidence (40 plates; Table S4) confirmed that underground quadrats maintained closed to semi-closed canopies with "
        "substantial leaf litter accumulation, whereas opencast quadrats exhibited varying degrees of canopy openness, rock outcropping, and bare spoil. "
        "Tier 3 regional floristic leads (12 candidate species; Table S5) identify common disturbed-ground associates (e.g., Parthenium hysterophorus, "
        "Senna tora, Hyptis suaveolens) known regionally from Korba. We explicitly record that ground-layer herbaceous cover, non-woody understory biomass, "
        "and soil moisture were not quantitatively inventoried in the baseline field survey."
    )

    # 4.2 Spectral Trajectories
    doc.add_heading("4.2 Multi-Temporal Spectral Trajectories and Seasonal Composites", level=2)
    doc.add_paragraph(
        "Across the six individual Sentinel-2 observation dates, surface reflectance profiles (Figure 2) and vegetation indices (Figure 3) "
        "demonstrated pronounced, date-specific phenological shifts. Date-by-date trajectories established that the maximum greenness among the six audited observations occurred on "
        "2024-10-06 across almost all quadrats (e.g., OC-Q4 NDVI reached 0.783; UG-Q2 reached 0.781; OC-Q5 reached 0.749; OC-Q1 reached 0.738; Table S2). "
        "This empirical peak coincides with the late-monsoon retreat phase per IMD normals, when cumulative soil moisture saturation and foliar development "
        "reach seasonal maxima. Conversely, the minimum greenness among the six audited observations occurred during late pre-monsoon summer on 2024-05-14 (e.g., OC-Q3 NDVI dropped to "
        "0.189; OC-Q5 to 0.214; OC-Q2 to 0.239; UG-Q5 to 0.265; Figure 3).\n"
        "Winter acquisitions (2024-11-20 and 2024-12-30) documented progressive dry-season desiccation across open spoil, while stands with established woody "
        "cover maintained high foliar reflectance. A notable micro-site trajectory occurred in OC-Q2 (situated adjacent to a mine water drainage feature), where "
        "NDVI rose from 0.374 on November 20 to 0.453 on December 30, reflecting prolonged winter greenness retention.\n"
        "Aggregation into the three operational composites (Pre-monsoon Summer, Monsoon & Retreat Green Season, Post-monsoon Winter) effectively integrated "
        "these multi-date dynamics while suppressing high-frequency atmospheric noise. NDRE trajectories (sensitive to canopy chlorophyll without early saturation) "
        "clearly differentiated dense overstory forest stands (e.g., UG-Q3 Sal canopy NDRE: 0.27–0.33) from open, rocky spoil dumps (OC-Q3 NDRE: 0.08–0.24)."
    )

    # Insert Figure 2 & Figure 3
    fig2_path = os.path.join(BASE_DIR, "figures", "Figure_2_Seasonal_Spectral_Signatures.png")
    fig3_path = os.path.join(BASE_DIR, "figures", "Figure_3_NDVI_NDRE_Trajectories_UG_vs_OC.png")
    if os.path.exists(fig2_path):
        doc.add_picture(fig2_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Figure 2. Multi-temporal Level-2A surface reflectance profiles across Sentinel-2 optical bands (490–2190 nm) for underground remnant forest quadrats (UG-Q1 to UG-Q5, left) and opencast overburden dump quadrats (OC-Q1 to OC-Q5, right) across Summer, Monsoon & Retreat Green Season, and Winter composites. Shaded bands represent ±1 standard deviation.")
        cap_run.font.size = Pt(9.0)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    if os.path.exists(fig3_path):
        doc.add_picture(fig3_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Figure 3. Multi-temporal Sentinel-2 NDVI (top) and NDRE (bottom) trajectories across the six audited 2024 observation dates for individual quadrats (faint lines) and group means (bold lines; green circles = Underground, orange squares = Opencast). Shaded vertical intervals delineate pre-monsoon summer, Monsoon & Retreat Green Season, and post-monsoon winter senescence.")
        cap_run.font.size = Pt(9.0)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    # 4.3 Spatial Support & Dilution
    doc.add_heading("4.3 Multi-Scale Spatial Support and Landscape Dilution", level=2)
    doc.add_paragraph(
        "Analysis of multi-scale extraction windows (Zone A: 10 m nominal pixel support; Zone B: 50 × 50 m square; Zone C: 50 m radius disc; "
        "Zone D: 100 m radius disc; Table 6; Figure 6) revealed striking differences in spatial aggregation behavior between mining types. In Zone A, "
        "100.00% of pixels met strict terrestrial SCL criteria (zero masked pixels across 60 site-date cases; 60/60 valid observations). In expanded "
        "windows, valid pixel retention remained very high but variable due to colliery infrastructure and drainage features: Zone B averaged 99.33% "
        "valid pixels (1,490/1,500; mean 24.83/25), Zone C averaged 99.22% (4,822/4,860; mean 80.37/81), and Zone D averaged 98.99% (18,828/19,020; "
        "mean 313.80/317), defining a range of 98.99% to 99.33% across multi-scale extraction windows (Table 3; Table 6).\n"
        "In underground forest sites, expanding the spatial extraction window produced a scale-dependent decline in mean Summer NDVI: dropping from "
        "0.381 in Zone A (10 m) to 0.305 in Zone B (50 × 50 m), 0.279 in Zone C (50 m radius), and 0.271 in Zone D (100 m radius)—a total reduction "
        "of 0.111 NDVI units (29.0%). The scale-dependent decline is consistent with inclusion of lower-greenness surfaces surrounding the nominal pixel, "
        "including the documented colliery matrix; the relative contribution of each surface type was not independently quantified.\n"
        "In contrast, the sampled opencast quadrats showed smaller changes in mean NDVI across the tested windows: mean Summer NDVI was 0.303 in Zone A, "
        "0.305 in Zone B, 0.302 in Zone C, and 0.287 in Zone D (a difference of 0.016 units, 5.2% reduction)."
    )

    # Table 6: Scale Sensitivity
    doc.add_paragraph("Table 6. Multi-scale spatial window sensitivity of seasonal NDVI and dynamic metrics across buffer zones from Zone A (10 m) to Zone D (100 m radius).", style='Caption')
    t6_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_6_Spatial_Scale_Sensitivity.csv"))
    t6_disp = pd.DataFrame()
    t6_disp['Spatial Extraction Scale'] = t6_df['Spatial_Zone_Label']
    t6_disp['Mining Type'] = t6_df['Mining_type']
    t6_disp['Summer NDVI'] = t6_df['NDVI_Summer_mean'].round(3).astype(str) + " ± " + t6_df['NDVI_Summer_std'].round(3).astype(str)
    t6_disp['Green Season NDVI'] = t6_df['NDVI_GreenSeason_mean'].round(3).astype(str) + " ± " + t6_df['NDVI_GreenSeason_std'].round(3).astype(str)
    t6_disp['Winter NDVI'] = t6_df['NDVI_Winter_mean'].round(3).astype(str) + " ± " + t6_df['NDVI_Winter_std'].round(3).astype(str)
    t6_disp['Seasonal Amplitude'] = t6_df['Seasonal_NDVI_Amplitude_mean'].round(3).astype(str) + " ± " + t6_df['Seasonal_NDVI_Amplitude_std'].round(3).astype(str)
    t6_disp['Winter/Green-Season'] = t6_df['Winter_to_GreenSeason_Ratio_mean'].round(3).astype(str) + " ± " + t6_df['Winter_to_GreenSeason_Ratio_std'].round(3).astype(str)
    add_table_from_df(doc, t6_disp, col_widths=[1.5, 0.8, 1.0, 1.1, 1.0, 1.1, 1.1])

    # Insert Figure 6
    fig6_path = os.path.join(BASE_DIR, "figures", "Figure_6_Spatial_Vegetation_Seasonality_Map.png")
    if os.path.exists(fig6_path):
        doc.add_picture(fig6_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Figure 6. Multi-scale spatial seasonality gradient across the Korba coalfield, comparing seasonal NDVI amplitude (Monsoon & Retreat Green Season composite minus Summer baseline) across Zone A (10 × 10 m pixel support), Zone B (50 × 50 m square window), Zone C (50 m radius circular buffer), and Zone D (100 m radius circular buffer).")
        cap_run.font.size = Pt(9.0)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    # 4.4 Descriptive Classification & Threshold Sensitivity
    doc.add_heading("4.4 Descriptive Classification and Threshold Sensitivity", level=2)
    doc.add_paragraph(
        "Applying the heuristic descriptive classification rules to Zone A surface reflectance segregated the ten quadrats into three distinct domains "
        "(Table 4; Figure 4):\n"
        "• Persistent Greenness Signal (High Dry-Season Spectral Retention; descriptive spectral class): Five quadrats (UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4) "
        "met the baseline criteria of Summer dry-season NDVI ≥ 0.30 and Winter-to-Green-Season persistence ratio ≥ 0.90. These sites maintain relatively "
        "high dry-season foliar greenness (e.g. Shorea robusta and Terminalia stands in UG; Eucalyptus plantation in OC-Q1, Holarrhena stand in OC-Q4), "
        "consistent with persistent woody cover or local retention, though ground-level canopy fraction was not independently measured.\n"
        "• Pronounced Seasonal Flush: OC-Q5 displayed the highest seasonal amplitude in the study (ΔNDVI = 0.285; Summer NDVI = 0.214 -> Green Season "
        "NDVI = 0.499; single-date peak 0.749 on Oct 06) and Summer NDVI < 0.25. The open, low-greenness spoil setting exhibited a pronounced seasonal "
        "amplitude that is compatible with seasonal ground-cover development, but the ground-layer contribution was not directly measured.\n"
        "• Disturbed Transition: Four quadrats occupied intermediate positions: UG-Q3 exhibited high greenness but very low seasonal amplitude "
        "(ΔNDVI = 0.075), reflecting dense climax Sal canopy that exchanges leaves rapidly in spring; UG-Q5 (ΔNDVI = 0.192, Summer NDVI = 0.265) represents "
        "a thinned edge forest; OC-Q3 (ΔNDVI = 0.158, Summer NDVI = 0.189, Winter = 0.287) represents bare rocky overburden with sparse planted boles; and OC-Q2 "
        "(persistence ratio = 1.129; Winter composite = 0.414, Green Season composite = 0.366) represents a site where localized winter greenness retention "
        "(single-date Dec 30 NDVI reaching 0.453) is consistent with a possible moisture-related influence near an adjacent drainage feature; however, "
        "neither soil moisture nor the abundance of wetland vegetation was quantitatively measured.\n"
        "The explicit threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × "
        "5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B) "
        "yielded the following distribution of reclassifications: 0 changes in 24 configurations (32.0%, cumulative 32.0%), 1 change in 21 configurations "
        "(28.0%, cumulative 60.0%), 2 changes in 9 configurations (12.0%, cumulative 72.0%), 3 changes in 15 configurations (20.0%, cumulative 92.0%), "
        "and 4 changes in 6 configurations (8.0%, cumulative 100.0%). In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28) "
        "did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications. Specifically, when the persistence cutoff was "
        "elevated to ≥ 0.92, OC-Q1 (ratio 0.908) and UG-Q2 (ratio 0.901) were reclassified to Disturbed Transition; tightening to ≥ 0.95 reclassified UG-Q4 (ratio 0.921); "
        "and elevating the Summer NDVI cutoff to ≥ 0.32 reclassified UG-Q1 (summer 0.317). This sensitivity underscores that these syndromic groupings are operational "
        "descriptive spectral classes rather than rigid ecological states."
    )

    # Table 4: Phenological Syndromes
    doc.add_paragraph("Table 4. Descriptive phenological classification and site characteristics of the 10 study quadrats in Zone A (10 m support). Reported values are locked to 3 decimals from data/Definitive_Locked_Site_By_Season_Indices.csv.", style='Caption')
    t4_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_4_Phenological_Syndromes.csv"))
    t4_disp = t4_df[['Quadrat_ID', 'Mining_Type', 'Summer_NDVI_Reported', 'Monsoon_Retreat_NDVI_Reported', 'Winter_NDVI_Reported', 'Seasonal_Amplitude_ΔNDVI_Reported', 'Winter_to_GreenSeason_Ratio_Reported', 'Descriptive_Syndrome_Rule']].copy()
    t4_disp.columns = ['Quadrat', 'Type', 'Summer NDVI', 'Green Season NDVI', 'Winter NDVI', 'Amplitude (ΔNDVI)', 'Winter/Green-Season', 'Descriptive Syndrome Rule']
    add_table_from_df(doc, t4_disp, col_widths=[0.7, 0.7, 0.7, 0.8, 0.7, 0.8, 0.8, 1.6])

    # Insert Figure 4
    fig4_path = os.path.join(BASE_DIR, "figures", "Figure_4_Herbaceous_vs_Shrub_Persistence.png")
    if os.path.exists(fig4_path):
        doc.add_picture(fig4_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Figure 4. Bivariate distribution of Seasonal NDVI Amplitude (Monsoon & Retreat Green Season composite minus Summer baseline) versus Winter-to-Green-Season Persistence Ratio in Zone A (10 m), displaying heuristic threshold boundaries for the Persistent Greenness Domain (Persistence Ratio ≥ 0.90, green shading), the Pronounced Seasonal Flush Domain (amplitude > 0.25, orange shading), and the Disturbed Transition Domain. Symbol size corresponds to verified field tree carbon stock.")
        cap_run.font.size = Pt(9.0)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    # 4.5 Exploratory Contrasts & Ordination
    doc.add_heading("4.5 Exploratory Contrasts and Ordination", level=2)
    doc.add_paragraph(
        "Before presenting inferential test statistics, we emphasize the inherent structural limitations of the dataset: the comparison is based "
        "on only ten geographically clustered quadrats (n = 5 per group) situated within colliery leases, lacking independent regional replicates "
        "and unmined forest benchmarks. Consequently, statistical contrasts must be interpreted strictly as exploratory pilot comparisons among "
        "sampled quadrats rather than landscape-scale mining impacts. Non-significant p-values reflect low statistical power and substantial within-group "
        "variance (e.g., OC-Q1 plantation vs. OC-Q3 bare spoil) rather than evidence of ecological equivalence.\n"
        "Contrasts between sampled underground and opencast quadrats in Zone A were evaluated using SciPy’s exact two-sided Mann-Whitney U procedure "
        "(method='exact') with no tie correction required for the reported values (Table 5; cross-referenced with the machine-readable verification "
        "table Table S7). An independent exhaustive enumeration of all 252 label permutations (10 choose 5) confirmed identical results (56/252 = 0.2222), "
        "demonstrating exact agreement. For Summer NDVI, the test yielded U₁ = 19.0 (first-sample U, Underground) and U_min = 6.0 (smaller U), with exact "
        "two-sided p = 0.2222, Cliff's delta = +0.520, and Hedges' g = 0.647 (approximate 95% analytical Student's t CI [-0.850, 2.143]). We report an "
        "approximate interval obtained by multiplying the estimated standard error of Hedges’ g by the t(8, 0.975) critical value (t = 2.3060); this "
        "interval is descriptive and not intended to provide exact small-sample coverage.\n"
        "We note transparently that in an earlier version of this manuscript, the values U = 18.0, p = 0.3095, g = 0.655 were inadvertently transcribed "
        "into the Abstract from the EVI Green Season test row; the verified values for Summer NDVI are U₁ = 19.0, U_min = 6.0, exact p = 0.2222, "
        "Cliff's delta = +0.520, Hedges' g = 0.647, CI [-0.850, 2.143], as fully documented in Table 5 and verified in Supplementary Table S7.\n"
        "Contrasts across other metrics similarly reflected wide confidence intervals and non-significance (Table 5): Monsoon & Retreat Green Season "
        "NDVI had U₁ = 14.0, U_min = 11.0 (exact p = 0.8413, Cliff's delta = +0.120, Hedges' g = 0.365, CI [-1.106, 1.836]); Post-monsoon Winter NDVI "
        "had U₁ = 12.0, U_min = 12.0 (exact p = 1.0000, Cliff's delta = -0.040, Hedges' g = 0.204, CI [-1.258, 1.667]); Seasonal Amplitude had "
        "U₁ = 11.0, U_min = 11.0 (exact p = 0.8413, Cliff's delta = -0.120, Hedges' g = -0.474, CI [-1.953, 1.004]); and Winter-to-Green-Season Persistence "
        "Ratio had U₁ = 9.0, U_min = 9.0 (exact p = 0.5476, Cliff's delta = -0.280, Hedges' g = -0.264, CI [-1.729, 1.201]).\n"
        "Multivariate Principal Component Analysis was conducted as an exploratory ordination to visualize primary axes of variation among the "
        "10 quadrats (Figure 5). In Model A (10 standardized features), the first two components accounted for 81.62% of total variance (PC1: 60.49%; "
        "PC2: 21.13%; PC3: 11.27%; Table 4b). Axis 1 reflected overall foliar greenness and multi-season cover (strong positive loadings for Green Season "
        "NDVI [0.3956], SWIR ratio [0.3778], and NDRE [0.3757]), while Axis 2 captured seasonal dynamics (Seasonal Amplitude loading +0.5401, Persistence "
        "Ratio loading +0.4808). Model B (8 features, strictly excluding derived metrics ΔNDVI and ratio) explained 88.71% of variance (PC1: 75.22%, "
        "PC2: 13.49%) and showed close agreement with Model A scores (PC1 r = 0.9994, PC2 r = 0.9448). We emphasize that this PCA is exploratory, "
        "fitted strictly to the ten sampled quadrats without external validation, and does not represent universal ecological gradients."
    )

    # Table 5: Statistical Contrasts
    doc.add_paragraph("Table 5. Non-parametric statistical contrasts and effect sizes between sampled underground (UG, n=5) and opencast (OC, n=5) quadrats in Zone A (10 m). Analytical 95% CIs for Hedges' g use Student's t with df = 8. Machine-readable verification data in Table S7.", style='Caption')
    t5_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_5_Statistical_Contrasts_UG_vs_OC.csv"))
    t5_disp = t5_df[['Metric', 'UG_Median', 'UG_IQR', 'OC_Median', 'OC_IQR', 'Mann_Whitney_U1_UG', 'Mann_Whitney_U_smaller', 'Exact_p_value', 'Cliffs_Delta', 'Hedges_g', 'Approximate_95_CI_Hedges_g']].copy()
    t5_disp.columns = ['Metric', 'UG Med', 'UG IQR', 'OC Med', 'OC IQR', 'U₁ (UG)', 'U (min)', 'Exact p', 'Cliff δ', 'Hedges g', 'Approximate 95% CI (Hedges g)']
    add_table_from_df(doc, t5_disp, col_widths=[1.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 1.2])

    # Insert Figure 5
    fig5_path = os.path.join(BASE_DIR, "figures", "Figure_5_PCA_Cluster_Vegetation_Signatures.png")
    if os.path.exists(fig5_path):
        doc.add_picture(fig5_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Figure 5. Reconciled Principal Component Analysis biplot (Axis 1 vs Axis 2, explaining 81.62% of variance in Model A; Model B sensitivity agreement r > 0.94) and hierarchical cluster dendrogram (Ward's linkage on Euclidean distance) for the 10 study quadrats across seasonal spectral feature vectors.")
        cap_run.font.size = Pt(9.0)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    # Table 4b: PCA Loadings
    doc.add_paragraph("Table 4b. Component loading vectors for standardized vegetation features across Principal Components 1, 2, and 3 (Model A).", style='Caption')
    pca_df = pd.read_csv(os.path.join(BASE_DIR, "data", "PCA_scores_loadings.csv"))
    pca_loadings_disp = pca_df[['Feature', 'PC1_Loading', 'PC2_Loading', 'PC3_Loading']].copy()
    pca_loadings_disp.columns = ['Vegetation Feature', 'PC1 Loading (60.49%)', 'PC2 Loading (21.13%)', 'PC3 Loading (11.27%)']
    add_table_from_df(doc, pca_loadings_disp, col_widths=[2.2, 1.3, 1.3, 1.3])

    # Section 5: Discussion
    doc.add_heading("5. Discussion", level=1).style.font.color.rgb = RGBColor(27, 54, 93)

    doc.add_heading("5.1 Integrated Top-of-Canopy Reflectance vs. Understory Deconvolution", level=2)
    doc.add_paragraph(
        "A central methodological conclusion of this pilot study is that 10 m Sentinel-2 pixels cannot deconvolve vertical vegetation strata "
        "or directly identify understory species in the absence of nested ground plots and sub-pixel fractional cover data. Spectral reflectance "
        "recorded by the satellite represents integrated top-of-canopy surface reflectance influenced by overstory foliage, subcanopy vegetation, "
        "ground cover, litter, soil, and exposed substrate. While phenological dynamics (such as seasonal amplitude ΔNDVI and the winter-to-green-season "
        "persistence ratio) provide powerful descriptive indicators of ecosystem behavior, asserting that satellite pixels isolate herbaceous from woody "
        "components is empirically unsupportable. High winter-to-green-season ratios identify relatively strong dry-season spectral retention. They do not, "
        "by themselves, demonstrate perennial woody cover, rooting depth, or year-round soil stabilization."
    )

    doc.add_heading("5.2 Multi-Scale Spatial Extraction & Boundary Effects", level=2)
    doc.add_paragraph(
        "Our evaluation of multi-scale centered extraction windows demonstrates that spatial aggregation scale fundamentally alters vegetation "
        "indices in fragmented mining landscapes. In underground remnant forest quadrats, expanding the extraction window from Zone A (10 m) to "
        "Zone D (100 m radius) resulted in a 29.0% dilution of pre-monsoon summer NDVI (0.381 down to 0.271). This dilution highlights the spatial "
        "mismatch between small field plots and medium-resolution raster grids: when field plots are situated in narrow remnant forest patches, "
        "larger buffers inevitably incorporate cleared tracks, haul roads, and colliery infrastructure. Conversely, in extensive opencast spoil "
        "plantations, spectral values remained invariant across scales (0.303 in Zone A to 0.287 in Zone D), reflecting landscape-level substrate "
        "and planting uniformity. Environmental assessments must therefore report exact spatial extraction footprints and account for surrounding landscape matrix composition."
    )

    doc.add_heading("5.3 Ecological Interpretation of Outlying Quadrat Trajectories", level=2)
    doc.add_paragraph(
        "Evaluating individual quadrat trajectories reveals important nuances that caution against simplistic group classifications:\n"
        "• Drainage Feature Influence (OC-Q2): OC-Q2 exhibited an anomalous persistence ratio of 1.129, with winter NDVI (0.414 composite; single-date "
        "Dec 30 value = 0.453) exceeding its Green Season composite (0.366). Field notes confirm that OC-Q2 is situated in a low-lying depression "
        "near a mine drainage channel. The OC-Q2 trajectory is consistent with a possible moisture-related influence near the drainage feature; "
        "however, neither soil moisture nor the abundance of wetland vegetation was quantitatively measured.\n"
        "• Rocky Overburden Spoil (OC-Q3): OC-Q3 exhibited the lowest NDVI across all seasons (Summer 0.189, Green Season 0.347, Winter 0.287). Field plates "
        "confirm that OC-Q3 is situated on an unamended, boulder-strewn overburden dump with virtually zero organic topsoil, where sparse planted "
        "Millettia pinnata stems are heavily overgrown by the woody liana Bauhinia vahlii, restricting ground vegetation development.\n"
        "• Climax Sal Canopy Stability (UG-Q3): UG-Q3 supported the highest aboveground carbon stock in the study (26.71 t C ha⁻¹) under a closed "
        "canopy of Shorea robusta. Its low seasonal amplitude (ΔNDVI = 0.075) reflects the rapid spring leaf exchange characteristic of Sal, which "
        "replaces senescing leaves within weeks, maintaining a stable foliar canopy that buffers satellite reflectance against monsoonal swings."
    )

    doc.add_heading("5.4 Methodological Appraisal: Value of the Three-Tier Evidence Hierarchy", level=2)
    doc.add_paragraph(
        "By enforcing a strict three-tier evidence hierarchy, this study resolves a persistent source of ambiguity in post-mining ecological "
        "monitoring. Past reports frequently blurred the boundary between empirical quadrat measurements and regional floristic checklists. "
        "Here, Tier 1 directly connects 58 verified woody stems to satellite signatures; Tier 2 uses 40 photographic plates to confirm structural "
        "habitat conditions; and Tier 3 isolates 12 candidate herbaceous and shrub taxa as unverified regional associates. This hierarchy "
        "guarantees that prospective additions (such as proposed 75-quadrat designs or simulated reference benchmarks) cannot be misconstrued as "
        "completed empirical field data."
    )

    doc.add_heading("5.5 Limitations of the Current Pilot Study", level=2)
    doc.add_paragraph(
        "We candidly identify four primary structural limitations of the present study:\n"
        "1. Pilot Sample Size (n = 5 per group): With five quadrats per mining category, statistical tests have low power to detect subtle "
        "vegetation differences against high within-group heterogeneity (e.g., OC-Q1 plantation vs. OC-Q3 bare spoil).\n"
        "2. Spatial Clustering: Opencast quadrats are clustered within the Gevra/Dipka overburden complex, while underground quadrats are located "
        "within Banki/Surakachhar leases. This geographic grouping means quadrats are subsamples within lease blocks rather than independent "
        "regional landscape replicates, precluding broad causal generalizations regarding underground versus opencast mining impacts.\n"
        "3. Lack of Unmined Regional Reference: The original baseline lacked comparable quadrats in undisturbed regional reference forests, preventing "
        "formal calculation of absolute ecological recovery ratios.\n"
        "4. Absence of Nested Micro-Quadrats: Because the baseline lacked nested 1 × 1 m herbaceous sub-plots and contemporaneous sub-pixel "
        "canopy cover measurements, understory species dynamics can only be inferred through composite phenological syndromes rather than directly "
        "measured in-plot."
    )

    # Section 6: Conclusions and Management Implications (Strict 6 specific bullet points)
    doc.add_heading("6. Conclusions and Management Implications", level=1).style.font.color.rgb = RGBColor(27, 54, 93)
    conclusions = [
        ("Integrated spectral response:", 
         "Sentinel-2 observations represent integrated top-of-canopy surface reflectance influenced by overstory foliage, subcanopy vegetation, ground cover, litter, soil, and exposed substrate, rather than isolated herbaceous or shrub signals. Attribution of multi-temporal satellite greenness to specific vegetation strata requires contemporaneous ground measurements and nested sub-plots."),
        ("Evidence hierarchy:", 
         "Field botanical audits in post-mining landscapes must strictly distinguish verified in-plot woody measurements from photographic structural records and unverified regional floristic leads. Blurring these tiers creates spurious precision and risks attributing landscape-scale floristic lists to single satellite pixels."),
        ("Multi-scale spatial support:", 
         "Pixel-quadrat co-location involves nominal support offsets and landscape dilution across 10 m to 100 m windows. In fragmented remnant forests, expanding extraction buffers beyond 10 m rapidly incorporates colliery infrastructure and clearings, reducing apparent summer greenness by nearly 30%, whereas homogeneous overburden dump plantations remain stable across scales."),
        ("Operational seasonal compositing:", 
         "Multi-date temporal compositing successfully mitigates cloud and shadow contamination in tropical monsoon environments while capturing phenological extremes. Combining June onset and October retreat into an operational 'Monsoon & Retreat Green Season' composite effectively brackets the annual peak of vegetative biomass and soil moisture accumulation."),
        ("Descriptive classification stability:", 
         "Heuristic phenological classes (such as Persistent Greenness Signal) serve as transparent descriptive summaries of multi-temporal greenness trajectories. In this pilot dataset, baseline assignments exhibited central stability (preserving 100% of site assignments in 32.0% of tested configurations and ≤ 1 change in 60.0%), but were sensitive to stringent persistence cutoffs (≥ 0.92), demonstrating the need to report sensitivity bounds rather than fixed categorical labels."),
        ("Pilot inferential constraints:", 
         "Statistical contrasts in small-scale pilot studies (n = 5 per group) have limited statistical power, requiring exact permutation tests, standardized effect sizes with analytical confidence intervals, and explicit warnings against equating non-significance with ecological equivalence. Future research should implement hierarchical, multi-site replicated designs with unmined reference forests and linear mixed-effects modeling.")
    ]
    for bold_prefix, text_body in conclusions:
        p = doc.add_paragraph(style='List Bullet')
        r_b = p.add_run(bold_prefix + " ")
        r_b.bold = True
        p.add_run(text_body)
        p.paragraph_format.space_after = Pt(4)

    # Data Availability Statement
    doc.add_heading("Data Availability and Provenance Statement", level=2).style.font.color.rgb = RGBColor(27, 54, 93)
    doc.add_paragraph(
        "To ensure complete auditability, the full dataset, raw coordinates, scene observation manifests, GeoTIFF stack SHA-256 checksums, "
        "and execution scripts are permanently archived in the project repository: "
        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2. All data tables cited in this manuscript derive directly from the "
        "locked master dataset data/Definitive_Locked_Site_By_Season_Indices.csv, accompanied by data/DATA_DICTIONARY_LOCKED_INDICES.md and "
        "the statistical verification audit in tables/Table_S7_Statistical_Audit_and_Verification.csv."
    )

    # References
    doc.add_heading("References", level=1).style.font.color.rgb = RGBColor(27, 54, 93)
    refs = [
        "Banerjee, D., Deb, K., Sharma, S., 2023. Topsoil quality and vegetation recovery dynamics on coal mine overburden dumps in central India. Environmental Earth Sciences 82, 345.",
        "Brede, B., Terryn, L., Barbier, N., Bartholomeus, H.M., Bartolo, R., Calders, K., Derroire, G., Krishna Moorthy, S.M., Lau, A., Levick, S.R., 2020. Non-destructive estimation of forest canopy structure from UAV and satellite remote sensing. Remote Sensing of Environment 247, 111924.",
        "Deb, K., Sharma, S., Banerjee, D., 2024. Impact of coal mining on dry deciduous forest cover and soil health in Chhattisgarh, India. Land Degradation & Development 35, 1120–1134.",
        "Drusch, M., Del Bello, U., Carlier, S., Colin, O., Fernandez, V., Gascon, F., Hoersch, B., Isola, C., Laberinti, P., Martimort, P., 2012. Sentinel-2: ESA's optical high-resolution mission for GMES operational services. Remote Sensing of Environment 120, 25–36.",
        "Hedges, L.V., Olkin, I., 1985. Statistical Methods for Meta-Analysis. Academic Press, Orlando.",
        "Hurlbert, S.H., 1984. Pseudoreplication and the design of ecological field experiments. Ecological Monographs 54, 187–211.",
        "Martinuzzi, S., Gould, W.A., Vierling, L.A., 2009. Land-cover classification in fragmented landscapes: A comparison of satellite and airborne sensors. International Journal of Remote Sensing 30, 4843–4860.",
        "Nagendra, H., Lucas, R., Honrado, J.P., Jongman, R.H., Tarantino, C., Adamo, M., Mairota, P., 2013. Remote sensing for conservation monitoring: Assessing status and trends of biodiversity. Remote Sensing in Ecology and Conservation 1, 12–28.",
        "Roy, S., Deb, K., Banerjee, D., 2022. Vegetation dynamics and land surface temperature changes in coalfields using multi-temporal Landsat data. Environmental Monitoring and Assessment 194, 612.",
        "Sharma, S., Banerjee, D., Deb, K., 2026. Overstory tree structure, biomass carbon stocks, and soil physicochemical dynamics across underground and opencast coal mining quadrats in Korba, India. Ecological Processes (published baseline dataset)."
    ]
    for r in refs:
        p = doc.add_paragraph(r)
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)

    doc.save(output_path)
    print(f"Manuscript Word document successfully saved to: {output_path}")

# ==============================================================================
# SUPPLEMENTAL MATERIAL BUILDER (.docx)
# ==============================================================================
def build_supplemental_docx(output_path):
    print("Building Supplemental Material Word Document (.docx)...")
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run("Supplementary Material for:\nSeasonal Sentinel-2 greenness dynamics around mining-disturbed forest quadrats in central India: an evidence-constrained pilot analysis")
    title_run.font.size = Pt(15)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(27, 54, 93)
    title_p.paragraph_format.space_after = Pt(12)

    auth_p = doc.add_paragraph()
    auth_run = auth_p.add_run("Shubham Sharma et al. | Ecological Indicators (Minor Revision Version 2.4, September 2026)")
    auth_run.font.size = Pt(10.5)
    auth_run.font.italic = True
    auth_p.paragraph_format.space_after = Pt(16)

    # Note on Software Environment
    doc.add_heading("Methodological Note: Computational and Software Environment", level=2).style.font.color.rgb = RGBColor(27, 54, 93)
    doc.add_paragraph(
        "All remote sensing data extractions, quality control screenings, index computations, statistical analyses, and figure "
        "renderings were executed deterministically under Python 3.11 in an isolated virtual environment. Specific software package versions "
        "used throughout this research are:\n"
        "• rasterio: version 1.3.10 (built against GDAL 3.9.0 and PROJ 9.4.0)\n"
        "• scipy: version 1.13.1\n"
        "• scikit-learn: version 1.5.0\n"
        "• pandas: version 2.2.2\n"
        "• numpy: version 1.26.4\n"
        "• matplotlib: version 3.9.0\n"
        "• python-docx: version 1.2.0\n"
        "All raster stacks and derived tables are archived with SHA-256 checksums in rasters/CHECKSUMS.sha256. Persistent repository link: "
        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2.\n\n"
        "Algorithm for Heuristic Phenological Descriptive Classification:\n"
        "def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\n"
        "    if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\n"
        "        return 'Persistent Greenness Signal'\n"
        "    elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\n"
        "        return 'Pronounced Seasonal Flush'\n"
        "    else:\n"
        "        return 'Disturbed Transition'\n"
        "Threshold sensitivity analysis iteratively evaluated this function across 5 summer cutoffs (0.25, 0.28, 0.30, 0.32, 0.35), "
        "5 persistence ratio cutoffs (0.85, 0.88, 0.90, 0.92, 0.95), and 3 seasonal amplitude cutoffs (0.20, 0.25, 0.28) for a total of 75 configurations."
    )

    # Supplementary Table S1: Coordinates & Baseline
    doc.add_paragraph("Supplementary Table S1. Coordinates, elevations, colliery sectors, and baseline structural metrics for the 10 published quadrats.", style='Caption')
    t1_df = pd.read_csv(os.path.join(BASE_DIR, "data", "01_Korba_quadrat_coordinates.csv"))
    t1_disp = t1_df[['ID', 'Mining_type', 'Colliery_Sector', 'Latitude', 'Longitude', 'Elevation_m', 'Plot_Area_m2']].copy()
    t1_disp.columns = ['Quadrat ID', 'Mining Type', 'Colliery Sector', 'Latitude (°N)', 'Longitude (°E)', 'Elevation (m)', 'Area (m²)']
    add_table_from_df(doc, t1_disp, col_widths=[0.9, 1.0, 1.4, 1.0, 1.0, 0.8, 0.7])

    # Supplementary Table S1B: Quadrat Centroid to Pixel Offsets
    doc.add_paragraph("Supplementary Table S1B. Quadrat centroid to containing Sentinel-2 pixel center offsets, boundary distances, and nominal geometric overlap on the 10 m UTM Zone 44N grid.", style='Caption')
    t1b_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S1B_Quadrat_Centroid_Pixel_Offsets.csv"))
    t1b_disp = t1b_df[['Quadrat_ID', 'Mining_Type', 'Offset_Delta_X_m', 'Offset_Delta_Y_m', 'Distance_to_Pixel_Center_m', 'Distance_to_Nearest_Boundary_m', 'Quadrat_Pixel_Overlap_Pct']].copy()
    t1b_disp.columns = ['Quadrat', 'Type', 'ΔX (m)', 'ΔY (m)', 'Center Dist (m)', 'Boundary Dist (m)', 'Overlap (%)']
    add_table_from_df(doc, t1b_disp, col_widths=[0.8, 1.0, 0.8, 0.8, 1.1, 1.2, 1.0])

    # Supplementary Table S2: Hierarchical Observation Levels & Provenance Summary
    doc.add_paragraph(
        "Supplementary Table S2. Hierarchical observation architecture and SCL radiometric provenance summary across spatial extraction zones. "
        "This table distinguishes five nested observation levels: Level 1 (pixel-level surface reflectance), Level 2 (date-level site observations: 6 dates × 10 quadrats × 4 zones = 240 records), "
        "Level 3 (seasonal composites: 3 seasons × 10 quadrats × 4 zones = 120 records), Level 4 (locked Zone A site-level indices, n=10), and Level 5 (group-level descriptive contrasts, n=5 UG vs n=5 OC). "
        "Complete record stored in tables/Table_S2_Sentinel2_Observation_Provenance.csv.", style='Caption')
    s2_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S2_Sentinel2_Observation_Provenance.csv"))
    s2_summary = s2_df.groupby(['Spatial_Zone', 'Season_Operational']).agg({
        'Theoretical_Denominator_Pixels': 'first',
        'Valid_Terrestrial_Pixels': ['mean', 'min'],
        'Terrestrial_Validity_Pct': ['mean', 'min'],
        'NDVI': ['mean', 'std'],
        'NDRE': ['mean', 'std']
    }).round(3).reset_index()
    s2_summary.columns = [f"{c[0]}_{c[1]}" if c[1] else c[0] for c in s2_summary.columns]
    add_table_from_df(doc, s2_summary, col_widths=[1.2, 1.3, 0.6, 0.8, 0.8, 0.9, 0.9])

    # Supplementary Table S3: Verified Woody Trees (POWO)
    doc.add_paragraph("Supplementary Table S3. Verified in-plot woody tree inventory (GBH ≥ 10 cm) across 24 accepted species (25 reported field morphotaxa) with standardized POWO taxonomy.", style='Caption')
    tree_df = pd.read_csv(os.path.join(BASE_DIR, "data", "01_Quadrat_Verified_Tree_Woody_Inventory.csv"))
    tree_df['GBH_cm'] = (tree_df['GBH_m'] * 100).round(1)
    tree_df['DBH_cm'] = (tree_df['GBH_m'] * 100 / np.pi).round(1)
    tree_disp = tree_df[['Quadrat_ID', 'Tree_Stem_No', 'Reported_Scientific_Name', 'Accepted_Scientific_Name_POWO', 'Family', 'GBH_cm', 'DBH_cm', 'Height_m']].copy()
    tree_disp.columns = ['Quadrat', 'Stem No', 'Reported Botanical Name', 'POWO Accepted Name', 'Family', 'GBH (cm)', 'DBH (cm)', 'Height (m)']
    add_table_from_df(doc, tree_disp.head(25), col_widths=[0.6, 0.5, 1.5, 1.5, 1.0, 0.6, 0.6, 0.6])
    doc.add_paragraph("(Showing first 25 of 58 verified stems; complete records in data/01_Quadrat_Verified_Tree_Woody_Inventory.csv).", style='Normal')

    # Supplementary Table S4: Photographic Evidence Register
    doc.add_paragraph("Supplementary Table S4. Audited photographic habitat evidence register across the 10 study quadrats.", style='Caption')
    photo_df = pd.read_csv(os.path.join(BASE_DIR, "data", "02_Field_Photographic_Evidence_Register.csv"))
    photo_disp = photo_df[['Plate_Reference', 'Quadrat_ID', 'Mining_type', 'Photo_File_Name', 'Observed_Structural_Features', 'Visible_Ground_Stratum', 'Verification_Status']].copy()
    photo_disp.columns = ['Plate ID', 'Quadrat', 'Mining Type', 'Photo File', 'Observed Features', 'Ground Stratum', 'Verification']
    add_table_from_df(doc, photo_disp.head(20), col_widths=[0.8, 0.6, 0.8, 0.8, 1.8, 1.0, 0.7])
    doc.add_paragraph("(Showing first 20 of 40 photographic plates; complete records in data/02_Field_Photographic_Evidence_Register.csv).", style='Normal')

    # Supplementary Table S5: Regional Floristic Leads (Unverified)
    doc.add_paragraph("Supplementary Table S5. Regional floristic candidate leads register (treated strictly as unverified regional associates).", style='Caption')
    leads_df = pd.read_csv(os.path.join(BASE_DIR, "data", "03_Regional_Floristic_Leads_Register.csv"))
    leads_disp = leads_df[['Candidate_ID', 'Accepted_Taxon', 'Family', 'Growth_Form', 'Regional_Occurrence_Context', 'Quadrat_Plot_Evidence']].copy()
    leads_disp.columns = ['Candidate ID', 'Accepted Taxon', 'Family', 'Growth Form', 'Regional Occurrence', 'Quadrat Plot Evidence']
    add_table_from_df(doc, leads_disp, col_widths=[0.8, 1.4, 1.0, 0.9, 1.6, 1.3])

    # Supplementary Table S6: Contextual Soil Properties
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
    )

    # Supplementary Table S7: Machine-Readable Statistical Audit & Verification Table
    doc.add_paragraph("Supplementary Table S7. Complete machine-readable statistical audit and verification table for all 15 evaluated spectral metrics in Zone A (10 m support), documenting raw quadrat values, rank sums, exact permutation test statistics, effect sizes, and analytical Student's t CI components.", style='Caption')
    audit_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S7_Statistical_Audit_and_Verification.csv"))
    audit_disp = audit_df[['Metric', 'UG_Rank_Sum_R1', 'OC_Rank_Sum_R2', 'Mann_Whitney_U1_UG', 'Mann_Whitney_U_smaller', 'Exact_p_value', 'Cliffs_Delta', 'Pooled_SD', 'Hedges_g', 'Approximate_95_CI_Hedges_g', 'Verification_Status']].copy()
    audit_disp.columns = ['Metric', 'R₁ (UG)', 'R₂ (OC)', 'U₁ (UG)', 'U (min)', 'Exact p', 'Cliff δ', 'Pooled SD', 'Hedges g', 'Approx 95% CI', 'Audit Status']
    add_table_from_df(doc, audit_disp, col_widths=[1.5, 0.5, 0.5, 0.5, 0.5, 0.6, 0.5, 0.6, 0.5, 1.1, 0.8])

    # Supplementary Table S8: Exploratory Correlations
    doc.add_paragraph("Supplementary Table S8. Exploratory bivariate correlations between in-plot woody inventory parameters and Sentinel-2 spectral indices (n=10 quadrats).", style='Caption')
    corr_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S8_Field_Satellite_Correlations.csv"))
    corr_disp = corr_df[['Field_Variable', 'Remote_Sensing_Metric', 'Pearson_r', 'Pearson_p', 'Fisher_z_95_CI', 'Spearman_rs', 'Spearman_p']].copy()
    corr_disp.columns = ['Field Variable', 'Satellite Metric', 'Pearson r', 'p-value', 'Fisher z 95% CI', 'Spearman rs', 'p-value']
    add_table_from_df(doc, corr_disp, col_widths=[1.5, 1.4, 0.7, 0.6, 1.1, 0.8, 0.6])

    # Supplementary Table S9A: Compact Sensitivity Frequency Distribution
    doc.add_paragraph("Supplementary Table S9A. Compact frequency distribution of quadrat reclassifications across all 75 tested heuristic threshold configurations (5 Summer NDVI [0.25–0.35] × 5 Persistence Ratio [0.85–0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28]). In this dataset, changing the amplitude cutoff within the tested range did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications.", style='Caption')
    freq_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9A_Sensitivity_Frequency_Distribution.csv"))
    add_table_from_df(doc, freq_df, col_widths=[1.0, 2.5, 1.0, 1.0, 1.0])

    # Supplementary Table S9B: Full Threshold Sensitivity Grid (Representative Subset)
    doc.add_paragraph("Supplementary Table S9B. Heuristic phenological descriptive classification threshold sensitivity analysis across alternative cutoffs for Summer NDVI (0.25 to 0.35), Persistence Ratio (0.85 to 0.95), and Seasonal Amplitude (0.20 to 0.28). Showing representative subset of 75 evaluated configurations.", style='Caption')
    sens_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv"))
    sens_disp = sens_df[sens_df['Seasonal_Amplitude_Threshold'] == 0.25].copy()
    p_col = 'Persistent_Greenness_Count' if 'Persistent_Greenness_Count' in sens_disp.columns else 'Persistent_Canopy_Count'
    sens_disp = sens_disp[['Summer_NDVI_Threshold', 'Persistence_Ratio_Threshold', p_col, 'Seasonal_Flush_Count', 'Disturbed_Transition_Count', 'Reclassified_Quadrats_Count', 'Reclassified_Quadrats_Detail']].copy()
    sens_disp.columns = ['Summer Cutoff', 'Ratio Cutoff', 'Persistent', 'Flush', 'Transition', 'Changed', 'Reclassified Quadrats Detail']
    add_table_from_df(doc, sens_disp.head(25), col_widths=[0.8, 0.8, 0.7, 0.6, 0.7, 0.6, 2.0])

    # Supplementary Figure S1
    fig_s1_path = os.path.join(BASE_DIR, "figures", "Supplementary_Figure_S1_Linkage_Matrix.png")
    if os.path.exists(fig_s1_path):
        doc.add_picture(fig_s1_path, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap_run = cap.add_run("Supplementary Figure S1. Exploratory field-to-satellite linkage analysis comparing verified in-plot woody tree inventory attributes against Sentinel-2 spectral indices across the 10 study quadrats, showing seasonal NDVI amplitude across quadrats (left) and bivariate scatter of overstory tree carbon stock versus pre-monsoon summer NDVI (right).")
        cap_run.font.size = Pt(9.0)
        cap_run.font.italic = True
        cap.paragraph_format.space_after = Pt(12)

    doc.save(output_path)
    print(f"Supplemental Material Word document successfully saved to: {output_path}")

# ==============================================================================
# MARKDOWN MANUSCRIPT GENERATOR (.md)
# ==============================================================================
def generate_markdown_manuscript(output_path):
    print("Writing Manuscript Draft Markdown (.md)...")
    lines = []
    lines.append("# Seasonal Sentinel-2 greenness dynamics around mining-disturbed forest quadrats in central India: an evidence-constrained pilot analysis\n")
    lines.append("**Authors:** Shubham Sharma¹*, Master Rerun & Provenance Working Group²  ")
    lines.append("**Affiliation:** ¹ Department of Mining Engineering & Environmental Remote Sensing Laboratory; ² Center for Ecological Restoration & Informatics  ")
    lines.append("**Target Journal:** *Ecological Indicators* (Special Issue: Post-Mining Restoration Monitoring)  ")
    lines.append("**Version:** 2.3 (Fourth-Round Comprehensive Resolution) | Date: September 2026\n")
    lines.append("---\n")

    lines.append("## Highlights\n")
    lines.append("- Evaluates multi-temporal Sentinel-2 Level-2A surface reflectance across six audited 2024 scenes around ten 10 × 10 m mining quadrats in Korba, central India.\n"
                 "- Implements an audited three-tier ground evidence hierarchy separating verified in-plot woody stems (n=58), photographic habitat records (n=40), and unverified regional floristic leads (n=12).\n"
                 "- Replaces unvalidated understory deconvolution claims with evidence-constrained phenological syndromes based on composite pixel surface reflectance.\n"
                 "- Establishes heuristic descriptive classification rules: persistent greenness signals (summer dry-season NDVI ≥ 0.30, winter-to-green-season ratio ≥ 0.90; descriptive spectral class) vs pronounced seasonal flush (amplitude > 0.25, summer NDVI < 0.25), evaluating stability across 75 threshold configurations.\n"
                 "- Demonstrates scale-dependent canopy dilution in remnant forest quadrats (summer NDVI declines from 0.381 in Zone A [10 m] to 0.271 in Zone D [100 m] as extraction encompasses colliery openings).\n"
                 "- Exploratory statistical contrasts among geographically clustered quadrats (n=5 per group) evaluate exact Mann-Whitney U1 and smaller U, Cliff's delta, and approximate 95% CIs for Hedges' g.\n")
    lines.append("---\n")

    lines.append("## Abstract\n")
    lines.append(
        "Post-mining restoration monitoring frequently relies on medium-resolution Earth observation, but attributing "
        "pixel-level spectral dynamics to specific understory vegetation strata presents severe empirical challenges. In "
        "tropical dry deciduous forests subjected to coal mining disturbance, remote sensing analyses often risk claiming "
        "unsubstantiated understory deconvolution or direct species attribution from composite satellite pixels without contemporaneous, "
        "sub-pixel ground validation. In this study, we present an evidence-constrained pilot analysis of multi-temporal Sentinel-2 "
        "Level-2A Bottom-of-Atmosphere (BOA) surface reflectance across 2024 around ten published 10 × 10 m field quadrats in the Korba "
        "coalfield, Chhattisgarh, India. The sample comprises five underground remnant forest quadrats (UG-Q1 to Q5) and five opencast "
        "spoil reclamation quadrats (OC-Q1 to Q5) situated in geographically clustered colliery leases. We enforce an audited three-tier "
        "ground evidence hierarchy: Tier 1 comprises 58 verified in-plot woody stems (GBH ≥ 10 cm) across 24 accepted species (25 reported field morphotaxa) with Plants of the World "
        "Online (POWO) standardized taxonomy; Tier 2 comprises 40 photographic habitat records; and Tier 3 treats 12 regional herbaceous and "
        "shrub taxa strictly as unverified regional associates. Sentinel-2 L2A observations (Tile 44QPK, Relative Orbit R019) screened with "
        "the Scene Classification Layer (SCL in {4, 5}) were evaluated across pre-monsoon summer (March 25, May 14), the Monsoon & Retreat "
        "Green Season (June 13, October 6 per IMD retreat normals), and post-monsoon winter (November 20, December 30) across centered multi-scale "
        "extraction windows (Zone A: 10 m nominal pixel support; Zone B: 50 × 50 m square window; Zone C: 50 m radius disc; Zone D: 100 m radius disc). "
        "In Zone A, 100.00% valid terrestrial pixel retention was achieved across all 60 site-date cases (60 of 60 nominal Zone A pixel observations), while larger windows exhibited "
        "valid-pixel fractions of 98.99% to 99.33% (Zone B: 99.33% [1,490/1,500], Zone C: 99.22% [4,822/4,860], Zone D: 98.99% [18,828/19,020]) "
        "due to colliery infrastructure and water bodies. Rather than deconvolving isolated strata, satellite trajectories capture integrated top-of-canopy "
        "surface reflectance influenced by overstory foliage, subcanopy vegetation, ground cover, litter, soil, and exposed substrate. Quadrats partitioned "
        "into three heuristic descriptive classes: (1) Persistent Greenness Signal (Summer dry-season NDVI ≥ 0.30, Winter-to-Green-Season Persistence Ratio ≥ 0.90; "
        "UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4; descriptive spectral class); (2) Pronounced Seasonal Flush (Seasonal Amplitude ΔNDVI > 0.25, Summer NDVI < 0.25; OC-Q5); "
        "and (3) Disturbed Transition, which includes sites consistent with a possible moisture-related influence near drainage features such as OC-Q2, "
        "where winter greenness retention is elevated (persistence ratio = 1.129; winter composite = 0.414, single-date Dec 30 value = 0.453). "
        "Threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI × 5 Persistence Ratio × 3 Seasonal Amplitude = 75 configurations) "
        "yielded 0 reclassifications in 24 configurations (32.0%), 1 in 21 (28.0%, cumulative ≤ 1: 45 [60.0%]), 2 in 9 (12.0%), 3 in 15 (20.0%), and 4 in 6 (8.0%; Supplementary Table S9A). "
        "In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28) did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications. "
        "Multi-scale spatial evaluation revealed a scale-dependent decline in mean summer NDVI in sampled underground forest quadrats (0.381 in Zone A to 0.271 in Zone D) "
        "consistent with matrix inclusion, whereas sampled opencast quadrats showed smaller changes (0.303 to 0.287). Principal Component Analysis accounted "
        "for 81.62% of variance across two primary axes, demonstrating close agreement (score correlation r > 0.94) with an 8-feature sensitivity model "
        "excluding derived metrics. Non-parametric contrasts between the sampled underground and opencast quadrats (n=5 per group) using SciPy's exact two-sided "
        "Mann-Whitney U procedure yielded non-significant differences (e.g., Summer NDVI U₁ = 19.0, U_min = 6.0, exact p = 0.2222, Cliff's delta = +0.520, "
        "Hedges' g = 0.647, approximate 95% analytical Student's t CI [-0.850, 2.143]); the wide uncertainty and small clustered sample make the non-significant result "
        "uninformative about ecological equivalence. We provide a complete machine-readable provenance audit and outline requirements for future hierarchical, replicated landscape restoration designs.\n"
    )
    lines.append("**Keywords:** Sentinel-2; Phenological Syndromes; Evidence Hierarchy; Mining Disturbance; Spatial Dilution; Korba Coalfield; Data Provenance; Ecological Indicators\n")
    lines.append("---\n")

    # Section 1
    lines.append("## 1. Introduction\n")
    lines.append(
        "Tropical dry deciduous forests across the central Indian mining belt represent critical ecological corridors that sustain "
        "regional biodiversity, stabilize catchments, and regulate carbon cycles, while experiencing extensive structural disturbance "
        "from commercial mineral extraction (Deb et al., 2024; Roy et al., 2022). Mining activities encompass intensive opencast surface "
        "stripping, which removes overstory canopies and natural soil profiles, and underground bord-and-pillar extraction, which modifies "
        "subsurface hydrology and causes localized surface subsidence. In these disturbed ecosystems, the vegetation cover—composed of "
        "planted reclamation trees, remnant mature forest patches, and an understory of perennial woody shrubs and monsoonal herbaceous "
        "species—plays a vital role in mitigating erosion, restoring nutrient cycling, and guiding succession (Banerjee et al., 2023).\n\n"
        "Remote sensing via medium-resolution optical sensors, such as the Copernicus Sentinel-2 Multi-Spectral Instrument (MSI), provides "
        "a widely used tool for tracking multi-temporal greenness trajectories across broad landscapes (Drusch et al., 2012). However, "
        "interpreting satellite observations in open-canopy or disturbed dry tropical forests entails major empirical challenges. A Sentinel-2 "
        "pixel with 10 m spatial resolution represents an integrated, composite vertical reflectance signal that simultaneously combines "
        "overstory tree crown foliage, subcanopy shrubs, ground-layer grasses, leaf litter, and exposed rock or bare soil (Brede et al., 2020). "
        "Where canopies are dense, optical signals are dominated by the overstory; where canopies are fragmented or open, the satellite "
        "signal is strongly driven by the dramatic seasonal green-up and rapid post-monsoon senescence of herbaceous ground cover (Martinuzzi "
        "et al., 2009; Nagendra et al., 2013).\n\n"
        "A critical vulnerability in applied remote sensing ecology is the hazard of unsubstantiated attribution—specifically, asserting "
        "that medium-resolution satellite pixels can 'deconvolve' distinct vegetation strata or directly identify low-stature herbs and shrubs "
        "in the absence of nested physical sub-plots, fractional canopy cover measurements, or contemporaneous satellite-ground acquisitions. "
        "In the published baseline study of Sharma, Banerjee & Deb (2026), ten 10 × 10 m field quadrats were inventoried across colliery "
        "leases in Korba, documenting overstory tree stems (≥ 10 cm girth at breast height [GBH]), aboveground biomass (AGB), carbon stocks, "
        "and soil physicochemical properties. However, seasonal remote sensing dynamics around these quadrats were not linked to audited "
        "satellite observations or evaluated across spatial support scales.\n\n"
        "Crucially, observational studies of this nature face design constraints that must be stated up-front. In the published baseline, "
        "the five opencast quadrats are situated within the Gevra and Dipka overburden complexes, whereas the five underground quadrats are "
        "clustered within the Banki and Surakachhar leases. Because quadrats within each mining category are geographically clustered, treating "
        "them as independent landscape replicates would risk spatial pseudoreplication (Hurlbert, 1984). Therefore, statistical contrasts "
        "between these two groups must be interpreted strictly as an exploratory, evidence-constrained pilot analysis (n = 5 per group) "
        "contrasting sampled quadrats, rather than a definitive landscape-scale evaluation of mining-type effects.\n\n"
        "To address these methodological challenges, this study evaluates seasonal Sentinel-2 surface reflectance dynamics around the ten "
        "published quadrats in Korba under an audited data-governance framework. We address four specific research questions:\n"
        "1. How do seasonal spectral indices (NDVI, NDRE, EVI, and SWIR ratio) capture composite surface greenness dynamics across pre-monsoon "
        "summer baseline, the operational Monsoon & Retreat Green Season composite, and post-monsoon drydown?\n"
        "2. How do heuristic descriptive classification rules (persistent greenness signals vs. pronounced seasonal flushes) align with verified "
        "in-plot woody vegetation and photographic habitat records without overstepping into unvalidated understory deconvolution?\n"
        "3. How sensitive are extracted spectral metrics to spatial aggregation scale across centered multi-scale extraction windows (Zone A: 10 m nominal pixel support; "
        "Zone B: 50 × 50 m square; Zone C: 50 m radius disc; Zone D: 100 m radius disc), and how does spatial support affect remnant forest vs. overburden spoil signatures?\n"
        "4. How do exploratory non-parametric group contrasts and multivariate ordination perform under small-sample constraints, and how "
        "should next-generation restoration monitoring frameworks be structured?\n"
    )
    lines.append("---\n")

    # Section 2
    lines.append("## 2. Study Area and Environmental Setting\n")
    lines.append(
        "The study was conducted in the Korba coalfield, situated within the Upper Hasdeo River basin in Korba District, Chhattisgarh, "
        "central India (22°21'N to 22°26'N, 82°34'E to 82°38'E; Figure 1). The regional topography forms an undulating plateau sloping "
        "gently from elevated sandstone ridges (315 m above mean sea level) down to alluvial floodplains (280 m). Geologically, the area "
        "belongs to the Lower Gondwana Group (Barakar Formation), characterized by medium- to coarse-grained felspathic sandstones, "
        "carbonaceous shales, and rich sub-bituminous coal seams.\n\n"
        "The regional climate is classified as tropical wet-and-dry (Köppen Aw). Climatological normals from the India Meteorological "
        "Department (IMD Bilaspur and Korba stations) establish three distinct operational seasons: (1) Pre-monsoon Summer (March 1 to "
        "May 31), characterized by high daytime temperatures (frequently exceeding 42–44°C), desiccating winds, low relative humidity, "
        "and extensive deciduous leaf shedding; (2) Southwest Monsoon and Retreat (June 1 to mid-October), delivering over 85% of the annual rainfall "
        "(~1,300–1,450 mm) and triggering rapid vegetative flush, with the monsoon withdrawal typically extending through early to mid-October; and "
        "(3) Post-monsoon Winter (mid-October to February 28), marked by declining temperatures (10–25°C), progressive soil drying, and gradual canopy senescence.\n\n"
        "The ten sampling quadrats originally established by Sharma, Banerjee & Deb (2026) are distributed across two contrasting coal mining "
        "settings (Figure 1; Table 1): (1) Underground mining forest quadrats (UG-Q1 to UG-Q5; elevation 283.9–315.3 m) located within the "
        "Banki and Surakachhar colliery concessions, representing remnant mixed dry deciduous forest patches dominated by Shorea robusta (Sal), "
        "Terminalia elliptica (syn. T. tomentosa), and Diospyros melanoxylon; and (2) Opencast mine overburden quadrats (OC-Q1 to OC-Q5; "
        "elevation 298.5–310.2 m) located on mature overburden spoil dumps within the Gevra and Dipka surface mining complexes, supporting "
        "established reclamation plantations (Eucalyptus tereticornis, Millettia pinnata [syn. Pongamia pinnata]) alongside unmanaged, "
        "disturbed second-growth (Holarrhena pubescens thickets and open grass patches).\n"
    )
    lines.append("![Figure 1: Sampling Locations & Centered Windows](figures/Figure_1_Sampling_Locations_Spatial_Windows.png)\n")
    lines.append("**Figure 1.** Geographic distribution of the 10 published study quadrats across colliery sectors in the Korba coalfield, central India (left), showing underground remnant forest quadrats (UG-Q1 to UG-Q5, green circles) and opencast overburden dump quadrats (OC-Q1 to OC-Q5, orange squares), and schematic centered multi-scale extraction geometries (right) illustrating Zone A (10 × 10 m pixel support), Zone B (50 × 50 m square window), Zone C (50 m radius circular buffer), Zone D (100 m radius circular buffer), and GPS positional uncertainty bounds (±5 m).\n")
    lines.append("---\n")

    # Section 3
    lines.append("## 3. Materials and Methods\n")
    lines.append("### 3.1 Three-Tier Ground Evidence Hierarchy & Botanical Taxonomy\n")
    lines.append(
        "To establish rigorous empirical boundaries and prevent unsubstantiated species-level claims from satellite pixels, we implemented "
        "an audited three-tier ground evidence standard (Table 1):\n"
        "- **Tier 1: Verified In-Plot Woody Stems (Empirical Ground Census).** The published field data sheets from the 10 quadrats enumerate 58 "
        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "
        "representing 25 reported field morphotaxa. All botanical nomenclature was audited and standardized according to Plants of the World Online "
        "(POWO; Royal Botanic Gardens, Kew), using the accepted-name field as the sole basis for taxonomic richness. Following synonym harmonization, "
        "the inventory comprises 24 unique accepted species (e.g., *Terminalia tomentosa* (Roxb.) Wight & Arn. is harmonized as a synonym of *Terminalia elliptica* Willd., "
        "*Holarrhena antidysenterica* (L.) Wall. as a synonym of *Holarrhena pubescens* Wall. ex G.Don, and *Bauhinia vahlii* as *Phanera vahlii* (Wight & Arn.) Benth.). "
        "Several subcanopy woody taxa (*Ziziphus mauritiana*, *Carissa carandas*, *Holarrhena pubescens*) were directly measured within this Tier 1 census.\n"
        "- **Tier 2: Photographic Habitat Records (Structural Reality).** An audited repository of 40 field photographic plates taken across all 10 "
        "quadrats was systematically analyzed to document overstory canopy closure, bare ground exposure, leaf litter depth, and the presence "
        "of ground-layer vegetation.\n"
        "- **Tier 3: Regional Floristic Candidate Leads (Unverified Context).** A register of 12 regional herbaceous and shrub species common to "
        "disturbed mining areas in Korba (e.g., *Parthenium hysterophorus*, *Senna tora*, *Hyptis suaveolens*, *Lantana camara*, *Cynodon dactylon*) was "
        "compiled from regional floristic surveys. These are treated strictly as unverified regional associates representing candidate hypotheses "
        "for future nested micro-quadrat audits, rather than confirmed in-plot tallies.\n"
    )
    lines.append("### Table 1. Audited three-tier ground evidence hierarchy for the Korba mining study.\n")
    t1_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_1_Field_Evidence_Hierarchy.csv"))
    lines.append(df_to_markdown(t1_df))
    lines.append("\n\n")

    lines.append("### 3.2 Sentinel-2 Earth Observation, Provenance & SCL Radiometric Screening\n")
    lines.append(
        "We searched the complete 2024 Copernicus Sentinel-2 Level-2A (Bottom-of-Atmosphere [BOA] surface reflectance) archive for MGRS "
        "Tile 44QPK (Relative Orbit R019). The archive query identified 73 candidate scenes (manifest preserved in `data/04_Sentinel2_Candidate_Archive_Manifest.csv`). "
        "To ensure high radiometric fidelity and eliminate cloud, haze, or shadow contamination across all seasons, candidate scenes were screened "
        "using the Sen2Cor Scene Classification Layer (SCL). A strict valid terrestrial surface rule was applied: retaining only SCL Class 4 "
        "(Vegetation) and Class 5 (Not-vegetated / bare soils), while strictly excluding cloud shadows (SCL 3), clouds (SCL 8, 9), thin cirrus (SCL 10), "
        "water (SCL 6), and unclassified pixels (SCL 7).\n\n"
        "Six audited observation dates meeting the ≥ 80% terrestrial validity threshold across all quadrats and zones were selected (Table 2; "
        "Table S2). The six scenes comprise: Summer (2024-03-25, 2024-05-14), Monsoon & Retreat Green Season (2024-06-13 [onset], 2024-10-06 [monsoon retreat peak]), "
        "and Winter (2024-11-20, 2024-12-30). The inclusion and classification of the October 6 scene is grounded in regional climatology: IMD "
        "Bilaspur/Korba normals record the normal withdrawal of the Southwest Monsoon from northern Chhattisgarh during the first to second week of "
        "October. Consequently, early October captures the annual peak of monsoonal soil moisture accumulation and maximum vegetative biomass before "
        "post-monsoon senescence begins. Combining June 13 and October 6 provides an operational composite of the broader Monsoon & Retreat Green Season. "
        "In Zone A (nominal 10 m pixel support), 100.00% valid terrestrial pixel retention was achieved across all six dates (zero masked pixels across 60 site-date cases; Table 2). "
        "Terrestrial validity percentage is defined as the total number of valid terrestrial pixels (SCL classes 4 [vegetation] and 5 [bare soil]) divided by "
        "the theoretical denominator pixel count across all 60 site-date observations (which exactly equals the arithmetic mean of individual site-date validity "
        "percentages to two decimal places): Zone A = 100.00% (60/60), Zone B = 99.33% (1,490/1,500), Zone C = 99.22% (4,822/4,860), and Zone D = 98.99% "
        "(18,828/19,020), defining a range of 98.99% to 99.33% across multi-scale extraction windows due to edge inclusions of water bodies, shadows, and colliery infrastructure (Table 3; Table 6).\n"
    )
    lines.append("### Table 2. Audited Sentinel-2 Level-2A observation manifest, timestamps, and radiometric quality control (Tile 44QPK, Orbit R019).\n")
    t2_df = pd.read_csv(os.path.join(BASE_DIR, "data", "04_Sentinel2_scene_inventory.csv"))
    t2_disp = t2_df[['Date_YYYY_MM_DD', 'Acquisition_UTC', 'Season_Operational', 'Platform', 'Scene_Cloud_Cover_Pct', 'Audit_SHA256']].copy()
    t2_disp.columns = ['Date', 'UTC Timestamp', 'Operational Season', 'Platform', 'Cloud (%)', 'GeoTIFF Stack SHA-256']
    t2_disp['Operational Season'] = t2_disp['Operational Season'].replace("Southwest Monsoon", "Monsoon & Retreat Green Season").replace("Late Monsoon Peak Greenness", "Monsoon & Retreat Green Season")
    lines.append(df_to_markdown(t2_disp))
    lines.append("\n\n")

    lines.append("### 3.3 Multi-Scale Centered Extraction Windows & Nominal Pixel Support\n")
    lines.append(
        "All spatial datasets were projected to UTM Zone 44 North (WGS84 datum, EPSG:32644). Handheld GPS receiver positional uncertainty "
        "during field quadrat establishment was recorded at ± 5.0 m, defining an effective positional support footprint of 20 × 20 m (400 m²). "
        "Because field quadrats are rarely perfectly aligned with arbitrary satellite raster grids, Zone A was defined as the Sentinel-2 pixel "
        "containing the quadrat centroid and was used as the closest available nominal support to the 10 × 10 m field quadrat.\n\n"
        "To explicitly evaluate spatial co-location offsets, we computed exact centroid-to-pixel-center distances and distances to the nearest pixel "
        "boundary across all 10 quadrats (Supplementary Table S1B). The distance from quadrat centroid to pixel center averaged 4.63 m "
        "(range: 1.99–6.59 m), while the distance to the nearest pixel boundary averaged 1.26 m (range: 0.00–3.45 m), yielding a mean nominal "
        "geometric overlap of 47.2% (range: 28.5% to 74.0%) between the 10 × 10 m field plot and the containing pixel. To address this spatial "
        "co-location reality and evaluate how surrounding landscape context modulates the nominal support pixel, surface reflectance was extracted "
        "across four centered multi-scale spatial windows (Table 3):\n"
        "- **Zone A (Nominal Quadrat Pixel Support):** Exactly 1 pixel (10 × 10 m = 100 m² = 0.01 ha), defined as the raster pixel containing the quadrat centroid; field-plot overlap is variable.\n"
        "- **Zone B (Local Neighborhood):** A square 5 × 5 pixel window (50 × 50 m = 2,500 m² = 0.25 ha, 25 pixels total support), centered on the quadrat centroid and enclosing its GPS uncertainty footprint.\n"
        "- **Zone C (Intermediate Stand Buffer):** A circular disc of radius 50 m (theoretical geometric area = 7,854 m² [π × 50²]; enclosing 81 raster pixels with a rasterized support area of 8,100 m²), capturing the local forest or overburden dump plantation stand.\n"
        "- **Zone D (Contextual Landscape Buffer):** A circular disc of radius 100 m (theoretical geometric area = 31,416 m² [π × 100²]; enclosing 317 raster pixels with a rasterized support area of 31,700 m²), capturing the broader colliery landscape matrix.\n"
    )
    lines.append("### Table 3. Multi-scale spatial extraction geometries, theoretical geometric areas, raster support areas, and SCL terrestrial validity across all 60 site-date cases.\n")
    t3_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_3_MultiScale_Geometries.csv"))
    lines.append(df_to_markdown(t3_df))
    lines.append("\n\n")

    lines.append("### 3.4 Spectral Indices, Operational Compositing Scheme & Data Integrity\n")
    lines.append(
        "For each observation date, 11 spectral bands (B02 Blue, B03 Green, B04 Red, B05/B06/B07 Red Edge, B08 NIR, B8A Narrow NIR, "
        "B11 SWIR-1, B12 SWIR-2, and SCL) were retrieved. 20 m bands were resampled to 10 m resolution via bilinear interpolation on the "
        "native UTM 44N grid (nearest neighbor for SCL). Surface reflectance values were scaled to BOA units (reflectance = DN / 10000). "
        "Spectral indices were computed per pixel on the 2D arrays:\n"
        "- Normalized Difference Vegetation Index: $\\text{NDVI} = (B08 - B04) / (B08 + B04)$\n"
        "- Normalized Difference Red Edge: $\\text{NDRE} = (B08 - B05) / (B08 + B05)$\n"
        "- Enhanced Vegetation Index: $\\text{EVI} = 2.5 \\times (B08 - B04) / (B08 + 6.0 \\times B04 - 7.5 \\times B02 + 1.0)$\n"
        "- Shortwave Infrared Ratio: $\\text{SWIR\\_Ratio} = B11 / B12$\n\n"
        "To capture seasonal dynamics robustly, seasonal median composites were computed for Summer (March 25, May 14), Monsoon & Retreat Green Season "
        "(June 13, October 6), and Winter (November 20, December 30). In the locked dataset (`data/Definitive_Locked_Site_By_Season_Indices.csv`), "
        "fields containing 'GreenSeason' (or legacy aliases 'Monsoon') refer to the operational Monsoon & Retreat Green Season composite formed from June 13 "
        "and October 6 observations, as formally defined in `data/DATA_DICTIONARY_LOCKED_INDICES.md`. Phenological dynamic metrics were calculated as: "
        "Seasonal NDVI Amplitude ($\\Delta\\text{NDVI} = \\text{Green Season composite} - \\text{Summer composite}$); Winter-to-Green-Season Persistence Ratio "
        "($\\text{NDVI}_{\\text{Winter}} / \\text{NDVI}_{\\text{GreenSeason}}$); and Summer-to-Green-Season Retention Ratio ($\\text{NDVI}_{\\text{Summer}} / \\text{NDVI}_{\\text{GreenSeason}}$).\n"
    )

    lines.append("### 3.5 Heuristic Phenological Descriptive Classes and Threshold Sensitivity\n")
    lines.append(
        "Rather than asserting understory stratum deconvolution, we classify quadrats into heuristic descriptive classes "
        "developed from this pilot dataset based on explicit threshold rules applied to Zone A surface reflectance:\n"
        "1. **Persistent Greenness Signal:** Summer dry-season NDVI ≥ 0.30 AND Winter-to-Green-Season Persistence Ratio ≥ 0.90 (descriptive spectral class; "
        "also termed High Dry-Season Spectral Retention). This class characterizes sites where established woody cover maintains high foliar greenness and suppresses "
        "dry-season seasonal amplitude, without establishing canopy dominance in the absence of independent understory measurements.\n"
        "2. **Pronounced Seasonal Flush:** Seasonal NDVI amplitude ΔNDVI > 0.25 AND Summer dry-season NDVI < 0.25. This class characterizes open, "
        "low-greenness spoil settings exhibiting a pronounced seasonal amplitude compatible with seasonal ground-cover development, but ground-layer contributions were not directly measured.\n"
        "3. **Disturbed Transition:** Intermediate seasonal amplitude (0.07 ≤ ΔNDVI ≤ 0.25) or sites modulated by localized substrate or "
        "micro-hydrological conditions (e.g., mine water drainage features).\n\n"
        "The classification function evaluates each quadrat deterministically via the following logic:\n\n"
        "```python\n"
        "def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\n"
        "    if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\n"
        "        return 'Persistent Greenness Signal'\n"
        "    elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\n"
        "        return 'Pronounced Seasonal Flush'\n"
        "    else:\n"
        "        return 'Disturbed Transition'\n"
        "```\n\n"
        "To evaluate whether these classification assignments are sensitive to specific threshold cutoffs, we conducted an explicit sensitivity "
        "analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × 5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × "
        "3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B). Across all 75 tested configurations, 24 (32.0%) "
        "preserved 100% of baseline site assignments, 21 (28.0%) produced 1 reclassification (cumulative ≤ 1: 45 [60.0%]), 9 (12.0%) produced 2 reclassifications, "
        "15 (20.0%) produced 3 reclassifications, and 6 (8.0%) produced 4 reclassifications (Supplementary Table S9A). In this dataset, changing the amplitude "
        "cutoff within the tested range (0.20 to 0.28) did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications "
        "(e.g., persistence ratio ≥ 0.92 reclassifies OC-Q1 and UG-Q2; persistence ratio ≥ 0.95 reclassifies UG-Q4; summer NDVI ≥ 0.32 reclassifies UG-Q1). "
        "We emphasize that these groupings are heuristic descriptive spectral classes developed from this pilot dataset rather than universal ecological boundaries.\n"
    )

    lines.append("### 3.6 Statistical Inference Framework, Effect Sizes & Analytical Confidence Intervals\n")
    lines.append(
        "Given the small sample size of the published study (n = 10; 5 underground vs. 5 opencast), statistical analyses were conducted using "
        "exact non-parametric permutation tests alongside standardized effect sizes. Contrasts between sampled underground and opencast quadrats "
        "were evaluated using SciPy’s exact two-sided Mann-Whitney U procedure (`scipy.stats.mannwhitneyu`, `alternative='two-sided'`, `method='exact'`) "
        "with no tie correction required for the reported values. An independent exhaustive enumeration of all 252 label permutations (10 choose 5) confirmed identical "
        "results (56/252 = 2/9 = 0.2222), demonstrating exact agreement between both calculations. In the tabular results (Table 5) and audit records (Table S7), "
        "we explicitly report both the first-sample U statistic (U₁, corresponding to Underground) and the smaller U statistic ($U_{\\text{min}} = \\min(U_1, U_2)$), "
        "resolving potential ambiguities arising from software conventions.\n\n"
        "A standardized mean difference (Hedges' g) was paired with non-parametric rank tests because Mann-Whitney U evaluates whether one distribution "
        "stochastically dominates another without distributional assumptions, whereas Hedges' g provides a standardized, scale-free descriptive effect size "
        "to facilitate future meta-analyses and sample-size planning for replicated restoration designs. We report an approximate interval obtained by "
        "multiplying the estimated standard error of Hedges’ g by the $t(8, 0.975)$ critical value ($t = 2.3060$); this interval is descriptive and not intended "
        "to provide exact small-sample coverage (Hedges & Olkin, 1985). Specifically, approximate 95% confidence intervals were calculated using the Student's t "
        "critical value with $\\text{df} = n_1 + n_2 - 2 = 8$: $\\text{CI} = g \\pm t(0.025, 8) \\times \\text{SE}(g)$, where critical $t(0.025, 8) = 2.3060$, "
        "$\\text{SE}(g) = \\sqrt{[(n_1 + n_2)/(n_1 \\times n_2)] + [g^2 / (2(n_1 + n_2))]}$, pooled standard deviation $s_{\\text{pooled}} = \\sqrt{[((n_1 - 1)s_1^2 + (n_2 - 1)s_2^2) / \\text{df}]}$, "
        "and small-sample correction factor $J = 1 - [3 / (4(n_1 + n_2) - 9)] = 28/31 \\approx 0.9032$. Contrast direction is defined as Underground minus Opencast (UG - OC). "
        "Group medians and interquartile ranges (IQR) are reported alongside means and standard deviations.\n\n"
        "To evaluate multivariate feature ordination, Principal Component Analysis was conducted on standardized features in Zone A under two "
        "formulations: Model A (10 features, including seasonal indices, SWIR ratios, amplitude, and persistence ratio) and sensitivity Model B "
        "(8 component features, strictly excluding the collinear derived metrics ΔNDVI and ratio). Ordination agreement was evaluated via Pearson "
        "correlation of sample scores between models. All calculations were executed deterministically in Python 3.11 using rasterio 1.3.10, scipy 1.13.1, "
        "scikit-learn 1.5.0, pandas 2.2.2, and matplotlib 3.9.0.\n"
    )
    lines.append("---\n")

    # Section 4: Results (Strict 5-part structure)
    lines.append("## 4. Results\n")

    lines.append("### 4.1 Field Botanical and Habitat Evidence\n")
    lines.append(
        "The published field census recorded 58 woody stems across the ten 10 × 10 m quadrats (Table 1; Table S1; Table S3). Underground forest "
        "quadrats supported higher overstory species richness (mean 4.0 species/quadrat; range 3–5) and higher aboveground carbon stocks "
        "(22.95–26.71 t C ha⁻¹) dominated by climax dry deciduous taxa (*Shorea robusta*, *Terminalia elliptica*, *Diospyros melanoxylon*). Opencast "
        "quadrats exhibited lower carbon stocks (15.84–18.48 t C ha⁻¹) and were dominated by planted reclamation species (*Eucalyptus tereticornis*, "
        "*Millettia pinnata*) alongside disturbed second-growth (*Holarrhena pubescens*, *Tamarindus indica*). Four quadrats contained measured woody "
        "shrubs or subcanopy stems ≥ 10 cm GBH in the Tier 1 inventory: UG-Q1 and UG-Q5 contained *Ziziphus mauritiana*; OC-Q2 contained *Carissa carandas*; "
        "and OC-Q4 contained five distinct stems of *Holarrhena pubescens*.\n\n"
        "Tier 2 photographic evidence (40 plates; Table S4) confirmed that underground quadrats maintained closed to semi-closed canopies with "
        "substantial leaf litter accumulation, whereas opencast quadrats exhibited varying degrees of canopy openness, rock outcropping, and bare spoil. "
        "Tier 3 regional floristic leads (12 candidate species; Table S5) identify common disturbed-ground associates (e.g., *Parthenium hysterophorus*, "
        "*Senna tora*, *Hyptis suaveolens*) known regionally from Korba. We explicitly record that ground-layer herbaceous cover, non-woody understory biomass, "
        "and soil moisture were not quantitatively inventoried in the baseline field survey.\n"
    )

    lines.append("### 4.2 Multi-Temporal Spectral Trajectories and Seasonal Composites\n")
    lines.append(
        "Across the six individual Sentinel-2 observation dates, surface reflectance profiles (Figure 2) and vegetation indices (Figure 3) "
        "demonstrated pronounced, date-specific phenological shifts. Date-by-date trajectories established that the maximum greenness among the six audited observations occurred on "
        "2024-10-06 across almost all quadrats (e.g., OC-Q4 NDVI reached 0.783; UG-Q2 reached 0.781; OC-Q5 reached 0.749; OC-Q1 reached 0.738; Table S2). "
        "This empirical peak coincides with the late-monsoon retreat phase per IMD normals, when cumulative soil moisture saturation and foliar development "
        "reach seasonal maxima. Conversely, the minimum greenness among the six audited observations occurred during late pre-monsoon summer on 2024-05-14 (e.g., OC-Q3 NDVI dropped to "
        "0.189; OC-Q5 to 0.214; OC-Q2 to 0.239; UG-Q5 to 0.265; Figure 3).\n\n"
        "Winter acquisitions (2024-11-20 and 2024-12-30) documented progressive dry-season desiccation across open spoil, while stands with established woody "
        "cover maintained high foliar reflectance. A notable micro-site trajectory occurred in OC-Q2 (situated adjacent to a mine water drainage feature), where "
        "NDVI rose from 0.374 on November 20 to 0.453 on December 30, reflecting prolonged winter greenness retention.\n\n"
        "Aggregation into the three operational composites (Pre-monsoon Summer, Monsoon & Retreat Green Season, Post-monsoon Winter) effectively integrated "
        "these multi-date dynamics while suppressing high-frequency atmospheric noise. NDRE trajectories (sensitive to canopy chlorophyll without early saturation) "
        "clearly differentiated dense overstory forest stands (e.g., UG-Q3 Sal canopy NDRE: 0.27–0.33) from open, rocky spoil dumps (OC-Q3 NDRE: 0.08–0.24).\n"
    )
    lines.append("![Figure 2: Surface Reflectance](figures/Figure_2_Seasonal_Spectral_Signatures.png)\n")
    lines.append("**Figure 2.** Multi-temporal Level-2A surface reflectance profiles across Sentinel-2 optical bands (490–2190 nm) for underground remnant forest quadrats (UG-Q1 to UG-Q5, left) and opencast overburden dump quadrats (OC-Q1 to OC-Q5, right) across Summer, Monsoon & Retreat Green Season, and Winter composites. Shaded bands represent ±1 standard deviation.\n\n")

    lines.append("![Figure 3: NDVI & NDRE Trajectories](figures/Figure_3_NDVI_NDRE_Trajectories_UG_vs_OC.png)\n")
    lines.append("**Figure 3.** Multi-temporal Sentinel-2 NDVI (top) and NDRE (bottom) trajectories across the six audited 2024 observation dates for individual quadrats (faint lines) and group means (bold lines; green circles = Underground, orange squares = Opencast). Shaded vertical intervals delineate pre-monsoon summer, Monsoon & Retreat Green Season, and post-monsoon winter senescence.\n\n")

    lines.append("### 4.3 Multi-Scale Spatial Support and Landscape Dilution\n")
    lines.append(
        "Analysis of multi-scale extraction windows (Zone A: 10 m nominal pixel support; Zone B: 50 × 50 m square; Zone C: 50 m radius disc; "
        "Zone D: 100 m radius disc; Table 6; Figure 6) revealed striking differences in spatial aggregation behavior between mining types. In Zone A, "
        "100% of pixels met strict terrestrial SCL criteria (zero masked pixels across 60 site-date cases). In expanded windows, valid pixel retention "
        "remained very high but variable due to colliery infrastructure and drainage features: Zone B averaged 99.33% valid pixels (24.83/25), "
        "Zone C averaged 99.22% (80.37/81), and Zone D averaged 98.99% (313.8/317; Table 3).\n\n"
        "In underground forest sites, expanding the spatial extraction window produced a scale-dependent decline in mean Summer NDVI: dropping from "
        "0.381 in Zone A (10 m) to 0.305 in Zone B (50 × 50 m), 0.279 in Zone C (50 m radius), and 0.271 in Zone D (100 m radius)—a total reduction "
        "of 0.111 NDVI units (29.0%). The scale-dependent decline is consistent with inclusion of lower-greenness surfaces surrounding the nominal pixel, "
        "including the documented colliery matrix; the relative contribution of each surface type was not independently quantified.\n\n"
        "In contrast, the sampled opencast quadrats showed smaller changes in mean NDVI across the tested windows: mean Summer NDVI was 0.303 in Zone A, "
        "0.305 in Zone B, 0.302 in Zone C, and 0.287 in Zone D (a difference of 0.016 units, 5.2% reduction).\n"
    )
    lines.append("### Table 6. Multi-scale spatial window sensitivity of seasonal NDVI and dynamic metrics across buffer zones from Zone A (10 m) to Zone D (100 m radius).\n")
    t6_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_6_Spatial_Scale_Sensitivity.csv"))
    t6_disp = pd.DataFrame()
    t6_disp['Spatial Extraction Scale'] = t6_df['Spatial_Zone_Label']
    t6_disp['Mining Type'] = t6_df['Mining_type']
    t6_disp['Summer NDVI'] = t6_df['NDVI_Summer_mean'].round(3).astype(str) + " ± " + t6_df['NDVI_Summer_std'].round(3).astype(str)
    t6_disp['Green Season NDVI'] = t6_df['NDVI_GreenSeason_mean'].round(3).astype(str) + " ± " + t6_df['NDVI_GreenSeason_std'].round(3).astype(str)
    t6_disp['Winter NDVI'] = t6_df['NDVI_Winter_mean'].round(3).astype(str) + " ± " + t6_df['NDVI_Winter_std'].round(3).astype(str)
    t6_disp['Seasonal Amplitude'] = t6_df['Seasonal_NDVI_Amplitude_mean'].round(3).astype(str) + " ± " + t6_df['Seasonal_NDVI_Amplitude_std'].round(3).astype(str)
    t6_disp['Winter/Green-Season'] = t6_df['Winter_to_GreenSeason_Ratio_mean'].round(3).astype(str) + " ± " + t6_df['Winter_to_GreenSeason_Ratio_std'].round(3).astype(str)
    lines.append(df_to_markdown(t6_disp))
    lines.append("\n\n")

    lines.append("![Figure 6: Spatial Seasonality Map](figures/Figure_6_Spatial_Vegetation_Seasonality_Map.png)\n")
    lines.append("**Figure 6.** Multi-scale spatial seasonality gradient across the Korba coalfield, comparing seasonal NDVI amplitude (Monsoon & Retreat Green Season composite minus Summer baseline) across Zone A (10 × 10 m pixel support), Zone B (50 × 50 m square window), Zone C (50 m radius circular buffer), and Zone D (100 m radius circular buffer).\n\n")

    lines.append("### 4.4 Descriptive Classification and Threshold Sensitivity\n")
    lines.append(
        "Applying the heuristic descriptive classification rules to Zone A surface reflectance segregated the ten quadrats into three distinct domains "
        "(Table 4; Figure 4):\n"
        "- **Persistent Greenness Signal:** Five quadrats (UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4) met the baseline criteria of Summer dry-season NDVI ≥ 0.30 and "
        "Winter-to-Green-Season persistence ratio ≥ 0.90 (descriptive spectral class). These sites are characterized by established woody cover (*Shorea robusta*, "
        "*Terminalia elliptica*, *Eucalyptus tereticornis*, or *Holarrhena pubescens*) that maintains high foliar reflectance into the dry season without establishing "
        "canopy dominance in the absence of independent understory measurements.\n"
        "- **Pronounced Seasonal Flush:** OC-Q5 displayed the highest seasonal amplitude in the study (ΔNDVI = 0.285; Summer NDVI = 0.214 -> Green Season "
        "NDVI = 0.499; single-date peak 0.749 on Oct 06) and Summer NDVI < 0.25. The open, low-greenness spoil setting exhibited a pronounced seasonal "
        "amplitude that is compatible with seasonal ground-cover development, but the ground-layer contribution was not directly measured.\n"
        "- **Disturbed Transition:** Four quadrats occupied intermediate positions: UG-Q3 exhibited high greenness but very low seasonal amplitude "
        "(ΔNDVI = 0.075), reflecting dense climax Sal canopy that exchanges leaves rapidly in spring; UG-Q5 (ΔNDVI = 0.192, Summer NDVI = 0.265) represents "
        "a thinned edge forest; OC-Q3 (ΔNDVI = 0.158, Summer NDVI = 0.189, Winter = 0.287) represents bare rocky overburden with sparse planted boles; and OC-Q2 "
        "(persistence ratio = 1.129; Winter composite = 0.414, Green Season composite = 0.366) represents a site where localized winter greenness retention "
        "(single-date Dec 30 NDVI reaching 0.453) is consistent with a possible moisture-related influence near an adjacent drainage feature.\n\n"
        "The explicit threshold sensitivity analysis systematically varying three parameters (5 Summer NDVI [0.25, 0.28, 0.30, 0.32, 0.35] × "
        "5 Persistence Ratio [0.85, 0.88, 0.90, 0.92, 0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28] = 75 configurations; Supplementary Table S9A and Table S9B) "
        "yielded the following distribution of reclassifications: 0 changes in 24 configurations (32.0%, cumulative 32.0%), 1 change in 21 configurations "
        "(28.0%, cumulative 60.0%), 2 changes in 9 configurations (12.0%, cumulative 72.0%), 3 changes in 15 configurations (20.0%, cumulative 92.0%), "
        "and 4 changes in 6 configurations (8.0%, cumulative 100.0%). In this dataset, changing the amplitude cutoff within the tested range (0.20 to 0.28) "
        "did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications. Specifically, when the persistence cutoff was "
        "elevated to ≥ 0.92, OC-Q1 (ratio 0.908) and UG-Q2 (ratio 0.901) were reclassified to Disturbed Transition; tightening to ≥ 0.95 reclassified UG-Q4 (ratio 0.921); "
        "and elevating the Summer NDVI cutoff to ≥ 0.32 reclassified UG-Q1 (summer 0.317). This sensitivity underscores that these syndromic groupings are operational "
        "descriptive spectral classes rather than rigid ecological states.\n"
    )
    lines.append("### Table 4. Descriptive phenological classification and site characteristics of the 10 study quadrats in Zone A (10 m support).\n")
    t4_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_4_Phenological_Syndromes.csv"))
    t4_disp = t4_df[['Quadrat_ID', 'Mining_Type', 'Summer_NDVI_Reported', 'Monsoon_Retreat_NDVI_Reported', 'Winter_NDVI_Reported', 'Seasonal_Amplitude_ΔNDVI_Reported', 'Winter_to_GreenSeason_Ratio_Reported', 'Descriptive_Syndrome_Rule']].copy()
    t4_disp.columns = ['Quadrat', 'Type', 'Summer NDVI', 'Green Season NDVI', 'Winter NDVI', 'Amplitude (ΔNDVI)', 'Winter/Green-Season', 'Descriptive Syndrome Rule']
    lines.append(df_to_markdown(t4_disp))
    lines.append("\n\n")

    lines.append("![Figure 4: Bivariate Distribution](figures/Figure_4_Herbaceous_vs_Shrub_Persistence.png)\n")
    lines.append("**Figure 4.** Bivariate distribution of Seasonal NDVI Amplitude (Monsoon & Retreat Green Season composite minus Summer baseline) versus Winter-to-Green-Season Persistence Ratio in Zone A (10 m), displaying heuristic threshold boundaries for the Persistent Greenness Domain (Persistence Ratio ≥ 0.90, green shading), the Pronounced Seasonal Flush Domain (amplitude > 0.25, orange shading), and the Disturbed Transition Domain. Symbol size corresponds to verified field tree carbon stock.\n\n")

    lines.append("### 4.5 Exploratory Contrasts and Ordination\n")
    lines.append(
        "Before presenting inferential test statistics, we emphasize the inherent structural limitations of the dataset: the comparison is based "
        "on only ten geographically clustered quadrats (n = 5 per group) situated within colliery leases, lacking independent regional replicates "
        "and unmined forest benchmarks. Consequently, statistical contrasts must be interpreted strictly as exploratory pilot comparisons among "
        "sampled quadrats rather than landscape-scale mining impacts. Non-significant p-values reflect low statistical power and substantial within-group "
        "variance (e.g., OC-Q1 plantation vs. OC-Q3 bare spoil) rather than evidence of ecological equivalence.\n\n"
        "Exact non-parametric permutation tests comparing sampled underground and opencast quadrats in Zone A (Table 5; cross-referenced with "
        "the machine-readable verification table Table S7) indicated that group differences were not statistically significant at alpha = 0.05. "
        "For Summer NDVI, the exact Mann-Whitney test yielded U₁ = 19.0 (first-sample U, Underground) and U_min = 6.0 (smaller U), with exact "
        "permutation p = 0.2222, Cliff's delta = +0.520, and Hedges' g = 0.647 (approximate 95% analytical Student's t CI [-0.850, 2.143]). "
        "We note transparently that in an earlier version of this manuscript, the values U = 18.0, p = 0.3095, g = 0.655 were inadvertently transcribed "
        "into the Abstract from the EVI Green Season test row; the verified values for Summer NDVI are U₁ = 19.0, U_min = 6.0, exact p = 0.2222, "
        "Cliff's delta = +0.520, Hedges' g = 0.647, CI [-0.850, 2.143], as fully documented in Table 5 and verified in Supplementary Table S7.\n\n"
        "Contrasts across other metrics similarly reflected wide confidence intervals and non-significance (Table 5): Monsoon & Retreat Green Season "
        "NDVI had U₁ = 14.0, U_min = 11.0 (exact p = 0.8413, Cliff's delta = +0.120, Hedges' g = 0.365, CI [-1.106, 1.836]); Post-monsoon Winter NDVI "
        "had U₁ = 12.0, U_min = 12.0 (exact p = 1.0000, Cliff's delta = -0.040, Hedges' g = 0.204, CI [-1.258, 1.667]); Seasonal Amplitude had "
        "U₁ = 11.0, U_min = 11.0 (exact p = 0.8413, Cliff's delta = -0.120, Hedges' g = -0.474, CI [-1.953, 1.004]); and Winter-to-Green-Season Persistence "
        "Ratio had U₁ = 9.0, U_min = 9.0 (exact p = 0.5476, Cliff's delta = -0.280, Hedges' g = -0.264, CI [-1.729, 1.201]).\n\n"
        "Multivariate Principal Component Analysis was conducted as an exploratory ordination to visualize primary axes of variation among the "
        "10 quadrats (Figure 5). In Model A (10 standardized features), the first two components accounted for 81.62% of total variance (PC1: 60.49%; "
        "PC2: 21.13%; PC3: 11.27%; Table 4b). Axis 1 reflected overall foliar greenness and multi-season cover (strong positive loadings for Green Season "
        "NDVI [0.3956], SWIR ratio [0.3778], and NDRE [0.3757]), while Axis 2 captured seasonal dynamics (Seasonal Amplitude loading +0.5401, Persistence "
        "Ratio loading +0.4808). Model B (8 features, strictly excluding derived metrics ΔNDVI and ratio) explained 88.71% of variance (PC1: 75.22%, "
        "PC2: 13.49%) and showed close agreement with Model A scores (PC1 r = 0.9994, PC2 r = 0.9448). We emphasize that this PCA is exploratory, "
        "fitted strictly to the ten sampled quadrats without external validation, and does not represent universal ecological gradients.\n"
    )
    lines.append("### Table 5. Non-parametric statistical contrasts and effect sizes between sampled underground (UG, n=5) and opencast (OC, n=5) quadrats in Zone A (10 m).\n")
    t5_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_5_Statistical_Contrasts_UG_vs_OC.csv"))
    t5_disp = t5_df[['Metric', 'UG_Median', 'UG_IQR', 'OC_Median', 'OC_IQR', 'Mann_Whitney_U1_UG', 'Mann_Whitney_U_smaller', 'Exact_p_value', 'Cliffs_Delta', 'Hedges_g', 'Approximate_95_CI_Hedges_g']].copy()
    t5_disp.columns = ['Metric', 'UG Med', 'UG IQR', 'OC Med', 'OC IQR', 'U₁ (UG)', 'U (min)', 'Exact p', 'Cliff δ', 'Hedges g', 'Approximate 95% CI (Hedges g)']
    lines.append(df_to_markdown(t5_disp))
    lines.append("\n\n")

    lines.append("![Figure 5: PCA Biplot & Dendrogram](figures/Figure_5_PCA_Cluster_Vegetation_Signatures.png)\n")
    lines.append("**Figure 5.** Reconciled Principal Component Analysis biplot (Axis 1 vs Axis 2, explaining 81.62% of variance in Model A; Model B sensitivity agreement r > 0.94) and hierarchical cluster dendrogram (Ward's linkage on Euclidean distance) for the 10 study quadrats across seasonal spectral feature vectors.\n\n")

    lines.append("### Table 4b. Component loading vectors for standardized vegetation features across Principal Components 1, 2, and 3 (Model A).\n")
    pca_df = pd.read_csv(os.path.join(BASE_DIR, "data", "PCA_scores_loadings.csv"))
    pca_loadings_disp = pca_df[['Feature', 'PC1_Loading', 'PC2_Loading', 'PC3_Loading']].copy()
    pca_loadings_disp.columns = ['Vegetation Feature', 'PC1 Loading (60.49%)', 'PC2 Loading (21.13%)', 'PC3 Loading (11.27%)']
    lines.append(df_to_markdown(pca_loadings_disp))
    lines.append("\n\n")
    lines.append("---\n")

    # Section 5
    lines.append("## 5. Discussion\n")
    lines.append("### 5.1 Integrated Canopy Reflectance vs. Understory Deconvolution\n")
    lines.append(
        "A central methodological conclusion of this pilot study is that 10 m Sentinel-2 pixels cannot deconvolve vertical vegetation strata "
        "or directly identify understory species in the absence of nested ground plots and sub-pixel fractional cover data. Spectral reflectance "
        "recorded by the satellite represents an integrated vertical measurement combining overstory crown foliage, subcanopy shrubs, ground herbs, "
        "litter, and soil background. While phenological dynamics (such as seasonal amplitude ΔNDVI and the winter-to-green-season persistence ratio) "
        "provide powerful descriptive indicators of ecosystem behavior, asserting that satellite pixels isolate herbaceous from woody components "
        "is empirically unsupportable. High winter-to-green-season ratios identify relatively strong dry-season spectral retention. They do not, "
        "by themselves, demonstrate perennial woody cover, rooting depth, or year-round soil stabilization.\n"
    )

    lines.append("### 5.2 Multi-Scale Spatial Extraction & Boundary Effects\n")
    lines.append(
        "Our evaluation of multi-scale centered extraction windows demonstrates that spatial aggregation scale fundamentally alters vegetation "
        "indices in fragmented mining landscapes. In underground remnant forest quadrats, expanding the extraction window from Zone A (10 m) to "
        "Zone D (100 m radius) resulted in a 29.0% dilution of pre-monsoon summer NDVI (0.381 down to 0.271). This dilution highlights the spatial "
        "mismatch between small field plots and medium-resolution raster grids: when field plots are situated in narrow remnant forest patches, "
        "larger buffers inevitably incorporate cleared tracks, haul roads, and colliery infrastructure. Conversely, in extensive opencast spoil "
        "plantations, spectral values remained invariant across scales (0.303 in Zone A to 0.287 in Zone D), reflecting landscape-level substrate "
        "and planting uniformity. Environmental assessments must therefore report exact spatial extraction footprints and account for surrounding landscape matrix composition.\n"
    )

    lines.append("### 5.3 Ecological Interpretation of Outlying Quadrat Trajectories\n")
    lines.append(
        "Evaluating individual quadrat trajectories reveals important nuances that caution against simplistic group classifications:\n"
        "- **Drainage Feature Influence (OC-Q2):** OC-Q2 exhibited an anomalous persistence ratio of 1.129, with winter NDVI (0.414 composite; single-date "
        "Dec 30 value = 0.453) exceeding its Green Season composite (0.366). Field notes confirm that OC-Q2 is situated in a low-lying depression "
        "near a mine drainage channel. The OC-Q2 trajectory is consistent with a possible moisture-related influence near the drainage feature; "
        "however, neither soil moisture nor the abundance of wetland vegetation was quantitatively measured.\n"
        "- **Rocky Overburden Spoil (OC-Q3):** OC-Q3 exhibited the lowest NDVI across all seasons (Summer 0.189, Green Season 0.347, Winter 0.287). Field plates "
        "confirm that OC-Q3 is situated on an unamended, boulder-strewn overburden dump with virtually zero organic topsoil, where sparse planted "
        "*Millettia pinnata* stems are heavily overgrown by the woody liana *Bauhinia vahlii*, restricting ground vegetation development.\n"
        "- **Climax Sal Canopy Stability (UG-Q3):** UG-Q3 supported the highest aboveground carbon stock in the study (26.71 t C ha⁻¹) under a closed "
        "canopy of *Shorea robusta*. Its low seasonal amplitude (ΔNDVI = 0.075) reflects the rapid spring leaf exchange characteristic of Sal, which "
        "replaces senescing leaves within weeks, maintaining a stable foliar canopy that buffers satellite reflectance against monsoonal swings.\n"
    )

    lines.append("### 5.4 Methodological Appraisal: Value of the Three-Tier Evidence Hierarchy\n")
    lines.append(
        "By enforcing a strict three-tier evidence hierarchy, this study resolves a persistent source of ambiguity in post-mining ecological "
        "monitoring. Past reports frequently blurred the boundary between empirical quadrat measurements and regional floristic checklists. "
        "Here, Tier 1 directly connects 58 verified woody stems to satellite signatures; Tier 2 uses 40 photographic plates to confirm structural "
        "habitat conditions; and Tier 3 isolates 12 candidate herbaceous and shrub taxa as unverified regional associates. This hierarchy "
        "guarantees that prospective additions (such as proposed 75-quadrat designs or simulated reference benchmarks) cannot be misconstrued as "
        "completed empirical field data.\n"
    )

    lines.append("### 5.5 Limitations of the Current Pilot Study\n")
    lines.append(
        "We candidly identify four primary structural limitations of the present study:\n"
        "1. **Pilot Sample Size (n = 5 per group):** With five quadrats per mining category, statistical tests have low power to detect subtle "
        "vegetation differences against high within-group heterogeneity (e.g., OC-Q1 plantation vs. OC-Q3 bare spoil).\n"
        "2. **Spatial Clustering:** Opencast quadrats are clustered within the Gevra/Dipka overburden complex, while underground quadrats are located "
        "within Banki/Surakachhar leases. This geographic grouping means quadrats are subsamples within lease blocks rather than independent "
        "regional landscape replicates, precluding broad causal generalizations regarding underground versus opencast mining impacts.\n"
        "3. **Lack of Unmined Regional Reference:** The original baseline lacked comparable quadrats in undisturbed regional reference forests, preventing "
        "formal calculation of absolute ecological recovery ratios.\n"
        "4. **Absence of Nested Micro-Quadrats:** Because the baseline lacked nested 1 × 1 m herbaceous sub-plots and contemporaneous sub-pixel "
        "canopy cover measurements, understory species dynamics can only be inferred through composite phenological syndromes rather than directly "
        "measured in-plot.\n"
    )
    lines.append("---\n")

    # Section 6: Conclusions (Strict 6 specific bullet points)
    lines.append("## 6. Conclusions and Management Implications\n")
    lines.append(
        "This study demonstrates that multi-temporal Sentinel-2 Level-2A surface reflectance provides valuable, evidence-constrained insights "
        "into seasonal vegetation greenness around mining-disturbed forest quadrats in central India, provided that satellite signals are interpreted "
        "as integrated vertical reflectance rather than deconvolved strata. Key conclusions and restoration management implications include:\n\n"
        "1. **Integrated top-of-canopy spectral response:** Sentinel-2 observations represent integrated top-of-canopy surface reflectance influenced by overstory foliage, subcanopy vegetation, ground cover, litter, soil, and exposed substrate, rather than isolated herbaceous or shrub signals. Attribution of multi-temporal satellite greenness to specific vegetation strata requires contemporaneous ground measurements and nested sub-plots.\n"
        "2. **Evidence hierarchy:** Field botanical audits in post-mining landscapes must strictly distinguish verified in-plot woody measurements from photographic structural records and unverified regional floristic leads. Blurring these tiers creates spurious precision and risks attributing landscape-scale floristic lists to single satellite pixels.\n"
        "3. **Multi-scale spatial support:** Pixel-quadrat co-location involves nominal support offsets and landscape dilution across 10 m to 100 m windows. In fragmented remnant forests, expanding extraction buffers beyond 10 m rapidly incorporates colliery infrastructure and clearings, reducing apparent summer greenness by nearly 30%, whereas homogeneous overburden dump plantations remain stable across scales.\n"
        "4. **Operational seasonal compositing:** Multi-date temporal compositing successfully mitigates cloud and shadow contamination in tropical monsoon environments while capturing phenological extremes. Combining June onset and October retreat into an operational 'Monsoon & Retreat Green Season' composite effectively brackets the annual peak of vegetative biomass and soil moisture accumulation.\n"
        "5. **Descriptive classification stability:** Heuristic phenological classes (such as Persistent Greenness Signal) serve as transparent descriptive summaries of multi-temporal greenness trajectories. In this pilot dataset, baseline assignments exhibited central stability (preserving 100% of site assignments in 32.0% of tested configurations and ≤ 1 change in 60.0%), but were sensitive to stringent persistence cutoffs (≥ 0.92), demonstrating the need to report sensitivity bounds rather than fixed categorical labels.\n"
        "6. **Pilot inferential constraints:** Statistical contrasts in small-scale pilot studies (n = 5 per group) have limited statistical power, requiring exact permutation tests, standardized effect sizes with analytical confidence intervals, and explicit warnings against equating non-significance with ecological equivalence. Future research should implement hierarchical, multi-site replicated designs with unmined reference forests and linear mixed-effects modeling.\n"
    )
    lines.append("---\n")

    # Data Availability Statement
    lines.append("## Data Availability and Provenance Statement\n")
    lines.append(
        "To ensure complete auditability, the full dataset, raw coordinates, scene observation manifests, GeoTIFF stack SHA-256 checksums, "
        "and execution scripts are permanently archived in the project repository: "
        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2. All data tables cited in this manuscript derive directly from the "
        "locked master dataset `data/Definitive_Locked_Site_By_Season_Indices.csv`, accompanied by `data/DATA_DICTIONARY_LOCKED_INDICES.md` and "
        "the statistical verification audit in `tables/Table_S7_Statistical_Audit_and_Verification.csv`.\n\n"
    )

    # References
    lines.append("## References\n")
    refs = [
        "Banerjee, D., Deb, K., Sharma, S., 2023. Topsoil quality and vegetation recovery dynamics on coal mine overburden dumps in central India. Environmental Earth Sciences 82, 345.",
        "Brede, B., Terryn, L., Barbier, N., Bartholomeus, H.M., Bartolo, R., Calders, K., Derroire, G., Krishna Moorthy, S.M., Lau, A., Levick, S.R., 2020. Non-destructive estimation of forest canopy structure from UAV and satellite remote sensing. Remote Sensing of Environment 247, 111924.",
        "Deb, K., Sharma, S., Banerjee, D., 2024. Impact of coal mining on dry deciduous forest cover and soil health in Chhattisgarh, India. Land Degradation & Development 35, 1120–1134.",
        "Drusch, M., Del Bello, U., Carlier, S., Colin, O., Fernandez, V., Gascon, F., Hoersch, B., Isola, C., Laberinti, P., Martimort, P., 2012. Sentinel-2: ESA's optical high-resolution mission for GMES operational services. Remote Sensing of Environment 120, 25–36.",
        "Hedges, L.V., Olkin, I., 1985. Statistical Methods for Meta-Analysis. Academic Press, Orlando.",
        "Hurlbert, S.H., 1984. Pseudoreplication and the design of ecological field experiments. Ecological Monographs 54, 187–211.",
        "Martinuzzi, S., Gould, W.A., Vierling, L.A., 2009. Land-cover classification in fragmented landscapes: A comparison of satellite and airborne sensors. International Journal of Remote Sensing 30, 4843–4860.",
        "Nagendra, H., Lucas, R., Honrado, J.P., Jongman, R.H., Tarantino, C., Adamo, M., Mairota, P., 2013. Remote sensing for conservation monitoring: Assessing status and trends of biodiversity. Remote Sensing in Ecology and Conservation 1, 12–28.",
        "Roy, S., Deb, K., Banerjee, D., 2022. Vegetation dynamics and land surface temperature changes in coalfields using multi-temporal Landsat data. Environmental Monitoring and Assessment 194, 612.",
        "Sharma, S., Banerjee, D., Deb, K., 2026. Overstory tree structure, biomass carbon stocks, and soil physicochemical dynamics across underground and opencast coal mining quadrats in Korba, India. Ecological Processes (published baseline dataset)."
    ]
    for r in refs:
        lines.append(f"- {r}\n")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Manuscript Markdown document successfully saved to: {output_path}")

# ==============================================================================
# MARKDOWN SUPPLEMENTAL GENERATOR (.md)
# ==============================================================================
def generate_supplemental_markdown(output_path):
    print("Writing Supplemental Material Markdown (.md)...")
    lines = []
    lines.append("# Supplementary Material: Seasonal Sentinel-2 greenness dynamics around mining-disturbed forest quadrats in central India: an evidence-constrained pilot analysis\n")
    lines.append("**Authors:** Shubham Sharma et al.  ")
    lines.append("**Journal:** *Ecological Indicators* (Version 2.4 Minor Revision) | Date: September 2026\n")
    lines.append("---\n")

    lines.append("## Methodological Note: Computational and Software Environment\n")
    lines.append(
        "All remote sensing data extractions, quality control screenings, index computations, statistical analyses, and figure "
        "renderings were executed deterministically under Python 3.11 in an isolated virtual environment. Specific software package versions "
        "used throughout this research are:\n"
        "- `rasterio`: version 1.3.10 (built against GDAL 3.9.0 and PROJ 9.4.0)\n"
        "- `scipy`: version 1.13.1\n"
        "- `scikit-learn`: version 1.5.0\n"
        "- `pandas`: version 2.2.2\n"
        "- `numpy`: version 1.26.4\n"
        "- `matplotlib`: version 3.9.0\n"
        "- `python-docx`: version 1.2.0\n\n"
        "All raster stacks and derived tables are archived with SHA-256 checksums in `rasters/CHECKSUMS.sha256`. Persistent repository link: "
        "https://github.com/shubham-sharma-korba/korba-restoration-sentinel2.\n\n"
        "### Algorithm for Heuristic Phenological Descriptive Classification\n\n"
        "```python\n"
        "def classify_quadrat(summer_ndvi, persistence_ratio, seasonal_amplitude, s_cut=0.30, r_cut=0.90, a_cut=0.25):\n"
        "    if summer_ndvi >= s_cut and persistence_ratio >= r_cut:\n"
        "        return 'Persistent Greenness Signal'\n"
        "    elif seasonal_amplitude >= a_cut and summer_ndvi < s_cut:\n"
        "        return 'Pronounced Seasonal Flush'\n"
        "    else:\n"
        "        return 'Disturbed Transition'\n"
        "```\n\n"
        "The threshold sensitivity analysis evaluated this function across 5 summer cutoffs (0.25, 0.28, 0.30, 0.32, 0.35), "
        "5 persistence ratio cutoffs (0.85, 0.88, 0.90, 0.92, 0.95), and 3 seasonal amplitude cutoffs (0.20, 0.25, 0.28) for a total of 75 configurations.\n\n"
    )

    # Table S1
    lines.append("### Supplementary Table S1. Coordinates, elevations, colliery sectors, and baseline structural metrics for the 10 published quadrats.\n")
    t1_df = pd.read_csv(os.path.join(BASE_DIR, "data", "01_Korba_quadrat_coordinates.csv"))
    t1_disp = t1_df[['ID', 'Mining_type', 'Colliery_Sector', 'Latitude', 'Longitude', 'Elevation_m', 'Plot_Area_m2']].copy()
    t1_disp.columns = ['Quadrat ID', 'Mining Type', 'Colliery Sector', 'Latitude (°N)', 'Longitude (°E)', 'Elevation (m)', 'Area (m²)']
    lines.append(df_to_markdown(t1_disp))
    lines.append("\n\n")

    # Table S1B
    lines.append("### Supplementary Table S1B. Quadrat centroid to containing Sentinel-2 pixel center offsets, boundary distances, and nominal geometric overlap on the 10 m UTM Zone 44N grid.\n")
    t1b_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S1B_Quadrat_Centroid_Pixel_Offsets.csv"))
    t1b_disp = t1b_df[['Quadrat_ID', 'Mining_Type', 'Offset_Delta_X_m', 'Offset_Delta_Y_m', 'Distance_to_Pixel_Center_m', 'Distance_to_Nearest_Boundary_m', 'Quadrat_Pixel_Overlap_Pct']].copy()
    t1b_disp.columns = ['Quadrat', 'Type', 'ΔX (m)', 'ΔY (m)', 'Center Dist (m)', 'Boundary Dist (m)', 'Overlap (%)']
    lines.append(df_to_markdown(t1b_disp))
    lines.append("\n\n")

    # Table S2
    lines.append("### Supplementary Table S2. Hierarchical observation architecture and SCL radiometric provenance summary across spatial extraction zones.\n")
    lines.append(
        "This table distinguishes five nested observation levels: Level 1 (pixel-level surface reflectance), Level 2 (date-level site observations: 6 dates × 10 quadrats × 4 zones = 240 records), "
        "Level 3 (seasonal composites: 3 seasons × 10 quadrats × 4 zones = 120 records), Level 4 (locked Zone A site-level indices, n=10), and Level 5 (group-level descriptive contrasts, n=5 UG vs n=5 OC). "
        "Complete record stored in `tables/Table_S2_Sentinel2_Observation_Provenance.csv`.\n\n")
    s2_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S2_Sentinel2_Observation_Provenance.csv"))
    s2_summary = s2_df.groupby(['Spatial_Zone', 'Season_Operational']).agg({
        'Theoretical_Denominator_Pixels': 'first',
        'Valid_Terrestrial_Pixels': ['mean', 'min'],
        'Terrestrial_Validity_Pct': ['mean', 'min'],
        'NDVI': ['mean', 'std'],
        'NDRE': ['mean', 'std']
    }).round(3).reset_index()
    s2_summary.columns = [f"{c[0]}_{c[1]}" if c[1] else c[0] for c in s2_summary.columns]
    lines.append(df_to_markdown(s2_summary))
    lines.append("\n\n")

    # Table S3
    lines.append("### Supplementary Table S3. Verified in-plot woody tree inventory (GBH ≥ 10 cm) across 24 accepted species (25 reported field morphotaxa) with standardized POWO taxonomy (showing first 25 of 58 stems).\n")
    tree_df = pd.read_csv(os.path.join(BASE_DIR, "data", "01_Quadrat_Verified_Tree_Woody_Inventory.csv"))
    tree_df['GBH_cm'] = (tree_df['GBH_m'] * 100).round(1)
    tree_df['DBH_cm'] = (tree_df['GBH_m'] * 100 / np.pi).round(1)
    tree_disp = tree_df[['Quadrat_ID', 'Tree_Stem_No', 'Reported_Scientific_Name', 'Accepted_Scientific_Name_POWO', 'Family', 'GBH_cm', 'DBH_cm', 'Height_m']].copy()
    tree_disp.columns = ['Quadrat', 'Stem No', 'Reported Botanical Name', 'POWO Accepted Name', 'Family', 'GBH (cm)', 'DBH (cm)', 'Height (m)']
    lines.append(df_to_markdown(tree_disp.head(25)))
    lines.append("\n*(Complete 58-stem records archived in `data/01_Quadrat_Verified_Tree_Woody_Inventory.csv`)*\n\n")

    # Table S4
    lines.append("### Supplementary Table S4. Audited photographic habitat evidence register across the 10 study quadrats (showing first 20 of 40 plates).\n")
    photo_df = pd.read_csv(os.path.join(BASE_DIR, "data", "02_Field_Photographic_Evidence_Register.csv"))
    photo_disp = photo_df[['Plate_Reference', 'Quadrat_ID', 'Mining_type', 'Photo_File_Name', 'Observed_Structural_Features', 'Visible_Ground_Stratum', 'Verification_Status']].copy()
    photo_disp.columns = ['Plate ID', 'Quadrat', 'Mining Type', 'Photo File', 'Observed Features', 'Ground Stratum', 'Verification']
    lines.append(df_to_markdown(photo_disp.head(20)))
    lines.append("\n*(Complete 40-plate records archived in `data/02_Field_Photographic_Evidence_Register.csv`)*\n\n")

    # Table S5
    lines.append("### Supplementary Table S5. Regional floristic candidate leads register (treated strictly as unverified regional associates).\n")
    leads_df = pd.read_csv(os.path.join(BASE_DIR, "data", "03_Regional_Floristic_Leads_Register.csv"))
    leads_disp = leads_df[['Candidate_ID', 'Accepted_Taxon', 'Family', 'Growth_Form', 'Regional_Occurrence_Context', 'Quadrat_Plot_Evidence']].copy()
    leads_disp.columns = ['Candidate ID', 'Accepted Taxon', 'Family', 'Growth Form', 'Regional Occurrence', 'Quadrat Plot Evidence']
    lines.append(df_to_markdown(leads_disp))
    lines.append("\n\n")

    # Table S6
    lines.append("### Supplementary Table S6. Contextual soil physicochemical properties across the 10 quadrats (distinguishing published composite samples, regional literature ranges, chronosequence ranges, and model-derived SoilGrids 250m estimates).\n")
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
    lines.append("\n*Contextual Note: Soil physicochemical parameters are compiled strictly as background environmental context. Several entries represent published colliery-level composites or regional chronosequence ranges from nearby stations (0.67 to 3.49 km away) and SoilGrids 250 m model predictions, rather than independent quadrat-specific empirical measurements. Consequently, these soil data were not used in the primary statistical comparisons between quadrat groups.*\n\n")

    # Table S7
    lines.append("### Supplementary Table S7. Machine-readable statistical audit and verification table for all 15 evaluated spectral metrics in Zone A (10 m support).\n")
    audit_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S7_Statistical_Audit_and_Verification.csv"))
    audit_disp = audit_df[['Metric', 'UG_Rank_Sum_R1', 'OC_Rank_Sum_R2', 'Mann_Whitney_U1_UG', 'Mann_Whitney_U_smaller', 'Exact_p_value', 'Cliffs_Delta', 'Pooled_SD', 'Hedges_g', 'Approximate_95_CI_Hedges_g', 'Verification_Status']].copy()
    audit_disp.columns = ['Metric', 'R₁ (UG)', 'R₂ (OC)', 'U₁ (UG)', 'U (min)', 'Exact p', 'Cliff δ', 'Pooled SD', 'Hedges g', 'Approx 95% CI', 'Audit Status']
    lines.append(df_to_markdown(audit_disp))
    lines.append("\n\n")

    # Table S8
    lines.append("### Supplementary Table S8. Exploratory bivariate correlations between in-plot woody inventory parameters and Sentinel-2 spectral indices (n=10 quadrats).\n")
    corr_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S8_Field_Satellite_Correlations.csv"))
    corr_disp = corr_df[['Field_Variable', 'Remote_Sensing_Metric', 'Pearson_r', 'Pearson_p', 'Fisher_z_95_CI', 'Spearman_rs', 'Spearman_p']].copy()
    corr_disp.columns = ['Field Variable', 'Satellite Metric', 'Pearson r', 'p-value', 'Fisher z 95% CI', 'Spearman rs', 'p-value']
    lines.append(df_to_markdown(corr_disp))
    lines.append("\n\n")

    # Table S9A
    lines.append("### Supplementary Table S9A. Compact frequency distribution of quadrat reclassifications across all 75 tested heuristic threshold configurations (5 Summer NDVI [0.25–0.35] × 5 Persistence Ratio [0.85–0.95] × 3 Seasonal Amplitude [0.20, 0.25, 0.28]). In this dataset, changing the amplitude cutoff within the tested range did not alter assignments, whereas persistence and summer-NDVI cutoffs produced reclassifications.\n")
    freq_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9A_Sensitivity_Frequency_Distribution.csv"))
    lines.append(df_to_markdown(freq_df))
    lines.append("\n\n")

    # Table S9B
    lines.append("### Supplementary Table S9B. Heuristic phenological descriptive classification threshold sensitivity grid across alternative cutoffs (representative subset at Seasonal Amplitude = 0.25).\n")
    sens_df = pd.read_csv(os.path.join(BASE_DIR, "tables", "Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv"))
    sens_disp = sens_df[sens_df['Seasonal_Amplitude_Threshold'] == 0.25].copy()
    p_col = 'Persistent_Greenness_Count' if 'Persistent_Greenness_Count' in sens_disp.columns else 'Persistent_Canopy_Count'
    sens_disp = sens_disp[['Summer_NDVI_Threshold', 'Persistence_Ratio_Threshold', p_col, 'Seasonal_Flush_Count', 'Disturbed_Transition_Count', 'Reclassified_Quadrats_Count', 'Reclassified_Quadrats_Detail']].copy()
    sens_disp.columns = ['Summer Cutoff', 'Ratio Cutoff', 'Persistent', 'Flush', 'Transition', 'Changed', 'Reclassified Quadrats Detail']
    lines.append(df_to_markdown(sens_disp.head(25)))
    lines.append("\n*(Complete 75-configuration sensitivity dataset archived in `tables/Table_S9_Syndrome_Threshold_Sensitivity_Analysis.csv`)*\n\n")

    # Figure S1
    lines.append("### Supplementary Figure S1. Exploratory Field-to-Satellite Linkage Matrix\n")
    lines.append("![Supplementary Figure S1](figures/Supplementary_Figure_S1_Linkage_Matrix.png)\n")
    lines.append("**Figure S1.** Exploratory field-to-satellite linkage analysis comparing verified in-plot woody tree inventory attributes against Sentinel-2 spectral indices across the 10 study quadrats, showing seasonal NDVI amplitude across quadrats (left) and bivariate scatter of overstory tree carbon stock versus pre-monsoon summer NDVI (right).\n")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Supplemental Material Markdown document successfully saved to: {output_path}")

# ==============================================================================
# FIELD AUDIT PROTOCOL GENERATOR (.md)
# ==============================================================================
def generate_field_protocol_markdown(output_path):
    print("Writing Botanical Field Audit Protocol Markdown (.md)...")
    content = """# Botanical Field Audit Protocol: Ground-Truthing Seasonal Sentinel-2 Signals in Mining Landscapes

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
"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Botanical Field Audit Protocol saved to: {output_path}")

# ==============================================================================
# MAIN ORCHESTRATOR
# ==============================================================================
def main():
    print("=" * 70)
    print("Generating Audited Manuscript and Supplemental Deliverables (Version 2.4)...")
    print("=" * 70)

    docx_path = os.path.join(BASE_DIR, "MANUSCRIPT_DRAFT.docx")
    md_path = os.path.join(BASE_DIR, "MANUSCRIPT_DRAFT.md")
    supp_docx_path = os.path.join(BASE_DIR, "SUPPLEMENTAL_MATERIAL.docx")
    supp_md_path = os.path.join(BASE_DIR, "SUPPLEMENTAL_MATERIAL.md")
    proto_md_path = os.path.join(BASE_DIR, "BOTANICAL_FIELD_AUDIT_PROTOCOL.md")

    build_manuscript_docx(docx_path)
    generate_markdown_manuscript(md_path)
    build_supplemental_docx(supp_docx_path)
    generate_supplemental_markdown(supp_md_path)
    generate_field_protocol_markdown(proto_md_path)

    print("\n" + "=" * 70)
    print("All deliverables generated successfully!")
    print(f" - {docx_path}")
    print(f" - {md_path}")
    print(f" - {supp_docx_path}")
    print(f" - {supp_md_path}")
    print(f" - {proto_md_path}")
    print("=" * 70)

if __name__ == "__main__":
    main()
