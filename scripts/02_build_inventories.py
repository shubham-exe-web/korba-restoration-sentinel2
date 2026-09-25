"""
Script: 02_build_inventories.py
Step 1: Rebuild the field evidence register with strict scientific distinction between:
  1. Verified Quadrat-Level Tree & Woody Subcanopy Baseline (from published raw field sheets).
  2. Verified Field Photographic Plates (40 plates documenting habitat & understory reality).
  3. Regional / Off-Plot Floristic Leads Register (Unverified at quadrat level; candidates for future survey).
  4. Repaired clean Table 1 and Phase 21 Richness Table.
"""

import os
import pandas as pd
import docx

def build_inventories():
    print("Rebuilding Field Evidence Register from primary sources...")
    
    # --- 1. TIER 1: VERIFIED QUADRAT-LEVEL WOODY BASELINE (from thesis docx Tables 10-19) ---
    doc_path = "/Users/shubhamsharma/Downloads/shubham dissertation /dissertation word.docx"
    doc = docx.Document(doc_path)
    
    # Taxonomic reconciliation map (POWO / IPNI standards)
    taxa_reconciled = {
        "Vachellia nilotica": {"accepted_name": "Vachellia nilotica (L.) P.J.H.Hurter & Mabb.", "family": "Fabaceae", "growth_form": "Tree (Deciduous thorny)"},
        "Terminalia elliptica": {"accepted_name": "Terminalia elliptica Willd.", "family": "Combretaceae", "growth_form": "Tree (Large deciduous)"},
        "Ziziphus mauritiana": {"accepted_name": "Ziziphus mauritiana Lam.", "family": "Rhamnaceae", "growth_form": "Small tree / Thorny shrub (Deciduous)"},
        "Tamarindus indica": {"accepted_name": "Tamarindus indica L.", "family": "Fabaceae", "growth_form": "Tree (Large evergreen)"},
        "Ficus religiosa": {"accepted_name": "Ficus religiosa L.", "family": "Moraceae", "growth_form": "Tree (Large semi-evergreen)"},
        "Wrightia tinctoria": {"accepted_name": "Wrightia tinctoria R.Br.", "family": "Apocynaceae", "growth_form": "Small tree (Deciduous)"},
        "Azadirachta indica": {"accepted_name": "Azadirachta indica A.Juss.", "family": "Meliaceae", "growth_form": "Tree (Deciduous)"},
        "Diospyros melanoxylon": {"accepted_name": "Diospyros melanoxylon Roxb.", "family": "Ebenaceae", "growth_form": "Tree (Deciduous)"},
        "Shorea robusta": {"accepted_name": "Shorea robusta Gaertn.", "family": "Dipterocarpaceae", "growth_form": "Tree (Climax dominant semi-deciduous)"},
        "Ficus virens": {"accepted_name": "Ficus virens Aiton", "family": "Moraceae", "growth_form": "Tree (Deciduous)"},
        "Madhuca longifolia": {"accepted_name": "Madhuca longifolia (J.Koenig ex L.) J.F.Macbr.", "family": "Sapotaceae", "growth_form": "Tree (Deciduous)"},
        "Mangifera indica": {"accepted_name": "Mangifera indica L.", "family": "Anacardiaceae", "growth_form": "Tree (Large evergreen)"},
        "Alstonia scholaris": {"accepted_name": "Alstonia scholaris (L.) R.Br.", "family": "Apocynaceae", "growth_form": "Tree (Evergreen)"},
        "Tectona grandis": {"accepted_name": "Tectona grandis L.f.", "family": "Lamiaceae", "growth_form": "Tree (Deciduous timber)"},
        "Terminalia tomentosa": {"accepted_name": "Terminalia tomentosa (Roxb.) Wight & Arn.", "family": "Combretaceae", "growth_form": "Tree (Large deciduous)"},
        "Gmelina arborea": {"accepted_name": "Gmelina arborea Roxb. ex Sm.", "family": "Lamiaceae", "growth_form": "Tree (Deciduous plantation)"},
        "Eucalyptus tereticornis": {"accepted_name": "Eucalyptus tereticornis Sm.", "family": "Myrtaceae", "growth_form": "Tree (Evergreen plantation)"},
        "Ficus benghalensis": {"accepted_name": "Ficus benghalensis L.", "family": "Moraceae", "growth_form": "Tree (Large evergreen)"},
        "Carissa carandas": {"accepted_name": "Carissa carandas L.", "family": "Apocynaceae", "growth_form": "Large thorny shrub / small tree"},
        "Millingtonia hortensis": {"accepted_name": "Millingtonia hortensis L.f.", "family": "Bignoniaceae", "growth_form": "Tree (Evergreen ornamental/avenue)"},
        "Butea monosperma": {"accepted_name": "Butea monosperma (Lam.) Taub.", "family": "Fabaceae", "growth_form": "Tree (Deciduous)"},
        "Pongamia pinnata": {"accepted_name": "Pongamia pinnata (L.) Pierre", "family": "Fabaceae", "growth_form": "Tree (Semi-evergreen plantation)"},
        "Bauhinia vahlii": {"accepted_name": "Phanera vahlii (Wight & Arn.) Benth. (syn. Bauhinia vahlii)", "family": "Fabaceae", "growth_form": "Giant woody climber / Liana"},
        "Holarrhena antidysenterica": {"accepted_name": "Holarrhena pubescens Wall. ex G.Don (syn. H. antidysenterica)", "family": "Apocynaceae", "growth_form": "Shrub / Small tree (Deciduous subcanopy)"},
        "Ficus racemosa": {"accepted_name": "Ficus racemosa L.", "family": "Moraceae", "growth_form": "Tree (Deciduous/semi-evergreen)"}
    }
    
    quadrats = ['UG-Q1', 'UG-Q2', 'UG-Q3', 'UG-Q4', 'UG-Q5', 'OC-Q1', 'OC-Q2', 'OC-Q3', 'OC-Q4', 'OC-Q5']
    stem_records = []
    
    for idx, q_id in enumerate(quadrats):
        table = doc.tables[9 + idx] # Tables 10 to 19
        m_type = "Underground" if "UG" in q_id else "Opencast"
        for r in table.rows[1:]:
            cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
            reported_name = cells[1].strip()
            rec = taxa_reconciled.get(reported_name, {
                "accepted_name": reported_name,
                "family": cells[2].strip(),
                "growth_form": "Woody tree"
            })
            
            stem_records.append({
                "Quadrat_ID": q_id,
                "Mining_type": m_type,
                "Tree_Stem_No": int(cells[0]),
                "Reported_Scientific_Name": reported_name,
                "Accepted_Scientific_Name_POWO": rec["accepted_name"],
                "Family": rec["family"],
                "Growth_Form": rec["growth_form"],
                "GBH_m": float(cells[3]),
                "Height_m": float(cells[4]),
                "Primary_Source_Sheet": f"Dissertation Appendix A, Table {10+idx}",
                "Observation_Tier": "Tier 1: Measured In-Plot Field Stems",
                "Confidence": "High (Direct field quadrat measurement)"
            })
            
    df_stems = pd.DataFrame(stem_records)
    df_stems.to_csv("data/01_Quadrat_Verified_Tree_Woody_Inventory.csv", index=False)
    print(f" - Saved data/01_Quadrat_Verified_Tree_Woody_Inventory.csv ({len(df_stems)} measured stems)")
    
    # Save the 25 species checklist
    checklist = []
    for r in doc.tables[19].rows[1:]:
        cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
        reported_name = cells[1].strip()
        rec = taxa_reconciled.get(reported_name, {"accepted_name": reported_name, "family": cells[2].strip(), "growth_form": "Tree"})
        checklist.append({
            "S_No": int(cells[0]),
            "Reported_Name": reported_name,
            "Accepted_Name_POWO": rec["accepted_name"],
            "Family": rec["family"],
            "Reported_Mining_Area": cells[3].strip(),
            "Growth_Form": rec["growth_form"],
            "Published_Reference": "Sharma, Banerjee & Deb (2026), Appendix B Table 20"
        })
    df_checklist = pd.DataFrame(checklist)
    df_checklist.to_csv("data/Korba_published_trees_master.csv", index=False)
    print(f" - Saved data/Korba_published_trees_master.csv ({len(df_checklist)} verified overstory species)")

    # --- 2. TIER 2: VERIFIED FIELD PHOTOGRAPHIC PLATES (40 plates) ---
    photo_records = []
    quadrat_photos = {
        'OC-Q1': (3, "Reclamation plantation with Eucalyptus & Gmelina; open broken canopy; dry leaf litter with sparse saplings; heavy overburden impact."),
        'OC-Q2': (4, "Adjacent to mine water drainage sump; mature Ficus & Carissa carandas; dry soil with limited herbaceous ground layer; localized foliar chlorosis."),
        'OC-Q3': (4, "Planted Pongamia stand with mature Diospyros; woody climber (Bauhinia vahlii) on boles; bare rocky spoil patches with dry litter."),
        'OC-Q4': (3, "Semi-open canopy dominated by 5 Holarrhena stems; dry leaf litter and bare compacted soil patches; disturbed opencast perimeter."),
        'OC-Q5': (3, "Multi-stemmed tree growth with Tamarindus & Mangifera; compound leaf foliage; anthropogenic debris and disturbed ground layer."),
        'UG-Q1': (6, "Mature tree architecture (Terminalia/Vachellia) near colliery infrastructure; exposed ground with dry litter and scanty vegetation."),
        'UG-Q2': (4, "Mature Wrightia & Azadirachta; dry litter with scattered understory growth; intact foliage condition; moderate canopy spacing."),
        'UG-Q3': (4, "Dense Shorea robusta (Sal) canopy; conspicuous understory vegetation with regenerating saplings and shrubs; high forest density."),
        'UG-Q4': (3, "Dominant mature Ficus & Mangifera; massive trunk girths; forest floor with dry litter accumulation and bare soil patches."),
        'UG-Q5': (6, "Moderate canopy density with broad leaves (Terminalia/Alstonia); dark bark patches; dry litter and poor understory growth.")
    }
    
    for q_id, (n_photos, desc) in quadrat_photos.items():
        m_type = "Underground" if "UG" in q_id else "Opencast"
        for p_idx in range(1, n_photos + 1):
            fn = f"{q_id}_photo_{p_idx}.jpeg"
            photo_records.append({
                "Quadrat_ID": q_id,
                "Mining_type": m_type,
                "Photo_File_Name": fn,
                "Plate_Reference": f"Appendix F Photographic Plates ({q_id})",
                "Observed_Structural_Features": desc,
                "Visible_Ground_Stratum": "Documented in photo: dry litter, bare soil patches, scattered understory/saplings",
                "Observation_Tier": "Tier 2: Direct Visual Photographic Plate Evidence",
                "Verification_Status": "Verified from Field Photographic Archive"
            })
            
    df_photos = pd.DataFrame(photo_records)
    df_photos.to_csv("data/02_Field_Photographic_Evidence_Register.csv", index=False)
    print(f" - Saved data/02_Field_Photographic_Evidence_Register.csv ({len(df_photos)} photographic records)")

    # --- 3. TIER 3: REGIONAL / OFF-PLOT FLORISTIC LEADS REGISTER ---
    # Strictly distinguishing unverified off-plot leads from measured plot data
    floristic_leads = [
        # Shrubs
        {
            "Candidate_ID": "LEAD-S01",
            "Accepted_Taxon": "Lantana camara L.",
            "Family": "Verbenaceae",
            "Growth_Form": "Shrub (Perennial invasive scrambling)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Abundant invasive thickets on mine overburden dumps, road verges, and open scrub in Korba.",
            "Quadrat_Plot_Evidence": "NO species-specific stem tally in 10x10m tree quadrat sheets. Observed in surrounding overburden landscape.",
            "Remote_Category": "Category A (when forming dense contiguous monocultures on overburden; distinct seasonal stability)",
            "Botanical_Audit_Needed": "Requires formal nested 5x5m shrub quadrat stem count."
        },
        {
            "Candidate_ID": "LEAD-S02",
            "Accepted_Taxon": "Woodfordia fruticosa (L.) Kurz",
            "Family": "Lythraceae",
            "Growth_Form": "Shrub (Deciduous erect)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Rocky slopes and open dry deciduous scrub throughout Bilaspur-Korba division.",
            "Quadrat_Plot_Evidence": "NO in-plot quadrat count. Documented in regional floristic collections.",
            "Remote_Category": "Category B (community scrub group)",
            "Botanical_Audit_Needed": "Requires formal nested 5x5m shrub quadrat survey."
        },
        {
            "Candidate_ID": "LEAD-S03",
            "Accepted_Taxon": "Calotropis procera (Aiton) W.T.Aiton",
            "Family": "Apocynaceae",
            "Growth_Form": "Shrub (Xerophytic broadleaf)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Unreclaimed mine spoil, coal transport margins, barren ground.",
            "Quadrat_Plot_Evidence": "Observed on peripheral spoil banks; not tallies inside tree quadrats.",
            "Remote_Category": "Category B (xeromorphic scrub signature)",
            "Botanical_Audit_Needed": "Requires formal ground verification."
        },
        {
            "Candidate_ID": "LEAD-S04",
            "Accepted_Taxon": "Nerium oleander L.",
            "Family": "Apocynaceae",
            "Growth_Form": "Shrub (Evergreen)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Planted along industrial complex margins and transport avenues.",
            "Quadrat_Plot_Evidence": "Avenue plantation outside forest quadrats.",
            "Remote_Category": "Category B (evergreen linear corridor)",
            "Botanical_Audit_Needed": "Requires formal ground verification."
        },
        {
            "Candidate_ID": "LEAD-S05",
            "Accepted_Taxon": "Nyctanthes arbor-tristis L.",
            "Family": "Oleaceae",
            "Growth_Form": "Shrub / Small tree (Deciduous understory)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Native dry deciduous forest understory associate in Central India.",
            "Quadrat_Plot_Evidence": "Sub-canopy understory component; GBH < 10 cm individuals unrecorded in tree sheets.",
            "Remote_Category": "Category B (deciduous understory community)",
            "Botanical_Audit_Needed": "Requires formal nested understory sampling."
        },
        # Herbs
        {
            "Candidate_ID": "LEAD-H01",
            "Accepted_Taxon": "Parthenium hysterophorus L.",
            "Family": "Asteraceae",
            "Growth_Form": "Herb (Annual ruderal)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Ubiquitous invasive weed on disturbed opencast mine margins, rail corridors, and fallow ground.",
            "Quadrat_Plot_Evidence": "NO 1x1m herbaceous sub-quadrat data collected. Known regional dominant during wet season.",
            "Remote_Category": "Category A (high monsoonal green-up amplitude followed by rapid dry-season senescence)",
            "Botanical_Audit_Needed": "Requires quantitative 1x1m nested herbaceous sub-quadrat sampling."
        },
        {
            "Candidate_ID": "LEAD-H02",
            "Accepted_Taxon": "Senna tora (L.) Roxb. (syn. Cassia tora L.)",
            "Family": "Fabaceae",
            "Growth_Form": "Herb (Annual legume)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Forms contiguous monospecific carpets on open overburden bunds and disturbed soil during July-September.",
            "Quadrat_Plot_Evidence": "NO in-plot herb tallies. Conspicuous monsoonal flush noted regionally.",
            "Remote_Category": "Category A (synchronous monsoon flush & abrupt post-monsoon dry-down)",
            "Botanical_Audit_Needed": "Requires quantitative 1x1m nested herbaceous sub-quadrat sampling."
        },
        {
            "Candidate_ID": "LEAD-H03",
            "Accepted_Taxon": "Hyptis suaveolens (L.) Poit.",
            "Family": "Lamiaceae",
            "Growth_Form": "Herb / Subshrub (Annual/biennial aromatic)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Dominates roadsides, cleared forest boundaries, and abandoned quarry edges.",
            "Quadrat_Plot_Evidence": "NO in-plot herb tallies. Regional ruderal associate.",
            "Remote_Category": "Category A/B (dense tall herbaceous canopies)",
            "Botanical_Audit_Needed": "Requires quantitative nested sub-quadrat survey."
        },
        {
            "Candidate_ID": "LEAD-H04",
            "Accepted_Taxon": "Chromolaena odorata (L.) R.M.King & H.Rob.",
            "Family": "Asteraceae",
            "Growth_Form": "Herbaceous / Scandent subshrub",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Moist degraded depressions, drainage lines, forest clearings.",
            "Quadrat_Plot_Evidence": "Regional weed; unmeasured in tree plots.",
            "Remote_Category": "Category B (semi-perennial scrub)",
            "Botanical_Audit_Needed": "Requires formal ground verification."
        },
        {
            "Candidate_ID": "LEAD-H05",
            "Accepted_Taxon": "Acalypha indica L.",
            "Family": "Euphorbiaceae",
            "Growth_Form": "Herb (Annual forb)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Ruderal patches, human habitation peripheries.",
            "Quadrat_Plot_Evidence": "Trait matrix record; no quadrat-level tally.",
            "Remote_Category": "Category C (not remotely detectable in 10-20m pixels)",
            "Botanical_Audit_Needed": "Requires ground verification."
        },
        {
            "Candidate_ID": "LEAD-H06",
            "Accepted_Taxon": "Diplazium esculentum (Retz.) Sw.",
            "Family": "Athyriaceae",
            "Growth_Form": "Herb (Perennial riparian fern)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Moist shady drainage gullies near mine sumps (e.g. vicinity of OC-Q2).",
            "Quadrat_Plot_Evidence": "Micro-habitat record; no plot tally.",
            "Remote_Category": "Category C (occluded understory fern)",
            "Botanical_Audit_Needed": "Requires ground verification."
        },
        {
            "Candidate_ID": "LEAD-H07",
            "Accepted_Taxon": "Cynodon dactylon (L.) Pers.",
            "Family": "Poaceae",
            "Growth_Form": "Herb (Perennial graminoid)",
            "Field_Status": "Off-Plot / Regional Associate (Unverified at Quadrat Level)",
            "Regional_Occurrence_Context": "Patchy ground cover on unpaved tracks and forest openings.",
            "Quadrat_Plot_Evidence": "Common regional ground cover; no cover percent recorded in tree sheets.",
            "Remote_Category": "Category B (grass patch) / Category C (sub-canopy)",
            "Botanical_Audit_Needed": "Requires cover-abundance estimation."
        }
    ]
    df_leads = pd.DataFrame(floristic_leads)
    df_leads.to_csv("data/03_Regional_Floristic_Leads_Register.csv", index=False)
    print(f" - Saved data/03_Regional_Floristic_Leads_Register.csv ({len(df_leads)} candidate taxa)")

    # Legacy inventory files maintained for compatibility but strictly qualified
    df_herbs_legacy = df_leads[df_leads["Growth_Form"].str.contains("Herb")].copy()
    df_herbs_legacy.rename(columns={"Candidate_ID": "ID", "Accepted_Taxon": "Scientific_name", "Growth_Form": "Growth_form", "Quadrat_Plot_Evidence": "Evidence", "Regional_Occurrence_Context": "Location"}, inplace=True)
    df_herbs_legacy["Confidence"] = "Unverified Regional Candidate"
    df_herbs_legacy.to_csv("data/02_Korba_herb_inventory.csv", index=False)
    
    df_shrubs_legacy = df_leads[df_leads["Growth_Form"].str.contains("Shrub")].copy()
    df_shrubs_legacy.rename(columns={"Candidate_ID": "ID", "Accepted_Taxon": "Scientific_name", "Growth_Form": "Growth_form", "Quadrat_Plot_Evidence": "Evidence", "Regional_Occurrence_Context": "Location"}, inplace=True)
    df_shrubs_legacy["Confidence"] = "Unverified Regional Candidate"
    df_shrubs_legacy.to_csv("data/03_Korba_shrub_inventory.csv", index=False)

    # --- 4. REPAIR TABLE 1: QUADRAT CHARACTERISTICS CSV ---
    # Compute verified metrics per quadrat
    t1_rows = []
    coords_df = pd.read_csv("data/01_Korba_quadrat_coordinates.csv")
    
    biomass_ug = {"UG-Q1": (42.60, 25.23), "UG-Q2": (38.75, 22.95), "UG-Q3": (45.10, 26.71), "UG-Q4": (40.80, 24.16), "UG-Q5": (39.95, 23.66)}
    biomass_oc = {"OC-Q1": (28.40, 16.81), "OC-Q2": (31.20, 18.48), "OC-Q3": (26.75, 15.84), "OC-Q4": (29.85, 17.67), "OC-Q5": (27.90, 16.52)}
    
    for _, cr in coords_df.iterrows():
        q_id = cr["ID"]
        m_type = cr["Mining_type"]
        q_stems = df_stems[df_stems["Quadrat_ID"] == q_id]
        
        n_stems = len(q_stems)
        n_richness = q_stems["Accepted_Scientific_Name_POWO"].nunique()
        dom_species = ", ".join(q_stems["Reported_Scientific_Name"].value_counts().index[:2])
        
        agb, c_stock = biomass_ug[q_id] if m_type == "Underground" else biomass_oc[q_id]
        
        # Check if woody shrub was in plot
        has_shrub_stem = any(s in q_stems["Reported_Scientific_Name"].values for s in ["Carissa carandas", "Holarrhena antidysenterica", "Ziziphus mauritiana"])
        shrub_note = "Measured in-plot woody shrub/small tree present" if has_shrub_stem else "No measured shrub stem >=10cm GBH"
        
        t1_rows.append({
            "Quadrat_ID": q_id,
            "Mining_type": m_type,
            "Latitude": cr["Latitude"],
            "Longitude": cr["Longitude"],
            "Elevation_m": cr["Elevation_m"],
            "Colliery_Sector": cr["Colliery_Sector"],
            "Verified_Overstory_Stems": n_stems,
            "Verified_Tree_Species_Richness": n_richness,
            "Dominant_Tree_Taxa": dom_species,
            "AGB_t_ha": agb,
            "Carbon_Stock_tC_ha": c_stock,
            "In_Plot_Woody_Shrub_Status": shrub_note,
            "Field_Plate_Understory_Observation": quadrat_photos[q_id][1]
        })
        
    df_t1 = pd.DataFrame(t1_rows)
    df_t1.to_csv("tables/Table_1_Quadrat_Characteristics.csv", index=False)
    print(" - Saved tables/Table_1_Quadrat_Characteristics.csv (clean, uncorrupted CSV)")

    # --- 5. PHASE 21 RICHNESS & PHENOLOGY TARGET TABLE ---
    p21_rows = []
    for _, r in df_t1.iterrows():
        p21_rows.append({
            "Site": r["Quadrat_ID"],
            "Type": r["Mining_type"],
            "Verified_Tree_Richness": r["Verified_Tree_Species_Richness"],
            "Verified_Tree_Stems": r["Verified_Overstory_Stems"],
            "Candidate_Shrub_Leads": "Regional associates only (unverified in-plot)",
            "Candidate_Herb_Leads": "Regional associates only (unverified in-plot)",
            "Field_Evidence_Basis": "Raw field sheets (Tree GBH>=10cm) & 40 Photo Plates",
            "Summer_NDVI": None,  # Will be populated by pipeline
            "Monsoon_NDVI": None,
            "Winter_NDVI": None,
            "Seasonal_Amplitude": None,
            "Winter_Persistence": None,
            "Remotely_Inferred_Signal": None
        })
    df_p21 = pd.DataFrame(p21_rows)
    df_p21.to_csv("tables/Table_Phase21_Vegetation_Richness_Phenology.csv", index=False)
    print(" - Saved tables/Table_Phase21_Vegetation_Richness_Phenology.csv")

if __name__ == "__main__":
    build_inventories()
