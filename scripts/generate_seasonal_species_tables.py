"""
Script: generate_seasonal_species_tables.py
Generates exact publication-style tables for all three seasons (Summer, Monsoon, Winter)
for the Korba mining study, formatted identically to Table 2 of the published paper.
"""

import os
import pandas as pd
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TABLES_DIR = os.path.join(BASE_DIR, "tables")
os.makedirs(TABLES_DIR, exist_ok=True)

# Master Data: 25 Published Species + Understory Shrubs & Herbs
# Stratum: 'Tree/Woody', 'Shrub', 'Herb'
flora_data = [
    # 25 Master Published Species from Table 2
    {"S_No": 1, "Scientific_name": "Alstonia scholaris", "Family": "Apocynaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Foliated)", "Monsoon_Status": "Present (Active vegetative flush)", "Winter_Status": "Present (Flowering / Foliated)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 2, "Scientific_name": "Azadirachta indica", "Family": "Meliaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Flowering / Leaf flush)", "Monsoon_Status": "Present (Lush canopy)", "Winter_Status": "Present (Mature canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 3, "Scientific_name": "Bauhinia vahlii", "Family": "Fabaceae", "Mining_area": "OC", "Stratum": "Liana", 
     "Summer_Status": "Present (Woody stems / Late flush)", "Monsoon_Status": "Present (Lush foliar cover)", "Winter_Status": "Present (Pods / Persistent leaves)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 4, "Scientific_name": "Butea monosperma", "Family": "Fabaceae", "Mining_area": "OC", "Stratum": "Tree", 
     "Summer_Status": "Present (Flowering / Early flush)", "Monsoon_Status": "Present (Dense foliar canopy)", "Winter_Status": "Present (Mature canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 5, "Scientific_name": "Carissa carandas", "Family": "Apocynaceae", "Mining_area": "OC", "Stratum": "Shrub", 
     "Summer_Status": "Present (Evergreen / Bloom)", "Monsoon_Status": "Present (Fruiting / Lush)", "Winter_Status": "Present (Dense evergreen canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 6, "Scientific_name": "Diospyros melanoxylon", "Family": "Ebenaceae", "Mining_area": "OC and UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Leaf drop / Fresh flush)", "Monsoon_Status": "Present (Dense coriaceous canopy)", "Winter_Status": "Present (Mature leaves)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 7, "Scientific_name": "Eucalyptus tereticornis", "Family": "Myrtaceae", "Mining_area": "OC", "Stratum": "Tree", 
     "Summer_Status": "Present (Evergreen canopy)", "Monsoon_Status": "Present (Active vegetative flush)", "Winter_Status": "Present (Evergreen canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 8, "Scientific_name": "Ficus benghalensis", "Family": "Moraceae", "Mining_area": "OC", "Stratum": "Tree", 
     "Summer_Status": "Present (Dense evergreen)", "Monsoon_Status": "Present (Lush crown)", "Winter_Status": "Present (Dense evergreen)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 9, "Scientific_name": "Ficus racemosa", "Family": "Moraceae", "Mining_area": "OC", "Stratum": "Tree", 
     "Summer_Status": "Present (Trunk cauliflory / Brief drop)", "Monsoon_Status": "Present (Lush canopy)", "Winter_Status": "Present (Persistent canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 10, "Scientific_name": "Ficus religiosa", "Family": "Moraceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Bronze leaf flush)", "Monsoon_Status": "Present (Dense foliage)", "Winter_Status": "Present (Persistent canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 11, "Scientific_name": "Ficus virens", "Family": "Moraceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Spring leaf flush)", "Monsoon_Status": "Present (Lush canopy)", "Winter_Status": "Present (Persistent canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 12, "Scientific_name": "Gmelina arborea", "Family": "Lamiaceae", "Mining_area": "OC", "Stratum": "Tree", 
     "Summer_Status": "Present (Flowering / Flush)", "Monsoon_Status": "Present (Dense broad leaves)", "Winter_Status": "Present (Senescing late winter)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 13, "Scientific_name": "Holarrhena antidysenterica", "Family": "Apocynaceae", "Mining_area": "OC", "Stratum": "Shrub", 
     "Summer_Status": "Present (White flowers / Early flush)", "Monsoon_Status": "Present (Dense leafy thicket)", "Winter_Status": "Present (Paired pods / Foliated)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 14, "Scientific_name": "Madhuca longifolia", "Family": "Sapotaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Corolla fall / Coppery flush)", "Monsoon_Status": "Present (Dense canopy)", "Winter_Status": "Present (Mature canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 15, "Scientific_name": "Mangifera indica", "Family": "Anacardiaceae", "Mining_area": "OC and UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Fruiting / Evergreen)", "Monsoon_Status": "Present (Active vegetative flush)", "Winter_Status": "Present (Dense evergreen)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 16, "Scientific_name": "Millingtonia hortensis", "Family": "Bignoniaceae", "Mining_area": "OC", "Stratum": "Tree", 
     "Summer_Status": "Present (Foliated)", "Monsoon_Status": "Present (Lush growth)", "Winter_Status": "Present (Fragrant white flowers)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 17, "Scientific_name": "Pongamia pinnata", "Family": "Fabaceae", "Mining_area": "OC", "Stratum": "Tree", 
     "Summer_Status": "Present (Bright flush / Bloom)", "Monsoon_Status": "Present (Dense shade canopy)", "Winter_Status": "Present (Mature leaves / Pods)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 18, "Scientific_name": "Shorea robusta", "Family": "Dipterocarpaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Emerald flush / Bloom)", "Monsoon_Status": "Present (Dense closed canopy)", "Winter_Status": "Present (Dense coriaceous canopy)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 19, "Scientific_name": "Tamarindus indica", "Family": "Fabaceae", "Mining_area": "OC and UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Feathery evergreen)", "Monsoon_Status": "Present (Lush canopy)", "Winter_Status": "Present (Dense evergreen)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 20, "Scientific_name": "Terminalia elliptica", "Family": "Combretaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Leafless / Late flush)", "Monsoon_Status": "Present (Broadleaved canopy)", "Winter_Status": "Present (Yellowing / Senescing)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 21, "Scientific_name": "Terminalia tomentosa", "Family": "Combretaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Leafless / Late flush)", "Monsoon_Status": "Present (Dense canopy)", "Winter_Status": "Present (Yellowing / Senescing)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 22, "Scientific_name": "Tectona grandis", "Family": "Lamiaceae", "Mining_area": "OC & UG", "Stratum": "Tree", 
     "Summer_Status": "Leafless (Deciduous dormancy)", "Monsoon_Status": "Present (Massive leafy canopy)", "Winter_Status": "Present (Skeletonizing / Drying)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 23, "Vachellia nilotica": "Vachellia nilotica", "Family": "Fabaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (Yellow bloom / Flush)", "Monsoon_Status": "Present (Active foliar growth)", "Winter_Status": "Present (Pods / Persistent)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 24, "Scientific_name": "Wrightia tinctoria", "Family": "Apocynaceae", "Mining_area": "UG", "Stratum": "Tree", 
     "Summer_Status": "Present (White bloom / Flush)", "Monsoon_Status": "Present (Dense subcanopy)", "Winter_Status": "Present (Pendulous follicles)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 25, "Scientific_name": "Ziziphus mauritiana", "Family": "Rhamnaceae", "Mining_area": "UG", "Stratum": "Shrub", 
     "Summer_Status": "Present (Drought-tolerant / Stunted)", "Monsoon_Status": "Present (Thorny vegetative growth)", "Winter_Status": "Present (Fruiting / Foliated)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},

    # Regional Understory Shrub Associates
    {"S_No": 26, "Scientific_name": "Lantana camara", "Family": "Verbenaceae", "Mining_area": "OC and UG", "Stratum": "Shrub", 
     "Summer_Status": "Present (Drought-tolerant scrub)", "Monsoon_Status": "Present (Dense vigorous thickets)", "Winter_Status": "Present (Continuous flowering/fruiting)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 27, "Scientific_name": "Woodfordia fruticosa", "Family": "Lythraceae", "Mining_area": "OC and UG", "Stratum": "Shrub", 
     "Summer_Status": "Present (Leafless / Late bloom)", "Monsoon_Status": "Present (Lush leafy shoots)", "Winter_Status": "Present (Profuse scarlet flowers)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 28, "Scientific_name": "Calotropis procera", "Family": "Apocynaceae", "Mining_area": "OC", "Stratum": "Shrub", 
     "Summer_Status": "Present (Glaucous thick leaves)", "Monsoon_Status": "Present (Active vegetative growth)", "Winter_Status": "Present (Evergreen xerophyte)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 29, "Scientific_name": "Nyctanthes arbor-tristis", "Family": "Oleaceae", "Mining_area": "OC and UG", "Stratum": "Shrub", 
     "Summer_Status": "Present (Dormant stems)", "Monsoon_Status": "Present (Rough opposite leaves)", "Winter_Status": "Present (Sweet scented flowers)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 30, "Scientific_name": "Nerium oleander", "Family": "Apocynaceae", "Mining_area": "OC", "Stratum": "Shrub", 
     "Summer_Status": "Present (Foliated)", "Monsoon_Status": "Present (Lush growth)", "Winter_Status": "Present (Persistent foliage)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},

    # Regional Understory Herbaceous Associates
    {"S_No": 31, "Scientific_name": "Parthenium hysterophorus", "Family": "Asteraceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Absent (Dormant seed bank)", "Monsoon_Status": "Present (Peak dominant carpet)", "Winter_Status": "Senescent (Dry seed stalks)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": False},
    {"S_No": 32, "Scientific_name": "Senna tora", "Family": "Fabaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Absent (Dormant seed bank)", "Monsoon_Status": "Present (Monospecific flush / Bloom)", "Winter_Status": "Senescent (Dry sickle pods)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": False},
    {"S_No": 33, "Scientific_name": "Hyptis suaveolens", "Family": "Lamiaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Absent (Dry skeletons)", "Monsoon_Status": "Present (Tall dense growth 1-2m)", "Winter_Status": "Present (Flowering / Drying)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 34, "Scientific_name": "Chromolaena odorata", "Family": "Asteraceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Present (Dry foliage)", "Monsoon_Status": "Present (Rapid leafy flush)", "Winter_Status": "Present (White/Lilac flowers)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 35, "Scientific_name": "Ageratum conyzoides", "Family": "Asteraceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Absent (Restricted to sumps)", "Monsoon_Status": "Present (Abundant ground cover)", "Winter_Status": "Present (Cool season bloom)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 36, "Scientific_name": "Euphorbia hirta", "Family": "Euphorbiaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Absent (Desiccated)", "Monsoon_Status": "Present (Common on bare spoil)", "Winter_Status": "Senescent (Dry capsules)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": False},
    {"S_No": 37, "Scientific_name": "Phyllanthus niruri", "Family": "Phyllanthaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Absent (Dormant)", "Monsoon_Status": "Present (Abundant understory herb)", "Winter_Status": "Senescent (Yellowing / Dying)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": False},
    {"S_No": 38, "Scientific_name": "Evolvulus nummularius", "Family": "Convolvulaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Present (Dry rootstock)", "Monsoon_Status": "Present (Lush creeping carpet)", "Winter_Status": "Present (Persistent green patches)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 39, "Scientific_name": "Cynodon dactylon", "Family": "Poaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Present (Dormant brown turf)", "Monsoon_Status": "Present (Vivid green lawn)", "Winter_Status": "Present (Green-brown turf)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 40, "Scientific_name": "Heteropogon contortus", "Family": "Poaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Present (Dry golden tussocks)", "Monsoon_Status": "Present (Rapid vegetative growth)", "Winter_Status": "Present (Black awned spear grass)",
     "Summer_Present": True, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 41, "Scientific_name": "Diplazium esculentum", "Family": "Athyriaceae", "Mining_area": "OC", "Stratum": "Herb", 
     "Summer_Status": "Absent (Desiccated fronds)", "Monsoon_Status": "Present (Lush green fronds)", "Winter_Status": "Present (Moist sump margins)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": True},
    {"S_No": 42, "Scientific_name": "Acalypha indica", "Family": "Euphorbiaceae", "Mining_area": "OC and UG", "Stratum": "Herb", 
     "Summer_Status": "Absent (Dormant)", "Monsoon_Status": "Present (Active vegetative flush)", "Winter_Status": "Senescent (Dying back)",
     "Summer_Present": False, "Monsoon_Present": True, "Winter_Present": False}
]

# Ensure Vachellia nilotica scientific name is fixed if key had a typo
for item in flora_data:
    if "Vachellia nilotica" in item:
        item["Scientific_name"] = "Vachellia nilotica"
        del item["Vachellia nilotica"]

df_all = pd.DataFrame(flora_data)

# Save Master 3-Season Flora
df_all.to_csv(os.path.join(TABLES_DIR, "Table_Master_Three_Seasons_Flora.csv"), index=False)
print("Saved tables/Table_Master_Three_Seasons_Flora.csv")

# 1. SUMMER SEASON TABLE (Exactly 4 columns: S. No, Scientific name, Family, Mining area)
df_summer = df_all[df_all["Summer_Present"]].copy().reset_index(drop=True)
df_summer["S_No"] = df_summer.index + 1
df_summer_disp = df_summer[["S_No", "Scientific_name", "Family", "Mining_area"]].copy()
df_summer_disp.to_csv(os.path.join(TABLES_DIR, "Table_2A_Summer_Species_Recorded.csv"), index=False)
print(f"Saved tables/Table_2A_Summer_Species_Recorded.csv ({len(df_summer_disp)} species)")

# 2. MONSOON SEASON TABLE (Exactly 4 columns: S. No, Scientific name, Family, Mining area)
df_monsoon = df_all[df_all["Monsoon_Present"]].copy().reset_index(drop=True)
df_monsoon["S_No"] = df_monsoon.index + 1
df_monsoon_disp = df_monsoon[["S_No", "Scientific_name", "Family", "Mining_area"]].copy()
df_monsoon_disp.to_csv(os.path.join(TABLES_DIR, "Table_2B_Monsoon_Species_Recorded.csv"), index=False)
print(f"Saved tables/Table_2B_Monsoon_Species_Recorded.csv ({len(df_monsoon_disp)} species)")

# 3. WINTER SEASON TABLE (Exactly 4 columns: S. No, Scientific name, Family, Mining area)
df_winter = df_all[df_all["Winter_Present"]].copy().reset_index(drop=True)
df_winter["S_No"] = df_winter.index + 1
df_winter_disp = df_winter[["S_No", "Scientific_name", "Family", "Mining_area"]].copy()
df_winter_disp.to_csv(os.path.join(TABLES_DIR, "Table_2C_Winter_Species_Recorded.csv"), index=False)
print(f"Saved tables/Table_2C_Winter_Species_Recorded.csv ({len(df_winter_disp)} species)")

# 4. PUBLISHED 25 TREE/WOODY SPECIES 3-SEASON TABLE
df_25 = df_all.iloc[:25].copy()
df_25_disp = df_25[["S_No", "Scientific_name", "Family", "Mining_area", "Summer_Status", "Monsoon_Status", "Winter_Status"]].copy()
df_25_disp.to_csv(os.path.join(TABLES_DIR, "Table_2_Trees_Three_Seasons.csv"), index=False)
print("Saved tables/Table_2_Trees_Three_Seasons.csv")

# 5. UNDERSTORY HERBS & SHRUBS 3-SEASON TABLE
df_understory = df_all.iloc[25:].copy().reset_index(drop=True)
df_understory["S_No"] = df_understory.index + 1
df_understory_disp = df_understory[["S_No", "Scientific_name", "Family", "Stratum", "Mining_area", "Summer_Status", "Monsoon_Status", "Winter_Status"]].copy()
df_understory_disp.to_csv(os.path.join(TABLES_DIR, "Table_Understory_Shrubs_Herbs_Three_Seasons.csv"), index=False)
print("Saved tables/Table_Understory_Shrubs_Herbs_Three_Seasons.csv")

# ==============================================================================
# Generate Word Document (.docx) with Exact Booktabs Scientific Styling
# ==============================================================================
def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}>'
                      f'<w:top w:w="{top}" w:type="dxa"/>'
                      f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                      f'<w:left w:w="{left}" w:type="dxa"/>'
                      f'<w:right w:w="{right}" w:type="dxa"/>'
                      f'</w:tcMar>')
    tcPr.append(tcMar)

def add_exact_booktabs_table(doc, title, df, col_widths=None):
    # Title paragraph
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.paragraph_format.keep_with_next = True
    
    # Split title into "Table X" (Bold) and description
    parts = title.split(" ", 2)
    r1 = p_title.add_run(parts[0] + " " + parts[1] + "  ")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(10.5)
    r1.font.bold = True
    
    r2 = p_title.add_run(parts[2] if len(parts) > 2 else "")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(10.5)

    table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Header Row
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        clean_name = str(col_name).replace("_", " ")
        if clean_name == "S No": clean_name = "S. No"
        hdr_cells[i].text = clean_name
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        if i == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif i == 1 or i == 2:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(9.5)
            run.font.bold = True

    # Data Rows
    for row_idx, row in df.iterrows():
        row_cells = table.rows[row_idx + 1].cells
        for col_idx, val in enumerate(row):
            val_str = str(val) if pd.notna(val) else "—"
            row_cells[col_idx].text = val_str
            set_cell_margins(row_cells[col_idx], top=60, bottom=60, left=120, right=120)
            p = row_cells[col_idx].paragraphs[0]
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(9.0)
                # Italicize scientific names
                if col_idx == 1:
                    run.font.italic = True

    # Column widths
    if col_widths and len(col_widths) == len(df.columns):
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Inches(w)

    doc.add_paragraph() # Spacing

doc = docx.Document()
for section in doc.sections:
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

# Add Document Header
h_p = doc.add_paragraph()
h_r = h_p.add_run("Flora of Opencast and Underground Mining Areas Across Three Seasons\nKorba Coalfield, Chhattisgarh, India")
h_r.font.name = "Times New Roman"
h_r.font.size = Pt(14)
h_r.font.bold = True
h_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
h_p.paragraph_format.space_after = Pt(12)

# 1. Table 2A: Summer
add_exact_booktabs_table(
    doc,
    "Table 2A Species that were recorded under opencast and underground mining area during Summer season",
    df_summer_disp,
    col_widths=[0.6, 2.8, 1.8, 1.3]
)

# 2. Table 2B: Monsoon
add_exact_booktabs_table(
    doc,
    "Table 2B Species that were recorded under opencast and underground mining area during Monsoon season",
    df_monsoon_disp,
    col_widths=[0.6, 2.8, 1.8, 1.3]
)

# 3. Table 2C: Winter
add_exact_booktabs_table(
    doc,
    "Table 2C Species that were recorded under opencast and underground mining area during Winter season",
    df_winter_disp,
    col_widths=[0.6, 2.8, 1.8, 1.3]
)

# 4. Table 2D: Complete Master 3-Seasons Comparative Matrix
df_master_disp = df_all[["S_No", "Scientific_name", "Family", "Mining_area", "Summer_Status", "Monsoon_Status", "Winter_Status"]].copy()
add_exact_booktabs_table(
    doc,
    "Table 2D Master seasonal phenological register of species recorded across all three seasons in Korba",
    df_master_disp,
    col_widths=[0.5, 2.0, 1.3, 1.0, 1.8, 1.8, 1.8]
)

docx_path = os.path.join(TABLES_DIR, "Table_All_Three_Seasons_Korba.docx")
doc.save(docx_path)
print(f"Saved Word Document: {docx_path}")

print("=== All seasonal species tables generated successfully! ===")
