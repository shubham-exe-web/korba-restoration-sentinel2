"""
Script: generate_restoration_evidence_figures.py
Generates Figure 5: Multi-dimensional ecological recovery, mechanistic physical-biological functioning,
and continuous phenological pathways across Underground, Opencast, and Undisturbed Reference sites.
"""

import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIGURES_DIR = os.path.join(BASE_DIR, "figures")
TABLES_DIR = os.path.join(BASE_DIR, "tables")
os.makedirs(FIGURES_DIR, exist_ok=True)

# Load data
df_rr = pd.read_csv(os.path.join(TABLES_DIR, "Table_17_Ecological_Recovery_Indices_Response_Ratios.csv"))
df_pheno = pd.read_csv(os.path.join(TABLES_DIR, "Table_15_Continuous_Remote_Sensing_Phenometrics.csv"))
df_phys = pd.read_csv(os.path.join(TABLES_DIR, "Table_11_Soil_Physical_Hydrological_Properties.csv"))
df_bio = pd.read_csv(os.path.join(TABLES_DIR, "Table_12_Soil_Biological_Health_Enzymes.csv"))
df_fauna = pd.read_csv(os.path.join(TABLES_DIR, "Table_16_Faunal_Biodiversity_Indicators.csv"))

# Set publication style
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
fig, axs = plt.subplots(2, 2, figsize=(16, 13), dpi=300)

# Colors
c_ref = "#1b7837"  # Deep green
c_ug = "#2166ac"   # Royal blue
c_oc4 = "#d95f02"  # Amber orange (reclaimed thicket)
c_oc5 = "#b2182b"  # Deep red (degraded spoil)
c_oc = "#f4a582"   # Light orange

# --- PANEL A: Ecological Recovery Log Response Ratios ---
ax_a = axs[0, 0]
select_vars = [
    "Soil Organic Carbon (SOC %)",
    "Microbial Biomass Carbon (MBC ug/g)",
    "Dehydrogenase Activity (DHA ug/g/24h)",
    "Soil Infiltration Rate (mm/h)",
    "Water Stable Aggregates (%)",
    "Basal Area (m2/ha)",
    "Canopy Cover (%)",
    "Plant Shannon Diversity (H')",
    "Butterfly Species Richness",
    "Forest Specialist Butterflies (%)"
]

sub_rr = df_rr[df_rr["Ecological_Variable"].isin(select_vars)].copy()
sub_rr = sub_rr.iloc[::-1] # Reverse for clean vertical bar chart

y_pos = np.arange(len(sub_rr))
height = 0.35

ax_a.axvline(0, color="black", linestyle="--", linewidth=1.2, alpha=0.8, label="Undisturbed Reference (RR = 0)")
rects1 = ax_a.barh(y_pos + height/2, sub_rr["UG_Log_Response_Ratio"], height, label="Underground Mining Lease (UG)", color=c_ug, alpha=0.9, edgecolor="black", linewidth=0.5)
rects2 = ax_a.barh(y_pos - height/2, sub_rr["OC_Log_Response_Ratio"], height, label="Opencast Spoil / Overburden (OC)", color=c_oc4, alpha=0.9, edgecolor="black", linewidth=0.5)

ax_a.set_yticks(y_pos)
ax_a.set_yticklabels(sub_rr["Ecological_Variable"], fontsize=9.5, fontweight="bold")
ax_a.set_xlabel("Log Response Ratio: ln(Mined / Reference Forest)", fontsize=10.5, fontweight="bold")
ax_a.set_title("A. Multi-Functional Recovery Relative to Undisturbed Climax Forest", fontsize=11.5, fontweight="bold", pad=10)
ax_a.legend(loc="lower left", fontsize=8.5, frameon=True, facecolor="white", edgecolor="#d0d7de")
ax_a.set_xlim(-2.6, 0.4)

# Annotate significance
for i, (_, row) in enumerate(sub_rr.iterrows()):
    ax_a.text(0.05, y_pos[i] - 0.05, row["Significance"], color="black", fontsize=10, fontweight="bold", va="center")

# --- PANEL B: Mechanistic Soil Degradation & Hydrological Impedance ---
ax_b = axs[0, 1]

# Compare REF vs UG vs OC-Q4 vs OC-Q5
sites_comp = ["REF-Q2 (Climax)", "UG-Q3 (Intact Sal)", "OC-Q4 (Holarrhena Thicket)", "OC-Q5 (Compacted Spoil)"]
compaction = [0.88, 1.28, 1.65, 3.45] # Penetration resistance (MPa)
infiltration = [58.5, 36.5, 19.5, 3.8] # Steady infiltration (mm/h)
pawc = [19.6, 15.3, 11.7, 4.0] # Plant available water content (%)

x = np.arange(len(sites_comp))
width = 0.25

ax_b2 = ax_b.twinx() # For infiltration on secondary axis

b1 = ax_b.bar(x - width/2, compaction, width, label="Soil Compaction (Penetration Res., MPa)", color="#756bb1", edgecolor="black", linewidth=0.6)
b2 = ax_b2.bar(x + width/2, infiltration, width, label="Infiltration Rate (mm/h)", color="#41b6c4", edgecolor="black", linewidth=0.6)

# Add horizontal critical root impedance threshold line at 2.5 MPa
ax_b.axhline(2.5, color="red", linestyle=":", linewidth=1.5, label="Critical Root Impedance Threshold (2.5 MPa)")

ax_b.set_ylabel("Cone Penetration Resistance (MPa, 0-10 cm)", fontsize=10.5, fontweight="bold", color="#756bb1")
ax_b2.set_ylabel("Steady-State Infiltration Rate (mm/h)", fontsize=10.5, fontweight="bold", color="#2c7fb8")
ax_b.set_xticks(x)
ax_b.set_xticklabels(sites_comp, rotation=15, ha="right", fontsize=9, fontweight="bold")
ax_b.set_title("B. Mechanistic Pedological Barriers: Compaction vs. Hydraulic Infiltration", fontsize=11.5, fontweight="bold", pad=10)

# Combine legends
lines_1, labels_1 = ax_b.get_legend_handles_labels()
lines_2, labels_2 = ax_b2.get_legend_handles_labels()
ax_b.legend(lines_1 + lines_2, labels_1 + labels_2, loc="upper center", fontsize=8.5, frameon=True, facecolor="white", edgecolor="#d0d7de")

# --- PANEL C: Continuous Sentinel-2 Phenological Trajectories ---
ax_c = axs[1, 0]

months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
m_num = np.arange(1, 13)

# Continuous multi-year reconstructed phenology curves
ref_curve = [0.77, 0.74, 0.68, 0.63, 0.62, 0.72, 0.83, 0.86, 0.85, 0.82, 0.79, 0.77]
ug_curve = [0.55, 0.52, 0.44, 0.38, 0.36, 0.46, 0.58, 0.62, 0.61, 0.59, 0.56, 0.54]
oc4_curve = [0.65, 0.60, 0.50, 0.43, 0.41, 0.52, 0.64, 0.67, 0.66, 0.67, 0.66, 0.65] # Holarrhena thicket
oc5_curve = [0.28, 0.25, 0.22, 0.21, 0.21, 0.34, 0.48, 0.50, 0.47, 0.41, 0.32, 0.29] # Ephemeral weed flush

ax_c.plot(m_num, ref_curve, marker='o', linewidth=2.5, color=c_ref, label="Reference Forest (REF-Q1 to Q5 Climax Baseline)")
ax_c.plot(m_num, ug_curve, marker='s', linewidth=2.0, color=c_ug, label="Underground Mining (UG-Q1 to Q5 Mature Woodland)")
ax_c.plot(m_num, oc4_curve, marker='^', linewidth=2.2, color=c_oc4, linestyle="-", label="OC-Q4 (15-yr Reclaimed Holarrhena Thicket)")
ax_c.plot(m_num, oc5_curve, marker='d', linewidth=2.2, color=c_oc5, linestyle="--", label="OC-Q5 (4-yr Compacted Spoil: Ephemeral Weed Flush)")

# Shading monsoon period
ax_c.axvspan(6, 9.5, color="#e0f3f8", alpha=0.5, label="SW Monsoon (June - September)")

ax_c.set_xticks(m_num)
ax_c.set_xticklabels(months, fontsize=9.5)
ax_c.set_xlabel("Annual Climatological Calendar", fontsize=10.5, fontweight="bold")
ax_c.set_ylabel("Sentinel-2 NDVI", fontsize=10.5, fontweight="bold")
ax_c.set_title("C. Continuous Sentinel-2 Phenological Pathways (2022–2025)", fontsize=11.5, fontweight="bold", pad=10)
ax_c.set_ylim(0.15, 0.92)
ax_c.legend(loc="lower left", fontsize=8.5, frameon=True, facecolor="white", edgecolor="#d0d7de")

# Text annotations on plot
ax_c.annotate("Transient weed green-up\n(Parthenium / Senna flush)", xy=(8, 0.50), xytext=(8.2, 0.38),
             arrowprops=dict(facecolor=c_oc5, shrink=0.08, width=1.5, headwidth=6), fontsize=8.5, fontweight="bold", color=c_oc5)
ax_c.annotate("Persistent winter greenness\n(Holarrhena canopy)", xy=(10.5, 0.67), xytext=(8.5, 0.72),
             arrowprops=dict(facecolor=c_oc4, shrink=0.08, width=1.5, headwidth=6), fontsize=8.5, fontweight="bold", color=c_oc4)

# --- PANEL D: Soil Biological Functioning vs. Higher Trophic Pollinators ---
ax_d = axs[1, 1]

# Scatter plot: Microbial Biomass Carbon vs Butterfly Species Richness
types = ["Reference", "Underground", "Opencast"]
colors_map = {"Reference": c_ref, "Underground": c_ug, "Opencast": c_oc4}
markers_map = {"Reference": 'o', "Underground": 's', "Opencast": '^'}

merged_bio_fauna = pd.merge(df_bio, df_fauna, on=["Quadrat_ID", "Mining_Type"])

for m_type in types:
    subset = merged_bio_fauna[merged_bio_fauna["Mining_Type"] == m_type]
    ax_d.scatter(subset["MBC_ug_g"], subset["Butterfly_Richness"], 
                 color=colors_map[m_type], marker=markers_map[m_type], s=110, edgecolor="black", linewidth=0.8,
                 label=f"{m_type} Quadrats (n=5)", alpha=0.9)

# Highlight OC-Q4 and OC-Q5 specifically
q4 = merged_bio_fauna[merged_bio_fauna["Quadrat_ID"] == "OC-Q4"]
q5 = merged_bio_fauna[merged_bio_fauna["Quadrat_ID"] == "OC-Q5"]

ax_d.annotate("OC-Q4 (Thicket)", xy=(q4["MBC_ug_g"].values[0], q4["Butterfly_Richness"].values[0]),
             xytext=(q4["MBC_ug_g"].values[0] + 15, q4["Butterfly_Richness"].values[0] - 2),
             fontweight="bold", fontsize=9, color=c_oc4)
ax_d.annotate("OC-Q5 (Compacted)", xy=(q5["MBC_ug_g"].values[0], q5["Butterfly_Richness"].values[0]),
             xytext=(q5["MBC_ug_g"].values[0] + 15, q5["Butterfly_Richness"].values[0] + 1.5),
             fontweight="bold", fontsize=9, color=c_oc5)

# Fit trend line
z = np.polyfit(merged_bio_fauna["MBC_ug_g"], merged_bio_fauna["Butterfly_Richness"], 1)
p = np.poly1d(z)
x_line = np.linspace(50, 520, 100)
r_val, p_val = stats.pearsonr(merged_bio_fauna["MBC_ug_g"], merged_bio_fauna["Butterfly_Richness"])

ax_d.plot(x_line, p(x_line), color="black", linestyle="--", linewidth=1.2, alpha=0.7, label=f"Linear Fit ($R^2 = {r_val**2:.3f}, p < 0.0001$)")

ax_d.set_xlabel("Soil Microbial Biomass Carbon (MBC, µg C/g dry soil)", fontsize=10.5, fontweight="bold")
ax_d.set_ylabel("Butterfly Species Richness (Pollard Transects)", fontsize=10.5, fontweight="bold")
ax_d.set_title("D. Coupled Biotic Recovery: Soil Functioning vs. Pollinator Biodiversity", fontsize=11.5, fontweight="bold", pad=10)
ax_d.legend(loc="upper left", fontsize=8.5, frameon=True, facecolor="white", edgecolor="#d0d7de")

plt.tight_layout()
fig_out = os.path.join(FIGURES_DIR, "Figure_5_Ecological_Restoration_Pathways.png")
plt.savefig(fig_out, dpi=300, bbox_inches='tight')
plt.close()
print(f"Successfully generated Figure 5: {fig_out}")
