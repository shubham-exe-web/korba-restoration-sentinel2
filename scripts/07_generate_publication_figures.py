"""
Script: 07_generate_publication_figures.py
Step 7: Regenerate publication figures (Figures 1 to 6 and Supplementary Figure S1) from the frozen, audited dataset at 300 DPI.
- Figure 1: Published sampling locations, GPS uncertainty (+/- 5m), and concentric spatial extraction zones (A: 10m/1-pix, B: 50x50m/25-pix, C: 50m radius/81-pix, D: 100m radius/317-pix).
- Figure 2: Seasonal multi-spectral surface reflectance profiles (Summer vs Monsoon vs Winter) across Level-2A BOA bands.
- Figure 3: Multi-temporal seasonal NDVI and NDRE trajectories with explicit IMD seasonal windows and peak retreat identification.
- Figure 4: Decoupling greenness persistence from seasonal amplitude with explicit quantitative syndrome boundaries.
- Figure 5: Reconciled PCA biplot and Hierarchical Cluster Dendrogram of vegetation signatures (Model A with Model B sensitivity note).
- Figure 6: Multi-scale spatial seasonality gradient across Korba coalfield (Zones A, B, C, D).
- Supplementary Figure S1: Exploratory field-satellite linkage matrix and carbon stock association.
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from scipy.cluster.hierarchy import linkage, dendrogram
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

plt.rcParams['font.sans-serif'] = 'Helvetica', 'Arial', 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#333333'
plt.rcParams['axes.linewidth'] = 0.8
plt.rcParams['grid.color'] = '#dddddd'
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.alpha'] = 0.6

os.makedirs("figures", exist_ok=True)

# -------------------------------------------------------------
# Figure 1: Published Sampling Locations & Spatial Geometry
# -------------------------------------------------------------
def plot_figure_1():
    print("Generating Figure 1: Sampling Locations & Spatial Geometry...")
    coords_df = pd.read_csv("data/01_Korba_quadrat_coordinates.csv")
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6), gridspec_kw={'width_ratios': [1.3, 1]})
    
    ug = coords_df[coords_df["Mining_type"] == "Underground"]
    oc = coords_df[coords_df["Mining_type"] == "Opencast"]
    
    ax1.scatter(ug["Longitude"], ug["Latitude"], color="#1b7837", s=130, edgecolor="black", linewidth=1.2, label="Underground Remnant Forest Quadrats (UG-Q1 to Q5)", zorder=4)
    ax1.scatter(oc["Longitude"], oc["Latitude"], color="#d95f02", s=130, edgecolor="black", linewidth=1.2, marker="s", label="Opencast Mine Spoil Quadrats (OC-Q1 to Q5)", zorder=4)
    
    for _, r in coords_df.iterrows():
        offset_y = 0.0025 if "UG" in r["ID"] else -0.0035
        offset_x = 0.001 if r["ID"] in ["UG-Q2", "OC-Q5"] else -0.003
        ax1.annotate(f"{r['ID']}\n({r['Elevation_m']:.0f}m)", (r["Longitude"] + offset_x, r["Latitude"] + offset_y),
                     fontsize=8.5, fontweight='bold', ha='center',
                     bbox=dict(boxstyle="round,pad=0.2", facecolor="white", alpha=0.85, edgecolor="#aaaaaa"))
        
    ax1.set_title("A. Published Field Sampling Locations in Korba Coalfield", fontsize=11, fontweight='bold', pad=12)
    ax1.set_xlabel("Longitude (°E)", fontsize=10, fontweight='bold')
    ax1.set_ylabel("Latitude (°N)", fontsize=10, fontweight='bold')
    ax1.grid(True)
    ax1.legend(loc="upper left", frameon=True, framealpha=0.9, fontsize=9)
    
    # Right: Spatial extraction geometry diagram
    ax2.set_xlim(-130, 130)
    ax2.set_ylim(-130, 130)
    ax2.set_aspect('equal')
    
    # Zone D: 100m radius circle (317 pixels)
    circle_d = plt.Circle((0, 0), 100, facecolor='#e0f3f8', alpha=0.5, edgecolor='#4575b4', linewidth=1.5, linestyle='--', label='Zone D: 100 m Radius Buffer (317 pixels, 3.17 ha)')
    # Zone C: 50m radius circle (81 pixels)
    circle_c = plt.Circle((0, 0), 50, facecolor='#fee090', alpha=0.6, edgecolor='#fdae61', linewidth=1.5, linestyle='-.', label='Zone C: 50 m Radius Buffer (81 pixels, 0.81 ha)')
    # Zone B: 50x50m box (25 pixels)
    rect_b = plt.Rectangle((-25, -25), 50, 50, facecolor='#abd9e9', alpha=0.7, edgecolor='#74add1', linewidth=1.5, label='Zone B: 50 × 50 m Window (25 pixels, 0.25 ha)')
    # GPS Uncertainty footprint: 20x20m box (+/- 5m jitter)
    rect_gps = plt.Rectangle((-10, -10), 20, 20, facecolor='none', edgecolor='#9970ab', linewidth=1.4, linestyle=':', label='GPS Positional Uncertainty (±5 m buffer)')
    # Zone A: 10x10m quadrat box (1 pixel)
    rect_a = plt.Rectangle((-5, -5), 10, 10, facecolor='#d73027', alpha=0.85, edgecolor='black', linewidth=1.8, label='Zone A: 10 × 10 m Quadrat Pixel Support (1 pixel, 0.01 ha)')
    
    ax2.add_patch(circle_d)
    ax2.add_patch(circle_c)
    ax2.add_patch(rect_b)
    ax2.add_patch(rect_gps)
    ax2.add_patch(rect_a)
    
    ax2.plot(0, 0, 'k+', markersize=10, markeredgewidth=2)
    ax2.annotate("Quadrat Center\n(Published GPS)", (0, 0), xytext=(15, -28),
                 arrowprops=dict(arrowstyle="->", color="black", lw=1), fontsize=8.5, fontweight='bold')
    
    ax2.set_title("B. Concentric Multi-Scale Spatial Extraction Windows", fontsize=11, fontweight='bold', pad=12)
    ax2.set_xlabel("Relative Distance East-West (m)", fontsize=10, fontweight='bold')
    ax2.set_ylabel("Relative Distance North-South (m)", fontsize=10, fontweight='bold')
    ax2.legend(loc="upper right", frameon=True, framealpha=0.9, fontsize=8)
    ax2.grid(True)
    
    plt.tight_layout()
    plt.savefig("figures/Figure_1_Sampling_Locations_Spatial_Windows.png", dpi=300)
    plt.close()
    print(" - Saved figures/Figure_1_Sampling_Locations_Spatial_Windows.png")

# -------------------------------------------------------------
# Figure 2: Seasonal Spectral Reflectance Profiles
# -------------------------------------------------------------
def plot_figure_2():
    print("Generating Figure 2: Seasonal Spectral Reflectance Profiles...")
    df_bands = pd.read_csv("data/06_Korba_seasonal_spectral_features.csv")
    zone_a = df_bands[df_bands["Spatial_Zone"] == "Zone_A_10m"]
    
    bands = ["B02", "B03", "B04", "B05", "B06", "B07", "B08", "B8A", "B11", "B12"]
    wl_nm = [490, 560, 665, 705, 740, 783, 842, 865, 1610, 2190]
    
    fig, axes = plt.subplots(1, 2, figsize=(13, 5), sharey=True)
    season_colors = {"Summer": "#d73027", "Monsoon": "#1b7837", "Winter": "#4575b4"}
    
    for ax, m_type, title in zip(axes, ["Underground", "Opencast"], ["A. Underground Remnant Forest Quadrats (UG-Q1 to Q5)", "B. Opencast Mine Overburden Quadrats (OC-Q1 to Q5)"]):
        sub = zone_a[zone_a["Mining_type"] == m_type]
        for season in ["Summer", "Monsoon", "Winter"]:
            s_data = sub[sub["Season"] == season]
            mean_vals = [s_data[f"{b}_median"].mean() for b in bands]
            std_vals = [s_data[f"{b}_median"].std() for b in bands]
            
            ax.plot(wl_nm, mean_vals, marker='o', linewidth=2, color=season_colors[season], label=f"{season} (Seasonal Composite)")
            ax.fill_between(wl_nm, np.array(mean_vals) - np.array(std_vals), np.array(mean_vals) + np.array(std_vals),
                            color=season_colors[season], alpha=0.15)
            
        ax.set_title(title, fontsize=10.5, fontweight='bold', pad=10)
        ax.set_xlabel("Wavelength (nm)", fontsize=10, fontweight='bold')
        ax.set_ylabel("Surface Reflectance (BOA)" if m_type == "Underground" else "", fontsize=10, fontweight='bold')
        ax.grid(True)
        ax.legend(loc="upper right", frameon=True, fontsize=9)
        
        ax.axvspan(700, 785, facecolor='#ffffbf', alpha=0.3, label='Red-Edge')
        ax.axvspan(785, 900, facecolor='#e6f598', alpha=0.3, label='NIR Plateau')
        
    plt.tight_layout()
    plt.savefig("figures/Figure_2_Seasonal_Spectral_Signatures.png", dpi=300)
    plt.close()
    print(" - Saved figures/Figure_2_Seasonal_Spectral_Signatures.png")

# -------------------------------------------------------------
# Figure 3: Multi-Temporal Phenology Trajectories
# -------------------------------------------------------------
def plot_figure_3():
    print("Generating Figure 3: Seasonal NDVI and NDRE Trajectories...")
    df_ts = pd.read_csv("data/08_Korba_phenology_timeseries.csv")
    zone_a = df_ts[df_ts["Spatial_Zone"] == "Zone_A_10m"].sort_values("Date")
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
    
    ug = zone_a[zone_a["Mining_type"] == "Underground"]
    oc = zone_a[zone_a["Mining_type"] == "Opencast"]
    dates = sorted(zone_a["Date"].unique())
    
    for q_id in zone_a["Quadrat_ID"].unique():
        q_data = zone_a[zone_a["Quadrat_ID"] == q_id]
        is_ug = "UG" in q_id
        c = "#1b7837" if is_ug else "#d95f02"
        ax1.plot(q_data["Date"], q_data["NDVI"], color=c, alpha=0.35, linewidth=1.2)
        ax2.plot(q_data["Date"], q_data["NDRE"], color=c, alpha=0.35, linewidth=1.2)
        
    ug_mean_ndvi = ug.groupby("Date")["NDVI"].mean()
    oc_mean_ndvi = oc.groupby("Date")["NDVI"].mean()
    ug_mean_ndre = ug.groupby("Date")["NDRE"].mean()
    oc_mean_ndre = oc.groupby("Date")["NDRE"].mean()
    
    ax1.plot(dates, [ug_mean_ndvi[d] for d in dates], color="#1b7837", linewidth=2.8, marker='o', label="Underground Mean (n=5)")
    ax1.plot(dates, [oc_mean_ndvi[d] for d in dates], color="#d95f02", linewidth=2.8, marker='s', label="Opencast Mean (n=5)")
    
    ax2.plot(dates, [ug_mean_ndre[d] for d in dates], color="#1b7837", linewidth=2.8, marker='o', label="Underground Mean (n=5)")
    ax2.plot(dates, [oc_mean_ndre[d] for d in dates], color="#d95f02", linewidth=2.8, marker='s', label="Opencast Mean (n=5)")
    
    for ax in [ax1, ax2]:
        ax.grid(True)
        ax.axvspan(-0.5, 1.5, facecolor='#d73027', alpha=0.1)
        ax.axvspan(1.5, 3.5, facecolor='#1b7837', alpha=0.1)
        ax.axvspan(3.5, 5.5, facecolor='#4575b4', alpha=0.1)
        
    ax1.text(0.5, 0.82, "Pre-monsoon Summer\n(Mar 25 & May 14)", ha='center', fontsize=9, fontweight='bold', color='#a50026')
    ax1.text(2.5, 0.82, "Monsoon & Retreat Green Season\n(Jun 13 & Oct 06)", ha='center', fontsize=9, fontweight='bold', color='#006837')
    ax1.text(4.5, 0.82, "Post-monsoon Winter\n(Nov 20 & Dec 30)", ha='center', fontsize=9, fontweight='bold', color='#313695')
    
    ax1.set_ylabel("NDVI", fontsize=11, fontweight='bold')
    ax1.set_title("A. Multi-Temporal Sentinel-2 NDVI Trajectories Across 10 Quadrats (Tile 44QPK)", fontsize=11, fontweight='bold', pad=10)
    ax1.legend(loc="upper left", frameon=True, fontsize=9)
    ax1.set_ylim(0.10, 0.90)
    
    ax2.set_ylabel("NDRE (Red-Edge Index)", fontsize=11, fontweight='bold')
    ax2.set_title("B. Multi-Temporal Sentinel-2 NDRE Trajectories (Canopy Chlorophyll Proxy)", fontsize=11, fontweight='bold', pad=10)
    ax2.set_xlabel("Observation Date (YYYY-MM-DD)", fontsize=10, fontweight='bold')
    ax2.tick_params(axis='x', rotation=20)
    ax2.set_ylim(0.00, 0.55)
    
    plt.tight_layout()
    plt.savefig("figures/Figure_3_NDVI_NDRE_Trajectories_UG_vs_OC.png", dpi=300)
    plt.close()
    print(" - Saved figures/Figure_3_NDVI_NDRE_Trajectories_UG_vs_OC.png")

# -------------------------------------------------------------
# Figure 4: Decoupling Greenness Persistence vs Phenological Amplitude
# -------------------------------------------------------------
def plot_figure_4():
    print("Generating Figure 4: Greenness Persistence vs Phenological Amplitude with Quantitative Boundaries...")
    df_link = pd.read_csv("data/09_Korba_field_satellite_linkage.csv")
    
    fig, ax = plt.subplots(figsize=(10, 7.5))
    ug = df_link[df_link["Mining_type"] == "Underground"]
    oc = df_link[df_link["Mining_type"] == "Opencast"]
    
    # Quantitative threshold boundaries
    # Domain 1: Persistent Greenness (Summer NDVI >= 0.30, Winter-to-Green-Season Ratio >= 0.90, Amplitude < 0.25)
    ax.axhspan(0.90, 1.20, xmin=0.0, xmax=0.55, facecolor="#1b7837", alpha=0.08, label="Persistent Greenness Domain (Ratio ≥ 0.90)")
    # Domain 2: Pronounced Seasonal Flush (Amplitude > 0.25, Summer NDVI < 0.25)
    ax.axvspan(0.25, 0.35, ymin=0.0, ymax=1.0, facecolor="#d95f02", alpha=0.08, label="Pronounced Seasonal Flush Domain (ΔNDVI > 0.25)")
    
    # Boundary threshold lines
    ax.axhline(0.90, color="#1b7837", linestyle="--", linewidth=1.5, alpha=0.7)
    ax.axvline(0.25, color="#d95f02", linestyle="--", linewidth=1.5, alpha=0.7)
    
    # Quadrats scatter
    ax.scatter(ug["ZoneA_Seasonal_Amplitude"], ug["ZoneA_Winter_to_Monsoon_Ratio"],
               s=ug["Carbon_Stock_tC_ha"]*16, color="#1b7837", edgecolor="black",
               linewidth=1.5, label="Underground Remnant Forest (UG)", alpha=0.85, zorder=4)
    ax.scatter(oc["ZoneA_Seasonal_Amplitude"], oc["ZoneA_Winter_to_Monsoon_Ratio"],
               s=oc["Carbon_Stock_tC_ha"]*16, color="#d95f02", edgecolor="black",
               linewidth=1.5, marker="s", label="Opencast Mine Spoil (OC)", alpha=0.85, zorder=4)
    
    for _, r in df_link.iterrows():
        offset_y = 0.015 if r["Quadrat_ID"] in ["OC-Q2", "UG-Q2"] else -0.02
        offset_x = 0.005 if r["Quadrat_ID"] not in ["UG-Q3"] else -0.018
        ax.annotate(f"{r['Quadrat_ID']}",
                    (r["ZoneA_Seasonal_Amplitude"] + offset_x, r["ZoneA_Winter_to_Monsoon_Ratio"] + offset_y),
                    fontsize=9, fontweight='bold',
                    bbox=dict(boxstyle="round,pad=0.2", facecolor="white", alpha=0.8, edgecolor="#cccccc"))
        
    # Domain Text Callouts
    ax.text(0.08, 1.05, "PERSISTENT GREENNESS DOMAIN\n(High Dry-Season Retention)\n• Winter-to-Green-Season Ratio ≥ 0.90\n• Summer NDVI ≥ 0.30\n(UG-Q1, UG-Q2, UG-Q4, OC-Q1, OC-Q4)",
            fontsize=9, fontweight='bold', color="#1b7837", bbox=dict(boxstyle="square,pad=0.3", facecolor="#f0f9e8", edgecolor="#1b7837", alpha=0.8))
    
    ax.text(0.255, 0.72, "PRONOUNCED SEASONAL FLUSH DOMAIN\n• Seasonal Amplitude ΔNDVI > 0.25\n• Summer NDVI < 0.25\n(OC-Q5: open bench, high seasonal amplitude)",
            fontsize=9, fontweight='bold', color="#d95f02", bbox=dict(boxstyle="square,pad=0.3", facecolor="#fee6ce", edgecolor="#d95f02", alpha=0.8))

    ax.text(0.12, 0.81, "DISTURBED TRANSITION\n• Intermediate amplitude (0.07–0.21)\n• Drainage feature trajectory (OC-Q2: Ratio 1.129)\n• Sal stand phenology (UG-Q3: ΔNDVI 0.075)",
            fontsize=8.5, fontstyle='italic', color="#555555", bbox=dict(boxstyle="round,pad=0.3", facecolor="#ffffff", edgecolor="#aaaaaa", alpha=0.8))

    ax.set_title("Seasonal Greenness Amplitude vs. Winter-to-Green-Season Persistence Ratio (Zone A, 10 m)", fontsize=11, fontweight='bold', pad=12)
    ax.set_xlabel("Seasonal Amplitude ΔNDVI (Monsoon & Retreat Composite − Summer Baseline)", fontsize=10, fontweight='bold')
    ax.set_ylabel("Winter-to-Green-Season Persistence Ratio (NDVI_Winter / NDVI_GreenSeason)", fontsize=10, fontweight='bold')
    ax.set_xlim(0.04, 0.32)
    ax.set_ylim(0.68, 1.18)
    ax.grid(True)
    
    legend1 = ax.legend(loc="upper left", frameon=True, fontsize=9)
    ax.add_artist(legend1)
    
    p1 = Line2D([0], [0], marker='o', color='w', label='Tree Carbon: 15 t C ha⁻¹', markerfacecolor='#888888', markersize=np.sqrt(15*16))
    p2 = Line2D([0], [0], marker='o', color='w', label='Tree Carbon: 25 t C ha⁻¹', markerfacecolor='#888888', markersize=np.sqrt(25*16))
    ax.legend(handles=[p1, p2], loc="lower left", frameon=True, fontsize=8.5, title="Symbol Size = Overstory Carbon Stock")
    
    plt.tight_layout()
    plt.savefig("figures/Figure_4_Herbaceous_vs_Shrub_Persistence.png", dpi=300)
    plt.close()
    print(" - Saved figures/Figure_4_Herbaceous_vs_Shrub_Persistence.png")

# -------------------------------------------------------------
# Figure 5: Reconciled PCA Biplot & Hierarchical Dendrogram
# -------------------------------------------------------------
def plot_figure_5():
    print("Generating Figure 5: PCA Biplot & Hierarchical Dendrogram...")
    idx_df = pd.read_csv("data/07_Korba_vegetation_indices.csv")
    zone_a = idx_df[idx_df["Spatial_Zone"] == "Zone_A_10m"].copy()
    
    pca_features = [
        "NDVI_Summer", "NDVI_Monsoon", "NDVI_Winter",
        "NDRE_Summer", "NDRE_Monsoon", "NDRE_Winter",
        "SWIR_Ratio_Summer", "SWIR_Ratio_Monsoon",
        "Seasonal_NDVI_Amplitude", "Winter_to_Monsoon_Ratio"
    ]
    
    X = zone_a[pca_features].values
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    pca = PCA(n_components=2)
    scores = pca.fit_transform(X_scaled)
    var_exp = pca.explained_variance_ratio_ * 100
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    
    for i, r in zone_a.reset_index().iterrows():
        c = "#1b7837" if r["Mining_type"] == "Underground" else "#d95f02"
        marker = 'o' if r["Mining_type"] == "Underground" else 's'
        ax1.scatter(scores[i, 0], scores[i, 1], color=c, marker=marker, s=130, edgecolor="black", zorder=4)
        ax1.annotate(r["Quadrat_ID"], (scores[i, 0] + 0.1, scores[i, 1] + 0.1), fontsize=9, fontweight='bold')
        
    loadings = pca.components_.T * 2.5
    for j, feat in enumerate(pca_features):
        ax1.arrow(0, 0, loadings[j, 0], loadings[j, 1], color="#555555", alpha=0.7, head_width=0.1, lw=1.2)
        ax1.text(loadings[j, 0] * 1.15, loadings[j, 1] * 1.15, feat.replace("_", " "), color="#333333", fontsize=7.5, ha='center', va='center')
        
    ax1.set_title(f"A. PCA Biplot of Seasonal Spectral Features\n(PC1: {var_exp[0]:.1f}%, PC2: {var_exp[1]:.1f}%; Model B Agreement: r > 0.94)", fontsize=11, fontweight='bold', pad=10)
    ax1.set_xlabel(f"Axis 1 ({var_exp[0]:.1f}% Variance) [Overall Foliar Greenness & Cover]", fontsize=10, fontweight='bold')
    ax1.set_ylabel(f"Axis 2 ({var_exp[1]:.1f}% Variance) [Dynamic Seasonal Amplitude vs Persistence]", fontsize=10, fontweight='bold')
    ax1.grid(True)
    ax1.axhline(0, color="gray", lw=0.8, linestyle=":")
    ax1.axvline(0, color="gray", lw=0.8, linestyle=":")
    
    Z = linkage(X_scaled, method="ward", metric="euclidean")
    dendrogram(Z, labels=zone_a["Quadrat_ID"].values, ax=ax2, color_threshold=3.5,
               above_threshold_color="#555555", leaf_font_size=10)
    ax2.set_title("B. Hierarchical Cluster Dendrogram (Ward's Linkage on Standardized Features)", fontsize=11, fontweight='bold', pad=10)
    ax2.set_xlabel("Quadrat Identifier", fontsize=10, fontweight='bold')
    ax2.set_ylabel("Euclidean Distance", fontsize=10, fontweight='bold')
    ax2.grid(True, axis='y')
    
    plt.tight_layout()
    plt.savefig("figures/Figure_5_PCA_Cluster_Vegetation_Signatures.png", dpi=300)
    plt.close()
    print(" - Saved figures/Figure_5_PCA_Cluster_Vegetation_Signatures.png")

# -------------------------------------------------------------
# Figure 6: Multi-Scale Spatial Seasonality Gradient Map
# -------------------------------------------------------------
def plot_figure_6():
    print("Generating Figure 6: Multi-scale Seasonality Gradient...")
    idx_df = pd.read_csv("data/07_Korba_vegetation_indices.csv")
    coords_df = pd.read_csv("data/01_Korba_quadrat_coordinates.csv")
    merged = idx_df.merge(coords_df, left_on="Quadrat_ID", right_on="ID")
    
    fig, axes = plt.subplots(2, 2, figsize=(13, 11), sharex=True, sharey=True)
    zones = ["Zone_A_10m", "Zone_B_50m", "Zone_C_50m", "Zone_D_100m"]
    titles = [
        "A. Zone A (10 × 10 m Pixel Support, 1 pixel)",
        "B. Zone B (50 × 50 m Square Window, 25 pixels)",
        "C. Zone C (50 m Radius Circular Buffer, 81 pixels)",
        "D. Zone D (100 m Radius Circular Buffer, 317 pixels)"
    ]
    
    for ax, z_name, title in zip(axes.flatten(), zones, titles):
        sub = merged[merged["Spatial_Zone"] == z_name]
        sc = ax.scatter(sub["Longitude"], sub["Latitude"], c=sub["Seasonal_NDVI_Amplitude"],
                        s=160, cmap="YlGn", edgecolor="black", linewidth=1.2, vmin=0.05, vmax=0.35, zorder=4)
        
        for _, r in sub.iterrows():
            ax.annotate(r["Quadrat_ID"], (r["Longitude"] + 0.0015, r["Latitude"] + 0.0015), fontsize=8, fontweight='bold')
            
        ax.set_title(title, fontsize=10, fontweight='bold', pad=8)
        ax.set_xlabel("Longitude (°E)", fontsize=9, fontweight='bold')
        ax.set_ylabel("Latitude (°N)", fontsize=9, fontweight='bold')
        ax.grid(True)
        
    cbar = fig.colorbar(sc, ax=axes.ravel().tolist(), orientation='horizontal', fraction=0.045, pad=0.08)
    cbar.set_label("Seasonal NDVI Amplitude (Monsoon & Retreat Composite − Summer Baseline)", fontsize=10, fontweight='bold')
    
    plt.savefig("figures/Figure_6_Spatial_Vegetation_Seasonality_Map.png", dpi=300)
    plt.close()
    print(" - Saved figures/Figure_6_Spatial_Vegetation_Seasonality_Map.png")

# -------------------------------------------------------------
# Supplementary Figure S1: Exploratory Field-Satellite Linkage Matrix
# -------------------------------------------------------------
def plot_supplementary_figure_s1():
    print("Generating Supplementary Figure S1: Exploratory Field-Satellite Linkage...")
    df_link = pd.read_csv("data/09_Korba_field_satellite_linkage.csv")
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6), gridspec_kw={'width_ratios': [1.2, 1]})
    colors = ["#1b7837" if "UG" in q else "#d95f02" for q in df_link["Quadrat_ID"]]
    
    ax1.barh(df_link["Quadrat_ID"], df_link["ZoneA_Seasonal_Amplitude"], color=colors, alpha=0.75, edgecolor="black", height=0.6)
    ax1.set_title("A. Seasonal NDVI Amplitude Across Verified Field Quadrats", fontsize=11, fontweight='bold', pad=10)
    ax1.set_xlabel("Seasonal Amplitude ΔNDVI (Monsoon − Summer)", fontsize=10, fontweight='bold')
    ax1.set_ylabel("Quadrat Identifier", fontsize=10, fontweight='bold')
    ax1.grid(True, axis='x')
    
    for i, r in df_link.iterrows():
        ax1.text(r["ZoneA_Seasonal_Amplitude"] + 0.005, i, f"{r['ZoneA_Seasonal_Amplitude']:.3f} | {r['Dominant_Tree_Taxa'][:20]}",
                 va='center', fontsize=8)

    ax2.scatter(df_link["Carbon_Stock_tC_ha"], df_link["ZoneA_NDVI_Summer"], c=colors, s=130, edgecolor="black", zorder=4)
    z = np.polyfit(df_link["Carbon_Stock_tC_ha"], df_link["ZoneA_NDVI_Summer"], 1)
    p = np.poly1d(z)
    x_vals = np.linspace(df_link["Carbon_Stock_tC_ha"].min(), df_link["Carbon_Stock_tC_ha"].max(), 50)
    ax2.plot(x_vals, p(x_vals), "k--", alpha=0.7, label=f"Exploratory Trend (r = 0.443, p = 0.200, 95% CI: [-0.259, 0.839])")
    
    for _, r in df_link.iterrows():
        ax2.annotate(r["Quadrat_ID"], (r["Carbon_Stock_tC_ha"] + 0.3, r["ZoneA_NDVI_Summer"] + 0.005), fontsize=8.5, fontweight='bold')
        
    ax2.set_title("B. Overstory Carbon Stock vs Pre-monsoon Summer NDVI (Zone A)", fontsize=11, fontweight='bold', pad=10)
    ax2.set_xlabel("Field Overstory Carbon Stock (t C ha⁻¹)", fontsize=10, fontweight='bold')
    ax2.set_ylabel("Pre-monsoon Summer NDVI (Zone A, 10 m)", fontsize=10, fontweight='bold')
    ax2.grid(True)
    ax2.legend(loc="lower right", frameon=True, fontsize=8.5)
    
    plt.tight_layout()
    plt.savefig("figures/Supplementary_Figure_S1_Linkage_Matrix.png", dpi=300)
    plt.close()
    print(" - Saved figures/Supplementary_Figure_S1_Linkage_Matrix.png")

if __name__ == "__main__":
    plot_figure_1()
    plot_figure_2()
    plot_figure_3()
    plot_figure_4()
    plot_figure_5()
    plot_figure_6()
    plot_supplementary_figure_s1()
    print("\nAll publication and supplementary figures generated successfully!")
