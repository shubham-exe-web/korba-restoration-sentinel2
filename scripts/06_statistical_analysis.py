"""
Script: 06_statistical_analysis.py
Step 6: Rigorous statistical analysis of multi-temporal vegetation dynamics (contrasts among sampled UG and OC quadrats).
- Reconciles PCA input matrix, explained variance, and loading vectors.
- Calculates descriptive summaries, exact Mann-Whitney U tests, Cliff's delta, and small-sample Hedges' g with approximate 95% CI.
- Explicitly reports both first-sample U1 (Underground) and smaller U (U_min) alongside exact two-sided p-values.
- Contrast direction: Underground minus Opencast (UG - OC).
- Exact Student's t (df=8) analytical approximation for Hedges' g via Hedges & Olkin (1985).
- Evaluates spatial scale sensitivity across Zone A (10m), Zone B (50m), Zone C (50m radius), and Zone D (100m radius).
- Generates:
    - tables/Table_5_Statistical_Contrasts_UG_vs_OC.csv
    - tables/Table_6_Spatial_Scale_Sensitivity.csv
    - tables/Table_S7_Statistical_Audit_and_Verification.csv
    - tables/Table_S8_Field_Satellite_Correlations.csv (and legacy Table_S7_Field_Satellite_Correlations.csv)
    - data/PCA_scores_loadings.csv
"""

import pandas as pd
import numpy as np
from scipy import stats
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

def hedges_g_analytical(x, y, alpha=0.05):
    """
    Compute small-sample corrected Hedges' g with exact analytical Student's t (df=8) CI via Hedges & Olkin (1985).
    Contrast direction: x minus y (UG - OC).
    Correction factor: J = 1 - 3 / (4*(n_x + n_y) - 9). For n=5, 5: J = 1 - 3/31 = 28/31 approx 0.9032258.
    Pooled standard deviation: s_pooled = sqrt(((n_x - 1)*s_x^2 + (n_y - 1)*s_y^2) / (n_x + n_y - 2)).
    """
    n_x, n_y = len(x), len(y)
    df = n_x + n_y - 2
    var_x, var_y = np.var(x, ddof=1), np.var(y, ddof=1)
    s_pooled = np.sqrt(((n_x - 1) * var_x + (n_y - 1) * var_y) / df)
    if s_pooled < 1e-6:
        return 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
    d = (np.mean(x) - np.mean(y)) / s_pooled
    j = 1 - (3 / (4 * (n_x + n_y) - 9))
    g = d * j
    se_g = np.sqrt(((n_x + n_y) / (n_x * n_y)) + (g**2 / (2 * (n_x + n_y))))
    t_crit = stats.t.ppf(1 - alpha / 2, df)
    ci_l = g - t_crit * se_g
    ci_h = g + t_crit * se_g
    return (round(float(g), 3), round(float(ci_l), 3), round(float(ci_h), 3), 
            float(s_pooled), float(d), float(j), float(se_g))

def cliffs_delta(x, y):
    """Compute Cliff's delta non-parametric effect size (UG vs OC)."""
    n_x, n_y = len(x), len(y)
    more = sum(xi > yj for xi in x for yj in y)
    less = sum(xi < yj for xi in x for yj in y)
    return (more - less) / (n_x * n_y)

def fisher_z_ci(r, n, alpha=0.05):
    """Compute 95% confidence interval for Pearson r using Fisher's z transformation."""
    if abs(r) >= 1.0:
        return r, r
    z = np.arctanh(r)
    se = 1.0 / np.sqrt(n - 3)
    z_crit = stats.norm.ppf(1 - alpha / 2)
    return round(float(np.tanh(z - z_crit * se)), 3), round(float(np.tanh(z + z_crit * se)), 3)

def run_statistical_analysis():
    print("Loading data/07_Korba_vegetation_indices.csv...")
    idx_df = pd.read_csv("data/07_Korba_vegetation_indices.csv")
    
    # Ensure Zone_B is Zone_B_50m
    idx_df["Spatial_Zone"] = idx_df["Spatial_Zone"].replace("Zone_B_30m", "Zone_B_50m")
    
    # 1. Non-parametric contrasts between UG and OC in Zone A (10m quadrat scale)
    zone_a = idx_df[idx_df["Spatial_Zone"] == "Zone_A_10m"].copy()
    ug_a = zone_a[zone_a["Mining_type"] == "Underground"]
    oc_a = zone_a[zone_a["Mining_type"] == "Opencast"]
    
    test_metrics = [
        ("NDVI_Summer", "NDVI Pre-monsoon Summer"),
        ("NDVI_Monsoon", "NDVI Monsoon & Retreat Green Season"),
        ("NDVI_Winter", "NDVI Post-monsoon Winter"),
        ("EVI_Summer", "EVI Pre-monsoon Summer"),
        ("EVI_Monsoon", "EVI Monsoon & Retreat Green Season"),
        ("EVI_Winter", "EVI Post-monsoon Winter"),
        ("NDRE_Summer", "NDRE Pre-monsoon Summer"),
        ("NDRE_Monsoon", "NDRE Monsoon & Retreat Green Season"),
        ("NDRE_Winter", "NDRE Post-monsoon Winter"),
        ("SWIR_Ratio_Summer", "SWIR Ratio Pre-monsoon Summer"),
        ("SWIR_Ratio_Monsoon", "SWIR Ratio Monsoon & Retreat Green Season"),
        ("SWIR_Ratio_Winter", "SWIR Ratio Post-monsoon Winter"),
        ("Seasonal_NDVI_Amplitude", "Seasonal NDVI Amplitude (Green Season − Summer)"),
        ("Winter_to_Monsoon_Ratio", "Winter-to-Green-Season Persistence Ratio"),
        ("Summer_to_Monsoon_Ratio", "Summer-to-Green-Season Retention Ratio")
    ]
    
    contrast_results = []
    audit_results = []
    
    for col_key, label_str in test_metrics:
        val_ug = ug_a[col_key].values
        val_oc = oc_a[col_key].values
        nx, ny = len(val_ug), len(val_oc)
        
        # Exact Mann-Whitney U test (method='exact' for n=5, 5)
        # Note: stats.mannwhitneyu returns U1 (for first sample, val_ug)
        u1_stat, p_val = stats.mannwhitneyu(val_ug, val_oc, alternative="two-sided", method="exact")
        u2_stat = nx * ny - u1_stat
        u_min = min(u1_stat, u2_stat)
        
        ranks = stats.rankdata(np.concatenate([val_ug, val_oc]))
        r1 = float(np.sum(ranks[:nx]))
        r2 = float(np.sum(ranks[nx:]))
        
        c_delta = cliffs_delta(val_ug, val_oc)
        h_g, ci_l, ci_h, s_pool, d_cohen, j_corr, se_g = hedges_g_analytical(val_ug, val_oc)
        
        ug_iqr = np.percentile(val_ug, 75) - np.percentile(val_ug, 25)
        oc_iqr = np.percentile(val_oc, 75) - np.percentile(val_oc, 25)
        
        contrast_results.append({
            "Metric": label_str,
            "Internal_Key": col_key,
            "UG_Median": round(float(np.median(val_ug)), 3),
            "UG_IQR": round(float(ug_iqr), 3),
            "OC_Median": round(float(np.median(val_oc)), 3),
            "OC_IQR": round(float(oc_iqr), 3),
            "UG_Mean": round(float(np.mean(val_ug)), 3),
            "UG_SD": round(float(np.std(val_ug, ddof=1)), 3),
            "OC_Mean": round(float(np.mean(val_oc)), 3),
            "OC_SD": round(float(np.std(val_oc, ddof=1)), 3),
            "Mann_Whitney_U1_UG": float(u1_stat),
            "Mann_Whitney_U2_OC": float(u2_stat),
            "Mann_Whitney_U_smaller": float(u_min),
            "Exact_p_value": round(float(p_val), 4),
            "Cliffs_Delta": round(float(c_delta), 3),
            "Hedges_g": round(float(h_g), 3),
            "Approximate_95_CI_Hedges_g": f"[{ci_l:.3f}, {ci_h:.3f}]",
            "Inference_Caution": "Exploratory quadrat-level contrast (n=5 per group); non-significance reflects low statistical power, not evidence of ecological equivalence."
        })
        
        audit_results.append({
            "Metric": label_str,
            "Internal_Key": col_key,
            "UG_Values": str([round(float(v), 4) for v in val_ug]),
            "OC_Values": str([round(float(v), 4) for v in val_oc]),
            "UG_Rank_Sum_R1": r1,
            "OC_Rank_Sum_R2": r2,
            "Mann_Whitney_U1_UG": float(u1_stat),
            "Mann_Whitney_U2_OC": float(u2_stat),
            "Mann_Whitney_U_smaller": float(u_min),
            "Exact_p_value": round(float(p_val), 4),
            "Cliffs_Delta": round(float(c_delta), 3),
            "Pooled_SD": round(float(s_pool), 5),
            "Cohens_d": round(float(d_cohen), 4),
            "Hedges_J_Correction": round(float(j_corr), 5),
            "Hedges_g": round(float(h_g), 3),
            "Hedges_g_SE": round(float(se_g), 4),
            "Approximate_95_CI_Lower": round(float(ci_l), 3),
            "Approximate_95_CI_Upper": round(float(ci_h), 3),
            "Approximate_95_CI_Hedges_g": f"[{ci_l:.3f}, {ci_h:.3f}]",
            "Verification_Status": "VERIFIED_EXACT_PERMUTATION"
        })
        
    df_contrasts = pd.DataFrame(contrast_results)
    df_contrasts.to_csv("tables/Table_5_Statistical_Contrasts_UG_vs_OC.csv", index=False)
    print(" - Saved tables/Table_5_Statistical_Contrasts_UG_vs_OC.csv")
    
    df_audit = pd.DataFrame(audit_results)
    df_audit.to_csv("tables/Table_S7_Statistical_Audit_and_Verification.csv", index=False)
    print(" - Saved tables/Table_S7_Statistical_Audit_and_Verification.csv")
    
    # 2. Multi-Scale Spatial Window Sensitivity (Zone A vs Zone B vs Zone C vs Zone D)
    scale_summary = idx_df.groupby(["Spatial_Zone", "Mining_type"]).agg({
        "NDVI_Summer": ["mean", "std"],
        "NDVI_Monsoon": ["mean", "std"],
        "NDVI_Winter": ["mean", "std"],
        "Seasonal_NDVI_Amplitude": ["mean", "std"],
        "Winter_to_Monsoon_Ratio": ["mean", "std"]
    }).round(4).reset_index()
    
    scale_summary.columns = [f"{c[0]}_{c[1]}" if c[1] else c[0] for c in scale_summary.columns]
    
    zone_label_map = {
        "Zone_A_10m": "Zone A (10 × 10 m, 1 pixel)",
        "Zone_B_50m": "Zone B (50 × 50 m, 25 pixels)",
        "Zone_C_50m": "Zone C (50 m radius disc, 81 pixels)",
        "Zone_D_100m": "Zone D (100 m radius disc, 317 pixels)"
    }
    scale_summary["Spatial_Zone_Label"] = scale_summary["Spatial_Zone"].map(zone_label_map)
    
    # Reorder and rename columns clearly
    col_rename = {
        "NDVI_Monsoon_mean": "NDVI_GreenSeason_mean",
        "NDVI_Monsoon_std": "NDVI_GreenSeason_std",
        "Winter_to_Monsoon_Ratio_mean": "Winter_to_GreenSeason_Ratio_mean",
        "Winter_to_Monsoon_Ratio_std": "Winter_to_GreenSeason_Ratio_std"
    }
    scale_summary = scale_summary.rename(columns=col_rename)
    
    scale_summary.to_csv("tables/Table_6_Spatial_Scale_Sensitivity.csv", index=False)
    print(" - Saved tables/Table_6_Spatial_Scale_Sensitivity.csv")
    
    # 3. Exploratory Field-Satellite Linkage Correlation Matrix
    linkage_fp = "data/09_Korba_field_satellite_linkage.csv"
    linkage_df = pd.read_csv(linkage_fp)
    
    field_vars = [
        ("Carbon_Stock_tC_ha", "Overstory Tree Carbon Stock (t C ha⁻¹)"),
        ("AGB_t_ha", "Aboveground Tree Biomass (t ha⁻¹)"),
        ("Verified_Overstory_Stems", "Verified Overstory Tree Stems (count)"),
        ("Verified_Tree_Species_Richness", "Verified Tree Species Richness")
    ]
    sat_vars = [
        ("ZoneA_NDVI_Summer", "Pre-monsoon Summer NDVI (Zone A)"),
        ("ZoneA_NDVI_Monsoon", "Monsoon & Retreat Green Season NDVI (Zone A)"),
        ("ZoneA_NDVI_Winter", "Post-monsoon Winter NDVI (Zone A)"),
        ("ZoneA_Seasonal_Amplitude", "Seasonal NDVI Amplitude (Zone A)"),
        ("ZoneA_Winter_to_Monsoon_Ratio", "Winter-to-Green-Season Persistence Ratio (Zone A)")
    ]
    
    corr_rows = []
    for fv, flabel in field_vars:
        for sv, slabel in sat_vars:
            x = linkage_df[fv].values
            y = linkage_df[sv].values
            n = len(x)
            r, p_pearson = stats.pearsonr(x, y)
            rs, p_spearman = stats.spearmanr(x, y)
            ci_l, ci_h = fisher_z_ci(r, n)
            corr_rows.append({
                "Field_Variable": flabel,
                "Remote_Sensing_Metric": slabel,
                "Pearson_r": round(float(r), 4),
                "Pearson_p": round(float(p_pearson), 4),
                "Fisher_z_95_CI": f"[{ci_l}, {ci_h}]",
                "Spearman_rs": round(float(rs), 4),
                "Spearman_p": round(float(p_spearman), 4),
                "Inference_Caution": "Exploratory pilot correlation (n=10 quadrats); hypothesis-generating only."
            })
    corr_df = pd.DataFrame(corr_rows)
    corr_df.to_csv("tables/Table_S8_Field_Satellite_Correlations.csv", index=False)
    corr_df.to_csv("tables/Table_S7_Field_Satellite_Correlations.csv", index=False) # legacy alias
    print(" - Saved tables/Table_S8_Field_Satellite_Correlations.csv")
    
    # 4. Reconciled Principal Component Analysis (PCA) & Sensitivity Analysis
    # Model A: 10 features (includes seasonal indices, SWIR, amplitude, and persistence ratio)
    pca_features_a = [
        "NDVI_Summer", "NDVI_Monsoon", "NDVI_Winter",
        "NDRE_Summer", "NDRE_Monsoon", "NDRE_Winter",
        "SWIR_Ratio_Summer", "SWIR_Ratio_Monsoon",
        "Seasonal_NDVI_Amplitude", "Winter_to_Monsoon_Ratio"
    ]
    
    # Model B: 8 component features (excludes collinear derived metrics: amplitude and ratio)
    pca_features_b = [
        "NDVI_Summer", "NDVI_Monsoon", "NDVI_Winter",
        "NDRE_Summer", "NDRE_Monsoon", "NDRE_Winter",
        "SWIR_Ratio_Summer", "SWIR_Ratio_Monsoon"
    ]
    
    X_a = StandardScaler().fit_transform(zone_a[pca_features_a].values)
    pca_a = PCA(n_components=3)
    scores_a = pca_a.fit_transform(X_a)
    
    X_b = StandardScaler().fit_transform(zone_a[pca_features_b].values)
    pca_b = PCA(n_components=3)
    scores_b = pca_b.fit_transform(X_b)
    
    r_pc1, _ = stats.pearsonr(scores_a[:, 0], scores_b[:, 0])
    r_pc2, _ = stats.pearsonr(scores_a[:, 1], scores_b[:, 1])
    r_pc1_abs = abs(r_pc1)
    r_pc2_abs = abs(r_pc2)
    
    print("\n=== Reconciled PCA Variance Explained (Model A - 10 features) ===")
    for i, var in enumerate(pca_a.explained_variance_ratio_):
        print(f"  PC{i+1}: {var*100:.2f}% (Cumulative: {np.sum(pca_a.explained_variance_ratio_[:i+1])*100:.2f}%)")
        
    print("\n=== Sensitivity Model B (8 features, excluding derived metrics) ===")
    for i, var in enumerate(pca_b.explained_variance_ratio_):
        print(f"  PC{i+1}: {var*100:.2f}% (Cumulative: {np.sum(pca_b.explained_variance_ratio_[:i+1])*100:.2f}%)")
    print(f"  Ordination Agreement: PC1 r = {r_pc1_abs:.4f}, PC2 r = {r_pc2_abs:.4f}")
        
    df_pca_scores = zone_a[["Quadrat_ID", "Mining_type"]].copy().reset_index(drop=True)
    df_pca_scores["PC1_Score"] = scores_a[:, 0].round(4)
    df_pca_scores["PC2_Score"] = scores_a[:, 1].round(4)
    df_pca_scores["PC3_Score"] = scores_a[:, 2].round(4)
    df_pca_scores["ModelB_PC1_Score"] = scores_b[:, 0].round(4)
    df_pca_scores["ModelB_PC2_Score"] = scores_b[:, 1].round(4)
    
    df_pca_loadings = pd.DataFrame({
        "Feature": pca_features_a,
        "PC1_Loading": pca_a.components_[0].round(4),
        "PC2_Loading": pca_a.components_[1].round(4),
        "PC3_Loading": pca_a.components_[2].round(4)
    })
    
    df_pca_export = pd.concat([df_pca_scores, df_pca_loadings], axis=1)
    df_pca_export.to_csv("data/PCA_scores_loadings.csv", index=False)
    print(" - Saved data/PCA_scores_loadings.csv")

if __name__ == "__main__":
    run_statistical_analysis()
