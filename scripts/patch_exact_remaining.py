"""
Clean script to update the remaining exact strings in scripts/08_generate_manuscript_and_supplemental.py.
"""

SCRIPT_PATH = "scripts/08_generate_manuscript_and_supplemental.py"

with open(SCRIPT_PATH, "r", encoding="utf-8") as f:
    text = f.read()

# 1. Species richness in abstract (both docx and md)
text = text.replace(
    'Tier 1 comprises 58 verified in-plot woody stems (GBH ≥ 10 cm) across 25 species with Plants of the World',
    'Tier 1 comprises 58 verified in-plot woody stems (GBH ≥ 10 cm) across 24 accepted species (25 reported field morphotaxa) with Plants of the World'
)

# 2. Annual maximum in Results 4.2 (both docx and md)
text = text.replace(
    'Date-by-date trajectories established that annual maximum greenness occurred on',
    'Date-by-date trajectories established that the maximum greenness among the six audited observations occurred on'
)

# 3. Permits in Results 4.4 (both docx and md)
old_p1 = 'The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush \n        followed by rapid post-monsoon senescence.'
old_p2 = 'The open, uncanopied spoil bench permits an intense monsoonal herbaceous flush followed by rapid post-monsoon senescence.'
new_p = 'The open, low-greenness spoil setting exhibited a pronounced seasonal amplitude that is compatible with seasonal ground-cover development, but the ground-layer contribution was not directly measured.'

text = text.replace(old_p1, new_p)
text = text.replace(old_p2, new_p)

# 4. Section 3.1 in Markdown
old_s31_block = (
    '        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "\n'
    '        "across 25 species. All botanical nomenclature was audited and standardized according to Plants of the World Online (POWO; Royal Botanic "\n'
    '        "Gardens, Kew), resolving standard taxonomic synonyms (e.g., *Shorea robusta* Gaertn., *Terminalia elliptica* Willd. [syn. *T. tomentosa* (Roxb.) "\n'
    '        "Wight & Arn.], *Holarrhena pubescens* Wall. ex G.Don [syn. *H. antidysenterica* (L.) Wall.], *Millettia pinnata* (L.) Panigrahi [syn. *Pongamia "\n'
    '        "pinnata* (L.) Pierre]). Several subcanopy woody taxa (*Ziziphus mauritiana*, *Carissa carandas*, *Holarrhena pubescens*) were directly measured "\n'
    '        "within this Tier 1 census.\\n"'
)

new_s31_block = (
    '        "individual woody tree and subcanopy shrub stems with girth at breast height (GBH) ≥ 10 cm (diameter at breast height [DBH] ≥ 3.18 cm) "\n'
    '        "representing 25 reported field morphotaxa. All botanical nomenclature was audited and standardized according to Plants of the World Online "\n'
    '        "(POWO; Royal Botanic Gardens, Kew), using the accepted-name field as the sole basis for taxonomic richness. Following synonym harmonization, "\n'
    '        "the inventory comprises 24 unique accepted species (e.g., *Terminalia tomentosa* (Roxb.) Wight & Arn. is harmonized as a synonym of *Terminalia elliptica* Willd., "\n'
    '        "*Holarrhena antidysenterica* (L.) Wall. as a synonym of *Holarrhena pubescens* Wall. ex G.Don, and *Bauhinia vahlii* as *Phanera vahlii* (Wight & Arn.) Benth.). "\n'
    '        "Several subcanopy woody taxa (*Ziziphus mauritiana*, *Carissa carandas*, *Holarrhena pubescens*) were directly measured within this Tier 1 census.\\n"'
)

text = text.replace(old_s31_block, new_s31_block)

# 5. Table 1 markdown row
text = text.replace(
    '| 58 stems (25 species) |',
    '| 58 stems (24 accepted species; 25 reported taxa) |'
)

with open(SCRIPT_PATH, "w", encoding="utf-8") as f:
    f.write(text)

print("Saved updated scripts/08_generate_manuscript_and_supplemental.py")
