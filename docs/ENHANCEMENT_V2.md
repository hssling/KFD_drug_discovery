# KFD Enhancement v2

This additive package strengthens the revised KFD manuscript without modifying the existing submission files.

## What was added

- Target-level random-effects meta-analysis for the 50-gene panel
- Approximate 95% confidence intervals for pooled log2 fold-change
- Heterogeneity metrics (`I2`, `tau2`)
- Evidence-tier classification:
  - `cross-cohort`
  - `single-cohort`
  - `mechanistic-only`
- A separate memo:
  - `manuscripts/KFD_Scientific_Enhancement_Memo_v2.docx`

## New outputs

- `outputs/enhanced_v2_tables/kfd_enhanced_v2_meta_targets.csv`
- `outputs/enhanced_v2_tables/kfd_enhanced_v2_evidence_summary.csv`
- `outputs/enhanced_v2_tables/kfd_enhanced_v2_translational_targets.csv`
- `outputs/enhanced_v2_figures/figure_v2_meta_priority.png`
- `outputs/enhanced_v2_figures/figure_v2_pathway_heterogeneity.png`

## Main interpretation

- The added meta-analysis reinforces that the current public-data evidence base is strongest for inflammatory targets.
- No gene in the 50-gene panel met a strict `cross-cohort` evidence tier using the current recurrence rule.
- Several translationally interesting endothelial/coagulation genes remain in the `mechanistic-only` tier.

## Why this still matters

This result is scientifically useful because it:

1. makes uncertainty explicit,
2. prevents overstatement,
3. shows exactly which claims are data-driven versus mechanism-driven,
4. provides a stronger platform for future KFD-specific validation.

## Recommended use

Use v2 when you want to:

- prepare a stronger future submission,
- write a grant or protocol for follow-up validation,
- justify why KFD-specific transcriptomic or biomarker data are now the most important next step.
