# KFD Transcriptomic HDT Revision

Reproducible revision workspace for a host-directed therapy manuscript on **Kyasanur Forest Disease (KFD)**.

This repository now reflects the revised analysis used to answer peer review. It should be read as a **cross-flaviviral transcriptomic prioritization framework**, not as a validated KFD-specific multi-omics discovery study.

## What This Revision Does

- Replaces the earlier opaque scoring workflow with a deterministic, documented ranking pipeline.
- Uses three public **human dengue severity** transcriptomic cohorts as cross-flaviviral proxy datasets because KFD-specific blood transcriptomes were not available in GEO.
- Rescores a prespecified 50-gene host-response panel spanning inflammatory, endothelial, coagulation, platelet, neurological, and oxidative-stress biology.
- Generates revised manuscript files, supplementary materials, figures, and a point-by-point reviewer response.

## Scientific Position

- The strongest transcriptomic signal in the public datasets is **inflammatory/cytokine signaling**.
- **Endothelial and coagulation pathways remain mechanistically relevant for KFD**, but in this revision they are weaker and often cohort-specific compared with the inflammatory signal.
- Drug outputs are **hypothesis-generating only**. They are not clinical treatment recommendations.

## Key Revision Outputs

### Analysis tables

- `outputs/revision_tables/kfd_revision_targets.csv`
- `outputs/revision_tables/kfd_revision_pathway_summary.csv`
- `outputs/revision_tables/kfd_revision_drug_candidates.csv`
- `outputs/revision_tables/kfd_revision_weight_sensitivity_summary.csv`

### Figures

- `outputs/revision_figures/figure1_discovery_cohorts.png`
- `outputs/revision_figures/figure2_target_ranking.png`
- `outputs/revision_figures/figure3_signature_heatmap.png`
- `outputs/revision_figures/figure4_pathway_scores.png`
- `outputs/revision_figures/figure5_candidate_table.png`

### Submission package

- `manuscripts/Manuscript_KFD_MJDYPV_Revised_Blinded.docx`
- `manuscripts/TitlePage_KFD_MJDYPV_Revised.docx`
- `manuscripts/CoverLetter_KFD_MJDYPV_Revised.docx`
- `manuscripts/Response_to_Reviewers_KFD_MJDYPV.docx`
- `manuscripts/Supplementary_Materials_KFD_Revised.docx`

## Repository Structure

```text
KFD_HDT_Pipeline/
├── data/
│   ├── gene_signature.csv
│   └── revision/
├── outputs/
│   ├── revision_figures/
│   └── revision_tables/
├── manuscripts/
├── scripts/
│   ├── rebuild_kfd_revision.py
│   └── generate_mjdypv_revision_package.py
└── README.md
```

## Reproducing the Revision

```bash
pip install -r requirements.txt
python scripts/rebuild_kfd_revision.py
python scripts/generate_mjdypv_revision_package.py
```

## Main Methods Summary

1. Download processed GEO series-matrix files for `GSE18090`, `GSE43777`, and `GSE51808`.
2. Map probe IDs to gene symbols using GPL annotations.
3. Collapse duplicate probes by highest mean expression.
4. Run within-dataset severe-versus-non-severe contrasts using Welch's t test and Benjamini-Hochberg correction.
5. Score the prespecified 50-gene panel using:

   `0.45 × omics + 0.20 × tractability + 0.20 × pathway relevance + 0.15 × clinical-phase relevance`

6. Check weight sensitivity using equal-weight and omics-heavy alternatives.

## Current High-Level Findings

- Highest pathway mean score: `cytokine_signaling`
- Next highest: `coagulation_fibrinolysis`, `endothelial_barrier`
- Top-ranked genes: `IL1B`, `CXCL10`, `IL6`, `TNF`, `CCL2`
- Weight sensitivity remained stable:
  - base vs equal-weight: Spearman rho `0.865`
  - base vs omics-heavy: Spearman rho `0.961`

## Limitations

- No KFD-specific transcriptomic dataset was available.
- Only blood-derived transcriptomic data were analyzed.
- The 50-gene panel is prespecified and mechanistic, not an unbiased genome-wide KFD signature.
- Drug links support prioritization only; they do not establish efficacy or safety in KFD.

## Author

**Dr. Siddalingaiah H S**  
Professor, Department of Community Medicine  
Shridevi Institute of Medical Sciences and Research Hospital  
Tumkur, Karnataka, India  
Email: `hssling@yahoo.com`

## License

MIT
