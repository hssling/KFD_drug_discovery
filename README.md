# KFD (Kyasanur Forest Disease) Host-Directed Therapy Pipeline

[![GitHub Actions](https://github.com/hssling/KFD_drug_discovery/workflows/CI/badge.svg)](https://github.com/hssling/KFD_drug_discovery/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview

An integrated computational pipeline for identifying host-directed therapy (HDT) targets for **Kyasanur Forest Disease (KFD)**, a tick-borne viral hemorrhagic fever endemic to Karnataka's Western Ghats, India.

## Why KFD Matters

| Statistic | Value |
|-----------|-------|
| Annual cases | 400-500 |
| Case fatality | 3-5% |
| Endemic districts | Shimoga, Chikkamagaluru, Uttara Kannada |
| Specific treatment | **NONE** |
| Vaccine efficacy | 50-60% |

## Key Findings

### Priority Targets

| Rank | Target | Pathway | Rationale |
|------|--------|---------|-----------|
| 1 | **ANGPT2** | Endothelial | Vascular leak marker |
| 2 | **TNF** | Cytokine | Inflammation driver |
| 3 | **IL6** | Cytokine | Acute phase |
| 4 | **F3** | Coagulation | DIC initiator |
| 5 | **VWF** | Endothelial | Damage marker |

### Priority Drug Candidates (Available in Karnataka)

| Drug | Cost | Target | Priority |
|------|------|--------|----------|
| **Atorvastatin** | $5/course | Endothelium | High |
| **Tranexamic acid** | $10/course | Fibrinolysis | High |
| **Vitamin K** | $2 | Coagulation | High |
| **Dexamethasone** | $2 | GR | Medium |
| **N-Acetylcysteine** | $10 | ROS | Medium |

## Project Structure

```
KFD_HDT_Pipeline/
├── config/kfd_config.yaml
├── data/gene_signature.csv       # 50-gene VHF signature
├── outputs/figures/, tables/
├── manuscripts/
│   └── Manuscript_KFD_HDT_ENHANCED.docx
├── scripts/
│   ├── run_pipeline.py
│   ├── generate_figures.py
│   └── generate_manuscript.py
├── tests/test_pipeline.py
└── docs/METHODOLOGY.md
```

## Quick Start

```bash
git clone https://github.com/hssling/KFD_drug_discovery.git
cd KFD_drug_discovery
pip install -r requirements.txt
python scripts/run_pipeline.py
python scripts/generate_figures.py
python scripts/generate_manuscript.py
```

## Clinical Recommendations

### Immediate (District Hospital Level)
- Atorvastatin 40mg daily (vascular protection)
- Tranexamic acid 1g TID (if bleeding)
- Vitamin K 10mg (coagulation support)
- NAC 600mg TID (antioxidant)

### Referral Center
- Fresh frozen plasma
- Platelet transfusion
- Consider ribavirin (research)

## Author

**Dr. Siddalingaiah H S**  
Professor, Department of Community Medicine  
Shridevi Institute of Medical Sciences, Tumkur, Karnataka, India  
Email: hssling@yahoo.com | ORCID: 0000-0002-4771-8285

## License

MIT License

## Citation

```bibtex
@article{siddalingaiah2026kfd_hdt,
  title={Host-Directed Therapy for Kyasanur Forest Disease},
  author={Siddalingaiah, H S},
  journal={Indian Journal of Medical Microbiology},
  year={2026}
}
```
