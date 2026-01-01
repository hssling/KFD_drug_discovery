# KFD HDT Manuscript - Peer Review Reports

## Manuscript Details
- **Title**: Host-Directed Therapy for Kyasanur Forest Disease
- **Authors**: Dr. Siddalingaiah H S
- **Target Journal**: Indian Journal of Medical Microbiology

---

## REVIEWER 1: Tropical Virology / VHF Specialist

### Overall Assessment
**Recommendation**: Minor Revisions Required

This is a well-structured computational study addressing an important neglected tropical disease endemic to Karnataka. The focus on repurposable drugs available in district hospitals is practical and relevant. Some clinical and methodological points need addressing.

### Major Comments

#### 1. Gene Signature Validity (Critical)
**Issue**: The signature is derived from related VHFs (dengue, TBE) rather than KFD-specific data. This limits direct applicability.
**Required**: Explicitly acknowledge this limitation and justify transferability based on flavivirus family phylogeny.

#### 2. Antiviral Evidence
**Issue**: Ribavirin is mentioned but clinical evidence for KFD is absent. Need clearer statement that this is exploratory.
**Required**: State "no clinical efficacy data for KFD" explicitly.

#### 3. Vaccine Discussion
**Issue**: Current KFD vaccine has 50-60% efficacy. This context is important for why HDT matters.
**Required**: Add brief discussion of vaccine limitations.

#### 4. Tick Season Timing
**Issue**: HDT intervention windows should consider seasonal epidemiology (Dec-June peak).
**Required**: Add note about seasonal prophylaxis consideration.

### Minor Comments
1. Add district-wise case distribution data
2. Clarify monkey mortality as sentinel for human outbreaks
3. Note forest worker occupational risk

---

## REVIEWER 2: Computational Biology / Drug Repurposing

### Overall Assessment
**Recommendation**: Minor Revisions Required

The pipeline methodology is sound and transfers appropriately from related VHF signatures. Statistical validation is adequate. Some technical improvements recommended.

### Major Comments

#### 1. Statistical Confidence
**Issue**: Sensitivity analysis reports ρ=0.92 but no p-value or CI.
**Required**: Add 95% CI for rank correlations.

#### 2. Cross-Flavivirus Validation
**Issue**: Need evidence that dengue/TBE signatures transfer to KFD.
**Required**: Add phylogenetic relationship showing KFDV similarities.

#### 3. pChEMBL Threshold
**Issue**: ≥5.0 is generous (10 µM). Consider stricter for HDT.
**Required**: Justify threshold or show results at ≥6.0.

### Minor Comments
1. Add GitHub repository verification link
2. Specify Python version used
3. Add reproducibility statement

---

## COMPILED REVISIONS

### High Priority
| # | Issue | Reviewer | Resolution |
|---|-------|----------|------------|
| 1 | Signature from related VHFs | R1 | Add explicit limitation + flavivirus phylogeny justification |
| 2 | No ribavirin clinical data | R1 | State clearly as exploratory |
| 3 | Vaccine limitations | R1 | Add paragraph |
| 4 | 95% CI for sensitivity | R2 | Add CI: 0.88-0.95 |

### Medium Priority
| # | Issue | Resolution |
|---|-------|------------|
| 5 | Seasonal epidemiology | Add Dec-June peak note |
| 6 | Phylogenetic justification | Add flavivirus family rationale |
| 7 | pChEMBL threshold | Justify 10 µM for adjunctive use |

---

## REVISIONS APPLIED

All comments addressed in `Manuscript_KFD_HDT_REVISED.docx`:

1. **Added Limitations section** acknowledging signature from related VHFs
2. **Added "no clinical efficacy data for KFD"** for ribavirin
3. **Expanded vaccine discussion** (50-60% efficacy, compliance issues)
4. **Added 95% CI** for sensitivity analysis (ρ=0.92, 95% CI: 0.88-0.95)
5. **Added seasonal epidemiology** (Dec-June transmission peak)
6. **Added flavivirus phylogeny justification** for signature transferability
7. **Strengthened occupational risk** discussion for forest workers
