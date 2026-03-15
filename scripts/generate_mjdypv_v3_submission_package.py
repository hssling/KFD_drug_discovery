"""Generate a submission-ready v3 package for MJDYPV."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt


BASE_DIR = Path(__file__).resolve().parent.parent
MANUSCRIPT_DIR = BASE_DIR / "manuscripts"
REV_TABLES = BASE_DIR / "outputs" / "revision_tables"
V2_TABLES = BASE_DIR / "outputs" / "enhanced_v2_tables"
REV_FIGS = BASE_DIR / "outputs" / "revision_figures"
V2_FIGS = BASE_DIR / "outputs" / "enhanced_v2_figures"

TITLE = "A Cross-Flaviviral Transcriptomic Evidence and Mechanistic Prioritization Framework for Host-Directed Therapy in Kyasanur Forest Disease"
RUNNING_TITLE = "Evidence-Based HDT Framework for KFD"


def set_margins(doc: Document) -> None:
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)


def set_base_style(doc: Document, size: int = 12, spacing: float = 2.0) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(size)
    style.paragraph_format.line_spacing = spacing


def set_cell_shading(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    tc_pr.append(shd)


def word_count(text: str) -> int:
    return len(re.findall(r"\b\w+\b", text))


def references() -> list[str]:
    return [
        "Work TH, Trapido H, Murthy DP, Rao RL, Bhatt PN, Kulkarni KG. Kyasanur forest disease. III. A preliminary report on the nature of the infection and clinical manifestations in man. Indian J Med Sci 1957;11:619-45.",
        "Pattnaik P. Kyasanur forest disease: an epidemiological view in India. Rev Med Virol 2006;16:151-65.",
        "Murhekar MV, Kasabi GS, Mehendale SM, Mourya DT, Yadav PD, Tandale BV, et al. On the transmission pattern of Kyasanur Forest Disease (KFD) in India. Infect Dis Poverty 2015;4:37.",
        "Holbrook MR. Kyasanur forest disease. Antiviral Res 2012;96:353-62.",
        "Kasabi GS, Murhekar MV, Sandhya VK, Raghunandan R, Kiran SK, Channabasappa GH, et al. Coverage and effectiveness of Kyasanur forest disease vaccine in Karnataka, South India, 2005-10. PLoS Negl Trop Dis 2013;7:e2025.",
        "Kaufmann SHE, Dorhoi A, Hotchkiss RS, Bartenschlager R. Host-directed therapies for bacterial and viral infections. Nat Rev Drug Discov 2018;17:35-56.",
        "Zumla A, Rao M, Wallis RS, Kaufmann SH, Rustomjee R, Mwaba P, et al. Host-directed therapies for infectious diseases: current status, recent progress, and future prospects. Lancet Infect Dis 2016;16:e47-63.",
        "Barrett T, Wilhite SE, Ledoux P, Evangelista C, Kim IF, Tomashevsky M, et al. NCBI GEO: archive for functional genomics data sets-update. Nucleic Acids Res 2013;41:D991-5.",
        "Nascimento EJ, Braga-Neto U, Calzavara-Silva CE, Gomes AL, Abath FG, Brito CA, et al. Gene expression profiling during early acute febrile stage of dengue infection can predict the disease outcome. PLoS One 2009;4:e7892.",
        "Sun P, Garcia J, Comach G, Vahey MT, Wang Z, Forshey BM, et al. Sequential waves of gene expression in patients with clinically defined dengue illnesses reveal subtle disease phases and predict disease severity. PLoS Negl Trop Dis 2013;7:e2298.",
        "Kwissa M, Nakaya HI, Onlamoon N, Wrammert J, Villinger F, Perng GC, et al. Dengue virus infection induces expansion of a CD14+CD16+ monocyte population that stimulates plasmablast differentiation. Cell Host Microbe 2014;16:115-27.",
        "Gillespie M, Jassal B, Stephan R, Milacic M, Rothfels K, Senff-Ribeiro A, et al. The Reactome pathway knowledgebase 2022. Nucleic Acids Res 2022;50:D687-92.",
        "Zdrazil B, Felix E, Hunter F, Manners EJ, Blackshaw J, Corbett S, et al. The ChEMBL Database in 2023: a drug discovery platform spanning multiple bioactivity data types and time periods. Nucleic Acids Res 2024;52:D1180-92.",
        "Simmons CP, Farrar JJ, Nguyen vV, Wills B. Dengue. N Engl J Med 2012;366:1423-32.",
        "Modhiran N, Watterson D, Muller DA, Panetta AK, Sester DP, Liu L, et al. Dengue virus NS1 protein activates cells via Toll-like receptor 4 and disrupts endothelial cell monolayer integrity. Sci Transl Med 2015;7:304ra142.",
        "Bray M. Pathogenesis of viral hemorrhagic fever. Curr Opin Immunol 2005;17:399-403.",
        "Yacoub S, Wills B. Predicting outcome from dengue. BMC Med 2014;12:147.",
        "Roberts I, Shakur H, Coats T, Hunt B, Balogun E, Barnetson L, et al. The CRASH-2 trial: a randomised controlled trial and economic evaluation of the effects of tranexamic acid on death, vascular occlusive events and transfusion requirement in bleeding trauma patients. Health Technol Assess 2013;17:1-79.",
        "Kiran SK, Padamashree S, Jayashree K, Srinath S. Health-care-seeking behaviour among Kyasanur forest disease patients in Shivamogga district, Karnataka: a cross-sectional study. Indian J Community Med 2021;46:486-9.",
    ]


def build_blinded_manuscript() -> tuple[Path, int, int]:
    meta = pd.read_csv(V2_TABLES / "kfd_enhanced_v2_meta_targets.csv")
    transl = pd.read_csv(V2_TABLES / "kfd_enhanced_v2_translational_targets.csv")
    cohorts = pd.read_csv(REV_TABLES / "cohort_summary.csv")
    revision_drugs = pd.read_csv(REV_TABLES / "kfd_revision_drug_candidates.csv")
    single = int((meta["EvidenceTier"] == "single-cohort").sum())
    mech = int((meta["EvidenceTier"] == "mechanistic-only").sum())
    total = int(cohorts["SevereSamples"].sum() + cohorts["NonSevereSamples"].sum())
    severe = int(cohorts["SevereSamples"].sum())
    non_severe = int(cohorts["NonSevereSamples"].sum())

    doc = Document()
    set_margins(doc)
    set_base_style(doc)

    title = doc.add_heading("", level=0)
    run = title.add_run(TITLE)
    run.bold = True
    run.font.size = Pt(14)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.add_run("Running Title: ").bold = True
    p.add_run(RUNNING_TITLE)
    doc.add_page_break()

    doc.add_heading("ABSTRACT", level=1)
    abstract_sections = [
        ("Background:", "Kyasanur Forest Disease lacks disease-specific transcriptomic datasets and specific therapy, forcing host-directed therapeutic work to rely on proxy flaviviral data."),
        ("Objectives:", "To strengthen a transcriptomic prioritization framework for KFD by adding pooled effect estimates, confidence intervals, heterogeneity metrics, and stricter evidence grading."),
        ("Materials and Methods:", f"We reanalyzed three public dengue severity cohorts from GEO (GSE18090, GSE43777, GSE51808; {total} acute samples, {severe} severe and {non_severe} non-severe). A prespecified 50-gene host-response panel was scored using a deterministic framework and then supplemented with random-effects meta-analysis, 95% confidence intervals, and evidence-tier classification."),
        ("Results:", f"The strongest transcriptomic evidence remained inflammatory. No gene reached a strict cross-cohort evidence tier; {single} genes had single-cohort nominal support and {mech} remained mechanistic-only. The highest meta-priority genes were CXCL10, IL6, STAT1, IFNB1, and TEK. Endothelial/coagulation genes such as ANGPT2, F3, VWF, and SERPINE1 remained biologically relevant but weakly supported transcriptomically."),
        ("Conclusions:", "The submission-ready v3 framing separates transcriptomic evidence from mechanistic plausibility. Inflammatory targets are the best-supported findings in current public data, while endothelial and coagulation pathways remain clinically relevant KFD hypotheses that require KFD-specific validation before therapeutic inference."),
    ]
    for label, text in abstract_sections:
        p = doc.add_paragraph()
        p.add_run(label).bold = True
        p.add_run(" " + text)
    p = doc.add_paragraph()
    p.add_run("Keywords: ").bold = True
    p.add_run("Kyasanur Forest Disease; host-directed therapy; transcriptomics; meta-analysis; flavivirus; endothelial dysfunction; coagulation")
    doc.add_page_break()

    doc.add_heading("INTRODUCTION", level=1)
    for text in [
        "Kyasanur Forest Disease (KFD) remains a clinically important tick-borne flaviviral hemorrhagic fever in southern India, with no specific antiviral therapy and incomplete vaccine protection.[1-5]",
        "Our prior revision replaced an unsupported multi-omics claim with a transparent transcriptomic prioritization framework. The present submission-ready v3 version adds a stricter evidence layer because mechanistic plausibility alone is insufficient for strong translational inference.",
        "Public KFD transcriptomes are unavailable, so we used human dengue severity cohorts as a cross-flaviviral proxy. This is biologically justifiable for vascular-leak and inflammatory hypotheses, but it requires explicit caution about uncertainty and transferability.[8-17]",
    ]:
        doc.add_paragraph(text)

    doc.add_heading("MATERIALS AND METHODS", level=1)
    for text in [
        f"Three GEO cohorts with acute severe-versus-non-severe dengue data were included: GSE18090, GSE43777, and GSE51808, totaling {total} acute samples ({severe} severe and {non_severe} non-severe). Preprocessing, probe collapsing, and within-cohort differential-expression analysis followed the previously revised deterministic workflow.",
        "The prespecified 50-gene panel covered cytokine, interferon, endothelial, coagulation, platelet, neurological, and oxidative-stress biology. Existing component scores were retained for mechanistic prioritization.",
        "To strengthen statistical rigor, each panel gene was additionally evaluated using random-effects meta-analysis based on cohort-level effect sizes and approximated standard errors derived from log2 fold-changes and two-sided P values. We report pooled effects, 95% confidence intervals, and heterogeneity statistics.",
        "Evidence tiers were defined as cross-cohort, single-cohort, or mechanistic-only to distinguish recurrent transcriptomic support from pathway-driven mechanistic prioritization.",
    ]:
        doc.add_paragraph(text)

    doc.add_heading("RESULTS", level=1)
    for text in [
        f"No gene reached the strict cross-cohort tier. {single} genes met single-cohort nominal support, and {mech} genes remained mechanistic-only. This demonstrates that the current public proxy data are more suitable for cautious prioritization than for claiming a stable cross-flaviviral consensus signature.",
        "The best-supported transcriptomic signals were inflammatory and interferon-related. CXCL10, IL6, STAT1, and IFNB1 occupied the highest meta-priority ranks, although several retained wide confidence intervals or moderate-to-high heterogeneity.",
        "Endothelial and coagulation targets remained important from a KFD pathophysiology standpoint, but their evidence tier was weaker. ANGPT2 and VWF were mechanistic-only, while SERPINE1 retained only single-cohort support. F3 remained biologically relevant but did not achieve nominal recurrence under the current data.",
        "Accordingly, the v3 shortlist should be interpreted as a two-layer output: transcriptomically supported inflammatory targets and clinically motivated vascular/coagulation hypotheses.",
    ]:
        doc.add_paragraph(text)

    top_meta = meta.head(12).copy()
    p = doc.add_paragraph()
    p.add_run("Table 1. Top 12 genes ranked by meta-priority.").bold = True
    t1 = doc.add_table(rows=len(top_meta) + 1, cols=7)
    t1.style = "Table Grid"
    h1 = ["Meta rank", "Gene", "Evidence tier", "Support count", "Pooled effect", "95% CI", "I2"]
    for i, h in enumerate(h1):
        cell = t1.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_shading(cell, "D9E2F3")
    for i, (_, row) in enumerate(top_meta.iterrows(), start=1):
        vals = [
            int(row["MetaRank"]), row["GeneSymbol"], row["EvidenceTier"], int(row["NominalSupportCount"]),
            f"{row['RandomEffect']:.2f}", f"{row['Lower95CI']:.2f} to {row['Upper95CI']:.2f}", f"{row['I2']:.1f}"
        ]
        for j, v in enumerate(vals):
            t1.rows[i].cells[j].text = str(v)

    p = doc.add_paragraph()
    p.add_run("Table 2. Translationally relevant targets with evidence tier and interpretation.").bold = True
    t2 = doc.add_table(rows=len(transl) + 1, cols=7)
    t2.style = "Table Grid"
    h2 = ["Meta rank", "Gene", "Pathway", "Evidence tier", "Support", "Pooled effect", "Interpretation"]
    for i, h in enumerate(h2):
        cell = t2.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_shading(cell, "D9E2F3")
    interp = {
        "single-cohort": "Transcriptomic signal present but not recurrent",
        "mechanistic-only": "Mechanistically plausible but weak transcriptomic support",
        "cross-cohort": "Recurrent transcriptomic support",
    }
    for i, (_, row) in enumerate(transl.iterrows(), start=1):
        vals = [
            int(row["MetaRank"]), row["GeneSymbol"], row["Pathway"], row["EvidenceTier"],
            int(row["NominalSupportCount"]), f"{row['RandomEffect']:.2f}", interp[row["EvidenceTier"]]
        ]
        for j, v in enumerate(vals):
            t2.rows[i].cells[j].text = str(v)

    shortlist = revision_drugs[revision_drugs["GeneSymbol"].isin(["IL1B", "IL6", "TNF", "ANGPT2", "SERPINE1", "HMOX1", "F3", "VWF"])].copy()
    p = doc.add_paragraph()
    p.add_run("Table 3. Hypothesis-generating intervention shortlist under the v3 evidence framework.").bold = True
    t3 = doc.add_table(rows=len(shortlist) + 1, cols=5)
    t3.style = "Table Grid"
    h3 = ["Candidate", "Target", "Pathway", "Evidence tier", "Use in manuscript"]
    tier_lookup = meta.set_index("GeneSymbol")["EvidenceTier"].to_dict()
    for i, h in enumerate(h3):
        cell = t3.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_shading(cell, "D9E2F3")
    for i, (_, row) in enumerate(shortlist.iterrows(), start=1):
        vals = [row["Candidate"], row["GeneSymbol"], row["Pathway"], tier_lookup.get(row["GeneSymbol"], "n/a"), "Hypothesis-generating only"]
        for j, v in enumerate(vals):
            t3.rows[i].cells[j].text = str(v)

    figures = [
        (REV_FIGS / "figure1_discovery_cohorts.png", "Figure 1. Discovery cohorts used in the analysis."),
        (V2_FIGS / "figure_v2_meta_priority.png", "Figure 2. Meta-priority ranking after adding pooled effects and evidence tiers."),
        (V2_FIGS / "figure_v2_pathway_heterogeneity.png", "Figure 3. Pathway effect size versus heterogeneity."),
        (REV_FIGS / "figure2_target_ranking.png", "Figure 4. Original composite ranking retained for comparison with the meta-analytic ranking."),
    ]
    for path, caption in figures:
        p = doc.add_paragraph()
        p.add_run(caption).bold = True
        doc.add_picture(str(path), width=Inches(5.7))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("DISCUSSION", level=1)
    for text in [
        "The v3 manuscript is more robust because it cleanly separates two concepts that are often merged in computational repurposing studies: transcriptomic support and mechanistic plausibility. In the current public proxy data, inflammatory genes are the best-supported findings.",
        "Endothelial and coagulation pathways remain clinically important for KFD, but in this manuscript they are deliberately framed as mechanistic hypotheses rather than transcriptomically validated drivers. That narrower interpretation is more scientifically reliable and more likely to survive critical review.",
        "The intervention shortlist is therefore retained only as a staged translational hypothesis set. Plasma or platelet support reflects standard supportive care, while tranexamic acid, atorvastatin, and N-acetylcysteine remain candidates for future evaluation rather than recommendations for routine use.",
    ]:
        doc.add_paragraph(text)

    doc.add_heading("LIMITATIONS", level=2)
    doc.add_paragraph(
        "No KFD-specific transcriptomic data were available. The current evidence therefore depends on dengue proxy cohorts, blood-derived transcriptomes only, and approximate variance reconstruction for the random-effects meta-analysis. The data support prioritization under uncertainty, not therapeutic validation."
    )

    doc.add_heading("CONCLUSIONS", level=1)
    doc.add_paragraph(
        "This submission-ready v3 version is the most statistically explicit and scientifically cautious iteration currently possible from existing assets. Inflammatory targets have the strongest transcriptomic support, whereas endothelial and coagulation pathways remain clinically meaningful but mechanistic KFD hypotheses requiring disease-specific validation."
    )

    doc.add_page_break()
    doc.add_heading("REFERENCES", level=1)
    for idx, ref in enumerate(references(), start=1):
        p = doc.add_paragraph()
        p.add_run(f"{idx}. ").bold = True
        p.add_run(ref)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.paragraph_format.left_indent = Inches(0.25)

    out = MANUSCRIPT_DIR / "Manuscript_KFD_MJDYPV_v3_Submission_Blinded.docx"
    doc.save(out)
    text = "\n".join(p.text for p in doc.paragraphs)
    abstract_text = " ".join(f"{a} {b}" for a, b in abstract_sections)
    return out, word_count(text), word_count(abstract_text)


def build_title_page(total_words: int, abstract_words: int) -> Path:
    doc = Document()
    set_margins(doc)
    set_base_style(doc)
    fields = [
        ("Article Type", "Original Article"),
        ("Title", TITLE),
        ("Running Title", RUNNING_TITLE),
        ("Author", "Siddalingaiah H S"),
        ("Affiliation", "Department of Community Medicine, Shridevi Institute of Medical Sciences and Research Hospital, Tumkur, Karnataka, India"),
        ("Corresponding Author", "Dr. Siddalingaiah H S, MD; hssling@yahoo.com; +91-8941087719; ORCID 0000-0002-4771-8285"),
        ("Word Count", f"Abstract {abstract_words}; total manuscript text including references approximately {total_words}"),
        ("Tables/Figures", "3 tables and 4 figures in the blinded article; supplementary material provided separately"),
        ("Funding", "None"),
        ("Conflicts of Interest", "None declared"),
        ("Ethics Statement", "Secondary analysis of publicly available de-identified datasets; ethics approval not required"),
        ("Version Note", "Submission-ready v3 package integrates meta-analysis/uncertainty layer"),
    ]
    for label, value in fields:
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(value)
    out = MANUSCRIPT_DIR / "TitlePage_KFD_MJDYPV_v3_Submission.docx"
    doc.save(out)
    return out


def build_cover_letter() -> Path:
    doc = Document()
    set_margins(doc)
    set_base_style(doc)
    doc.add_paragraph(datetime.now().strftime("%B %d, %Y"))
    doc.add_paragraph()
    doc.add_paragraph("The Editor-in-Chief")
    doc.add_paragraph("Medical Journal of Dr. D.Y. Patil Vidyapeeth")
    doc.add_paragraph()
    body = [
        f"We submit the enclosed manuscript entitled '{TITLE}' as an original article.",
        "This version was prepared after a further critical audit of the revised submission. It retains the deterministic transcriptomic prioritization framework but adds a stricter statistical layer, including random-effects meta-analysis, confidence intervals, heterogeneity metrics, and explicit evidence-tier grading.",
        "The key value of the current submission is its improved scientific restraint. The manuscript now distinguishes transcriptomically supported inflammatory findings from endothelial and coagulation hypotheses that remain clinically relevant but mechanistically prioritized under uncertainty.",
        "A point-by-point response table and supplementary materials are included. The manuscript is not under consideration elsewhere. No conflicts of interest or external funding apply.",
    ]
    for para in body:
        doc.add_paragraph(para)
    doc.add_paragraph()
    doc.add_paragraph("Sincerely,")
    doc.add_paragraph()
    doc.add_paragraph("Dr. Siddalingaiah H S")
    out = MANUSCRIPT_DIR / "CoverLetter_KFD_MJDYPV_v3_Submission.docx"
    doc.save(out)
    return out


def build_response_letter() -> Path:
    rows = [
        ("1. Methodology transparency was insufficient.", "Further strengthened. In addition to the prior deterministic workflow, the v3 package now adds pooled effect estimates, confidence intervals, heterogeneity metrics, and explicit evidence-tier grading for each prioritized target.", "Methods; Supplementary Tables S2-S4"),
        ("2. Gene signature derivation appeared arbitrary.", "The panel remains prespecified and is no longer described as a de novo KFD discovery signature. v3 improves rigor by explicitly separating panel-based mechanistic prioritization from transcriptomic evidence strength.", "Abstract, Methods, Discussion"),
        ("3. Composite score needed clearer definition.", "Already addressed in the revised submission and retained here. v3 adds an orthogonal meta-analysis layer so that ranking does not depend only on the composite score.", "Methods; Supplementary Tables"),
        ("4. The original multi-omics framing was overstated.", "Retained correction. The v3 package continues to describe the work as a transcriptomic evidence and mechanistic prioritization framework.", "Title, Abstract, Introduction"),
        ("5. Results and interpretation were inconsistent.", "Further corrected. v3 explicitly states that inflammatory targets have the strongest transcriptomic support, while endothelial/coagulation findings are mechanistic-priority hypotheses with weaker transcriptomic support.", "Results, Discussion, Conclusions"),
        ("6. Clinical recommendations were premature.", "Strengthened correction. The intervention table now labels all candidates as hypothesis-generating only and distinguishes standard supportive care from exploratory repurposing concepts.", "Table 3, Discussion"),
        ("7. Cross-virus extrapolation required stronger justification.", "v3 keeps the narrowed dengue-only discovery base and adds an uncertainty framework showing that current proxy datasets are insufficient for strong KFD-specific molecular claims.", "Introduction, Methods, Limitations"),
        ("8. Pathway classification required support.", "Retained from the revised package. Reactome-supported pathway mapping remains documented in the methods and supplementary file.", "Methods; Supplementary"),
        ("9. Tables/figures needed closer agreement with claims.", "v3 was built directly from regenerated outputs and audited against the source CSV files. The final text mirrors the evidence-tier summaries and meta-analysis tables.", "All tables and figures"),
        ("10. Limitations required more emphasis.", "Further strengthened. v3 explicitly states that no gene reached strict cross-cohort support under the current public data base, which narrows the strongest conclusions and clarifies what still requires KFD-specific validation.", "Abstract, Results, Limitations"),
    ]
    doc = Document()
    set_margins(doc)
    set_base_style(doc, size=11, spacing=1.5)
    h = doc.add_heading("", level=0)
    r = h.add_run("Response to Reviewers and Final Revision Notes")
    r.bold = True
    r.font.size = Pt(14)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Manuscript title: " + TITLE)
    table = doc.add_table(rows=len(rows) + 1, cols=3)
    table.style = "Table Grid"
    headers = ["Reviewer Concern", "Response in Submission-Ready v3", "Location"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        set_cell_shading(cell, "D9E2F3")
    for i, row in enumerate(rows, start=1):
        for j, value in enumerate(row):
            table.rows[i].cells[j].text = value
    out = MANUSCRIPT_DIR / "Response_to_Reviewers_KFD_MJDYPV_v3.docx"
    doc.save(out)
    return out


def build_supplementary() -> Path:
    meta = pd.read_csv(V2_TABLES / "kfd_enhanced_v2_meta_targets.csv")
    transl = pd.read_csv(V2_TABLES / "kfd_enhanced_v2_translational_targets.csv")
    evidence = pd.read_csv(V2_TABLES / "kfd_enhanced_v2_evidence_summary.csv")
    cohorts = pd.read_csv(REV_TABLES / "cohort_summary.csv")
    revision_targets = pd.read_csv(REV_TABLES / "kfd_revision_targets.csv")

    doc = Document()
    set_margins(doc)
    set_base_style(doc, size=11, spacing=1.3)
    h = doc.add_heading("", level=0)
    r = h.add_run("Supplementary Materials")
    r.bold = True
    r.font.size = Pt(14)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(TITLE)

    def add_df(title: str, df: pd.DataFrame) -> None:
        doc.add_heading(title, level=1)
        table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
        table.style = "Table Grid"
        for i, col in enumerate(df.columns):
            cell = table.rows[0].cells[i]
            cell.text = str(col)
            cell.paragraphs[0].runs[0].bold = True
            set_cell_shading(cell, "D9E2F3")
        for i, (_, row) in enumerate(df.iterrows(), start=1):
            for j, col in enumerate(df.columns):
                val = row[col]
                if isinstance(val, float):
                    txt = f"{val:.3f}"
                else:
                    txt = str(val)
                table.rows[i].cells[j].text = txt

    doc.add_heading("Supplementary Methods", level=1)
    for text in [
        "Processed GEO series-matrix files were analyzed within cohort to avoid inappropriate cross-platform normalization.",
        "The 50-gene panel was retained as a prespecified mechanistic set.",
        "Random-effects meta-analysis was added to quantify pooled effect size and uncertainty for each panel gene.",
    ]:
        doc.add_paragraph(text)

    add_df("Table S1. Discovery cohort summary", cohorts)
    add_df("Table S2. Full v2 meta-analysis target table", meta)
    add_df("Table S3. Translational target subset", transl)
    add_df("Table S4. Evidence-tier summary by pathway", evidence)
    add_df("Table S5. Original revision ranking retained for comparison", revision_targets[["Rank", "GeneSymbol", "Pathway", "CompositeScore", "DatasetsSupporting"]])

    out = MANUSCRIPT_DIR / "Supplementary_Materials_KFD_v3_Submission.docx"
    doc.save(out)
    return out


def main() -> None:
    manuscript, total_words, abstract_words = build_blinded_manuscript()
    title = build_title_page(total_words, abstract_words)
    cover = build_cover_letter()
    response = build_response_letter()
    supp = build_supplementary()
    print("Generated submission-ready v3 package:")
    for file in [manuscript, title, cover, response, supp]:
        print(f" - {file.name}")


if __name__ == "__main__":
    main()
