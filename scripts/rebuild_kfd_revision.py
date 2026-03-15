"""Rebuild the KFD revision analysis with transparent, deterministic methods."""

from __future__ import annotations

import csv
import gzip
import io
import math
import re
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
import seaborn as sns
from scipy import stats


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data" / "revision"
TABLE_DIR = BASE_DIR / "outputs" / "revision_tables"
FIG_DIR = BASE_DIR / "outputs" / "revision_figures"

for directory in (DATA_DIR, TABLE_DIR, FIG_DIR):
    directory.mkdir(parents=True, exist_ok=True)

warnings.filterwarnings("ignore", message="Precision loss occurred in moment calculation")


@dataclass(frozen=True)
class DatasetConfig:
    accession: str
    matrix_url: str
    platform: str
    group_parser: Callable[[dict[str, list[list[str]]]], pd.DataFrame]
    citation_label: str


DATASETS = [
    DatasetConfig(
        accession="GSE18090",
        matrix_url="https://ftp.ncbi.nlm.nih.gov/geo/series/GSE18nnn/GSE18090/matrix/GSE18090_series_matrix.txt.gz",
        platform="GPL570",
        citation_label="Brazilian PBMC cohort; DHF vs DF",
        group_parser=lambda meta: pd.DataFrame(
            {
                "sample_id": meta["Sample_geo_accession"][0],
                "title": meta["Sample_title"][0],
            }
        ).assign(
            severity=lambda df: np.where(
                df["title"].str.contains("DHF"), "severe",
                np.where(df["title"].str.contains(r"\bDF\b"), "non_severe", "exclude"),
            ),
            phase="acute",
        ),
    ),
    DatasetConfig(
        accession="GSE51808",
        matrix_url="https://ftp.ncbi.nlm.nih.gov/geo/series/GSE51nnn/GSE51808/matrix/GSE51808_series_matrix.txt.gz",
        platform="GPL13158",
        citation_label="Thai whole-blood cohort; DHF vs DF",
        group_parser=lambda meta: pd.DataFrame(
            {
                "sample_id": meta["Sample_geo_accession"][0],
                "title": meta["Sample_title"][0],
            }
        ).assign(
            severity=lambda df: np.where(
                df["title"].str.contains("Hemorrhagic"), "severe",
                np.where(df["title"].str.contains("Dengue Fever"), "non_severe", "exclude"),
            ),
            phase="acute",
        ),
    ),
    DatasetConfig(
        accession="GSE43777",
        matrix_url="https://ftp.ncbi.nlm.nih.gov/geo/series/GSE43nnn/GSE43777/matrix/GSE43777-GPL570_series_matrix.txt.gz",
        platform="GPL570",
        citation_label="Venezuelan PBMC longitudinal cohort; acute DHF vs DF",
        group_parser=lambda meta: pd.DataFrame(
            {
                "sample_id": meta["Sample_geo_accession"][0],
                "title": meta["Sample_title"][0],
                "phase": _extract_characteristic(meta, "phase"),
                "severity_text": _extract_characteristic(meta, "severity"),
            }
        ).assign(
            severity=lambda df: np.where(
                df["severity_text"].str.contains("DHF", case=False), "severe",
                np.where(df["severity_text"].str.contains(r"\bDF\b", case=False), "non_severe", "exclude"),
            )
        ),
    ),
]


PATHWAY_DEFINITIONS = {
    "cytokine_signaling": {
        "score": 1.00,
        "phase": "febrile",
        "reactome": "Cytokine Signaling in Immune system / Interferon Signaling",
        "genes": {
            "IL1B", "IL6", "IL10", "TNF", "CXCL10", "CCL2", "CCL8", "CXCL11",
            "STAT1", "STAT2", "IRF7", "ISG15", "OAS1", "OAS2", "IFI27", "IFI44L",
            "IFIT1", "IFIT2", "IFIT3", "MX1", "MX2", "RSAD2", "XAF1", "EPSTI1",
            "SIGLEC1", "GBP1", "GBP5", "IFI6", "IFI44", "IFI35", "ISG20",
        },
    },
    "coagulation_fibrinolysis": {
        "score": 0.95,
        "phase": "hemorrhagic",
        "reactome": "Platelet activation, signaling and aggregation / Hemostasis",
        "genes": {"SERPINE1", "F3", "VWF", "THBD", "FGA", "PLAU", "PLAT", "TFPI"},
    },
    "endothelial_barrier": {
        "score": 0.95,
        "phase": "hemorrhagic",
        "reactome": "Cell junction organization / VEGFA-VEGFR2 pathway",
        "genes": {"ANGPT2", "ANGPT1", "KDR", "TEK", "VEGFA", "EDN1", "SELE", "ICAM1", "VCAM1"},
    },
    "platelet_activation": {
        "score": 0.90,
        "phase": "hemorrhagic",
        "reactome": "Platelet activation, signaling and aggregation",
        "genes": {"ITGA2B", "ITGB3", "GP1BA", "PF4", "PPBP", "SELP", "TREML1"},
    },
    "monocyte_innate_activation": {
        "score": 0.85,
        "phase": "febrile",
        "reactome": "Innate Immune System / Toll-like Receptor Cascades",
        "genes": {"S100A8", "S100A9", "LILRB1", "FCGR1A", "FCGR1B", "TLR7", "TLR8", "AIM2"},
    },
    "oxidative_stress": {
        "score": 0.80,
        "phase": "recovery",
        "reactome": "Cellular responses to stress",
        "genes": {"HMOX1", "NQO1", "SOD2", "TXN", "TXNIP", "NFE2L2"},
    },
    "neurological_barrier": {
        "score": 0.75,
        "phase": "neurological",
        "reactome": "Neuronal System / Tight junction interactions",
        "genes": {"AQP4", "CLDN5", "OCLN", "TJP1", "BDNF", "NGF"},
    },
}

PATHWAY_ALIAS = {
    "cytokine": "cytokine_signaling",
    "interferon": "cytokine_signaling",
    "endothelial": "endothelial_barrier",
    "coagulation": "coagulation_fibrinolysis",
    "platelet": "platelet_activation",
    "oxidative": "oxidative_stress",
    "neurological": "neurological_barrier",
    "neuroprotection": "neurological_barrier",
}

PHASE_SCORES = {
    "hemorrhagic": 1.00,
    "febrile": 0.85,
    "neurological": 0.75,
    "recovery": 0.60,
    "protective": 0.60,
    "both": 0.75,
}

DRUGGABILITY_SCORES = {
    "high": 1.00,
    "moderate": 0.65,
    "supportive": 0.45,
    "low": 0.20,
}

TARGET_DRUGS = {
    "SERPINE1": [("Tranexamic acid", "supportive", "Hypothesis-generating bleeding-control adjunct aligned to fibrinolysis imbalance")],
    "IL1B": [("Anakinra", "moderate", "Target-matched anti-inflammatory biologic; specialist or research setting only")],
    "IL6": [("Tocilizumab", "moderate", "Target-matched anti-inflammatory biologic; specialist or research setting only")],
    "TNF": [("Pentoxifylline", "moderate", "Indirect TNF-modulating repurposing candidate")],
    "CXCL10": [("No direct low-cost repurposed agent", "low", "Biomarker-priority target rather than immediate repurposing candidate")],
    "ANGPT2": [("Atorvastatin", "moderate", "Mechanistic endothelial-stabilization hypothesis rather than direct antagonism")],
    "VWF": [("Fresh frozen plasma / platelet support", "supportive", "Supportive care rather than target-specific inhibition")],
    "F3": [("Fresh frozen plasma / coagulation monitoring", "supportive", "Pathway-level supportive strategy, not direct blockade")],
    "HMOX1": [("N-acetylcysteine", "moderate", "Redox-modulating adjunct hypothesis")],
}


def _extract_characteristic(meta: dict[str, list[list[str]]], label: str) -> list[str]:
    for row in meta.get("Sample_characteristics_ch1", []):
        if row and row[0].lower().startswith(f"{label}:"):
            return [value.split(":", 1)[1].strip() if ":" in value else value for value in row]
    return [""] * len(meta["Sample_geo_accession"][0])


def fetch_text(url: str) -> str:
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    if url.endswith(".gz"):
        return gzip.decompress(response.content).decode("utf-8", errors="replace")
    return response.text


def parse_series_matrix(config: DatasetConfig) -> tuple[dict[str, list[list[str]]], pd.DataFrame]:
    text = fetch_text(config.matrix_url)
    lines = text.splitlines()
    meta: dict[str, list[list[str]]] = {}
    table_start = None
    table_end = None

    for index, line in enumerate(lines):
        if line.startswith("!series_matrix_table_begin"):
            table_start = index + 1
            continue
        if line.startswith("!series_matrix_table_end"):
            table_end = index
            break
        if table_start is None and line.startswith("!Sample_"):
            parts = next(csv.reader([line], delimiter="\t"))
            key = parts[0][1:]
            values = [item.strip('"') for item in parts[1:]]
            meta.setdefault(key, []).append(values)

    if table_start is None or table_end is None:
        raise RuntimeError(f"Could not locate expression table for {config.accession}")

    expr = pd.read_csv(io.StringIO("\n".join(lines[table_start:table_end])), sep="\t")
    return meta, expr


def parse_annotation(platform: str) -> pd.DataFrame:
    prefix = platform[:-3] + "nnn"
    url = f"https://ftp.ncbi.nlm.nih.gov/geo/platforms/{prefix}/{platform}/annot/{platform}.annot.gz"
    text = fetch_text(url)
    lines = text.splitlines()
    start = lines.index("!platform_table_begin") + 1
    end = lines.index("!platform_table_end")
    annot = pd.read_csv(io.StringIO("\n".join(lines[start:end])), sep="\t", low_memory=False)
    annot.columns = [column.strip() for column in annot.columns]

    id_column = next(column for column in annot.columns if column.lower() == "id")
    gene_column = next(
        column
        for column in annot.columns
        if column.lower() in {"gene symbol", "gene_symbol", "symbol"}
    )
    return annot[[id_column, gene_column]].rename(columns={id_column: "ID_REF", gene_column: "GeneSymbol"})


def benjamini_hochberg(pvalues: pd.Series) -> pd.Series:
    order = np.argsort(pvalues.values)
    ranked = pvalues.values[order]
    n = len(ranked)
    adjusted = ranked * n / np.arange(1, n + 1)
    adjusted = np.minimum.accumulate(adjusted[::-1])[::-1]
    adjusted = np.clip(adjusted, 0, 1)
    result = pd.Series(index=pvalues.index[order], data=adjusted)
    return result.reindex(pvalues.index)


def collapse_to_genes(expression: pd.DataFrame, annotation: pd.DataFrame, sample_columns: list[str]) -> pd.DataFrame:
    merged = expression.merge(annotation, on="ID_REF", how="left")
    merged = merged.dropna(subset=["GeneSymbol"]).copy()
    merged["GeneSymbol"] = (
        merged["GeneSymbol"]
        .astype(str)
        .str.split(r" /// | // |///|//", regex=True)
        .str[0]
        .str.strip()
    )
    merged = merged[(merged["GeneSymbol"] != "") & (merged["GeneSymbol"] != "---")]

    values = merged[sample_columns].apply(pd.to_numeric, errors="coerce")
    if values.max().max() > 50:
        values = np.log2(values.clip(lower=1) + 1)
    merged[sample_columns] = values
    merged["probe_mean"] = merged[sample_columns].mean(axis=1, skipna=True)
    merged = merged.sort_values("probe_mean", ascending=False).drop_duplicates("GeneSymbol")
    return merged[["GeneSymbol", *sample_columns]]


def differential_expression(gene_matrix: pd.DataFrame, metadata: pd.DataFrame) -> pd.DataFrame:
    sample_columns = metadata["sample_id"].tolist()
    severe_samples = metadata.loc[metadata["severity"] == "severe", "sample_id"].tolist()
    non_severe_samples = metadata.loc[metadata["severity"] == "non_severe", "sample_id"].tolist()

    records = []
    for _, row in gene_matrix.iterrows():
        severe_values = row[severe_samples].astype(float).values
        non_severe_values = row[non_severe_samples].astype(float).values
        t_stat, pvalue = stats.ttest_ind(severe_values, non_severe_values, equal_var=False, nan_policy="omit")
        log_fc = float(np.nanmean(severe_values) - np.nanmean(non_severe_values))
        records.append(
            {
                "GeneSymbol": row["GeneSymbol"],
                "log2FC": log_fc,
                "pvalue": pvalue if not math.isnan(pvalue) else 1.0,
                "mean_severe": float(np.nanmean(severe_values)),
                "mean_non_severe": float(np.nanmean(non_severe_values)),
            }
        )

    deg = pd.DataFrame(records)
    deg["fdr"] = benjamini_hochberg(deg["pvalue"])
    deg["direction"] = np.where(deg["log2FC"] >= 0, "up", "down")
    return deg.sort_values(["fdr", "pvalue", "log2FC"], ascending=[True, True, False])


def classify_gene(gene: str) -> tuple[str, float, str]:
    for pathway, info in PATHWAY_DEFINITIONS.items():
        if gene in info["genes"]:
            return pathway, info["score"], info["phase"]
    if gene.startswith(("IFI", "IFIT", "ISG", "OAS", "MX", "GBP", "RSAD", "XAF", "SIGLEC")):
        return "cytokine_signaling", PATHWAY_DEFINITIONS["cytokine_signaling"]["score"], "febrile"
    return "host_response_other", 0.70, "febrile"


def load_candidate_panel() -> pd.DataFrame:
    panel = pd.read_csv(BASE_DIR / "data" / "gene_signature.csv")
    panel = panel.rename(columns={"Symbol": "GeneSymbol"}).copy()
    panel["PathwayRaw"] = panel["Pathway"].str.lower()
    panel["Pathway"] = panel["PathwayRaw"].map(PATHWAY_ALIAS).fillna(panel["PathwayRaw"])
    panel["PhaseBucket"] = panel["Phase_Relevance"].str.lower()
    panel["PathwayScore"] = panel["Pathway"].map(
        lambda pathway: PATHWAY_DEFINITIONS.get(pathway, {"score": 0.70})["score"]
    )
    panel["PhaseScore"] = panel["PhaseBucket"].map(PHASE_SCORES).fillna(0.70)
    panel["ReactomeModule"] = panel["Pathway"].map(
        lambda pathway: PATHWAY_DEFINITIONS.get(pathway, {"reactome": "Immune System"})["reactome"]
    )
    panel["TractabilityScore"] = panel["Druggability"].str.lower().map(DRUGGABILITY_SCORES).fillna(0.20)
    return panel


def build_target_table(candidate_panel: pd.DataFrame, dataset_results: dict[str, pd.DataFrame]) -> pd.DataFrame:
    all_stats = []
    for _, row in candidate_panel.iterrows():
        per_dataset = {}
        for accession, deg in dataset_results.items():
            hit = deg[deg["GeneSymbol"] == row["GeneSymbol"]]
            if not hit.empty:
                per_dataset[accession] = {
                    "log2FC": float(hit.iloc[0]["log2FC"]),
                    "pvalue": float(hit.iloc[0]["pvalue"]),
                    "fdr": float(hit.iloc[0]["fdr"]),
                }

        logfcs = [stats_row["log2FC"] for stats_row in per_dataset.values()]
        if logfcs:
            consensus_direction = "up" if np.nanmedian(logfcs) >= 0 else "down"
            supporting = [
                accession
                for accession, stats_row in per_dataset.items()
                if stats_row["pvalue"] <= 0.05
                and abs(stats_row["log2FC"]) >= 0.30
                and ((stats_row["log2FC"] >= 0) == (consensus_direction == "up"))
            ]
            median_abs_log2fc = float(np.median(np.abs(logfcs)))
            best_pvalue = min(stats_row["pvalue"] for stats_row in per_dataset.values())
            best_fdr = min(stats_row["fdr"] for stats_row in per_dataset.values())
        else:
            consensus_direction = "not_observed"
            supporting = []
            median_abs_log2fc = 0.0
            best_pvalue = 1.0
            best_fdr = 1.0

        recurrence_score = len(supporting) / len(dataset_results)
        effect_score = min(median_abs_log2fc / 1.5, 1.0)
        omics_score = 0.65 * recurrence_score + 0.35 * effect_score

        therapies = TARGET_DRUGS.get(row["GeneSymbol"], [("No direct repurposed agent", "low", "Biomarker-priority target")])
        druggability_label = therapies[0][1]
        tractability_score = DRUGGABILITY_SCORES[druggability_label]

        composite_score = (
            0.45 * omics_score
            + 0.20 * tractability_score
            + 0.20 * row["PathwayScore"]
            + 0.15 * row["PhaseScore"]
        )

        all_stats.append(
            {
                "GeneSymbol": row["GeneSymbol"],
                "GeneName": row["Gene"],
                "ConsensusDirection": consensus_direction,
                "DatasetsSupporting": ",".join(supporting) if supporting else "none",
                "DatasetCount": len(supporting),
                "MedianAbsLog2FC": median_abs_log2fc,
                "BestPValue": best_pvalue,
                "BestFDR": best_fdr,
                "Pathway": row["Pathway"],
                "ReactomeModule": row["ReactomeModule"],
                "PhaseBucket": row["PhaseBucket"],
                "OmicsScore": omics_score,
                "TractabilityScore": tractability_score,
                "PathwayScore": row["PathwayScore"],
                "PhaseScore": row["PhaseScore"],
                "CompositeScore": composite_score,
                "RepurposingLead": therapies[0][0],
                "RepurposingNote": therapies[0][2],
                "PerDatasetLog2FC": "; ".join(
                    f"{accession}:{stats_row['log2FC']:.2f} (p={stats_row['pvalue']:.3g})"
                    for accession, stats_row in per_dataset.items()
                ),
            }
        )

    targets = pd.DataFrame(all_stats).sort_values("CompositeScore", ascending=False).reset_index(drop=True)
    targets["Rank"] = np.arange(1, len(targets) + 1)
    return targets


def build_drug_table(targets: pd.DataFrame) -> pd.DataFrame:
    records = []
    for gene, therapies in TARGET_DRUGS.items():
        target_row = targets[targets["GeneSymbol"] == gene]
        if target_row.empty:
            continue
        for drug_name, priority, rationale in therapies:
            records.append(
                {
                    "GeneSymbol": gene,
                    "PriorityRank": int(target_row.iloc[0]["Rank"]),
                    "Candidate": drug_name,
                    "EvidenceTier": priority,
                    "Rationale": rationale,
                    "Pathway": target_row.iloc[0]["Pathway"],
                }
            )
    drugs = pd.DataFrame(records)
    order = {"supportive": 0, "moderate": 1, "low": 2}
    if not drugs.empty:
        drugs["sort_order"] = drugs["EvidenceTier"].map(order).fillna(9)
        drugs = drugs.sort_values(["sort_order", "PriorityRank", "Candidate"]).drop(columns="sort_order")
    return drugs


def run_weight_sensitivity(targets: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    schemes = {
        "base": {"OmicsScore": 0.45, "TractabilityScore": 0.20, "PathwayScore": 0.20, "PhaseScore": 0.15},
        "equal_weight": {"OmicsScore": 0.25, "TractabilityScore": 0.25, "PathwayScore": 0.25, "PhaseScore": 0.25},
        "omics_heavy": {"OmicsScore": 0.60, "TractabilityScore": 0.15, "PathwayScore": 0.15, "PhaseScore": 0.10},
    }

    sensitivity = targets[["GeneSymbol", "OmicsScore", "TractabilityScore", "PathwayScore", "PhaseScore"]].copy()
    rank_columns = []
    for scheme_name, weights in schemes.items():
        score_column = f"{scheme_name}_score"
        rank_column = f"{scheme_name}_rank"
        sensitivity[score_column] = sum(sensitivity[col] * weight for col, weight in weights.items())
        sensitivity[rank_column] = sensitivity[score_column].rank(method="min", ascending=False)
        rank_columns.append(rank_column)

    summary_rows = []
    base_rank = sensitivity["base_rank"]
    for rank_column in rank_columns[1:]:
        rho, pvalue = stats.spearmanr(base_rank, sensitivity[rank_column])
        summary_rows.append(
            {
                "Comparison": f"base_vs_{rank_column.replace('_rank', '')}",
                "SpearmanRho": rho,
                "PValue": pvalue,
            }
        )
    return sensitivity, pd.DataFrame(summary_rows)


def save_figures(candidate_panel: pd.DataFrame, targets: pd.DataFrame, dataset_results: dict[str, pd.DataFrame], metadata_map: dict[str, pd.DataFrame]) -> None:
    sns.set_theme(style="whitegrid")

    counts = []
    for accession, meta in metadata_map.items():
        counts.append(
            {
                "Dataset": accession,
                "Severe": int((meta["severity"] == "severe").sum()),
                "Non-severe": int((meta["severity"] == "non_severe").sum()),
            }
        )
    count_df = pd.DataFrame(counts).melt(id_vars="Dataset", var_name="Group", value_name="Samples")
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(data=count_df, x="Dataset", y="Samples", hue="Group", ax=ax, palette=["#9b1d20", "#2f6690"])
    ax.set_ylabel("Sample count")
    ax.set_title("Figure 1. Discovery cohorts included in the severe-vs-nonsevere meta-signature")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "figure1_discovery_cohorts.png", dpi=300)
    plt.close(fig)

    top_targets = targets.head(20).copy()
    fig, ax = plt.subplots(figsize=(9, 7))
    sns.barplot(data=top_targets, y="GeneSymbol", x="CompositeScore", hue="Pathway", dodge=False, ax=ax)
    ax.set_xlabel("Composite priority score")
    ax.set_ylabel("Gene")
    ax.set_title("Figure 2. Top ranked host targets from the consensus flaviviral severity signature")
    ax.legend(loc="lower right", fontsize=8, title="Pathway")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "figure2_target_ranking.png", dpi=300)
    plt.close(fig)

    recurrence = pd.DataFrame(index=candidate_panel["GeneSymbol"])
    for accession, deg in dataset_results.items():
        recurrence[accession] = (
            candidate_panel["GeneSymbol"]
            .map(deg.set_index("GeneSymbol")["log2FC"])
            .astype(float)
            .values
        )
    fig, ax = plt.subplots(figsize=(7, 10))
    sns.heatmap(recurrence, cmap="coolwarm", center=0, ax=ax, cbar_kws={"label": "log2 fold-change"})
    ax.set_title("Figure 3. Directional consistency of the 50-gene revision signature")
    ax.set_xlabel("Dataset")
    ax.set_ylabel("Gene")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "figure3_signature_heatmap.png", dpi=300)
    plt.close(fig)

    pathway_summary = (
        targets.groupby("Pathway")
        .agg(Targets=("GeneSymbol", "count"), MeanScore=("CompositeScore", "mean"), SD=("CompositeScore", "std"))
        .reset_index()
        .sort_values("MeanScore", ascending=False)
    )
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(data=pathway_summary, x="MeanScore", y="Pathway", ax=ax, color="#577590")
    ax.set_xlabel("Mean composite score")
    ax.set_ylabel("Pathway module")
    ax.set_title("Figure 4. Pathway-level prioritization across the revision target panel")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "figure4_pathway_scores.png", dpi=300)
    plt.close(fig)

    drug_table = build_drug_table(targets).head(10)
    if drug_table.empty:
        return
    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.axis("off")
    table = ax.table(
        cellText=drug_table[["Candidate", "GeneSymbol", "EvidenceTier", "Pathway"]].values,
        colLabels=["Candidate", "Target", "Evidence tier", "Pathway"],
        cellLoc="left",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.4)
    ax.set_title("Figure 5. Hypothesis-generating repurposing candidates linked to top-ranked targets", pad=16)
    fig.tight_layout()
    fig.savefig(FIG_DIR / "figure5_candidate_table.png", dpi=300)
    plt.close(fig)


def main() -> None:
    metadata_map: dict[str, pd.DataFrame] = {}
    dataset_results: dict[str, pd.DataFrame] = {}
    cohort_rows = []

    for config in DATASETS:
        meta, expression = parse_series_matrix(config)
        metadata = config.group_parser(meta)
        metadata = metadata[metadata["severity"].isin({"severe", "non_severe"})].copy()
        if "phase" in metadata.columns:
            metadata = metadata[~metadata["phase"].str.contains("Conval", case=False, na=False)].copy()
        metadata_map[config.accession] = metadata

        annotation = parse_annotation(config.platform)
        sample_columns = metadata["sample_id"].tolist()
        gene_matrix = collapse_to_genes(expression[["ID_REF", *sample_columns]], annotation, sample_columns)
        deg = differential_expression(gene_matrix, metadata)
        dataset_results[config.accession] = deg

        deg.to_csv(TABLE_DIR / f"{config.accession}_deg_results.csv", index=False)
        cohort_rows.append(
            {
                "Dataset": config.accession,
                "Description": config.citation_label,
                "Platform": config.platform,
                "SevereSamples": int((metadata["severity"] == "severe").sum()),
                "NonSevereSamples": int((metadata["severity"] == "non_severe").sum()),
                "GenesTested": int(deg.shape[0]),
                "P_0.05_DEGs": int(((deg["pvalue"] <= 0.05) & (deg["log2FC"].abs() >= 0.30)).sum()),
            }
        )

    candidate_panel = load_candidate_panel()
    targets = build_target_table(candidate_panel, dataset_results)
    drugs = build_drug_table(targets)
    sensitivity_table, sensitivity_summary = run_weight_sensitivity(targets)

    pd.DataFrame(cohort_rows).to_csv(TABLE_DIR / "cohort_summary.csv", index=False)
    candidate_panel.to_csv(TABLE_DIR / "kfd_revision_signature.csv", index=False)
    targets.to_csv(TABLE_DIR / "kfd_revision_targets.csv", index=False)
    drugs.to_csv(TABLE_DIR / "kfd_revision_drug_candidates.csv", index=False)
    sensitivity_table.to_csv(TABLE_DIR / "kfd_revision_weight_sensitivity.csv", index=False)
    sensitivity_summary.to_csv(TABLE_DIR / "kfd_revision_weight_sensitivity_summary.csv", index=False)

    pathway_summary = (
        targets.groupby("Pathway")
        .agg(Targets=("GeneSymbol", "count"), MeanScore=("CompositeScore", "mean"), SD=("CompositeScore", "std"))
        .reset_index()
        .sort_values("MeanScore", ascending=False)
    )
    pathway_summary.to_csv(TABLE_DIR / "kfd_revision_pathway_summary.csv", index=False)

    save_figures(candidate_panel, targets, dataset_results, metadata_map)

    print("Revision analysis completed.")
    print(f"Panel genes: {len(candidate_panel)}")
    print(targets.head(15)[["Rank", "GeneSymbol", "Pathway", "CompositeScore"]].to_string(index=False))


if __name__ == "__main__":
    main()
