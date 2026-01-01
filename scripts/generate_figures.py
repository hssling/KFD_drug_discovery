"""
Generate Publication-Quality Figures for KFD HDT Pipeline
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent

plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.labelsize': 12,
    'figure.dpi': 300
})

# Phase colors for KFD
PHASE_COLORS = {
    'Hemorrhagic': '#8B0000',
    'Febrile': '#FF6347',
    'Neurological': '#4169E1',
    'Protective': '#228B22',
    'Both': '#9370DB'
}

def figure1_target_prioritization():
    print("Generating Figure 1: Target Prioritization...")
    
    df = pd.read_csv(BASE_DIR / 'outputs' / 'tables' / 'targets_ranked.csv')
    top20 = df.head(20).copy()
    
    fig, ax = plt.subplots(figsize=(12, 8))
    
    colors = [PHASE_COLORS.get(s, '#808080') for s in top20['Phase_Relevance']]
    bars = ax.barh(range(len(top20)), top20['Composite_Score'], color=colors, edgecolor='black')
    
    ax.set_yticks(range(len(top20)))
    ax.set_yticklabels(top20['Symbol'])
    ax.invert_yaxis()
    ax.set_xlabel('Composite Score')
    ax.set_title('Top 20 Host-Directed Therapy Targets for KFD', fontweight='bold')
    
    handles = [plt.Rectangle((0,0),1,1, facecolor=c, edgecolor='black') 
               for c in PHASE_COLORS.values()]
    ax.legend(handles, PHASE_COLORS.keys(), title='Disease Phase', 
              loc='lower right', framealpha=0.9)
    
    plt.tight_layout()
    plt.savefig(BASE_DIR / 'outputs' / 'figures' / 'figure1_target_prioritization.png', 
                dpi=300, bbox_inches='tight')
    plt.close()
    print("  Saved figure1_target_prioritization.png")

def figure2_compound_distribution():
    print("Generating Figure 2: Compound Distribution...")
    
    df = pd.read_csv(BASE_DIR / 'outputs' / 'tables' / 'compounds_ranked.csv')
    
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    # Panel A: By evidence type
    evidence_counts = df['Evidence'].value_counts()
    colors_a = plt.cm.Set3(np.linspace(0, 1, len(evidence_counts)))
    axes[0].pie(evidence_counts.head(6), labels=evidence_counts.head(6).index, autopct='%1.0f%%',
                colors=colors_a, startangle=90)
    axes[0].set_title('A. Compounds by Mechanism', fontweight='bold')
    
    # Panel B: By target category
    categories = {
        'Endothelial': ['ANGPT2', 'VWF', 'VEGFA', 'EDN1'],
        'Coagulation': ['F3', 'PLAT', 'F2', 'SERPINE1'],
        'Cytokine': ['TNF', 'IL6', 'IL1B'],
        'Platelet': ['ITGA2B', 'ITGB3'],
        'Antiviral': ['IFNA1', 'STAT1'],
        'Supportive': ['HMOX1', 'SOD2', 'EPO', 'AQP4']
    }
    
    cat_counts = []
    for cat, genes in categories.items():
        count = df[df['Related_Gene'].isin(genes)].shape[0]
        cat_counts.append((cat, max(count, 1)))
    
    cat_df = pd.DataFrame(cat_counts, columns=['Category', 'Count'])
    colors_b = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
    axes[1].bar(cat_df['Category'], cat_df['Count'], color=colors_b, edgecolor='black')
    axes[1].set_xlabel('Target Category')
    axes[1].set_ylabel('Number of Compounds')
    axes[1].set_title('B. Compounds by Target Category', fontweight='bold')
    axes[1].tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    plt.savefig(BASE_DIR / 'outputs' / 'figures' / 'figure2_compound_distribution.png',
                dpi=300, bbox_inches='tight')
    plt.close()
    print("  Saved figure2_compound_distribution.png")

def figure3_target_potency():
    print("Generating Figure 3: Target Potency...")
    
    df = pd.read_csv(BASE_DIR / 'outputs' / 'tables' / 'compounds_ranked.csv')
    df = df[df['pChEMBL'] > 0]  # Filter out supportive care
    
    potency = df.groupby('Related_Gene')['pChEMBL'].max().sort_values(ascending=False).head(15)
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    colors = plt.cm.RdYlGn(np.linspace(0.2, 0.8, len(potency)))[::-1]
    ax.bar(potency.index, potency.values, color=colors, edgecolor='black')
    
    ax.axhline(y=6.0, color='red', linestyle='--', linewidth=1.5, label='1 µM threshold')
    ax.axhline(y=8.0, color='green', linestyle='--', linewidth=1.5, label='10 nM threshold')
    
    ax.set_xlabel('Target Gene')
    ax.set_ylabel('Maximum pChEMBL Value')
    ax.set_title('Compound Potency by Target Gene', fontweight='bold')
    ax.legend(loc='upper right')
    ax.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    plt.savefig(BASE_DIR / 'outputs' / 'figures' / 'figure3_target_potency.png',
                dpi=300, bbox_inches='tight')
    plt.close()
    print("  Saved figure3_target_potency.png")

def figure4_pathway_heatmap():
    print("Generating Figure 4: Pathway Analysis...")
    
    df = pd.read_csv(BASE_DIR / 'outputs' / 'tables' / 'targets_ranked.csv')
    
    pathway_stats = df.groupby('Pathway').agg({
        'Composite_Score': ['count', 'mean', 'std']
    }).reset_index()
    pathway_stats.columns = ['Pathway', 'Count', 'Mean', 'SD']
    pathway_stats = pathway_stats.sort_values('Mean', ascending=False)
    
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    
    colors = plt.cm.Reds(np.linspace(0.3, 0.9, len(pathway_stats)))
    axes[0].barh(pathway_stats['Pathway'], pathway_stats['Count'], color=colors, edgecolor='black')
    axes[0].set_xlabel('Number of Targets')
    axes[0].set_title('A. Targets per Pathway', fontweight='bold')
    axes[0].invert_yaxis()
    
    axes[1].barh(pathway_stats['Pathway'], pathway_stats['Mean'], 
                 xerr=pathway_stats['SD'].fillna(0), color=colors, edgecolor='black',
                 capsize=3)
    axes[1].set_xlabel('Mean Composite Score')
    axes[1].set_title('B. Mean Score by Pathway', fontweight='bold')
    axes[1].invert_yaxis()
    
    plt.tight_layout()
    plt.savefig(BASE_DIR / 'outputs' / 'figures' / 'figure4_pathway_heatmap.png',
                dpi=300, bbox_inches='tight')
    plt.close()
    print("  Saved figure4_pathway_heatmap.png")

def figure5_kfd_timeline():
    print("Generating Figure 5: KFD Disease Timeline...")
    
    fig, ax = plt.subplots(figsize=(14, 8))
    
    # KFD phases
    phases = [
        (0, 7, 'Febrile Phase\n(Days 1-7)', '#FFD700', 0.6),
        (7, 12, 'Hemorrhagic\n(Days 7-12)', '#FF4500', 0.8),
        (12, 16, 'Neurological\n(if progresses)', '#8B0000', 0.9),
        (16, 28, 'Recovery\nPhase', '#87CEEB', 0.5)
    ]
    
    for start, end, label, color, alpha in phases:
        ax.axvspan(start, end, alpha=alpha, color=color, label=label)
        ax.text((start+end)/2, 0.95, label, ha='center', va='top', fontsize=10,
                transform=ax.get_xaxis_transform(), fontweight='bold')
    
    # Fever curve
    days = np.linspace(0, 28, 100)
    fever = np.where(days < 14, 40 - 0.1*(days-7)**2 + np.sin(days)*0.5, 37 + 0.1*(days-14))
    fever = np.clip(fever, 36.5, 41)
    ax.plot(days, fever, 'r-', linewidth=3, label='Temperature')
    
    # Platelet count (inverse pattern)
    platelets = np.where(days < 14, 250 - 20*days, 50 + 15*(days-14))
    platelets = np.clip(platelets, 20, 250)
    ax2 = ax.twinx()
    ax2.plot(days, platelets, 'b--', linewidth=2, label='Platelets (×10³/µL)')
    ax2.set_ylabel('Platelet Count (×10³/µL)', color='blue')
    ax2.tick_params(axis='y', labelcolor='blue')
    
    # HDT windows
    ax.annotate('Supportive\nCare', xy=(3, 39.5), fontsize=9, ha='center',
                bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.8))
    ax.annotate('Coagulation\nSupport', xy=(9, 39), fontsize=9, ha='center',
                bbox=dict(boxstyle='round', facecolor='orange', alpha=0.8))
    ax.annotate('Neuro-\nprotection', xy=(14, 38.5), fontsize=9, ha='center',
                bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.8))
    
    ax.set_xlim(0, 28)
    ax.set_ylim(36, 42)
    ax.set_xlabel('Days After Symptom Onset', fontsize=12)
    ax.set_ylabel('Temperature (°C)', fontsize=12, color='red')
    ax.tick_params(axis='y', labelcolor='red')
    ax.set_title('KFD Disease Progression and HDT Intervention Windows', 
                 fontsize=14, fontweight='bold')
    
    plt.tight_layout()
    plt.savefig(BASE_DIR / 'outputs' / 'figures' / 'figure5_kfd_timeline.png',
                dpi=300, bbox_inches='tight')
    plt.close()
    print("  Saved figure5_kfd_timeline.png")

def main():
    print("="*60)
    print("GENERATING KFD HDT FIGURES")
    print("="*60)
    
    figure1_target_prioritization()
    figure2_compound_distribution()
    figure3_target_potency()
    figure4_pathway_heatmap()
    figure5_kfd_timeline()
    
    print("\nAll 5 figures generated successfully!")

if __name__ == '__main__':
    main()
