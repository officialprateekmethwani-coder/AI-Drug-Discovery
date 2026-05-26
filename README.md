# AI for Drug Discovery

**Compact Course — 26 Teaching Units | 4 Teaching Days + Exam**
**Instructor: Dr. Étienne Serbe-Kamp**

## Course Overview

This repository contains hands-on Jupyter notebooks for the "AI for Drug Discovery" course. All notebooks are designed for **Google Colab** — just upload and run.

## Course Schedule

| Day | Date | Topic | Notebook |
|-----|------|-------|----------|
| 1 | 15 April 2026 | Introduction, Drug Discovery Pipeline, AI Overview | — |
| 2 | 22 April 2026 | Molecular Representation, QSAR, Model Evaluation | `day2/` |
| 3 | 5 May 2026 | Deep Learning, GNNs, Molecular Graphs | `day3/` |
| 4 | 19 May 2026 | AlphaFold, Protein Structure, Structure-Based Drug Design | `day4/` |
| **Exam** | **27 May 2026** | **Group Presentations / Posters** | — |

## Repository Structure

```
day2/
  ├── Day2_Practical_QSAR.ipynb           # QSAR with Random Forest & XGBoost
  └── Day2_Practical_Evaluation.ipynb     # Scaffold splits, SHAP interpretability

day3/
  └── Day3_Practical_GNN.ipynb            # Graph Neural Networks with DeepChem

day4/
  └── Day4_Practical_AlphaFold.ipynb      # AlphaFold protein structure exploration
```

## Practical Descriptions

### Day 2: Molecular ML (QSAR)
- **QSAR Notebook**: Build a complete Quantitative Structure-Activity Relationship pipeline — load the Delaney solubility dataset, generate Morgan fingerprints, calculate molecular descriptors, train Random Forest and XGBoost models, and evaluate performance.
- **Evaluation Notebook**: Scaffold-based train/test splitting, SHAP feature importance, applicability domain analysis, and model comparison.

### Day 3: Deep Learning (GNNs)
- **GNN Notebook**: Represent molecules as graphs, build and train Graph Neural Networks using DeepChem, and compare GNN performance to traditional fingerprint-based models.

### Day 4: AlphaFold & Protein Structure
- **AlphaFold Notebook**: Access the AlphaFold Protein Structure Database, visualize predicted protein structures in 3D, analyze pLDDT confidence scores, compare with experimental PDB structures, and assess druggability of pharmacological targets.

## How to Use

1. Open [Google Colab](https://colab.research.google.com/)
2. Upload any notebook from this repository
3. Run all cells — each notebook installs its own dependencies

All notebooks include detailed comments explaining every step, making them suitable for students with basic Python knowledge.
