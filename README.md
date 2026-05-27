# 🧬 AI for Drug Discovery

> **Compact Course — 26 Teaching Units | 4 Teaching Days + Exam Day**  
> **Instructor:** Dr. Étienne Serbe-Kamp  
> **Term:** April–May 2026

---

## 🎯 What Is This Course?

Learn how **Artificial Intelligence** is transforming drug discovery — from predicting molecular properties to understanding protein structures. This repository contains everything you need: **working Jupyter notebooks**, reading materials, and exam information.

**All notebooks run in Google Colab — no local installation required.**

---

## 🚀 Quick Start

```
1. Go to https://colab.research.google.com/
2. File → Upload notebook
3. Pick any .ipynb from this repo
4. Click "Run all" (▶▶) — dependencies install automatically
```

That's it. Every notebook is self-contained and includes step-by-step explanations.

---

## 📅 Course Schedule

| Day | Date | Topic | Materials |
|:---:|------|-------|:---------:|
| 1 | 15 Apr 2026 | Drug Discovery Pipeline, AI/ML Fundamentals | [📖 day1/](day1/) |
| 2 | 22 Apr 2026 | Molecular Representations, QSAR, Model Evaluation | [💻 day2/](day2/) |
| 3 | 5 May 2026 | Deep Learning, Graph Neural Networks | [💻 day3/](day3/) |
| 4 | 27 May 2026 | AlphaFold, Protein Structure & Presentation Preparation | [💻 day4/](day4/) |
| 🎓 | **2 Jun 2026** | **Exam: Group Presentations** | [📋 Exam](resources/exam_info.md) · [📁 Presentations](Exam%20Presentations/) |

---

## 📂 Repository Structure

```
AI-Drug-Discovery/
│
├── day1/                              # Day 1: Introduction (lecture only)
│   └── README.md                      #   Key concepts, reading list
│
├── day2/                              # Day 2: Molecular ML
│   ├── Day2_Practical_QSAR.ipynb      #   🔬 QSAR with Random Forest & XGBoost
│   └── Day2_Practical_Evaluation.ipynb#   📊 Scaffold splits, SHAP, model comparison
│
├── day3/                              # Day 3: Deep Learning
│   └── Day3_Practical_GNN.ipynb       #   🧠 Graph Neural Networks with DeepChem
│
├── day4/                              # Day 4: Protein Structure & Presentation Prep
│   └── Day4_Practical_AlphaFold.ipynb #   🧬 AlphaFold structure exploration
│
├── Exam Presentations/                # 🎓 Exam Day (2 Jun): Upload presentations here
│   ├── EMG fatigue/                   #   Project: EMG Fatigue
│   ├── visual EEG/                    #   Project: Visual EEG
│   └── EKG drugs/                     #   Project: EKG Drugs
│
├── resources/                         # Supplemental materials
│   ├── reading_list.md                #   📚 Key papers & online resources
│   └── exam_info.md                   #   📋 Exam format, criteria & project ideas
│
└── README.md                          # ← You are here
```

---

## 🔬 What You'll Learn

### Day 1 — Foundations
> *No notebook — lecture day*

- The drug discovery pipeline (target → lead → clinic)
- Machine learning fundamentals (supervised learning, validation, overfitting)
- How molecules are represented for computers (SMILES, fingerprints, graphs)

### Day 2 — Molecular Machine Learning
> `day2/Day2_Practical_QSAR.ipynb` + `day2/Day2_Practical_Evaluation.ipynb`

- Build a **QSAR model** to predict molecular solubility
- Generate **Morgan fingerprints** and molecular descriptors with RDKit
- Train **Random Forest** and **XGBoost** models
- Evaluate with **scaffold splitting** (realistic train/test separation)
- Interpret predictions with **SHAP** (which molecular features matter?)

### Day 3 — Graph Neural Networks
> `day3/Day3_Practical_GNN.ipynb`

- Represent molecules as **graphs** (atoms = nodes, bonds = edges)
- Build and train a **Graph Convolutional Network** using DeepChem
- Compare GNN performance against traditional fingerprint methods
- Understand why learning representations can beat hand-crafted features

### Day 4 — AlphaFold & Protein Structure + Presentation Preparation (27 May)
> `day4/Day4_Practical_AlphaFold.ipynb` · AlphaFold Practical + time to prepare group presentations

- Access the **AlphaFold Protein Structure Database** via its API
- Visualize protein structures in **interactive 3D** (colored by confidence)
- Compare multiple drug targets (AChE, COX-2, EGFR, GABA-A, Dopamine D2)
- Compare AlphaFold predictions with **experimental PDB structures**
- Assess **druggability** of protein targets

---

## 🛠️ Technical Requirements

| Requirement | Details |
|-------------|---------|
| **Platform** | Google Colab (free tier is sufficient) |
| **Browser** | Chrome or Firefox recommended |
| **Python knowledge** | Basic (variables, loops, functions) |
| **Local install** | Not needed — everything runs in the cloud |

All dependencies are installed automatically at the top of each notebook:
- `rdkit` — Cheminformatics toolkit
- `scikit-learn` — Machine learning
- `xgboost` — Gradient boosting
- `shap` — Model interpretability
- `deepchem` — Deep learning for chemistry
- `py3Dmol` — 3D protein visualization
- `biopython` — Protein structure handling

---

## 📚 Resources

- **[Reading List](resources/reading_list.md)** — Key papers for each topic, organized by day
- **[Exam Information](resources/exam_info.md)** — Format, evaluation criteria, and project suggestions
- **[Day 1 Overview](day1/README.md)** — Concepts and vocabulary from the introductory lecture

---

## ❓ Troubleshooting

| Problem | Solution |
|---------|----------|
| "ModuleNotFoundError" | Re-run the first cell (pip install) and restart runtime |
| Notebook won't load | Try downloading and re-uploading, or open via GitHub URL in Colab |
| Slow execution | Use GPU runtime: Runtime → Change runtime type → GPU |
| Plot not showing | Run the cell again — some visualizations need a second execution |

---

## 📄 License

Educational materials for the AI for Drug Discovery course. Feel free to use and adapt for teaching purposes with attribution.

---

<p align="center">
  <i>Built with ❤️ for teaching the next generation of computational drug designers</i>
</p>
