# AI for Drug Discovery

**University Course — 34 Teaching Hours | 7 Weeks + 1 Project Week**

## Course Overview

This repository contains all teaching materials for the "AI for Drug Discovery" course, including PowerPoint presentations with detailed presenter notes and hands-on Jupyter notebooks for Google Colab.

## Course Structure

| Week | Topic | Teaching Hours |
|------|-------|---------------|
| 1 | Introduction & Drug Discovery Pipeline | 5 TH |
| 2 | Molecular Representation & Baseline ML | 5 TH |
| 3 | Model Evaluation & Interpretability | 5 TH |
| 4 | Deep Learning & Graph Neural Networks | 5 TH |
| 5 | Generative AI for Molecule Design | 5 TH |
| 6 | Protein Targets & Binding Prediction | 4.5 TH |
| 7 | Ethics, Regulation & Project Workshop | 4.5 TH |
| 8 | **Project Exam** (Presentation + Code) | — |

## Repository Structure

```
generate_slides.py                              # Script to regenerate all .pptx files
requirements_slides.txt                         # Python dependencies for slide generation

week1_introduction/
  ├── slides.pptx                               # Lecture slides (20 slides)
  └── Week1_Practical.ipynb                     # Practical: Explore molecules with RDKit

week2_molecular_representation/
  ├── slides.pptx                               # Lecture slides (20 slides)
  └── Week2_Practical.ipynb                     # Practical: Predict solubility (RF/XGBoost)

week3_evaluation/
  ├── slides.pptx                               # Lecture slides (16 slides)
  └── Week3_Practical.ipynb                     # Practical: Scaffold splits, SHAP

week4_deep_learning_gnn/
  ├── slides.pptx                               # Lecture slides (16 slides)
  └── Week4_Practical.ipynb                     # Practical: GNN with DeepChem
```

## How to Use

### Presentations
All `.pptx` files include **detailed presenter/speaker notes** on every slide with talking points, timing suggestions, and engagement prompts.

To regenerate the slides:
```bash
pip install -r requirements_slides.txt
python generate_slides.py
```

### Notebooks
All notebooks are designed for **Google Colab** — just upload and run. They install their own dependencies.

- **Week 1**: Load molecules, visualize with RDKit, calculate properties, Lipinski's Rule of Five
- **Week 2**: Morgan fingerprints, QSAR with Random Forest & XGBoost on Delaney solubility dataset
- **Week 3**: Scaffold splitting, SHAP interpretability, applicability domain analysis
- **Week 4**: Graph Neural Networks with DeepChem on BACE dataset

## Assessment (Week 8)

**Group Project (3-4 students)** — Choose one:
1. QSAR prediction model
2. GNN-based activity prediction
3. Generative molecule design
4. Binding affinity modeling

**Grading**: Implementation 40% | Scientific reasoning 25% | Evaluation 20% | Presentation 15%

## Key References

- DiMasi, J.A. et al. (2016). Innovation in the pharmaceutical industry. *J. Health Economics* 47:20-33
- Rogers, D. & Hahn, M. (2010). Extended-Connectivity Fingerprints. *JCIM* 50:742-754
- Gilmer, J. et al. (2017). Neural Message Passing for Quantum Chemistry. *ICML*
- Lundberg, S.M. & Lee, S. (2017). A Unified Approach to Interpreting Model Predictions. *NeurIPS*
- Wu, Z. et al. (2018). MoleculeNet: A Benchmark for Molecular Machine Learning. *Chemical Science* 9:513-530
- Yang, K. et al. (2019). Analyzing Learned Molecular Representations. *JCIM* 59:3370-3388
- Stokes, J.M. et al. (2020). A Deep Learning Approach to Antibiotic Discovery. *Cell* 180:688-702
