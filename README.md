# AI for Drug Discovery

**Compact Course — 26 Teaching Units | 4 Teaching Days + Exam**
**Instructor: Dr. Étienne Serbe-Kamp** — Neuroscientist, Drosophila visual circuits & physiology, Backyard Brains

## Course Overview

This repository contains all teaching materials for the "AI for Drug Discovery" course, including PowerPoint presentations with detailed presenter notes and hands-on Jupyter notebooks for Google Colab.

The course uniquely bridges **neuroscience** and **drug discovery**, leveraging the instructor's research background in Drosophila visual motion detection, cellular physiology (calcium imaging, electrophysiology, optogenetics), and open-source neuroscience tools (Backyard Brains SpikerBox/SpikerBot). A key theme throughout the course is **physiology as the final readout of drug action** — the same physiological techniques Dr. Serbe-Kamp uses in his research are the gold standard for measuring drug effects on neurons and ion channels.

**Important note on Dr. Serbe-Kamp's 2016 Neuron paper**: This paper ("Comprehensive Characterization of the Major Presynaptic Elements to the Drosophila OFF Motion Detector", Serbe et al., 2016) is a **physiology paper** — it used calcium imaging, electrophysiology, and optogenetics to characterize the functional responses of neurons. Dr. Serbe-Kamp's **current** connectomics work is with the MESH repository (separate project).

## Course Schedule

| Day | Date | Units | Hours | Topic |
|-----|------|-------|-------|-------|
| 1 | 15 April 2026 | 6 | 4.5 h | Introduction, Drug Discovery Pipeline, AI Overview, **Project Proposals** |
| 2 | 22 April 2026 | 8 | 6 h | Molecular Representation, QSAR, RF/XGBoost, Evaluation, SHAP |
| 3 | 5 May 2026 | 8 | 6 h | Deep Learning, GNNs, Generative AI, AlphaFold, Protein Targets |
| 4 | 19 May 2026 | 4 | 3 h | Ethics, Physiology as Drug Readout, Project Finalization |
| **Exam** | **27 May 2026** | — | — | **Group Presentations / Posters** |

*1 unit = 45 minutes. Every day includes breaks and hands-on practical sessions.*

## Repository Structure

```
generate_slides.py                              # Script to regenerate all .pptx files
requirements_slides.txt                         # Python dependencies for slide generation

day1_introduction/
  ├── slides.pptx                               # Day 1 lecture slides (23 slides)
  └── Day1_Practical.ipynb                      # Practical: RDKit + Hodgkin-Huxley simulation

day2_molecular_ml/
  ├── slides.pptx                               # Day 2 lecture slides (28 slides)
  ├── Day2_Practical_QSAR.ipynb                 # Practical 1: QSAR with RF & XGBoost
  └── Day2_Practical_Evaluation.ipynb           # Practical 2: Scaffold splits, SHAP

day3_deep_learning/
  ├── slides.pptx                               # Day 3 lecture slides (24 slides)
  └── Day3_Practical_GNN.ipynb                  # Practical: GNN with DeepChem

day4_ethics_physiology/
  └── slides.pptx                               # Day 4 lecture slides (17 slides)

publications/                                   # Upload PDFs for students
  ├── README.md                                 # Full publication list organized by topic
  ├── instructor/                               # Dr. Serbe-Kamp's publications
  ├── drug_discovery_ai/                        # Core AI drug discovery papers
  ├── molecular_ml/                             # Fingerprints, QSAR, evaluation
  ├── gnns_deep_learning/                       # GNN and deep learning papers
  ├── connectomics_neuroscience/                # Connectomics & brain mapping
  ├── neuropharmacology/                        # Ion channels, drug targets
  ├── ethics_regulation/                        # AI ethics and regulatory papers
  └── project_references/                       # Additional project-specific papers
```

## Student Projects (Exam: 27 May 2026)

Groups of 3–4 students choose from 7 proposed projects or propose their own:

### Hardware / Physiology Projects (Poster format — collect your own data)
1. **Computational SpikerBox** — Simulate/record neural signals, classify spike patterns with ML, model drug effects on ion channels
2. **SpikerBot** — Program stimulation patterns, record physiological responses, model dose-response relationships
3. **BYB Human Signals** — Record EMG/ECG/EEG, apply ML classification, relate to pharmacological effects
4. **Eye-Tracking** — Visual processing + ML on gaze data, connect to CNS drug effects on attention/saccades
5. **VR Setup** — VR-based neuro experiments, behavioral data + ML, connect to drug effects on perception

### Computational Projects (Presentation format — reproduce/explain published work)
6. **Classical QSAR** — Build QSAR model for a neuroscience drug target (GABA-A, 5-HT, dopamine D2, hERG) from ChEMBL
7. **AlphaFold Structure** — Predict a neuroscience target structure, analyze binding sites, reproduce key figures

**Grading**: Implementation 40% | Scientific reasoning 25% | Evaluation 20% | Presentation 15%

## Key Theme: Physiology as Drug Readout

A central thread of this course is that **physiology is the final readout of drug action**:

1. **In silico** (AI): QSAR, GNN, docking predict binding affinity → Days 2-3
2. **In vitro** (cell): electrophysiology, calcium imaging show functional effects → Dr. Serbe-Kamp's 2016 Neuron paper
3. **Tissue/organ**: cardiac QT, neural circuits → BYB ECG, EMG recordings
4. **Organism**: behavioral outcomes → eye-tracking, VR experiments

The BYB SpikerBox/SpikerBot make step 2 accessible — "DIY drug readout." Students can measure real physiological effects of ion channel drugs.

## Neuroscience Integration

- **Day 1**: Instructor physiology background, neurotransmitter SMILES, ion channel pharmacology, GluCl/ivermectin, Hodgkin-Huxley model, SpikerBot
- **Day 2**: CNS drug properties (BBB penetration), GABA-A scaffold bias, neuroscience SHAP examples
- **Day 3**: Drosophila connectome as graph data, GNNs on brain connectivity, AlphaFold for neuro targets
- **Day 4**: Physiology as drug readout, BYB gear demonstrations, drug safety testing (hERG)

## How to Use

### Presentations
All `.pptx` files include **detailed presenter/speaker notes** with timing, background info, neuroscience context, and references.

To regenerate:
```bash
pip install -r requirements_slides.txt
python generate_slides.py
```

### Notebooks
All notebooks are designed for **Google Colab** — just upload and run.

### Publications
Upload PDFs into the `publications/` subfolders. See `publications/README.md` for the full organized list.

## Publications List

See [`publications/README.md`](publications/README.md) for the complete list of 34 publications organized by topic.
