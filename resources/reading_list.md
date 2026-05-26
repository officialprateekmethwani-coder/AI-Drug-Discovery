# Recommended Reading & Resources

A curated list of key papers and resources for each course topic.

---

## General AI in Drug Discovery

| Paper | Year | Key Contribution |
|-------|------|-----------------|
| Vamathevan et al. "Applications of ML in drug discovery" (*Nat Rev Drug Discov*) | 2019 | Comprehensive review of ML applications across the drug pipeline |
| Schneider et al. "Rethinking drug design in the AI era" (*Nat Rev Drug Discov*) | 2020 | Perspective on how AI changes the drug design paradigm |
| Dara et al. "Machine Learning in Drug Discovery: A Review" (*Artif Intell Rev*) | 2022 | Recent overview of methods and challenges |

## Molecular Representations & QSAR (Day 2)

| Paper | Year | Key Contribution |
|-------|------|-----------------|
| Rogers & Hahn. "Extended-Connectivity Fingerprints" (*J Chem Inf Model*) | 2010 | Introduced Morgan/ECFP fingerprints |
| Cherkasov et al. "QSAR modeling: Where have you been?" (*J Med Chem*) | 2014 | Best practices for QSAR |
| Lundberg & Lee. "A Unified Approach to Interpreting Model Predictions" (NeurIPS) | 2017 | SHAP values for model interpretability |

## Graph Neural Networks (Day 3)

| Paper | Year | Key Contribution |
|-------|------|-----------------|
| Gilmer et al. "Neural Message Passing for Quantum Chemistry" (ICML) | 2017 | Message-passing framework for molecular graphs |
| Yang et al. "Analyzing Learned Molecular Representations" (*J Chem Inf Model*) | 2019 | Chemprop — directed MPNN for molecular property prediction |
| Wu et al. "MoleculeNet: A Benchmark for Molecular ML" (*Chem Sci*) | 2018 | Standard benchmarks for molecular property prediction |

## AlphaFold & Protein Structure (Day 4)

| Paper | Year | Key Contribution |
|-------|------|-----------------|
| Jumper et al. "Highly accurate protein structure prediction with AlphaFold" (*Nature*) | 2021 | AlphaFold2 — breakthrough in protein structure prediction |
| Varadi et al. "AlphaFold Protein Structure Database" (*Nucleic Acids Res*) | 2022 | Public database of 200M+ predicted structures |
| Abramson et al. "Accurate structure prediction of biomolecular interactions with AlphaFold 3" (*Nature*) | 2024 | AlphaFold3 — extends to complexes, nucleic acids, ligands |

## Online Resources

- **RDKit Documentation**: https://www.rdkit.org/docs/
- **DeepChem Tutorials**: https://deepchem.io/tutorials/
- **AlphaFold DB**: https://alphafold.ebi.ac.uk/
- **Protein Data Bank**: https://www.rcsb.org/
- **Google Colab**: https://colab.research.google.com/
- **ChEMBL Database**: https://www.ebi.ac.uk/chembl/

## Tools Used in This Course

| Tool | Purpose | Day |
|------|---------|-----|
| RDKit | Cheminformatics — fingerprints, descriptors, visualization | 2 |
| scikit-learn | Machine learning — Random Forest, cross-validation | 2 |
| XGBoost | Gradient boosting models | 2 |
| SHAP | Model interpretability | 2 |
| DeepChem | Deep learning for chemistry — GNNs, featurizers | 3 |
| PyTorch | Deep learning framework (via DeepChem) | 3 |
| py3Dmol | 3D molecular visualization | 4 |
| BioPython | Protein structure parsing | 4 |
