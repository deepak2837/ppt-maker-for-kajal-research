# Quantum Chaotic Encryption — PhD Research Deck

Generation of a research presentation on quantum chaotic encryption, together with the chaotic-map figures used in the slides.

## Contents
- `generate_phd_ppt.py` — builds `Quantum_Chaotic_Encryption_PhD.pptx` from the research plan
- `generate_sre_ppt.py` — companion deck generator
- `Quantum Chaotic Encryption PhD Plan.docx` — source research plan
- `bifurcation_logistic.png`, `bifurcation_tent.png` — bifurcation diagrams for the logistic and tent maps
- `histogram_logistic.png`, `histogram_tent.png` — output distribution histograms

## Running
```bash
pip install python-pptx matplotlib numpy
python generate_phd_ppt.py
```
