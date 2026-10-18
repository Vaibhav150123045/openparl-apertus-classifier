# OpenParl-Apertus-Classifier

> Multilingual topic classification for Swiss parliamentary affairs powered by Swiss-AI's **Apertus-v1.5** and **OpenParlData.ch**.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.12](https://img.shields.io/badge/python-3.12-blue.svg)](https://www.python.org/downloads/release/python-3120/)
[![Model: Apertus-v1.5](https://img.shields.io/badge/Model-Apertus--v1.5--8B-red.svg)](https://huggingface.co/swiss-ai/Apertus-v1.5-8B)

---

## 📌 Overview

**OpenParlData.ch** aggregates more than 300,000 political affairs and 900,000 documents from over 90 federal, cantonal, and municipal parliaments across Switzerland. However, navigating this corpus currently relies on raw full-text search, which requires users to anticipate specialized terminology across four official languages.

**OpenParl-Apertus-Classifier** provides an open, reusable classification engine that automatically categorizes parliamentary texts into:
1. **Standardized Taxonomies:** Multi-level classification schemes (such as the *Année Politique Suisse* / APS framework).
2. **Ad-Hoc Custom Topics:** Dynamic topic assignment (e.g., `"Digitalisation"`, `"Covid-19"`, `"Renewable Energy"`) without needing model retraining.
3. **Multilingual Ingestion:** Native processing in German (`de`), French (`fr`), Italian (`it`), and Romansh (`rm`).

This project is developed as part of **Hack Apertus 2026** in collaboration with OpenData.ch and Digital Administration Switzerland (DVS).

---

## ✨ Key Features

- **Multilingual Support:** Seamlessly handles multi-cantonal parliamentary documents in German, French, and Italian.
- **Dual Classification Modes:**
  - *Standard*: Maps affairs to hierarchical APS policy categories.
  - *Ad-Hoc / Zero-Shot*: Detects user-specified keywords and thematic concepts on demand.
- **Apertus Backbone:** Utilizes `swiss-ai/Apertus-v1.5-8B`
- **Explainability & Transparency:** Outputs matched topics alongside confidence scores and key excerpt rationales to maintain political neutrality and support human-in-the-loop validation.
- **RESTful API & CLI:** Simple endpoints to fetch directly from the OpenParlData API or classify custom document payloads.

---

## 🚀 Quick Start

### Prerequisites
- Python 3.12+
- Access to [swiss-ai/Apertus-v1.5-8B](https://huggingface.co/swiss-ai/Apertus-v1.5-8B) on Hugging Face (gated repository).
- NVIDIA GPU with CUDA support (or Mac MPS / CPU for inference).

### Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/Vaibhav150123045/openparl-apertus-classifier.git](https://github.com/Vaibhav150123045/openparl-apertus-classifier.git)
   cd openparl-apertus-classifier
