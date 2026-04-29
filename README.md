# OCR-and-Translate-using-Ollama
A high-performance Python pipeline designed to convert complex medical PDF documents into professionally translated Romanian Word documents. This script is optimized for Vast.ai GPU instances, utilizing Ollama to run state-of-the-art Vision and Language models locally.
🚀 Key Features
Dual-Model Architecture: Uses Qwen2-VL (8B) for high-accuracy medical OCR and TranslateGemma (12B) for nuanced medical translation into Romanian.

Intelligent Blank Page Detection: Employs image analysis (mean brightness and edge detection) to skip empty pages, saving compute time and GPU costs.

Vast.ai Optimization: Pre-configured for local Ollama API endpoints and optimized for high-VRAM cards like the NVIDIA RTX 5090.

Robust Progress System: Saves OCR and translation results to a JSON checkpoint after every page, allowing for seamless resumes after connection drops or timeouts.

Parallel Translation: Features a multi-threaded translation worker system to maximize GPU throughput during the text-to-text phase.

Medical Formatting: Outputs directly to .docx, preserving the structured flow of medical reports and lab results.

🛠️ Technical Configuration
OCR Engine: qwen2-vl:8b

Translation Engine: translategemma:12b

DPI Settings: 300 DPI for high-fidelity text extraction from scans.

Concurrency: Configurable workers for both OCR and Translation phases to balance stability and speed.

📋 Requirements
Python 3.10+

Ollama (running locally on port 21434)

System dependencies: poppler-utils (for PDF rendering)
