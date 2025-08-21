# 📑 PDFToolkit

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20Mac-lightgrey)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)

**PDFToolkit** is a lightweight and user-friendly desktop app built with **Python (Tkinter GUI)** for managing and processing PDF files.  
It combines essential PDF operations with **AI-powered document parsing, OCR, and table extraction** to make working with PDFs fast and efficient.  

---

## ✨ Features

- 📄 Convert between **PDF ↔ Word, Excel, Images**
- 📑 **Merge, Split, Reorder** PDF pages
- 📉 **Compress PDFs & Images** (supports multiple formats)
- 🔍 **OCR for scanned PDFs** using [Tesseract](https://github.com/tesseract-ocr/tesseract)
- 📊 **AI Smart Extract** → convert structured PDF data into **Excel (.xlsx)**
- 🖼️ Supports most popular image types for merging & conversion
- 🎨 Modern UI with **Light/Dark mode**
- ⚡ Optimized for both small and **large multi-page PDFs**

---

## 🛠️ Tech Stack

- **Python 3.8+**
- **Tkinter** (GUI Framework)
- **PyPDF2** – PDF manipulation (merge/split)
- **pdfplumber** – text extraction
- **pdf2docx** – PDF to Word
- **pytesseract** – OCR for scanned PDFs
- **PIL / OpenCV** – image handling & compression
- **Hugging Face Models** – LayoutLM / TableFormer for AI-powered extraction

---

## 🚀 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/pdf-toolkit.git
   cd pdf-toolkit
2. Create a Virtual Env:
   python -m venv venv
   source venv/bin/activate   # On Linux/Mac
   venv\Scripts\activate      # On Windows
3. Install Dependencies: 
   pip install -r requirements.txt
4. Run the App:
      python main.py
## 📌 Roadmap

- [ ] Batch PDF processing  
- [ ] Drag-and-drop support  
- [ ] Cloud integration (Google Drive / Dropbox)  
- [ ] Advanced AI models for invoices & receipts  
- [ ] More export formats (CSV, JSON)  
- [ ] Batch OCR for multi-document scanning  
- [ ] Windows `.exe` packaged release  
- [ ] Cross-platform installer (Linux, MacOS)  
- [ ] UI improvements with modern themes  
- [ ] Plugin system for community extensions  

---

## 🤝 Contributing

Contributions are always welcome! 🎉  

1. Fork the repo  
2. Create your feature branch (`git checkout -b feature/amazing-feature`)  
3. Commit your changes (`git commit -m 'Add amazing feature'`)  
4. Push to the branch (`git push origin feature/amazing-feature`)  
5. Open a Pull Request  

### Contribution Ideas
- Add support for new file conversions  
- Improve UI/UX with better layouts  
- Enhance OCR accuracy with custom models  
- Write unit tests for stability  
- Add documentation & tutorials  

---

## 📜 License

This project is licensed under the **MIT License** – see the [LICENSE](LICENSE) file for details.  

---

## 🙌 Acknowledgements

- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)  
- [Hugging Face](https://huggingface.co/) for AI models  
- [PyPDF2](https://pypi.org/project/PyPDF2/)  
- [pdfplumber](https://github.com/jsvine/pdfplumber)  
- [pdf2docx](https://github.com/dothinking/pdf2docx)  
- [Pillow](https://pypi.org/project/Pillow/) & [OpenCV](https://opencv.org/)  
- All open-source contributors who make this project possible ❤️  

---
