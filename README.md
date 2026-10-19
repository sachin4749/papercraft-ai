# 🎓 PaperCraft AI

**PaperCraft AI** is an AI-inspired, enterprise-grade examination paper generator designed specifically for professors, teachers, and academic administrators. Built with Python and Streamlit, it features a glassmorphism UI/UX design, rich metadata management, dynamic multi-section building, SQLite question banking, multi-set randomization, and PDF/Word generation engines.

---

## ✨ Features

- **Glassmorphism UI/UX**: Custom CSS styling with deep mesh gradients, translucent panels, glowing accent controls, and crisp typography.
- **Institutional Branding & Metadata**: Upload school/university logos and set headers, subject codes, exam durations, maximum marks, and instructions.
- **Sample Paper Parser**: Native parsing engine for `.pdf` and `.docx` files using regex heuristic extraction.
- **Dynamic Multi-Section Builder**: Create Sections A through F+, specify per-section marks, toggle MCQ modes, and reorder questions interactively.
- **Question Bank Repository**: Persistent SQLite database storage supporting tags for Difficulty and Bloom's Taxonomy levels.
- **Multi-Set Randomizer**: Generate Set A, Set B, and Set C automatically with scrambled question and option ordering.
- **Automated Answer Key Engine**: Option to generate paired answer keys and solution rubrics.
- **High-Precision Export Engine**:
  - **PDF Export**: ReportLab Platypus rendering engine with double borders, header logos, right-aligned mark allocations, and page numbering.
  - **Word Export**: `python-docx` export maintaining headers, metadata grids, section headers, and option alignments.

---

## 🛠️ Local Installation & Setup

1. **Clone or download the project folder**.
2. **Create and activate a virtual environment**:
   ```bash
   python -m venv venv
   # Windows: venv\Scripts\activate
   # macOS/Linux: source venv/bin/activate