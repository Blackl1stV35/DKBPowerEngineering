# DKB Power Engineering — Company Profile Dashboard

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.40%2B-red.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A professional Streamlit-based application for generating comprehensive, multi-language Company Profile documents (.docx) with project portfolios, case studies, and custom content sections.

## 🎯 Overview

DKB Power Engineering Company Profile Dashboard is a specialized tool designed to help companies create polished, professional Word documents that showcase their business information, project history, and portfolio. The application supports **dual-language content (Thai & English)** with proper font rendering for complex scripts.

### Key Features

✨ **Dynamic Company Profile Generation**
- Beautiful, multi-page Word documents (.docx) with professional styling
- Thai and English bilingual support with proper Unicode font rendering
- Automated timestamp versioning for output files

📋 **Project Management**
- Add and manage unlimited projects with bilingual descriptions
- Upload and organize project photos with automatic resizing
- Comprehensive project reference lists and case study pages
- Project details: customer info, location, budget, period, deliverables

📂 **Custom Content Sections**
- Pre-built templates: Preventive Maintenance, Construction, Factory References, Solar/Renewable Energy, Residential Projects
- Flexible section creation with custom titles, descriptions, and photo galleries
- Dynamic section ordering in final document

🏢 **Company Information Management**
- Edit company details: name, registration, capital, address, contact
- Upload and display company logo
- Custom company history and mission statements
- Services and scope definitions

📸 **Professional Media Handling**
- Image optimization: automatic EXIF rotation, format conversion, compression
- Photo gallery layouts with responsive grid display
- Support for JPG, PNG, and WebP formats

🎨 **Advanced Document Styling**
- Custom color schemes with navy blue and accent colors
- Professional header/footer bars with contact information
- Multi-column layouts for balanced information display
- Responsive table designs with alternating row colors

## 🚀 Quick Start

### Prerequisites

- Python 3.8+
- pip or similar package manager

### Installation

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd DKB
   ```

2. **Create a Python virtual environment:**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

### Running the Application

Launch the Streamlit app:

```bash
streamlit run dkb_profile_dashboard.py
```

The application will open in your default browser at `http://localhost:8501`

## 📖 Usage Guide

### Tab 1: Add New Project ➕
- Fill in project details in Thai and English
- Upload project photos (multiple files supported)
- Submit to save project and automatically store images

### Tab 2: View All Projects 📋
- Browse all saved projects with collapsible details
- View project photos in thumbnail grid
- Add additional photos to existing projects
- Delete projects if needed

### Tab 3: Custom Sections 📂
- Create special portfolio sections beyond individual projects
- Choose from preset templates or create custom sections
- Add descriptions and photo galleries
- Examples: Preventive Maintenance, Under-Construction Projects, Factory References

### Tab 4: Company Information 🏢
- Update company registration details
- Edit bilingual company history and mission
- Upload company logo
- Manage contact information

### Generate Document 📄
- Click "สร้าง Company Profile" to generate a Word document
- Download button appears after successful generation
- Auto-timestamped files saved to `/output/` folder

## 📁 Project Structure

```
DKB/
├── dkb_profile_dashboard.py    # Main Streamlit application
├── requirements.txt             # Python dependencies
├── streamlit_config.toml        # Streamlit configuration
├── company.json                 # Company information (persisted)
├── projects.json                # Project data (persisted)
├── custom_sections.json         # Custom section data (persisted)
├── photos/                      # Project & section photos directory
├── output/                      # Generated .docx files
├── assets/                      # Logo and asset files
└── README.md                    # This file
```

## 🛠️ Tech Stack

| Component | Version | Purpose |
|-----------|---------|---------|
| **Streamlit** | ≥1.40.0 | Web UI framework |
| **python-docx** | ≥1.1.2 | Word document generation |
| **Pillow** | ≥10.3.0 | Image processing |
| **Python** | ≥3.8 | Runtime |

## 🎨 Document Features

### Generated .docx Structure
1. **Cover Page** — Company name, profile title, services summary, year watermark
2. **Company Information** — Registration details, company history, scope of services
3. **Project Reference List** — Table with all projects, customers, values, descriptions
4. **Case Studies** — One page per project with detailed info and photo gallery
5. **Custom Sections** — User-defined sections with descriptions and photos
6. **Footer Bar** — Contact information and page numbers on every page

### Styling Details
- **Font:** TH Sarabun New (Thai), Calibri (Latin)
- **Color Scheme:** Navy blue (#0A3C7A), accent blue (#2979C8), light backgrounds
- **Unicode Support:** All Thai characters properly rendered with multi-script font attributes

## 🌐 Language Support

- **Thai (ไทย):** Full Unicode support with proper font handling
- **English:** Professional business English
- **Bilingual Layout:** Parallel Thai/English content on every page

## 📊 Data Persistence

Application data is stored in JSON files:
- `company.json` — Company details
- `projects.json` — Project portfolio
- `custom_sections.json` — Custom content sections

JSON structure is human-readable and can be manually edited if needed.

## 🔐 Data Management

- **Local Storage:** All data stored locally in JSON files and `/photos/` directory
- **No Cloud Upload:** 100% data privacy maintained
- **Backup:** Export generated .docx files to external storage as needed

## 🚧 Advanced Customization

### Modifying Colors
Edit the `C` dictionary in the main script (around line 142):
```python
C = {
    "navy"     : "0A3C7A",    # Primary dark blue
    "blue"     : "2979C8",    # Accent blue
    "ltblue"   : "EDF4FC",    # Light background
    # ... more colors
}
```

### Streamlit Configuration
Customize behavior in `streamlit_config.toml`:
```toml
[server]
maxUploadSize = 200  # Max file upload in MB

[theme]
primaryColor = "#0A3C7A"
```

### Photo Optimization
Adjust image quality in `_resize_image()` function:
- `max_px=900` — Maximum image resolution
- `quality=85` — JPEG compression quality

## 📝 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## 📞 Support

For issues or questions about this application, please open an issue in the Git repository.

## ✨ Version History

### v3.0 (Current)
- Thai font rendering: Fixed w:rFonts with ascii+hAnsi+cs+eastAsia attributes
- Removed all placeholder text from generated documents
- Streamlit 1.40+ compatibility (use_container_width → width)
- Dynamic custom sections support
- Spacious, elegant document layout

### v2.0
- Initial multi-language support
- Basic project management

### v1.0
- Initial release

---

**DKB POWER ENGINEERING CO., LTD.**  
บริษัท ดีเคบี เพาเวอร์ เอนจิเนียริ่ง จำกัด  
📞 082-234-4680 | ✉️ kasama.D.K.B@gmail.com
