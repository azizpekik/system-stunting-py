# SiTrack Stunting

**Sistem Validitas Data Anak Berbasis WHO** - Aplikasi web untuk mengkonversi file Excel data pertumbuhan balita menjadi format JSON dengan assessment status gizi berdasarkan standar WHO.

## 🌟 Fitur Utama

- **Multi-Format Excel Support**: Mendukung berbagai format Excel (PRD Format, Header Format, Direct Data)
- **WHO Assessment**: Implementasi lengkap assessment status gizi berdasarkan standar WHO untuk anak 0-24 bulan
- **Height Validation**: Validasi rasionalitas tinggi badan antar pengukuran
- **Template Validation**: Validasi kepatuhan template dengan download template reference
- **Complete Data Display**: Menampilkan assessment lengkap untuk SEMUA data (lengkap dan tidak lengkap)
- **Responsive UI**: Interface modern dan user-friendly dengan drag-and-drop support

## 🚀 Quick Start

### Prerequisites

- Python 3.7+
- pip package manager

### Installation

1. **Clone Repository**
```bash
git clone https://github.com/[USERNAME]/sitrack-stunting.git
cd sitrack-stunting
```

2. **Create Virtual Environment**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. **Install Dependencies**
```bash
pip install flask pandas openpyxl
```

4. **Run Application**
```bash
python app.py
```

5. **Open Browser**
Navigate to `http://localhost:5002`

## 📊 Cara Penggunaan

### 1. Upload File Excel
- Klik area upload atau drag-and-drop file Excel
- Format yang didukung: `.xlsx` atau `.xls`
- Maksimal ukuran file: 16MB

### 2. Template Reference
Download template reference "Data Test.xlsx" untuk format yang benar:
- Header: TGL UKUR, UMUR, BERAT, TINGGI, CARA UKUR
- Data identitas: NO, NIK, NAMA ANAK, TANGGAL LAHIR, JENIS KELAMIN

### 3. Hasil Assessment
Aplikasi akan menampilkan:
- **Status Gizi**: NORMAL, KURANG, LEBIH (berat badan)
- **Status Tinggi**: NORMAL, PENDEK, TINGGI (tinggi badan)
- **WHO Reference**: Rentang BB dan TB ideal sesuai umur & jenis kelamin
- **Validasi Tinggi**: Status rasionalitas pertumbuhan tinggi badan
- **Complete Display**: Informasi lengkap untuk data lengkap dan tidak lengkap

## 🏥 WHO Assessment Features

### Nutritional Status Assessment
- **Weight Assessment**: KURANG/NORMAL/LEBIH berdasarkan standar WHO
- **Height Assessment**: PENDEK/NORMAL/TINGGI berdasarkan standar WHO
- **Age Range**: 0-24 bulan dengan rentang nilai per bulan

### Height Rationality Validation
- **NORMAL**: Pertumbuhan tinggi badan normal
- **DANGER**: Penurunan tinggi badan tidak rasional
- **NO_BASELINE**: Tidak ada data sebelumnya
- **AMBIGU_METHODOLOGY**: Perbedaan metode ukur

### Data Processing
- **Complete Data**: Pengukuran dengan berat/tinggi lengkap
- **Incomplete Data**: Pengukuran tanpa berat/tinggi (tetap ditampilkan dengan assessment lengkap)
- **Data Filtering**: Data tidak lengkap difilter dari count tetapi tetap ditampilkan untuk transparansi

## 📁 Project Structure

```
sitrack-stunting/
├── app.py                    # Flask application main file
├── excel_to_json_anak.py     # Core conversion & assessment logic
├── templates/
│   └── index.html           # Web interface template
├── data master/
│   └── Tabel_Pertumbuhan_Anak_0-2_Tahun.csv  # WHO reference data
├── data test/
│   └── Data Test.xlsx       # Template reference file
├── uploads/                 # Temporary upload folder
├── venv/                    # Virtual environment
├── requirements.txt         # Python dependencies
├── .gitignore              # Git ignore file
└── README.md               # This file
```

## 🛠️ API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/` | Main application page |
| POST | `/upload` | Upload and process Excel file |
| GET | `/files` | List uploaded files |
| GET | `/download-template` | Download template reference |

## 📋 Format Excel yang Didukung

### 1. PRD Format
- Header di baris 1
- Sub-header di baris 2
- Data di baris 3+

### 2. Header Format (Recommended)
- Header langsung di baris 1: TGL UKUR, UMUR, BERAT, TINGGI, CARA UKUR
- Data di baris 2+

### 3. Direct Data Format
- Data langsung tanpa header

## 🔧 Configuration

### Port Configuration
Default port: `5002`
Ubah di `app.py`:
```python
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5002)
```

### WHO Reference Data
File: `data master/Tabel_Pertumbuhan_Anak_0-2_Tahun.csv`
- Rentang umur: 0-59 bulan
- Jenis kelamin: Laki-laki (L) & Perempuan (P)
- Standar: WHO growth standards

## 🐛 Troubleshooting

### Common Issues

1. **Template Validation Failed**
   - Download template reference "Data Test.xlsx"
   - Pastikan header sesuai format
   - Periksa kelengkapan data identitas

2. **Port Already in Use**
   - Kill existing process: `pkill -f "python app.py"`
   - Change port in app.py

3. **WHO Assessment Not Working**
   - Pastikan file WHO reference ada di `data master/`
   - Check format data umur dan jenis kelamin

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

## 👥 Contributors

- Developer: [Your Name]
- WHO Standards: World Health Organization

## 📞 Support

For issues and questions:
1. Check existing [Issues](https://github.com/[USERNAME]/sitrack-stunting/issues)
2. Create new issue with detailed description
3. Include sample data if applicable

---

**Note**: This application is designed for defensive security analysis and child growth monitoring purposes only.