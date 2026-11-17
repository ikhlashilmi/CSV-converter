# LSP Attendance Converter (CSV & XLSX to PDF / Excel)

A clean, modern, client-side web application built with HTML5, Tailwind CSS, and Vanilla JavaScript that converts raw attendance logs (CSV and Excel `.xlsx`) into an interactive dashboard, filterable data table, and cleanly formatted export reports (Excel `.xlsx` and landscape PDF).

## ✨ Features

- **Multi-Format Ingestion**:
  - Drag-and-drop or file picker for CSV (`.csv`) and Excel (`.xlsx`, `.xls`).
  - One-click buttons to load local sample files (`Data_Presensi_LSP_Pengadaan.csv` and `Laporan_Presensi_LSP_Lengkap.xlsx`).
  - Robust parser supporting standard CSVs, quoted multi-line/comma-separated strings, and packed single-column rows.
  - Extracts embedded images directly from `.xlsx` drawing packages via `JSZip`.

- **Interactive Dashboard & Filtering**:
  - Live KPI Summary Cards: Total Records, Di Kantor (WFO), Luar Kantor (Field/Out of office), and Unique Employees.
  - Granular Filters: Employee Name, Assignment Status, Date Range, and Keyword Search on Activity/Logbook.
  - Column Visibility Controls: Toggle any column on/off, with quick presets to hide technical metadata (Latitude, Longitude, Photo Links).

- **Responsive Data Table**:
  - Sticky headers with column sorting (ascending/descending).
  - Status badges (Green for "Di Kantor", Amber for "Luar Kantor").
  - Clean text wrapping for long activity descriptions and detected addresses.
  - Photo preview modal for embedded Excel photos and external link buttons for photo URLs.

- **Export Reports**:
  - **Styled Excel (.xlsx)**: Exports currently filtered rows and visible columns using SheetJS with auto-calculated column widths.
  - **Landscape PDF (.pdf)**: Generates an A4 landscape PDF report with `jsPDF` and `jspdf-autotable`, including header titles, active filter summaries, alternating row colors, cell image rendering, and footer pagination (`Halaman X dari Y`).

- **100% Client-Side**:
  - No server dependencies or backend database required. Everything is processed entirely inside your browser using CDN-delivered libraries.

## 🚀 Getting Started

Simply open `index.html` in your web browser. 

To use the **Load Project Sample** buttons without browser local `file://` CORS security restrictions, serve the directory via any local HTTP server:

```bash
# Using Python
python -m http.server 8000

# Or using Node.js
npx serve
```

Then navigate to `http://localhost:8000` in your web browser.

## 🛠️ Tech Stack & Libraries

- **Tailwind CSS** (via CDN)
- **Vanilla JavaScript** (ES6+)
- [PapaParse](https://www.papaparse.com/) — Fast in-browser CSV parsing
- [SheetJS (xlsx)](https://sheetjs.com/) — Excel spreadsheet reading & writing
- [JSZip](https://stuk.github.io/jszip/) — ZIP extraction for embedded Excel drawing media
- [jsPDF](https://github.com/parallax/jsPDF) & [jspdf-autotable](https://github.com/simonbengtsson/jsPDF-AutoTable) — Landscape PDF generation

## 📄 License

This project is licensed under the [MIT License](LICENSE).
