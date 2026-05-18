# VI-ROI-Computation-Macro

> **Open-source Python software for flexible region-of-interest (ROI)
> definition and vegetation index (VI) extraction from non-georeferenced
> multispectral turfgrass imagery.**

This tool is described in the following publication:

> Ku, K., Tuan, T. T., & Ahn, J. (2025). Development of a Flexible ROI
> Analysis Software for Improving Spatial Consistency in Turfgrass
> Time-Series Phenotyping. *Crop Science*.

---

## Table of Contents

1. [Overview](#1-overview)
2. [Repository Structure](#2-repository-structure)
3. [Requirements](#3-requirements)
4. [Installation](#4-installation)
5. [Usage](#5-usage)
   - [5.1 GUI — Grass VI Computation Macro](#51-gui--grass-vi-computation-macro)
   - [5.2 Batch Script — JSON to VI Computation](#52-batch-script--json-to-vi-computation)
6. [Input / Output Formats](#6-input--output-formats)
7. [Workflow Summary](#7-workflow-summary)
8. [Citation](#8-citation)
9. [License](#9-license)
10. [References](#10-references)

---

## 1. Overview

Time-series multispectral phenotyping requires consistent plot-boundary
definitions across repeated image acquisitions. When camera position
shifts between sessions, fully automated region tracking is unreliable,
and re-delineating regions independently for each vegetation index
introduces inter-index spatial mismatches.

This software provides a **two-part workflow**:

| Component | Script | Purpose |
|---|---|---|
| **GUI** | `Grass VI Computation Macro.py` | Interactively place, rotate, and subdivide ROI boxes on VI rasters; export per-ROI mean VI values and JSON coordinates |
| **Batch script** | `JSON to VI Computation.py` | Apply saved JSON ROI coordinates to any additional VI rasters from the same time point — no re-delineation needed |

**Key capabilities:**

- Rotatable, resizable rectangular ROIs aligned to oblique plot rows
- One-step *N*-way subdivision of a parent box into equal-width child
  subplots (e.g., 4 replicates per bed)
- Pixel-level spatial identity across NDVI, NDRE, SAVI, and any other
  precomputed VI from the same acquisition
- Operates directly on **non-georeferenced** pixel-coordinate TIFF
  rasters — no GIS software or CRS metadata required
- Exports per-ROI mean VI values to Excel (`.xlsx`) and ROI coordinates
  to JSON for audit and reuse
- Command-line interface for integration into automated pipelines

---

## 2. Repository Structure

```
VI-ROI-Computation-Macro/
│
├── Grass VI Computation Macro.py   # Component 1: Interactive PyQt5 GUI
├── JSON to VI Computation.py       # Component 2: Batch post-processing script
├── JSON to VI Computation-v2...    # Legacy version (retained for reference)
├── test.py                         # Unit tests / development tests
│
├── sample/                         # Sample data for getting started
│   ├── ndvi/                       # Example single-band NDVI TIFF rasters
│   ├── output/                     # Example output files (Excel, JSON, PNG)
│   └── README.md                   # Sample data usage notes
│
├── requirements.txt                # Python dependency list (pip-installable)
├── LICENSE                         # MIT License
└── README.md                       # This file
```

### Component 1: `Grass VI Computation Macro.py`

The interactive GUI application. Built with **PyQt5** and **OpenCV**.

Core classes and responsibilities:

| Class / Module | Responsibility |
|---|---|
| `MainWindow` | Top-level QMainWindow; manages file list, parameter fields, and canvas layout |
| `GraphicsScene` | QGraphicsScene subclass; handles mouse events for ROI drag and drop |
| `ROIRectItem` | QGraphicsRectItem subclass; stores `(x_c, y_c, w, h, θ)` parameters and renders the rotated box |
| `SubdivisionManager` | Computes *N* equal-width child ROI geometries from a parent ROI |
| `VIExtractor` | Rasterizes each ROI as a binary mask and computes mean pixel value from the floating-point raster |
| `ExportManager` | Serializes ROI coordinates to JSON and VI statistics to `.xlsx` |

### Component 2: `JSON to VI Computation.py`

The headless batch processing script. Accepts CLI arguments.

Core functions:

| Function | Responsibility |
|---|---|
| `load_roi_json(path)` | Parses the JSON coordinate file exported by the GUI |
| `apply_rois_to_image(img, rois)` | Rasterizes all ROIs onto a new VI raster and returns per-ROI means |
| `export_excel(results, out_path)` | Writes per-ROI VI statistics to `.xlsx` |
| `main()` | Parses CLI arguments (`-j`, `-i`, `-o`) and orchestrates the batch loop |

---

## 3. Requirements

| Package | Minimum version | Tested version |
|---|---|---|
| Python | 3.5 | 3.8.12 |
| PyQt5 | 5 | 5.15.4 |
| opencv-contrib-python | 4 | 4.5.2.54 |
| XlsxWriter | latest | 1.2.6 |
| pandas | latest | 1.2.4 |

**Operating system:** Windows (primary development platform).
Linux and macOS are expected to work but have not been formally tested.

**GPU:** Not required (CUDA: None).

---

## 4. Installation

```bash
# 1. Clone the repository
git clone https://github.com/kibonku/VI-ROI-Computation-Macro.git
cd VI-ROI-Computation-Macro

# 2. (Recommended) Create and activate a virtual environment
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS / Linux:
source venv/bin/activate

# 3. Install all dependencies
pip install -r requirements.txt
```

---

## 5. Usage

### 5.1 GUI — Grass VI Computation Macro

Launch the interactive application:

```bash
python "Grass VI Computation Macro.py"
```

**Step-by-step workflow:**

| Step | Action | GUI element |
|---|---|---|
| A | Select the folder containing your VI TIFF rasters | *Input Folder* button |
| B | Select an output folder for results | *Output Folder* button |
| C | Set ROI parameters: width (px), height (px), subdivision count *N*, name prefix | Parameter fields |
| D | Click *Create* (or `Ctrl+C`) to place a new ROI; drag to target plot | Canvas |
| E | Rotate the ROI to align with oblique plot rows via slider or `Ctrl+MouseWheel` | Rotate slider |
| F | Click *Save VI and ROI* to compute per-ROI mean values and save outputs | Save button |
| G | Processed files are highlighted in blue; click the next file to continue | File list |
| H | Click *Export DB* when all images are done to write the final Excel + JSON files | Export button |

**ROI identifier format:**

Each ROI is named `{BoxName}{index}` (e.g., `test1`, `test2`).  
Each parent ROI subdivided into *N* parts produces child identifiers
`{BoxName}{index}` for replicates 1 through *N*.

---

### 5.2 Batch Script — JSON to VI Computation

Apply ROI coordinates saved from a GUI session to additional VI rasters
(e.g., NDRE, SAVI) from the same acquisition:

```bash
python "JSON to VI Computation.py" \
    -j "/path/to/output/ndvi.json" \
    -i "/path/to/ndre_rasters/" \
    -o "/path/to/output/"
```

**Arguments:**

| Flag | Description |
|---|---|
| `-j` | Path to the JSON coordinate file exported by the GUI |
| `-i` | Directory containing the target VI raster files (`.tif` / `.tiff`) |
| `-o` | Output directory for the resulting Excel file(s) |

Because the batch script applies stored pixel-coordinate vertices
directly, the extracted values for NDVI, NDRE, and SAVI are
**guaranteed to originate from identical pixel sets** — no
re-delineation is performed.

---

## 6. Input / Output Formats

### Inputs

| File type | Description |
|---|---|
| `*.tif` / `*.tiff` | Single-band floating-point VI rasters (NDVI, NDRE, SAVI, etc.); read with `cv2.IMREAD_UNCHANGED` to preserve 32/64-bit precision |
| `*.jpg` / `*.png` | RGB images for visual inspection in the GUI canvas |
| `*.json` | ROI coordinate file produced by a previous GUI session (batch script only) |

### Outputs

| File type | Description |
|---|---|
| `*.xlsx` | Per-ROI mean VI values; one row per ROI, one column per image file |
| `*.json` | ROI coordinate metadata: identifier, four corner vertices (top-left, top-right, bottom-right, bottom-left in pixel coordinates), and parameterized fields `(x_c, y_c, w, h, θ)` |
| `*_screenshot.png` | Canvas screenshot with overlaid ROI boxes for visual audit |
| `*_mask.png` | Binary mask image containing only the ROI footprints; useful for downstream texture analysis or quality control |

---

## 7. Workflow Summary

```
Multispectral VI rasters (NDVI / NDRE / SAVI *.tif)
        │
        ▼
┌─────────────────────────────────────┐
│  Grass VI Computation Macro.py      │  ← Interactive GUI
│  • Load VI raster                   │
│  • Place + rotate ROI boxes         │
│  • N-way subdivide into subplots    │
│  • Compute per-ROI mean VI          │
│  • Save VI + ROI (screenshot, mask) │
│  • Export DB → Excel + JSON         │
└────────────────┬────────────────────┘
                 │  ndvi.json  (ROI coordinates)
                 ▼
┌─────────────────────────────────────┐
│  JSON to VI Computation.py          │  ← Headless batch script
│  • Load JSON coordinates            │
│  • Apply to NDRE / SAVI rasters     │
│  • Compute per-ROI mean VI          │
│  • Export Excel                     │
└─────────────────────────────────────┘
        │
        ▼
  Excel workbook (.xlsx)
  per-ROI VI statistics
  pixel-identical across all indices
```

---

## 8. Citation

If you use this software in your research, please cite:

```bibtex
@article{ku2025viroi,
  author  = {Ku, Kibon and Tuan, Thai Thanh and Ahn, Jinhyun},
  title   = {Development of a Flexible {ROI} Analysis Software for
             Improving Spatial Consistency in Turfgrass Time-Series
             Phenotyping},
  journal = {Crop Science},
  year    = {2025},
  url     = {https://github.com/kibonku/VI-ROI-Computation-Macro}
}
```

---

## 9. License

This project is licensed under the terms of the **MIT License**.
See [LICENSE](LICENSE) for details.

---

## 10. References

- Qt documentation: https://doc.qt.io/qt-5/qtexamplesandtutorials.html
- OpenCV documentation: https://docs.opencv.org/4.5.2/
- Python argparse: https://docs.python.org/3/library/argparse.html
- XlsxWriter documentation: https://xlsxwriter.readthedocs.io/
- Ku et al. (2023). Identification of new cold tolerant Zoysia grass
  species using high-resolution RGB and multi-spectral imaging.
  *Scientific Reports*, 13, 13209.
