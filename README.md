# TEM Vacuole Area Analyzer

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22113727.svg)](https://doi.org/10.5281/zenodo.22113727)

A semi-automatic Python workflow for **cell segmentation, manual cell review, vacuole review, and vacuole-area quantification in transmission electron microscopy (TEM) images of *Chlorella vulgaris*.**

**Author:** Adriana Pereira C. Sánchez  
**Repository:** `Chlorella-TEM-Vacuole-Analyzer`  
**Version:** 1.0.0  
**Archived release:** Zenodo, DOI: [10.5281/zenodo.22113727](https://doi.org/10.5281/zenodo.22113727)  
**License:** MIT  
**Associated manuscript:** *Microfluidic separation for microalgae: what matters is the inside of the cell*  
**Manuscript status:** Unpublished

## Purpose

TEM Vacuole Area Analyzer was developed to quantify intracellular vacuolar area in *Chlorella vulgaris* cells from TEM images. The workflow combines automatic image preprocessing and segmentation with **two sequential interactive review stages**: first the cell boundaries are reviewed and corrected, and only then are vacuoles detected and reviewed.

This design ensures that vacuole analysis is performed only inside user-approved cell masks, including images in which automatic cell detection fails completely.

For each reviewed cell, the software reports:

- cell area;
- number of vacuoles;
- summed vacuole area;
- vacuole-to-cell area ratio (%);
- cell and vacuole areas in µm² when spatial calibration is available.

The software also creates annotated images and a PowerPoint report for visual quality control.

## Example of the analysis

The example below shows a representative *Chlorella vulgaris* TEM image before and after analysis with the TEM Vacuole Area Analyzer.

![Representative TEM vacuole analysis](examples/tem_vacuole_analyzer_example.png)

```text
Results example:
Cells: 1
N vacuoles: 3
Cell area µm²: 5.255
Ratio vacuole/cell: 54%
```

**Representative analysis example.** (A) Original TEM image of a *Chlorella vulgaris* cell. (B) Reviewed segmentation generated using TEM Vacuole Area Analyzer. The **green contour** represents the reviewed cell boundary, while **red contours** represent the reviewed vacuole boundaries.

## Analysis workflow

```text
TEM images
    ↓
brightness / contrast normalization
    ↓
scale-bar detection and OCR
    ├── automatic calibration when successful
    └── manual calibration or pixel-only processing when needed
    ↓
automatic cell pre-detection
    ↓
CELL REVIEW — green contours
    ├── A: add a missing cell by clicking near its center
    ├── B: manually draw a cell boundary
    ├── C: erase an incorrect cell contour
    └── D: split merged cells using one or more cut lines
    ↓
approved cell masks
    ↓
automatic vacuole pre-detection inside each approved cell
    ↓
VACUOLE REVIEW — red contours
    ├── A: add a missing vacuole from a seed click
    ├── B: manually draw a vacuole
    └── C: erase an unwanted vacuole contour
    ↓
quantification
    ↓
Excel + annotated images + PowerPoint + processing log
```

## Requirements

Python 3.10 or newer is recommended.

Install the Python dependencies with:

```bash
pip install -r requirements.txt
```

The current workflow uses NumPy, OpenCV, tifffile, imagecodecs, pytesseract, pandas, openpyxl, python-pptx, Pillow, and Matplotlib.

### Tesseract OCR

Automatic reading of the scale label requires the **Tesseract OCR application**, in addition to the Python package `pytesseract`.

On Windows, install Tesseract and make sure `tesseract.exe` is either available in the system PATH or in a location searched by the script. The program also supports a manually specified Tesseract path through `CONFIG.TESSERACT_CMD`.

If Tesseract is unavailable, or if OCR fails for a particular image, the workflow can continue using manual calibration. Calibration failures can also be processed in pixel units when that option is selected.

## Input images

Supported formats are `.tif`, `.tiff`, `.png`, `.jpg`, and `.jpeg`.

The program is designed for grayscale TEM images. TIFF files containing microscopy metadata can also be loaded. Place the images to be analyzed in a single folder. The results folder created by the program is ignored during subsequent input-file discovery.

## How to run

From a terminal opened in the repository folder:

```bash
python tem_vacuole_analyzer.py
```

The program opens a folder-selection window. Select the folder containing the TEM images. A parameter window is then displayed. The default values correspond to the dataset for which the workflow was developed, and parameters can be adjusted without editing the source code.

Each image then follows two mandatory review stages: **cell review first, vacuole review second**.

## Cell review — green contours

The cell-review window opens **for every image**, even if automatic cell detection finds zero cells. At least one cell must be approved before the workflow can continue to vacuole analysis.

### A — Add a missing cell from a center click

Press `A`, then click near the center of a missing cell. The program attempts to recover the full outer cell contour using local grayscale information and several candidate segmentation thresholds.

This option is useful when the automatic detector misses a cell but the cell still has enough contrast relative to the background.

### B — Draw a cell manually

Press `B`, then hold the left mouse button and draw around the cell. The contour does **not** need to return exactly to the starting point: when the mouse button is released, the program automatically closes the last point to the first point, fills the region, and applies light smoothing to create the final green cell contour.

This option should be used when the click-based cell recovery is unsuccessful or when the automatic green boundary is clearly incorrect.

### C — Erase an incorrect cell

Press `C`, then click inside an unwanted green contour to remove that cell from the current review.

### D — Split merged cells

Press `D` when two or more cells have been detected as one connected green object.

Draw a cut line completely across the merged contour at the desired separation point. **Multiple cut lines can be drawn before applying them**, which allows three or more connected cells to be separated in the same review step.

For example, if three cells are connected, draw two separation lines and then press `Enter`. The program applies all pending cuts and attempts to generate three independent cell contours. The review window remains open so the result can be inspected before continuing.

### Cell-review controls

| Control | Action |
|---|---|
| `A` | Add/recover a missing cell by clicking near its center |
| `B` | Manually draw a complete cell boundary |
| `C` | Erase an incorrect green cell contour |
| `D` | Draw one or more split lines across merged cells |
| `Enter` in D mode | Apply all pending split lines |
| `Enter` outside D mode | Approve the current cell selection and continue to vacuoles |
| `Delete` or right click | Undo the last added cell or the last pending split line |
| `S` | Clear all cell selections from the current image |

## Vacuole review — red contours

After the green cell contours are approved, automatic vacuole pre-detection is performed **inside each approved cell mask**. A second review window then opens for vacuole correction.

### A — Add a missing vacuole from a seed click

Press `A`, then click inside a missing vacuole. The program samples the local grayscale intensity around the click and attempts to grow a connected region within the cell.

This method depends on local grayscale contrast between the vacuole and surrounding cytoplasm. It can be less successful in very dark or low-contrast cells; in those cases, use manual drawing with `B`.

### B — Draw a vacuole manually

Press `B`, then hold the left mouse button and draw the desired vacuole boundary. The drawn region is restricted to the corresponding approved cell mask before it is accepted.

### C — Erase an unwanted vacuole

Press `C`, then click inside a red vacuole contour to remove it.

### Vacuole-review controls

| Control | Action |
|---|---|
| `A` | Seed-guided vacuole addition |
| `B` | Manual vacuole drawing |
| `C` | Erase a selected vacuole |
| `Enter` | Approve the current vacuole selection |
| `Delete` or right click | Undo the last selection |
| `S` | Clear all vacuole selections from the current image |

## Cell segmentation

Cells are initially proposed as regions that are darker than the surrounding TEM background. Morphological filtering is used to remove noise, reconnect interrupted cell regions, fill internal holes, and refine the external cell silhouette.

Automatic cell segmentation is treated as a **pre-selection step only**. The final green cell contours are established during the cell-review stage using A/B/C/D controls.

A particularly important parameter when nearby cells are incorrectly merged is:

```python
CLOSE_KERNEL_CELL
```

Reducing this value makes the automatic detector less likely to connect neighboring cells. Increasing it helps reconnect separated pieces of the same cell. Even when automatic cells are merged, the `D` split tool can be used during manual review.

## Vacuole detection

The automatic stage proposes intracellular vacuole candidates using complementary strategies based on contour geometry, local grayscale contrast, texture, and morphology. One strategy attempts to preserve more detailed vacuole boundaries, while a complementary shape-based detector favors rounded or oval intracellular structures when the more detailed proposal is not suitable.

These automatic detections are **pre-selections**. The final accepted vacuole segmentation is established during the red-contour review.

Vacuoles may be round or elongated and may occupy a large fraction of the cell. Seed-guided Mode A can fail more often when the local contrast between the vacuole and cytoplasm is low, particularly in dark TEM images; Mode B provides a manual fallback in these cases.

## Brightness normalization

TEM images from different acquisitions may have different overall brightness levels. When brightness normalization is enabled, each image is normalized using configurable lower and upper intensity percentiles before segmentation.

If an image remains dark after normalization, optional automatic gamma correction can increase its brightness for analysis. The original image file is not overwritten.

## Scale calibration

The program searches for scale information in the lower-left region of the image. When automatic calibration succeeds, it detects the horizontal scale bar when possible, reads the printed scale value using OCR, recognizes µm or nm units, converts the value to micrometers, and calculates spatial calibration in µm/pixel.

If automatic OCR fails, manual calibration is available. The workflow can alternatively process calibration failures in pixel units.

## Main adjustable parameters

### Cell parameters

| Parameter | Function |
|---|---|
| `CELL_REENTRANCE_CLOSE_PX` | Closes inward indentations in the external cell boundary. |
| `CELL_OUTER_MAX_EXPANSION` | Limits expansion caused by external-boundary refinement. |
| `CLOSE_KERNEL_CELL` | Reconnects separated regions of the same cell; reduce it if nearby cells merge. |
| `OPEN_KERNEL_CELL` | Removes thin bridges and small artifacts. |
| `MIN_CELL_AREA_FRAC` | Minimum accepted automatic cell area relative to the image. |
| `MAX_CELL_AREA_FRAC` | Maximum accepted automatic cell area relative to the image. |
| `MIN_CELL_SOLIDITY` | Minimum automatic cell-contour compactness. |
| `MAX_CELL_ASPECT_RATIO` | Maximum accepted automatic cell elongation. |
| `MAX_CELLS_PER_IMAGE` | Maximum number of cells accepted during automatic pre-detection. |

### Interactive vacuole parameters

| Parameter | Function |
|---|---|
| `MANUAL_SEED_PATCH_RADIUS` | Neighborhood used to estimate local vacuole intensity after a click. |
| `MANUAL_SEED_MIN_AREA_FRAC` | Minimum accepted seed-guided vacuole area relative to the cell. |
| `MANUAL_SEED_MAX_AREA_FRAC` | Maximum accepted vacuole area relative to the cell. |
| `MANUAL_SEED_MIN_CIRCULARITY` | Minimum circularity for seed-guided candidates. |
| `MANUAL_SEED_MIN_SOLIDITY` | Rejects strongly indented or fragmented seed-guided candidates. |
| `MANUAL_SEED_MAX_ASPECT_RATIO` | Maximum accepted elongation. |
| `MANUAL_SEED_SMOOTH_FRAC` | Boundary smoothing relative to cell diameter. |
| `MAX_VACUOLES_PER_CELL` | Maximum number of reviewed vacuoles per cell. |

## Recommended parameter-adjustment order

For a new TEM dataset, avoid changing many parameters simultaneously. A practical sequence is:

1. inspect the automatic **green cell contours**;
2. correct missing or incorrect cells with A/B/C/D;
3. reduce `CLOSE_KERNEL_CELL` if automatic detection systematically merges nearby cells;
4. approve the final cell masks;
5. inspect the automatic **red vacuole proposals**;
6. correct individual vacuoles with A/B/C;
7. change vacuole or brightness parameters only if a systematic problem occurs across multiple images.

## Output

The program creates a results folder inside the selected input directory containing annotated images, an Excel workbook, a PowerPoint quality-control report, and a processing log.

The Excel workbook contains cell-level measurements and an image-level summary, including source file, cell number, number of vacuoles, cell area, total vacuole area, vacuole/cell area ratio, and spatially calibrated values when available.

## Quality control

The annotated images and PowerPoint report are intended for visual verification of both cell and vacuole segmentation. Automatic TEM segmentation should not be considered ground truth. TEM contrast can change with sample preparation, staining, section thickness, acquisition settings, intracellular composition, and image brightness.

For this reason, the workflow now requires **cell review before vacuole review**, ensuring that vacuole measurements are calculated only from user-approved green cell masks.

Users applying the software to a different dataset should validate the segmentation parameters for their own images.

## Reproducibility

For analyses reported in a publication, record the software version or Git commit, any parameters changed from the repository defaults, whether scale calibration was automatic or manual, whether any images were analyzed only in pixel units, and the reviewed annotated images used for quality control.

A tagged release should be used for the version associated with the manuscript.

## Citation

If you use this software, please cite the archived software release:

> **Pereira C. Sánchez, A. (2026).** *TEM Vacuole Area Analyzer* (Version 1.0.0) [Computer software]. Zenodo. https://doi.org/10.5281/zenodo.22113727

**DOI:** [10.5281/zenodo.22113727](https://doi.org/10.5281/zenodo.22113727)

The repository also includes a `CITATION.cff` file so that GitHub can provide a **Cite this repository** option.

## Associated manuscript

*Microfluidic separation for microalgae: what matters is the inside of the cell*

Status: **unpublished**.

The repository will be updated with the final article citation and DOI after publication.

## License

Copyright (c) 2026 Adriana Pereira C. Sánchez.

This software is distributed under the **MIT License**. See `LICENSE` for the complete license text.

## Disclaimer

This software was developed for research use and for a specific TEM image-analysis workflow. Results should be visually verified before scientific interpretation or publication.
