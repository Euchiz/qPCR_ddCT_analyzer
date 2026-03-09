# qPCR ddCT Analysis Framework

A modern, robust, and extendable Python GUI tool for analyzing qPCR (Quantitative Polymerase Chain Reaction) experiments using the standard $\Delta\Delta C_T$ (ddCT) method.

## Features

- **Intuitive Modern UI:** Built with Tkinter and `ttkthemes` for a clean, user-friendly data mapping experience.
- **Dynamic Control Mapping:** Auto-detects Targets and Samples. Allows mapping of a global Control Reference or unique Ref Gene/Ref Sample pairs **for every single Target independently**.
- **Automated Outlier Removal:** Implements an advanced, biologically-aware Median Absolute Deviation (MAD) logic tailored for small-n groups (like biological triplicates) to automatically flag and drop extreme CT outliers before averaging.
- **Custom Analysis Tasks:** Group samples by experimental subgroups (e.g., TPP vs Control) automatically based on metadata columns in your input sheet.
- **Manual Data Omission:** Need to drop a specific Target-Sample combination for biological reasons? Easily exclude it via an interactive checkbox grid!
- **Data Export:** Instantly outputs all intermediate calculations ($\Delta C_T$, $\Delta\Delta C_T$, RQ, and propagated asymmetric standard deviations) as a neat `.tsv` file.
- **Publication-Ready Plots:** Auto-generates Relative Quantification (RQ) grouped barcharts with error bars using `matplotlib`.
- **Config Templates:** Save your specific Target-Sample mappings to a `.json` configuration file, allowing you to load your complex analysis layouts instantly for future runs!

## Prerequisites

Ensure you have Python 3.8+ installed.

```bash
pip install -r requirements.txt
```

### Dependencies
- `pandas`
- `numpy`
- `matplotlib`
- `openpyxl`
- `ttkthemes`

## Usage

1. Run the interface:
   ```bash
   python qpcr_analyzer_gui.py
   ```
2. **Step 1:** Load your `.xlsx` exported CT results.
3. **Step 2:** Select which columns represent your `Sample`, `Target`, and `CT` values. Hit **Apply Mapping**.
4. **Step 3:** A scrollable grid will populate. Select the Reference Target (e.g., GAPDH, Actin) and Reference Sample for each Target. If most use the same controls, use the **Bulk Set All** menu.
5. *(Optional)* Click **Select Omissions** to drop specific Target/Sample pairs from being analyzed.
6. *(Optional)* Save your entire configuration mapping using **Save Config Template** at the top.
7. Click **Analyze & Generate Output**.

A TSV calculations table and a PNG Bar Chart will be saved right next to your input Excel file!
