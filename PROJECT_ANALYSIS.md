### MOWeD-LAC Project Analysis

This document provides a detailed analysis of the MOWeD-LAC repository based on the current workspace contents.

## Overview

- **Project name**: MOWeD-LAC (Molecular-wide Electron Density (De)localization Atomic Counts)
- **Type**: Desktop GUI application (Python, CustomTkinter)
- **Primary goal**: Parse computational chemistry outputs (AIMAll-generated .atomicfiles, .wfx) and FALDI-derived summaries to produce structured Excel workbooks summarizing electron localization/delocalization metrics at atomic and molecular levels.
- **Executables bundled**: FALDI command-line tools in `assets/` used to generate `.sum` data files from user inputs and `.wfx` sources.

## Repository Structure

- `MOWeD-LAC.py`: Main Python application (~4,900+ lines). Contains the GUI, processing logic, and Excel report generation.
- `assets/`: Runtime assets
  - Executables: `FALDI_pops.exe`, `FALDI_grid_12_02_2024.exe`, legacy variants, and `msys-2.0.dll`
  - UI media: `icon.ico`, screenshots, and illustration images
  - Documentation: `FALDIAdvancedOptionsDescription.(pdf|docx)`, suggested literature docs
- `MOWeD-LAC_12-07-2023.exe`: A prebuilt Windows executable of the app
- `Readme.md`: High-level description, features, usage, installation

No Python dependency files (e.g., `requirements.txt`, `pyproject.toml`) are present in the repo snapshot, though the README references `requirements.txt`.

## Key Dependencies (in-code imports)

- GUI: `customtkinter` (Tkinter wrapper), `CTkMessagebox`
- Files and system: `os`, `sys`, `pickle` (`dump`, `load`), `random`
- Images: `Pillow` (`PIL.Image.open` as `open_PIL`, `PIL.ImageTk.PhotoImage`)
- Excel: `openpyxl` (styles, images, workbook/sheets management)
- External processes: `os.system` to call included FALDI executables

The code supports PyInstaller-style frozen mode, using `sys._MEIPASS` for `base_path` and optional `pyi_splash` handling.

## Application Architecture

### Entrypoint and Base Path

- Entrypoint at bottom of `MOWeD-LAC.py`:
  - Instantiates `ATOMIC_App()` and starts Tk mainloop
  - Closes `pyi_splash` if running as a frozen executable
- `base_path` is determined by runtime:
  - Frozen: `sys._MEIPASS`
  - Normal: `os.path.abspath('.')`
  - Used to resolve assets (icons, images, FALDI executables, PDF help)

### Major GUI Classes

- `ATOMIC_App(customtkinter.CTk)`: The main window and frame container; manages navigation between pages.
- `ATOMIC_HomeFrame(CTkFrame)`: Landing page with theme/scaling controls, info, and primary actions:
  - Open FALDI 3D dialog, generate FALDI-pops.sum, choose FALDI-pops.sum
- `ATOMIC_InputFrame(CTkFrame)`: Collects molecule system name, counts, atom names, and thresholds; displays parsed info.
- `ATOMIC_Page2Frame(CTkFrame)`, `ATOMIC_Page3Frame(CTkFrame)`: Subsequent steps post-parse and pre-export.
- `ScrollableCheckBoxFrame(CTkScrollableFrame)`: UI helper.
- `HelpWindow(CTkToplevel)`: In-app multi-page help, including screenshots and references; keyboard navigation; segmented buttons.
- `FaldiWindow(CTkToplevel)`: Accepts advanced/custom FALDI commands; writes to `inp.add` for the external tool.
- `CustomSpinbox(CTkFrame)`: Small numeric input with +/- buttons, validation using `CTkMessagebox` for errors.

### FALDI Integration and Data Flow

There are two primary flows involving FALDI:

1) Use an existing `FALDI-pops.sum`
   - Action: "Choose FALDI-pops.sum file"
   - Reads file with `filedialog.askopenfilename`
   - Parses and splits into structured lists/dicts used throughout the app

2) Generate `FALDI-pops.sum` from `.wfx`
   - Action: "Generate FALDI-pops.sum file"
   - Collects `.wfx` via file dialog; writes `inp.add` in current working directory
   - Builds a command line using `assets/FALDI_pops.exe` (or grid variant), appending any custom commands from `FaldiWindow`
   - Executes via `os.system(cmd)`, showing modal progress using `ProgressCTkMessagebox`
   - On completion, reads the output `.sum` file from the working directory and parses into memory structures

`assets/FALDI_grid_12_02_2024.exe` is similarly invoked for grid-based processing.

### Parsing and In-Memory Model

The application parses the `.sum` file into:

- Scalars and lists: `number_of_atoms`, `atoms_names`, `total_electron`, `loc_electron`, `deloc_electron`, `intra_molecule`, `inter_molecule`, `atoms_pairs`, `total_deloc_electron`, etc.
- Derived mappings and aggregations:
  - Per-molecule totals (`total_for_each_molecule`)
  - Intra- and inter-molecular dictionaries (`intra_mol_dict`, `inter_mol_dict`, `inter_mol_dict_pair`)
  - Sorted views for reporting (e.g., `intra_sorted_according_to_total_deloc_electron`, `inter_sorted_according_to_contri`)
- Utilities:
  - `get_numbers(string)`: parses ranges like `1-3,5,7-8` into integer lists with validation
  - `get_names_list(numbers, atom_names)`
  - `add_cell_value(...)`: helper to set the value and multiple cell properties in one call

On successful parse, a summary is displayed and the app navigates to input confirmation and next steps.

### Excel Report Generation (openpyxl)

The application produces a multi-sheet workbook. Observed sheet names and creation points include:

- Active sheet (initial data)
- `atom-pairs(A,B)`
- `atoms-molecule-pairs`
- `intra-Mol(A,B)_cov`
- `intra-Mol(A,B)_non-cov`
- `inter-Mol(A,B)`

Formatting features include:

- Bold fonts for totals/headers; center alignment and wrap text
- Programmatic sums for row/column totals using Python-side aggregation before writing
- Image insertion into a sheet via `openpyxl.drawing.image.Image` with height normalization to ~10 cm
- Number formatting via a shared `number_properties` dict passed to `add_cell_value`

At save time:

- A `.xlsx` file is saved to a user-chosen location with default name `<system_name>_MOWeD-LAC.xlsx`
- A parallel `.pickle` file with the same stem is written, containing a dictionary of all intermediate data structures for reproducibility/post-processing

### Error Handling and UX

- Extensive use of `CTkMessagebox` and `ProgressCTkMessagebox` for feedback: errors, progress, success
- Defensive parsing with exception handling; on error, shows warning and returns to a safe state
- Appearance mode (System/Light/Dark) and UI scaling controls exposed on the home frame
- Help window with embedded screenshots, a references page, and keyboard navigation

## Runtime Considerations

- Windows-centric executables: The bundled `.exe` tools and `os.startfile` will not function on non-Windows systems without alternatives. Linux/macOS use would require corresponding binaries and conditionals.
- Working directory dependence: FALDI runs and writes `inp.add`/`.sum` to the current working directory. The app expects and reads from there; this may surprise users if the CWD differs from the project/data folder.
- Packaging: The app checks `sys._MEIPASS` to locate assets when frozen (PyInstaller). Splash support via `pyi_splash` is optional.

## Security and Robustness Notes

- External command invocation uses `os.system(cmd)`. Consider `subprocess.run([...], check=True)` for safer argument handling and better error capture.
- Paths and filenames from dialogs should be sanitized and validated; large files and non-ASCII paths may need explicit handling.
- The pickle output is for trusted contexts only; loading pickles is unsafe with untrusted sources.

## Gaps Observed in Repo Snapshot

- No `requirements.txt` even though README mentions it. Based on imports, expected packages include: `customtkinter`, `CTkMessagebox`, `Pillow`, `openpyxl`.
- No tests. Given complex parsing/aggregation, unit tests would add confidence.
- No CI/CD configuration.

## Recommendations

- Replace `os.system` with `subprocess.run` and capture stdout/stderr for user-visible logs.
- Factor parsing and workbook generation into separate modules for maintainability and testability.
- Add a `requirements.txt` and a minimal developer setup guide.
- Add a configuration option for output directory and working directory isolation (e.g., temporary job folder per run).
- Add platform guards or cross-platform binaries for non-Windows environments, or disable FALDI-run buttons on unsupported OS.
- Persist user preferences (theme, scaling, last-used directories) via a small JSON in user data dir.

## How to Run (from source)

1) Install dependencies (approximate):
   - `pip install customtkinter CTkMessagebox pillow openpyxl`
2) Launch: `python MOWeD-LAC.py`
3) Provide `.wfx` and AIMAll `.atomicfiles` data, or choose an existing `FALDI-pops.sum`
4) Follow the UI to generate and save the Excel report

## Outputs

- Excel workbook with multiple domain-specific sheets summarizing electron density localization/delocalization metrics
- Companion pickle file capturing the full computation state for reproducibility

## Assets and Documentation

- In-app help and `assets/FALDIAdvancedOptionsDescription.pdf`
- Screenshots in README and help
- Suggested literature links embedded in `Readme.md` and the help window

