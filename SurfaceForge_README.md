# SurfaceForge — Surface Area Report Macro for SolidWorks

> Every Surface. Every Formula. One Click.

**Free SolidWorks VBA macro. Run once. Get a professional surface area report instantly.**

---

## What It Does

SurfaceForge automatically analyses every surface in your SolidWorks part or assembly and generates a professional HTML report — with Excel and PDF export built in.

No plugins. No installers. One `.bas` file.

---

## Features

| Feature | Detail |
|---------|--------|
| **Auto Surface Classification** | Classifies every face into Front / Rear / Top / Bottom / Left / Right using surface normals |
| **5 Surface Types** | Planar · Cylindrical · Conical · Spherical · Toroidal — correct formula for each |
| **Formula Generation** | Generates the geometric formula with actual dimensions for every face |
| **Material Detection** | Reads material from SolidWorks properties. Detects sheet metal gauge automatically |
| **Prep Notes** | Auto-generates surface preparation instructions per material (Steel, SS, Aluminum, etc.) |
| **Screenshot Capture** | Embeds an isometric screenshot of your model directly in the report |
| **Multi-Format Export** | HTML (interactive) · Excel (.xls) · PDF (print to PDF) |
| **Editable Report** | Enter company name, revision, author, paint color, and prep notes directly in the report |
| **Smart Mode Detection** | Automatically detects: Selected Faces / Selected Components / Isolated View / Full Assembly / Single Part |

---

## Install in 30 Seconds

**Method 1 — Quick Run (no setup)**
```
1. Open your Part or Assembly in SolidWorks
2. Tools → Macro → Run
3. Browse to SurfaceAreaReport__1_.bas → Open
4. Select Main → Run
```

**Method 2 — Toolbar Button (recommended)**
```
1. Tools → Macro → New → save as SurfaceForge.swp
2. In VBA Editor: File → Import File → select the .bas file
3. Close VBA Editor
4. Tools → Customize → Commands → Macro
5. Drag "New Macro Button" to toolbar → point to SurfaceForge.swp → Method: Main
```

**Method 3 — Keyboard Shortcut**
```
1. Complete Method 2 first
2. Tools → Customize → Keyboard tab
3. Find Macros row → assign your shortcut (e.g. Shift + S)
```

---

## How It Works

When you run `Main()`:

1. Asks which plane is your **Front** (1 = Front/Z, 2 = Top/Y, 3 = Right/X)
2. Rebuilds the model to ensure geometry is current
3. Detects what you have open and selected
4. Scans every face — classifies type, side, calculates area
5. Captures an isometric screenshot
6. Generates and opens the HTML report in your browser automatically

---

## Report Modes

| What You Have | Report Generated |
|---------------|-----------------|
| Faces selected | Selected Faces Report |
| Components selected in assembly | Selected Components Report |
| Assembly with hidden parts | Isolated View Report |
| Full assembly | Full Assembly Report |
| Single part file | Part Report |

---

## Compatibility

Works with **SolidWorks 2014 through 2025** — all editions.
Windows only (SolidWorks requirement).
No additional software or plugins required.

---

## Output

The macro generates a standalone HTML file and opens it in your browser automatically. The report includes:

- Editable company name (gold branding strip)
- Revision, author, and project fields
- Summary cards: part count, surface count, total area (sq in + sq ft)
- Expandable part rows with face-by-face breakdown
- Formula breakdown per face with actual dimensions
- Paint color input field per part
- Editable surface preparation notes per part
- Export to Excel and PDF from inside the report

---

**Created by Ar. Kishan S. Lakhani**
Architectural Designer · SolidWorks · AutoCAD · BIM
© 2025 — Free & Open Source
Do not redistribute without permission.
