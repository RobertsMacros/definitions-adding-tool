# Robert's Definitions Tool

**Roberts Macros: no macro too micro.**

An Excel VBA macro that automatically inserts the correct **Relevant Definitions** under clauses extracted from agreements. Instead of searching manually, the tool reads the definitions column you specify, matches the defined terms used in each clause, and inserts only the definitions that apply — formatted correctly.

---

## Installation

1. Download **`Relevant Definitions - Macro.xlam`** from this repository.
2. Open Excel and go to **File → Options → Add-ins**.
3. At the bottom, set the **Manage** dropdown to **Excel Add-ins** and click **Go**.
4. Click **Browse**, navigate to where you saved the `.xlam` file, select it, and click **OK**.
5. Make sure the checkbox next to the add-in is ticked and click **OK**.

The tool is now installed and available in every workbook. Press **Alt + F8**, select **`Launch_RobertsMacro_UI`**, and click **Run** to open it.

---

## How to use the tool

1. Open your workbook in Excel and press **Alt + F8**.
2. Select **`Launch_RobertsMacro_UI`** and click **Run**.
3. The input form will open.

### What to fill in

| Field | What to enter | Examples |
|---|---|---|
| **Definitions column** | The column letter containing the block of definitions for each row | `AL` |
| **Clause columns** | The columns to search for clause text | `H:AK` · `H,J,M` · `H,J:K,N` |
| **Rows to process** | The rows to apply the tool to | `2:500` · `2,5,10` · `2,5,10:15` |

### Buttons

| Button | What it does |
|---|---|
| **Save choices** | Saves the three input fields into a hidden sheet so they auto-load next time. |
| **Run** | Runs the tool using your inputs. |
| **Close** | Closes the form. |

---

## Definition format

Each definition in the definitions column must follow the pattern:

```
<Term> means <definition text>.
```

Multiple definitions in a single cell must be separated by a pipe character `|`:

```
Fun means Contract Review.|Contract Review means work.|Work means fun.
```

> **Tip:** Ask an AI assistant to format your definitions list correctly.

---

## What the tool does when you click Run

1. Reads the definitions from the column(s) you specified.
2. Scans each clause cell in the columns and rows you specified.
3. Matches any defined terms that appear in the clause text.
4. Inserts a **`Relevant Definitions:`** block at the bottom of each matching cell, listing only the definitions that are actually used in that clause.
5. Bolds and italicises the defined terms within the definitions block.

If you run the tool again, the existing definitions block is replaced rather than duplicated.

---

## Tips

- Keep everything in **plain text**.
- The macro **overwrites all existing rich formatting** (bold, italics, font colour) in affected cells — run it before applying manual formatting.
- If possible, run the macro in your working document rather than a local copy, so that formatting is not lost on copy-paste.
- If columns or rows do not respond, check for typos in the input boxes.

---

## Editing the source code

The source files are in the [`src/`](src/) folder. You can import the `.bas` and `.frm` files directly into the VBA editor as described above.

> **Note on `.frx` files:** A `.frx` file is a binary companion file that Excel generates alongside a `.frm` UserForm. It is **not included** in this repository because it cannot be meaningfully imported on its own — it stores binary layout data that is machine-generated and not human-editable. The `.frm` file contains all the actual code and control definitions; Excel regenerates the `.frx` automatically when you import and save the form. However, you may need to recreate the userform from scratch, and then add the `.frm` code to the userform. To edit the underlying source code, it is easiest to import modules to the .xlam directly. Userforms cannot be created on Mac. 

---

## Source files

| File | Purpose |
|---|---|
| `src/modHelpers.bas` | Column/row parsing utilities |
| `src/modLaunchAndCore.bas` | UI entry point and input validation |
| `src/modCore.bas` | Core engine: definition matching and cell writing |
| `src/modSettings.bas` | Persist and load user choices in a hidden worksheet |
| `src/frmRobertsMacro.frm` | UserForm definition and event handlers |
