# Robert's Definitions Tool

An Excel VBA macro that automatically inserts the correct **Relevant Definitions** under the clauses in agreements. Instead of searching manually, the tool reads the definitions column you specify, matches the defined terms used in each clause, and inserts only the definitions that apply — formatted correctly.

---

## How to open the tool

1. Open your workbook in Excel and press **Alt + F8**.
2. Select **`Launch_RobertsMacro_UI`** and click **Run**.
3. The input form will open.

---

## What to fill in

| Field | What to enter | Examples |
|---|---|---|
| **Definitions column** | The column letter containing the block of definitions for each row | `AL` |
| **Clause columns** | The columns to search for clause text | `H:AK` · `H,J,M` · `H,J:K,N` |
| **Rows to process** | The rows to apply the tool to | `2:500` · `2,5,10` · `2,5,10:15` |

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

> **Tip:** Ask an AI assistant to format your definitions list correctly if needed.

---

## Buttons

| Button | What it does |
|---|---|
| **Save choices** | Saves the three input fields into a hidden sheet so they auto-load next time. |
| **Run RobertsMacro** | Runs the tool using your inputs. |
| **Report a bug** | Opens a pre-addressed email to report issues. |
| **Close** | Closes the form. |

---

## What the tool does when you click Run

1. Reads the definitions from the column you specified.
2. Scans each clause cell in the columns and rows you specified.
3. Matches any defined terms that appear in the clause text.
4. Inserts a **`Relevant Definitions:`** block at the bottom of each matching cell, listing only the definitions that are actually used in that clause.
5. Bolds the `Relevant Definitions:` label.

If you run the tool again, the existing definitions block is replaced rather than duplicated.

---

## Tips

- Keep everything in **plain text**.
- The macro **overwrites all existing rich formatting** (bold, italics, font colour) in affected cells — run it before applying manual formatting.
- If possible, run the macro in your working document rather than a local copy, so that formatting is not lost on copy-paste.
- If columns or rows do not respond, check for typos in the input boxes.

---

## File structure

| File | Purpose |
|---|---|
| `modHelpers.bas` | Column/row parsing utilities and URL encoder |
| `modLaunchAndCore.bas` | UI entry point and input validation |
| `modCore.bas` | Core engine: definition matching and cell writing |
| `modSettings.bas` | Persist and load user choices in a hidden worksheet |
| `frmRobertsMacro.frm` | UserForm definition and event handlers |
| `frmRobertsMacro.frx` | UserForm binary resource (required alongside the `.frm`) |

To import into your workbook, open the VBA editor (**Alt + F11**), right-click the project, choose **Import File**, and import each `.bas` and `.frm` file.

> **Important:** When importing `frmRobertsMacro.frm`, the matching `frmRobertsMacro.frx` file **must** be in the same folder. VBA reads the `.frx` automatically when you import the `.frm` — if it is missing you will get *"Class MSForms.Label was not a loaded control class"* errors for every control.
