# CLAUDE.md — AI Assistant Guide

This file provides context for AI assistants (e.g., Claude Code) working on this repository.

---

## Project Overview

**Pandas-Inspired-DataFrame-for-VBA** is an open-source VBA library that brings Pandas-style in-memory DataFrame functionality to Microsoft Excel. It targets corporate environments where Python/external tools may be unavailable, giving analysts powerful data manipulation capabilities purely in VBA.

**Language:** VBA (Visual Basic for Applications), runs inside Excel
**Target:** Excel 2016+ (Windows), tested on Office 365
**No build toolchain** — all development happens inside the Excel VBA IDE

---

## Repository Structure

```
/
├── DataFrame.cls           # Main class — the entire DataFrame implementation (~2,200 lines)
├── DataFrameTests.bas      # Manual test module (~440 lines)
└── README.md               # End-user documentation with examples and known limits
```

The project is intentionally minimal. There is no package manager, no CI pipeline, and no compiled output. VBA source files (`.cls`, `.bas`) are imported directly into an Excel workbook.

---

## Key Files

### `DataFrame.cls`

The entire library lives in this single class file. It is organized into clearly marked sections using `[SEC]` markers (searchable with Ctrl+F):

| Section | Contents |
|---|---|
| `[SEC] Lifecycle` | `Class_Initialize`, `Class_Terminate` |
| `[SEC] Properties (Read-Only)` | `RowsCount`, `ColsCount`, `Shape`, `header` |
| `[SEC] Properties (Read-Write)` | `NullToken`, `DebugMode`, `Keys` |
| `[SEC] Loaders` | `LoadFromArray`, `LoadFromListObject`, `LoadFromRange` |
| `[SEC] Core Operations` | `Filter`, `Sort`, `Dedup`, `Project`, `Rename`, `Append`, `JoinRight`, `Clean`, `InferTypes` |
| `[SEC] Input/Output` | `AsArray`, `WriteToRange`, `WriteToListObject`, `AppendTo`, `Metrics` |
| `[SEC] Diagnostics` | `WithDebug` |
| `[SEC] Helpers` | Private internal utilities |

**Internal data structures:**
- `mData` — core data as a 2D `Variant` array (1-based indexing)
- `mHeader` — column headers as a 1D `Variant` array (1-based)
- `mKeys` — key column names/indices for dedup and join
- `mMetrics` — `Scripting.Dictionary` for diagnostics and timing
- `mAliases` — `Scripting.Dictionary` for column alias resolution

### `DataFrameTests.bas`

Manual test module. Each function is self-contained and prints `PASS` or `FAIL` to the VBA Immediate window. There is no test runner — each function must be invoked manually from the IDE or via a wrapper macro.

**Test entry points:**
- `Test_LoadFromArray_Basic`
- `Test_LoadFromRange_WithHeader`
- `Test_Project_And_Rename`
- `Test_Append_HeaderUnion`
- `Test_Sort_MultiColumn`
- `Test_Filter_Contains`
- `Test_Dedup_ByKeys`
- `Test_Clean_And_InferTypes`

---

## Code Conventions

### Naming

| Pattern | Used for |
|---|---|
| `PascalCase` | Public methods (`LoadFromArray`, `WriteToRange`) |
| `camelCase` | Public properties (`RowsCount`, `header`) |
| `m` prefix | Private member variables (`mData`, `mHeader`, `mKeys`) |
| `Arr_` prefix | Private array utility helpers |
| `Hdr_` prefix | Private header manipulation helpers |
| `G_` prefix | General private helpers |

### Error Handling

- Custom error enum `DFErr` with codes in range `11010–11999`
- All errors are raised via `RaiseDf(source, code, message)` helper
- Callers use `On Error GoTo <label>` with labelled cleanup sections

### Data Indexing

- All internal arrays are **1-based** (standard VBA convention)
- Public API accepts both 1-based and name-based column references

### Fluent Interface

Most methods return `Me` (the `DataFrame` instance) to enable chaining:

```vba
df.Filter("Status = Active").Sort("Date", False).Project("Date,Name,Amount").WriteToRange(ws.Range("A1"))
```

I/O methods (`WriteToRange`, `WriteToListObject`, `AppendTo`) also return `Me` for consistency.

### Section Navigation

Use `' [SEC] SectionName` comments as bookmarks. In the VBA IDE, Ctrl+F and searching for `[SEC]` jumps through major sections.

### Comments

- Inline comments appear in both **Italian and English** (the project was developed in collaboration with an Italian Excel community — `ForumExcel.it`)
- Do not remove or translate existing comments
- New comments should be in English

---

## Public API Summary

### Loaders

```vba
' Load from a 2D Variant array
df.LoadFromArray(data As Variant, header As Variant)

' Load from an Excel ListObject (table)
df.LoadFromListObject(lo As ListObject)

' Load from a worksheet Range with rich boundary options
df.LoadFromRange(rng As Range, _
    Optional hasHeader As Boolean = True, _
    Optional headerAt As DFHeaderAt = dfRow, _
    Optional bounds As DFBounds = dfliteral, _
    Optional maxExtendRows As Long = 50)
```

### Core Operations

```vba
df.Filter(condition As String)          ' e.g. "Status = Active", "Amount > 100", "Name contains Jo"
df.Sort(cols As Variant, _
    Optional ascending As Variant)      ' multi-column; per-column direction array
df.Dedup(Optional policy As String = "keep_first")
df.Project(cols As Variant)             ' column subset/reorder; array or comma-string
df.Rename(map As Variant)               ' Dictionary or "OldName=NewName,..." string
df.Append(other As DataFrame)           ' schema union append
df.JoinRight(right As DataFrame, _
    keys As Variant, _
    Optional how As String = "inner", _
    Optional suffixes As Variant)       ' MVP: inner/right join
df.Clean(Optional trimSpaces As Boolean = True, _
    Optional collapseSpaces As Boolean = True, _
    Optional softCoerce As Boolean = True)
df.InferTypes()                         ' statistical type detection per column
```

### Output

```vba
df.AsArray()                            ' returns 2D Variant array
df.WriteToRange(dest As Range, Optional includeHeader As Boolean = True)
df.WriteToListObject(lo As ListObject)
df.AppendTo(lo As ListObject)
df.Metrics()                            ' returns Dictionary of diagnostics
```

### Properties

```vba
df.RowsCount    ' Long, read-only
df.ColsCount    ' Long, read-only
df.Shape        ' String, e.g. "12×5", read-only
df.header       ' Variant array copy, read-only
df.NullToken    ' String (default ""), read-write
df.DebugMode    ' Boolean, read-write
df.Keys         ' Variant, set key columns for Dedup/JoinRight
```

---

## Enums

```vba
Enum DFErr        ' Error codes: 11010–11999
Enum DFHeaderAt   ' dfRow=1, dfColumn=2
Enum DFBounds     ' dfliteral=1, dfCurrentRegion=2, dfTrimOuterEmpty=3, dfSmart=4
```

---

## Known Limitations (Do Not Work Around Without Checking README)

1. **Filter:** Single condition only — no `AND`/`OR` compound conditions yet
2. **JoinRight (MVP):** Left DataFrame keys must be unique; right DataFrame can have duplicates
3. **Sort:** Uses insertion sort — not optimized for very large datasets (>10k rows may be slow)
4. **LoadFromRange `dfSmart` bounds:** Designed for header-down layout; other layouts may need `dfliteral`
5. **InferTypes:** Statistical threshold-based — may misclassify mixed columns
6. **No persistence:** DataFrame is in-memory only; data must be re-loaded each session
7. **Excel dependency:** Library requires `Scripting.Dictionary` (Microsoft Scripting Runtime reference)

---

## Development Workflow

### Setting Up

1. Open Excel and press `Alt+F11` to open the VBA IDE
2. Import `DataFrame.cls` via **File → Import File**
3. Import `DataFrameTests.bas` via **File → Import File**
4. Ensure **Microsoft Scripting Runtime** is referenced: **Tools → References → Microsoft Scripting Runtime**

### Making Changes

- Edit `DataFrame.cls` directly in the IDE or in a text editor
- After editing in a text editor, re-import the file into the workbook
- Export updated files via **File → Export File** before committing to git

### Running Tests

- In the Immediate window: `Call Test_LoadFromArray_Basic`
- Or place cursor inside a test function and press `F5`
- Results appear in the Immediate window (`Ctrl+G`)

### Committing

- Export `.cls` and `.bas` files from the VBA IDE before staging
- Commit messages follow imperative mood: `"Fix filter operator for numeric comparison"`, `"Add InferTypes statistical threshold"`
- Do not commit `.xlsm` or `.xlsx` Excel files (binary, not diff-friendly)

---

## Things to Avoid

- **Do not add external dependencies** — the library must work with only standard VBA + Scripting Runtime
- **Do not change 1-based array indexing** — all internal arrays use 1-based indexing per VBA convention
- **Do not remove `[SEC]` markers** — they are the navigation system for this large file
- **Do not add MsgBox calls** — use `Debug.Print` or the `mMetrics` dictionary for diagnostics
- **Do not break the fluent interface** — all operation methods must return `Me`
- **Do not add Italian comments** — new comments should be in English; existing Italian comments are intentional

---

## Feature Status

| Feature | Status |
|---|---|
| `LoadFromArray` | Stable |
| `LoadFromListObject` | Stable |
| `LoadFromRange` | Stable (4 boundary modes) |
| `Filter` | MVP (single condition) |
| `Sort` | Stable (multi-column stable sort) |
| `Dedup` | Stable |
| `Project` | Stable |
| `Rename` | Stable |
| `Append` | Stable (schema union) |
| `JoinRight` | MVP (inner/right, unique left keys) |
| `Clean` | Stable |
| `InferTypes` | Stable (statistical) |
| `WriteToRange` | Stable |
| `WriteToListObject` | Stable |
| `AppendTo` | Stable |
| Compound Filter (`AND`/`OR`) | Planned |
| Left/Full Outer Join | Planned |
| GroupBy / Aggregate | Planned |
| Pivot | Planned |

---

## Attribution

- Portions of code were initially drafted with ChatGPT assistance and are under ongoing review
- Community collaboration via [ForumExcel.it](https://www.forumexcel.it)
