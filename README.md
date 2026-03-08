# Pandas-Inspired-DataFrame-for-VBA_Simplified-Data-Handling
An open‑source VBA library that provides a memory‑centric DataFrame for VBA. It lets you load, filter, sort, append and merge data using concise methods, hiding loops behind a clean interface and integrating seamlessly with ranges and tables.

## Motivation and objectives
In many corporate environments, Excel is the only permitted tool for data analysis, and integrating external programming languages like Python may require IT approval or be unavailable on older business versions of Office.
Even when Python is available in Excel, it cannot modify existing ListObjects (tables) and often returns data in spill format rather than altering structures in place. Consequently, this project was conceived to deliver Pandas‑style functionality directly in VBA.

Current goals:
- Provide a fully in‑memory data layer for Excel ranges, arrays and tables.
- Simplify data manipulation in VBA by hiding loops behind fluent methods.
- Maintain full integration with Excel for users without Python support.
- Offer optional diagnostics and performance metrics for debug and optimisation.

## Core structure
The library centres on a `DataFrame` class that encapsulates a 2‑D Variant array (1‑based indices) with metadata such as headers, keys and diagnostics.

Implemented areas:
- **Loaders**: `LoadFromArray`, `LoadFromListObject`, `LoadFromRange`.
- **Core ops (available)**: `Project`, `Rename`, `Append`, `Filter`, `Sort`, `Dedup`, `JoinRight` (MVP), `Clean`, `InferTypes`.
- **I/O**: `AsArray`, `WriteToRange`, `Metrics`.
- **Core ops (available)**: `Project`, `Rename`, `Append`.
- **Core ops (planned)**: `Filter`, `Sort`, `Dedup`, `JoinRight`, `Clean`, `InferTypes`.
- **I/O**: `AsArray`, `WriteToRange`.

## Usage examples (copy/paste ready)

### 1) Project columns (subset + order)
```vb
Sub Example_Project()
    Dim data(1 To 3, 1 To 3) As Variant
    Dim hdr(1 To 3) As Variant
    data(1, 1) = 1: data(1, 2) = "Anna": data(1, 3) = "IT"
    data(2, 1) = 2: data(2, 2) = "Bruno": data(2, 3) = "HR"
    data(3, 1) = 3: data(3, 2) = "Carla": data(3, 3) = "IT"
    hdr(1) = "id": hdr(2) = "name": hdr(3) = "dept"

    Dim df As New DataFrame
    Dim out As DataFrame
    df.LoadFromArray data, hdr

    Set out = df.Project("dept,name")
    out.WriteToRange Sheet1.Range("A1"), True
End Sub
```

### 2) Filter + sort
```vb
Sub Example_Filter_Sort()
### 2) Rename columns
```vb
Sub Example_Rename()
    Dim df As New DataFrame
    ' ...load df...

    Dim out As DataFrame
    Set out = df.Filter("city contains rom").Sort("city,id", "asc,desc")
End Sub
```

### 3) Dedup by key
```vb
Sub Example_Dedup()
    Dim df As New DataFrame
    ' ...load df...

    df.Keys = "CustomerID"
    Dim out As DataFrame
    Set out = df.Dedup("keep_first")
End Sub
```

### 4) JoinRight MVP
```vb
Sub Example_JoinRight()
    Dim leftDf As New DataFrame
    Dim rightDf As New DataFrame
    ' ...load both dataframes with key "id"...

    Dim out As DataFrame
    Set out = leftDf.JoinRight(rightDf, "id", "right", "_L,_R")
End Sub
```

### 5) Clean + InferTypes
```vb
Sub Example_Clean_InferTypes()
    Dim df As New DataFrame
    ' ...load df...

    Dim out As DataFrame
    Set out = df.Clean(True, True, True).InferTypes()
    Set out = df.Rename("dept:team,name:full_name")
End Sub
```

### 3) Append rows by header alignment
```vb
Sub Example_Append()
    Dim leftDf As New DataFrame
    Dim rightDf As New DataFrame
    ' ...load both dataframes...

    Dim merged As DataFrame
    Set merged = leftDf.Append(rightDf)
End Sub
```

## Known limits and edge cases
- `Filter` currently supports only a single condition (`=`, `<>`, `>`, `<`, `contains`).
- `Sort` is stable but uses insertion sort (good for small/medium sets, not yet optimised for very large datasets).
- `JoinRight` MVP requires unique keys in left dataframe; duplicate keys raise error.
- `AppendTo` and `WriteToListObject` are still TODO.
- `Project` rejects duplicated column specifications (e.g. `"name,name"`).
- `Rename` currently supports mapping via string (`"old:new"`) or `Scripting.Dictionary`.
- `Append` esegue unione schema per nome colonna: mantiene l'ordine della tabella sinistra e aggiunge eventuali colonne extra della destra, riempiendo i mancanti con stringa vuota.
- `Filter`, `Dedup`, `JoinRight`, `Clean`, `InferTypes` are still stubs.

## Manual test module
A repeatable manual test module is included in `DataFrameTests.bas` with these entry points:
- `Test_LoadFromArray_Basic`
- `Test_LoadFromRange_WithHeader`
- `Test_Project_And_Rename`
- `Test_Append_HeaderUnion`
- `Test_Sort_MultiColumn`
- `Test_Filter_Contains`
- `Test_Dedup_ByKeys`
- `Test_Sort_MultiColumn`
- `Test_Clean_And_InferTypes`

Each test prints `PASS/FAIL` details in the Immediate window.

## Notes
This project is currently under development. Collaboration has been proposed on ForumExcel.it (https://www.forumexcel.it/forum/threads/creazione-di-pandas-per-vba-titolo-accattivante.83207/#post-683135, you're welcome to join the thread).

Portions of this project were drafted with ChatGPT. The code is progressively under review and testing as development advances.
