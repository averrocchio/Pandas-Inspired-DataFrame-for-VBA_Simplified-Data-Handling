# Pandas-Inspired-DataFrame-for-VBA_Simplified-Data-Handling
An open‑source VBA library that provides a memory‑centric DataFrame for VBA. It lets you load, filter, sort, append and merge data using concise methods, hiding loops behind a clean interface and integrating seamlessly with ranges and tables.

## Motivation and objectives
In many corporate environments, Excel is the only permitted tool for data analysis, and integrating external programming languages like Python may require IT approval or be unavailable on older business versions of Office.
Even when Python is available in Excel, it cannot modify existing ListObjects (tables) and often returns data in spill format rather than altering structures in place. Consequently, this project was conceived to deliver Pandas‑style functionality directly in VBA. The main objectives are to:
Provide a fully in‑memory data layer for Excel ranges, arrays and tables, similar to a pandas DataFrame.
Simplify data manipulation in VBA by hiding loops behind intuitive class methods for loading, filtering, sorting, appending and concatenating data.
Maintain full integration with Excel, allowing users without Python support to perform advanced data operations while staying within the VBA environment.
Offer optional diagnostics and performance metrics to help debug and optimise data operations.
## Core structure:
The library centres on a DataFrame class that encapsulates a 2‑D Variant array (with 1‑based indices) and maintains metadata such as column headers, optional keys and type information.
The class is organised into several sections, each providing a specific set of functionalities:
Properties (read‑only and read/write): Methods such as RowsCount, ColsCount and Shape return the dimensions of the data, while properties like NullToken and DebugMode allow users to customise how blanks are treated and to enable or disable diagnostic output.
Loaders: Functions to load data from arrays, ListObjects (Excel tables) and ranges (LoadFromArray, LoadFromListObject, LoadFromRange). These methods normalise headers, ensure unique column names and read data efficiently into the internal 2‑D array.
Core operations: High‑level methods (currently stubs) such as Filter, Sort, Dedup, Project, Rename, Append, JoinRight, Clean and InferTypes. These are designed to mirror familiar pandas operations, allowing users to filter rows by conditions, sort by multiple columns, remove duplicates, project or rename columns, append rows or join DataFrames and perform basic data cleaning.
I/O functions: Methods to convert the DataFrame back to an array (AsArray), write or append data to an Excel table (WriteTo, AppendTo) and retrieve performance metrics (Metrics), with support for chaining operations through WithDebug.
Helpers and diagnostics: Internal utilities handle array transformations, header normalisation and optional diagnostic outputs for debugging and performance measurement.

## Notes:
This project is currently under development. Collaboration has been proposed on ForumExcel.it (https://www.forumexcel.it/forum/threads/creazione-di-pandas-per-vba-titolo-accattivante.83207/#post-683135, you're welcome to join the thread). If you’re interested in collaborating testing, opening issues or proposing features you’re very welcome to join the project.

Portions of this project were drafted with ChatGPT. The code is progressively under reviewing and testing as development advances.



