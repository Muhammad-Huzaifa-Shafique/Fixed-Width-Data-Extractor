# Fixed-Width Data Extractor

A Python script that turns a fixed-width text report into a clean Excel table. It finds the columns from the report's header, rebuilds rows that wrap across several lines, and saves a formatted `.xlsx` file next to the input.

I wrote it for a graduate data-wrangling course (GRT713) together with a classmate. We compared our outputs on the course file and they matched. The first lines of the script note this and link the ChatGPT conversation I used while working on it.

## Why it matters for ML work

Most machine-learning projects start with data that nobody designed for analysis. This script is a small example of that first step: recover the structure of a report, handle values that wrap across lines, and write a table that pandas and scikit-learn can load.

## What it does

1. **Finds the columns.** The first three lines of the file are headers. The header line with the most gaps of four or more spaces is the template, and each gap marks a column boundary.
2. **Builds the column names.** The three header lines are merged piece by piece, so a name that wraps over three lines becomes one name. Filler such as dots and `&` is removed.
3. **Finds the rows.** A new record starts when column 19 (index 18) is not blank. A line where it is blank continues the previous record.
4. **Rejoins wrapped text.** If the previous piece filled its whole column, the next piece is joined with no space (a word cut in the middle). Otherwise it is joined with one space.
5. **Writes the Excel file.** `<input name>_clean.xlsx` appears in the same folder, with a bold header row, wrapped text, column widths fitted to the content (capped at 60 characters), a frozen header row and an auto-filter.

## Try it

```bash
pip install pandas openpyxl
python fixed_width_to_excel.py
```

A file dialog opens so you can choose the TXT file. If no dialog is available, the script asks for the path in the terminal. To try it without your own data, choose `sample_input.txt` from this repository.

The sample input has three header lines and three records, two of which wrap:

```text
Item..    Customer............    Product.............    Unit......
No....    Name................    Description.........    Price.....
......    ....................    ....................    (USD).....
1001      Northwind Traders       Hex bolt, zinc          0.35
                                  plated
1002      Contoso Hardware        Pipe wrench 18 inch     12.90
1003      Fabrikam-Supply Co      Stainless steel anch    3.10
                                  or bolt
```

The result in `sample_input_clean.xlsx`:

| Item No | Customer Name | Product Description | Unit Price (USD) |
|---|---|---|---|
| 1001 | Northwind Traders | Hex bolt, zinc plated | 0.35 |
| 1002 | Contoso Hardware | Pipe wrench 18 inch | 12.90 |
| 1003 | Fabrikam-Supply Co | Stainless steel anchor bolt | 3.10 |

Record 1001 wrapped at a space, so its pieces are joined with a space. Record 1003 wrapped in the middle of a word ("anch" and "or"), so its pieces are joined with none.

## Notes

- **Version history.** The first commit is the script exactly as I submitted it. The second changes one line: `DataFrame.applymap` was removed in pandas 3, which made the script crash there, so it now calls `DataFrame.map` (and falls back to `applymap` on pandas older than 2.1).
- Every cell is written as text, including numbers such as `12.90`. Convert numeric columns before analysis.
- The layout rules (three header lines, the row marker in column 19, columns separated by four or more spaces) fit the course report. For a different marker column, change `ROW_START_COL` at the top of the script.
- `sample_input.txt` is a small synthetic report made for this repository. It is not the course data.

## What I would add next

- Command-line options for the number of header lines and the marker column, instead of constants.
- Unit tests for edge cases such as a value that wraps twice or an empty column.

## Author

Made by [Muhammad Huzaifa Shafique](https://github.com/Muhammad-Huzaifa-Shafique).
