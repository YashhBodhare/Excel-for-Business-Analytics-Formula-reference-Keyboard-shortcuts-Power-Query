# Excel for Business Analytics

### Formula reference · Keyboard shortcuts · Power Query

A practical revision guide for turning spreadsheet data into useful business information: look up records, summarise performance, clean exports, and prepare repeatable reports.

**Curated by Yash Bodhare** · [Visit my BA portfolio](https://thebayash.com)

Use this guide while practising in Excel. Each section connects a technique with a business task, so you can explain both **how it works** and **why you would use it**.

> **Resource status:** The cheat sheet is available below. The full interview question bank will be added after review; it is not included yet.

## Contents

- [Getting started](#getting-started)
- [1. Lookups and references](#1-lookups-and-references)
- [2. Logic and error handling](#2-logic-and-error-handling)
- [3. Aggregation and business metrics](#3-aggregation-and-business-metrics)
- [4. Text functions](#4-text-functions)
- [5. Dates and working days](#5-dates-and-working-days)
- [6. Modern Excel functions](#6-modern-excel-functions)
- [7. Keyboard shortcuts](#7-keyboard-shortcuts)
- [8. Power Query transformations](#8-power-query-transformations)
- [9. Combine monthly CSV files](#9-combine-monthly-csv-files)
- [10. Power Query working habits](#10-power-query-working-habits)
- [11. Practice challenges](#11-practice-challenges)
- [Interview question bank](#interview-question-bank)
- [Credits and references](#credits-and-references)

## Getting started

**Example context:** a sales analyst combines monthly transactions, adds customer information, and reports revenue by region.

Examples use an Excel Table named `SalesData`, with columns `OrderID`, `Date`, `CustomerID`, `Region`, `Product`, `Quantity`, `UnitPrice`, `Revenue`, and `Status`. A second table, `Customers`, contains `CustomerID` and `CustomerName`.

- Create tables with **Insert → Table** and rename them under **Table Design**.
- `SalesData[Revenue]` means the table's Revenue column; `[@Quantity]` means Quantity in the current row.
- Ordinary cell references such as `A2` are used in small examples. Named ranges such as `Holidays` must be defined before use.
- Formulas below use English function names and commas. Regional settings may require semicolons.
- Function availability depends on your Excel edition. `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `LET`, and `XLOOKUP` are available in Excel 2021 and later and Microsoft 365. `LAMBDA` requires a supported newer edition, such as Microsoft 365 or Excel 2024; do not assume Excel 2021 supports it.

## 1. Lookups and references

| Function | Purpose | Working example | Business use |
| --- | --- | --- | --- |
| `VLOOKUP` | Search the first column and return a value to its right | `=VLOOKUP(B2,Customers[[CustomerID]:[CustomerName]],2,FALSE)` | Match a customer ID to a name |
| `XLOOKUP` | Search one range and return from another | `=XLOOKUP(B2,Customers[CustomerID],Customers[CustomerName],"Not found")` | Enrich a sales report |
| `INDEX` + `MATCH` | Find a position, then return the corresponding value | `=INDEX(Customers[CustomerName],MATCH(B2,Customers[CustomerID],0))` | Lookup in either direction in older workbooks |
| `INDIRECT` | Convert text into a reference | `=INDIRECT("'"&A2&"'!B2")` | Read B2 from a sheet named in A2 |

<details>
<summary><strong>Lookup choices and common mistakes</strong></summary>

- `VLOOKUP` uses approximate matching when its last argument is omitted. Use `FALSE` for exact customer or product IDs.
- `XLOOKUP` defaults to exact matching, allows left or right lookups, and has an optional not-found result.
- `INDEX` + `MATCH` is not automatically a two-dimensional lookup. For a row-and-column intersection, use two `MATCH` expressions: `=INDEX(B2:E10,MATCH(H2,A2:A10,0),MATCH(I2,B1:E1,0))`.
- Exact-match lookups usually return the first match. Investigate duplicate keys before relying on the result.
- Check whether IDs are stored consistently as text or numbers and whether unwanted spaces exist.
- `INDIRECT` is volatile and can make large workbooks slower. Prefer direct references or structured tables when possible.

**Wildcard lookup:**

```excel
=XLOOKUP("*"&B2&"*",Customers[CustomerName],Customers[CustomerID],"Not found",2)
```

Here `2` enables wildcard matching. It does not reverse the search. To search from the last record toward the first, add `-1` as the sixth argument:

```excel
=XLOOKUP(B2,Customers[CustomerID],Customers[CustomerName],"Not found",0,-1)
```

</details>

## 2. Logic and error handling

| Function | Purpose | Example |
| --- | --- | --- |
| `IF` | Choose between two results | `=IF([@Revenue]>=10000,"High value","Standard")` |
| `IFS` | Return the result for the first true condition | `=IFS(A2>=90,"A",A2>=70,"B",TRUE,"Needs improvement")` |
| `AND` | Require all conditions to be true | `=IF(AND([@Revenue]>=10000,[@Status]="Paid"),"Eligible","Not eligible")` |
| `OR` | Require at least one condition to be true | `=IF(OR([@Region]="South",[@Region]="West"),"Priority region","Other")` |
| `NOT` | Reverse a logical result | `=NOT(ISBLANK(A2))` |
| `IFERROR` | Return a fallback for any formula error | `=IFERROR(A2/B2,"Check inputs")` |

**Analyst habit:** distinguish missing data, invalid inputs, and genuine zero values. Returning `0` for every error can make a broken calculation appear valid. A cell containing a formula that returns `""` is not truly blank to `ISBLANK`.

## 3. Aggregation and business metrics

| Function | Purpose | Example |
| --- | --- | --- |
| `SUMIFS` | Sum values matching multiple conditions | `=SUMIFS(SalesData[Revenue],SalesData[Region],"South",SalesData[Status],"Paid")` |
| `COUNTIFS` | Count rows matching multiple conditions | `=COUNTIFS(SalesData[Region],"South",SalesData[Status],"Paid")` |
| `AVERAGEIFS` | Average values matching conditions | `=AVERAGEIFS(SalesData[Revenue],SalesData[Region],"South")` |
| `SUMPRODUCT` | Multiply corresponding values and sum the products | `=SUMPRODUCT(SalesData[Quantity],SalesData[UnitPrice])` |

**Revenue for a calendar year:** use date boundaries, including an exclusive upper bound. This also handles date-time values correctly.

```excel
=SUMIFS(SalesData[Revenue],SalesData[Date],">="&DATE(2026,1,1),SalesData[Date],"<"&DATE(2027,1,1))
```

**Revenue in a calculated table column:**

```excel
=[@Quantity]*[@UnitPrice]
```

**Weighted average unit price:**

```excel
=IFERROR(SUMPRODUCT(SalesData[Quantity],SalesData[UnitPrice])/SUM(SalesData[Quantity]),"Check quantity")
```

Keep criteria ranges aligned. Use tables or bounded ranges for `SUMPRODUCT`; full-column array calculations can be expensive. An average of transaction rows is not necessarily an average per order if one order has multiple lines.

## 4. Text functions

| Function | Purpose | Example |
| --- | --- | --- |
| `LEFT` | Extract characters from the beginning | `=LEFT(A2,4)` |
| `RIGHT` | Extract characters from the end | `=RIGHT(A2,4)` |
| `MID` | Extract characters from a specified position | `=MID(A2,3,5)` |
| `LEN` | Count characters, including spaces | `=LEN(A2)` |
| `SEARCH` | Find a substring without case sensitivity | `=SEARCH("@",A2)` |
| `FIND` | Find a case-sensitive substring | `=FIND("SKU",A2)` |
| `SUBSTITUTE` | Replace matching text | `=SUBSTITUTE(A2,"old","new")` |
| `TEXT` | Format a value as text | `=TEXT(B2,"₹#,##0.00")` |
| `TEXTJOIN` | Combine text with a delimiter | `=TEXTJOIN(", ",TRUE,A2:A10)` |

**Additional cleaning helpers:** `=TRIM(A2)` removes excess ordinary spaces; `=CLEAN(A2)` removes certain nonprinting characters. Neither removes every possible imported Unicode character.

**Business example:** clean customer names before matching records. Keep amounts numeric for analysis; use cell number formats for display when you still need calculations. `TEXT` returns text, not a numeric amount.

## 5. Dates and working days

| Function | Purpose | Example |
| --- | --- | --- |
| `TODAY` | Current date | `=TODAY()` |
| `NOW` | Current date and time | `=NOW()` |
| `DATE` | Build a date from year, month, and day | `=DATE(2026,4,1)` |
| `YEAR` / `MONTH` / `DAY` | Extract date components | `=YEAR([@Date])` / `=MONTH([@Date])` / `=DAY([@Date])` |
| `EOMONTH` | Last date of an offset month | `=EOMONTH(A2,0)` |
| `NETWORKDAYS` | Count working days, excluding weekends and listed holidays | `=NETWORKDAYS(A2,B2,Holidays)` |
| `DATEDIF` | Calculate completed years, months, or days between dates | `=DATEDIF(A2,B2,"Y")` |

- `TODAY` and `NOW` recalculate; they are not fixed timestamps. Use the keyboard shortcuts below for a static entry.
- `NETWORKDAYS` includes qualifying start and end dates and assumes Saturday/Sunday weekends. Use `NETWORKDAYS.INTL` for other weekend patterns.
- `DATEDIF` is a legacy compatibility function documented by Microsoft. It has known limitations, particularly the `"MD"` unit; choose units deliberately and validate edge cases.
- Convert text dates to real dates before calculating. For international imports, check the intended day/month order.

## 6. Modern Excel functions

| Function | Purpose | Example |
| --- | --- | --- |
| `FILTER` | Return rows matching a condition | `=FILTER(SalesData,SalesData[Region]="South","No matching rows")` |
| `SORT` | Sort an array by a column index | `=SORT(SalesData,8,-1)` |
| `SORTBY` | Sort an array using another aligned array | `=SORTBY(SalesData,SalesData[Revenue],-1)` |
| `UNIQUE` | Return distinct values | `=UNIQUE(SalesData[Region])` |
| `LET` | Name values within a formula | `=LET(rev,SUM(SalesData[Revenue]),tax,rev*0.18,rev+tax)` |
| `LAMBDA` | Define a reusable calculation | `=LAMBDA(qty,price,qty*price)(5,200)` |

The `SORT` example assumes Revenue is the eighth column, as in the example schema. `SORTBY` avoids reliance on that column position. The `LET` tax rate is an illustrative assumption, not a universal tax rule.

<details>
<summary><strong>Spill ranges and reusable functions</strong></summary>

Dynamic arrays return results into neighbouring cells. Put spill formulas outside Excel Tables, leave the output area clear, and check for obstructions if `#SPILL!` appears.

To reuse a `LAMBDA`, open **Formulas → Name Manager**, create a name such as `LineRevenue`, and set **Refers to** to:

```excel
=LAMBDA(qty,price,qty*price)
```

Then call it in a cell:

```excel
=LineRevenue(5,200)
```

A `LAMBDA` definition entered directly in a cell without being called can return `#CALC!`.

</details>

## 7. Keyboard shortcuts

**Notation:** `+` means press together; `→` in a key sequence means press in order. These are desktop shortcuts, primarily for English/US keyboards. Mac settings, function keys, and keyboard layouts can affect behaviour; add `Fn` where your keyboard requires it.

### Navigation and selection

| Action | Windows | Mac |
| --- | --- | --- |
| Move to the edge of a data region | `Ctrl + Arrow` | `Cmd + Arrow` |
| Go to the beginning of the sheet | `Ctrl + Home` | `Control + Home` or `Control + Fn + Left` |
| Go to the last used cell | `Ctrl + End` | `Control + End` or `Control + Fn + Right` |
| Open Go To | `Ctrl + G` or `F5` | `Control + G` |
| Next / previous worksheet | `Ctrl + Page Down / Page Up` | `Option + Right / Left` |
| Move one screen down / up | `Page Down / Page Up` | `Fn + Down / Up` |
| Extend selection one cell | `Shift + Arrow` | `Shift + Arrow` |
| Extend selection to a data-region edge | `Ctrl + Shift + Arrow` | `Cmd + Shift + Arrow` |
| Select current region | `Ctrl + Shift + *` | `Cmd + A` (repeat to expand) |
| Select a column | `Ctrl + Space` | `Control + Space` |
| Select a row | `Shift + Space` | `Shift + Space` |
| Add a nonadjacent selection | `Ctrl + Click` | `Cmd + Click` |
| Extend selection with a click | `Shift + Click` | `Shift + Click` |
| Select visible cells only | `Alt + ;` | `Shift + Cmd + *` |
| Switch applications | `Alt + Tab` | `Cmd + Tab` |

Click the **Name Box** beside the formula bar to jump to a cell or named range. `Ctrl + A` selection depends on context; inside a populated region, repeated presses can expand from the region to the whole sheet.

### Editing, formatting, and cell flow

| Action | Windows | Mac |
| --- | --- | --- |
| Format Cells | `Ctrl + 1` | `Cmd + 1` |
| Bold / italic / underline | `Ctrl + B / I / U` | `Cmd + B / I / U` |
| Create an Excel Table | `Ctrl + T` | `Cmd + T` or `Control + T` |
| Insert hyperlink | `Ctrl + K` | `Cmd + K` |
| Insert static date | `Ctrl + ;` | `Control + ;` |
| Insert static time | `Ctrl + Shift + :` | `Cmd + ;` |
| Apply currency / percent format | `Ctrl + Shift + $ / %` | `Control + Shift + $ / %` |
| Insert cells, rows, or columns | `Ctrl + Shift + +` | `Control + Shift + +` |
| Delete cells, rows, or columns | `Ctrl + -` | `Control + -` |
| Hide row / column | `Ctrl + 9 / 0` | `Control + 9 / 0` |
| Fill down / right | `Ctrl + D / R` | `Cmd + D / R` or `Control + D / R` |
| Fill selected cells with the same entry | `Ctrl + Enter` | `Cmd + Return` or `Control + Return` |
| Flash Fill | `Ctrl + E` | `Control + E` |
| Copy the formula from the cell above | `Ctrl + '` | `Control + '` |
| Paste Special dialog | `Ctrl + Alt + V` | `Control + Cmd + V` |
| Cancel an edit | `Esc` | `Esc` |

For **Paste Values**, use Paste Special and choose Values. For borders, use Format Cells → Border. This avoids memorising layout-dependent symbol shortcuts.

### Formulas, PivotTables, and Power Query

| Action | Windows | Mac or menu route |
| --- | --- | --- |
| Edit active cell | `F2` | `Control + U` |
| Cycle relative / absolute references while editing | `F4` | `Cmd + T` |
| Recalculate workbook formulas | `F9` | `F9` (add `Fn` if needed) |
| Show / hide formulas | `Ctrl + grave accent` | `Control + grave accent` |
| Insert Function dialog | `Shift + F3` | `Shift + F3` (add `Fn` if needed) |
| AutoSum | `Alt + =` | `Shift + Cmd + T` |
| Complete a suggested function | `Tab` | `Tab` |
| Legacy array formula | `Ctrl + Shift + Enter` | `Control + Shift + Return` |
| Insert PivotTable | **Insert → PivotTable** | **Insert → PivotTable** |
| Refresh selected PivotTable | `Alt + F5` | **PivotTable Analyze → Refresh** |
| Refresh all workbook data | `Ctrl + Alt + F5` | **Data → Refresh All** |
| Drill into a PivotTable value | Double-click the value | Double-click the value |
| Group / ungroup worksheet outline | `Alt + Shift + Right / Left` | **Data → Group / Ungroup** |
| Open Power Query | **Data → Get Data** | **Data → Get Data**, if supported |
| Edit an existing query | **Data → Queries & Connections → right-click → Edit** | Use available Get Data/query controls |
| Edit full M code | Power Query **Home → Advanced Editor** | Power Query **Home → Advanced Editor** |
| Finish and load a query | Power Query **Home → Close & Load** | Use the available load controls |

Windows ribbon KeyTips are sequential and may vary by version or language: press `Alt` and follow the displayed letters. Windows Alt sequences are not portable Mac shortcuts. Power Query features and connectors also vary by platform.

`F9` recalculates formulas; refreshing data connections or PivotTables is a separate action. Most modern dynamic-array formulas use ordinary Enter rather than legacy array entry.

## 8. Power Query transformations

Power Query prepares data through repeatable steps. It is useful for cleaning exports, combining monthly files, and reshaping data before analysis.

In the following M examples, `Source` means the preceding table step. These are individual expressions, not a complete query.

| Task | Typical editor route | M expression |
| --- | --- | --- |
| Remove duplicate rows | Select relevant columns → Home → Remove Rows → Remove Duplicates | `Table.Distinct(Source, {"OrderID"})` |
| Change column type | Column header type icon | `Table.TransformColumnTypes(Source, {{"Revenue", type number}})` |
| Filter rows | Column filter dropdown | `Table.SelectRows(Source, each [Region] = "South")` |
| Replace text within a column | Transform → Replace Values | `Table.ReplaceValue(Source, "old", "new", Replacer.ReplaceText, {"Product"})` |
| Split by delimiter | Transform → Split Column → By Delimiter | `Table.SplitColumn(Source, "FullName", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"FirstName", "Remainder"})` |
| Fill down | Transform → Fill → Down | `Table.FillDown(Source, {"Region"})` |
| Fill up | Transform → Fill → Up | `Table.FillUp(Source, {"Region"})` |
| Add calculated column | Add Column → Custom Column | `Table.AddColumn(Source, "Revenue", each [Quantity] * [UnitPrice], type number)` |
| Aggregate by region | Home → Group By | `Table.Group(Source, {"Region"}, {{"TotalRevenue", each List.Sum([Revenue]), type number}})` |
| Merge two tables | Home → Merge Queries | `Table.NestedJoin(Orders, {"CustomerID"}, Customers, {"CustomerID"}, "CustomerData", JoinKind.LeftOuter)` |
| Append tables | Home → Append Queries | `Table.Combine({January, February})` |
| Unpivot month columns | Select identifier columns → Transform → Unpivot Other Columns | `Table.UnpivotOtherColumns(Source, {"Product"}, "Month", "Revenue")` |
| Pivot categories into columns | Transform → Pivot Column | `Table.Pivot(Source, List.Distinct(Source[Month]), "Month", "Revenue", List.Sum)` |
| Promote first row to headers | Home or Transform → Use First Row as Headers | `Table.PromoteHeaders(Source, [PromoteAllScalars=true])` |
| Remove columns | Select columns → Remove Columns | `Table.RemoveColumns(Source, {"Notes", "UnusedCode"})` |
| Rename a column | Double-click column header | `Table.RenameColumns(Source, {{"Amount", "Revenue"}})` |

**Merge vs append:** merge adds related information by matching keys; append stacks rows. After a nested merge, expand the nested column to select the fields you need.

**Data-quality checks:** define the record grain before deduplicating. Confirm join-key uniqueness to prevent row multiplication. Fill operations target `null` values, so convert empty strings where necessary. A simple name split may not handle compound names correctly.

## 9. Combine monthly CSV files

**Business problem:** monthly sales exports arrive separately. Prepare one clean table that can be refreshed when new files arrive.

**UI route:** Data → Get Data → From File → From Folder → Combine & Transform Data. Inspect the generated sample-file transformation and apply appropriate data types.

<details>
<summary><strong>Complete M example: combine CSVs with a header in every file</strong></summary>

This example assumes comma-delimited UTF-8 CSV files with consistent `Date` and `Revenue` columns. Dates use the `en-GB` parsing culture. Replace the folder and culture for your environment. `Folder.Files` includes subfolders.

```powerquery
let
    Source = Folder.Files("C:\Sales"),
    VisibleCSVFiles = Table.SelectRows(
        Source,
        each Text.Lower([Extension]) = ".csv"
            and [Attributes]?[Hidden]? <> true
    ),
    ParsedFiles = Table.AddColumn(
        VisibleCSVFiles,
        "Data",
        each
            let
                FileName = [Name],
                Parsed = Csv.Document(
                    [Content],
                    [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
                ),
                Headers = Table.PromoteHeaders(
                    Parsed,
                    [PromoteAllScalars = true]
                ),
                WithFileName = Table.AddColumn(
                    Headers,
                    "SourceFile",
                    each FileName,
                    type text
                )
            in
                WithFileName
    ),
    Combined = Table.Combine(ParsedFiles[Data]),
    Typed = Table.TransformColumnTypes(
        Combined,
        {{"Date", type date}, {"Revenue", type number}, {"SourceFile", type text}},
        "en-GB"
    )
in
    Typed
```

| Step | Purpose |
| --- | --- |
| `Source` | Read file metadata and binary contents |
| `VisibleCSVFiles` | Keep non-hidden CSV files, including uppercase extensions |
| `ParsedFiles` | Parse each file, promote its own header, and retain its filename |
| `Combined` | Stack the prepared tables |
| `Typed` | Convert columns using an explicit culture |

**Why headers are promoted inside each file:** combining raw CSV tables first and promoting only the first overall row can leave later files' headers among the records.

**Validate before reporting:** compare file count, row counts, and revenue totals against the source exports. Check missing columns, unexpected extra columns, conversion errors, duplicate records, and empty files. This is a teaching example, not a schema-validation framework; it assumes at least one valid file and does not suppress malformed-file errors.

</details>

## 10. Power Query working habits

- **Use explicit types and culture.** Check dates, decimal separators, and identifiers with leading zeros.
- **Name steps by purpose.** Names such as `FilteredPaidOrders` explain the transformation better than `Filtered Rows3`.
- **Understand query folding.** Some connectors can push transformations to the source. CSV folder imports do not behave like foldable database queries. Where supported, inspect View Native Query; its unavailability alone does not prove that folding has stopped.
- **Filter and select early when appropriate.** Reduce unnecessary data, while preserving required business logic.
- **Use reference queries for shared logic.** A reference inherits another query's steps; a duplicate creates a separate copy. References do not guarantee a single cached execution.
- **Keep staging queries connection-only.** Load only the outputs you need to the worksheet or Data Model.
- **Respect privacy levels.** Configure them according to the sources and the intended data combination. There is no universal “5× faster” privacy setting.
- **Measure before buffering.** `Table.Buffer` may help or slow a query and prevents downstream folding. Use it only for a specific, measured reason; placement alone does not make it beneficial.
- **Choose refresh behaviour deliberately.** Background refresh changes how refresh runs; disabling it does not guarantee that Excel stays responsive or that queries become faster.
- **Reconcile the output.** Confirm totals and record counts after merges, appends, filtering, and deduplication.

## 11. Practice challenges

These exercises were added for this guide. They are practice prompts, not a completed interview question bank.

| Challenge | Skills to practise | Completion check |
| --- | --- | --- |
| Add customer names to a sales export | `XLOOKUP` or `INDEX` + `MATCH` | Missing IDs and duplicate customer keys are investigated |
| Report paid revenue for South region | `SUMIFS`, `COUNTIFS` | Formula totals agree with filtered source records |
| Build a working-day turnaround measure | `NETWORKDAYS`, holiday list | Inclusive counting and weekend assumptions are documented |
| Create a dynamic regional view | `FILTER`, `SORTBY`, `UNIQUE` | Changing the selected region updates the results |
| Reshape a Jan–Dec sales matrix | Power Query unpivot | One row represents one product-month observation |
| Combine monthly CSVs | Folder import, data types, source filenames | New files appear on refresh without repeated header rows |

**Five-day revision plan:** navigation and selection → formulas and lookups → text and dates → Power Query → an end-to-end sales reporting task.

For each task, practise explaining: **the business need → the chosen approach → the validation → the decision the output supports.**

## Interview question bank

**Planned addition:** topic-based Excel interview questions, clear answers, and practical scenarios. This section will be expanded when the question material is supplied and reviewed.

The cheat sheet is useful on its own today. Once the question bank grows, it can move into a separate Markdown file while this README remains the repository's entry point.

## Credits and references

This guide is adapted from three user-supplied **Learn & Excel — Excel Interview Crack Pack** cheat sheets:

- `CS1_formula_reference.pdf` — Formula Reference.
- `CS2_shortcuts_master.pdf` — Shortcuts Master.
- `CS3_power_query_quickref.pdf` — Power Query Quick-Ref.

Original cheat sheets credit **Learn & Excel**, [learnandexcel.in](https://learnandexcel.in). Organisation, business examples, practice prompts, and technical clarifications in this README were added for Yash Bodhare's portfolio resource. Original branding and course content are not presented as Yash's authorship.

Official documentation for checking details and platform availability:

- [Microsoft: Keyboard shortcuts in Excel](https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)
- [Microsoft: Excel functions by category](https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)
- [Microsoft: Query folding basics](https://learn.microsoft.com/en-us/power-query/query-folding-basics)
- [Microsoft: Power Query M reference](https://learn.microsoft.com/en-us/powerquery-m/)
- [Microsoft: Table.Buffer](https://learn.microsoft.com/en-us/powerquery-m/table-buffer)

---

**Explore more:** [Yash Bodhare's Business Analysis portfolio](https://thebayash.com)
