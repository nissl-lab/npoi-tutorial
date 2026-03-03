# NPOI Examples (Advanced Example Subscription) 

---

## OOXML

| Example | Description |
|---|---|
| `CreateBasicOOXMLFile` | Create a basic OOXML document/package from scratch (OOXML container basics). |
| `ModifyExistingOOXMLFile` | Open an existing OOXML package and modify its contents/parts. |

---

## POIFS

| Example | Description |
|---|---|
| `CreateCustomProperties` | Write custom document properties into an OLE2/POIFS container. |
| `CreatePOIFSFile` | Create a basic POIFS (OLE2) file/compound document. |
| `CreatePOIFSFileWithProperties` | Create a POIFS file and include document properties/metadata. |
| `ReadThumbsDB` | Read/parse a `Thumbs.db` file (OLE2/POIFS-based) to extract information. |

---

## SXSSF

| Example | Description |
|---|---|
| `CreateWorkbook` | Create a streaming XLSX workbook using SXSSF (suited for large data with low memory usage). |

---

## SS (common spreadsheet / SS usermodel examples)

| Example | Description |
|---|---|
| `BusinessPlan` | Generate a “business plan” style spreadsheet report with formatting. |
| `CalendarDemo` | Generate a calendar-like spreadsheet layout (dates, cells, formatting). |
| `ColorfulMatrixTable` | Create a matrix/table with color styling to visualize structure/values. |
| `CopyRowAndEvaluate` | Copy rows and re-evaluate formulas/values after copying. |
| `LinkedDropDownLists` | Create dependent/linked data validation drop-down lists. |
| `LoanCalculator` | Build a loan calculator spreadsheet (formulas + layout). |
| `MeringCells` | Demonstrate merging cells/regions in a sheet. |
| `MonthlySalaryReport` | Generate a monthly salary report spreadsheet with formatting/formulas. |
| `MultplicationTable` | Generate a multiplication table in a sheet. |
| `ReadAndPrintData` | Read an existing spreadsheet and print/output cell data. |
| `SetCellValues` | Basic demo of writing values to cells (strings/numbers/dates). |
| `SetPrintArea` | Set workbook/sheet print area settings. |
| `TimeSheetDemo` | Create a timesheet spreadsheet template/report. |
| `UseBasicFormula` | Demonstrate creating and using basic Excel formulas. |
| `WorkbookFactoryDemo` | Use `WorkbookFactory` to open workbooks without caring about XLS vs XLSX. |

---

## XWPF (Word .docx)

| Example | Description |
|---|---|
| `ChangeOrientation` | Change page orientation (portrait/landscape) in a DOCX document. |
| `ComplexTableLayout` | Build a more complex DOCX table layout (merged cells, styling, etc.). |
| `CreateBulletList` | Create bullet lists in a DOCX document. |
| `CreateComments` | Add comments/annotations to a DOCX document. |
| `CreateEmptyDocument` | Create a minimal empty DOCX document. |
| `CreateFieldRun` | Insert field runs (e.g., dynamic fields) into a DOCX paragraph. |
| `CreateHeaderAndFooter` | Create headers and footers in a DOCX document. |
| `CreateHighlightRun` | Apply text highlighting to a run in a DOCX document. |
| `CreateHyperlink` | Insert hyperlinks into a DOCX document. |
| `CreateNestedTable` | Create nested tables (a table inside another table) in DOCX. |
| `CreateOMathFormula` | Create OMML (Office Math) formulas in DOCX. |
| `CreateSingleLineWithAlignments` | Create a paragraph/single line with different alignments/styling. |
| `CreateWatermark` | Add a watermark to a DOCX document. |
| `CreateWordTableMerge` | Merge cells in a Word table. |
| `CreateWordTableWithBulletList` | Combine tables and bullet lists in a DOCX document. |
| `GeneragePageNumber` | Insert/generate page numbers in header/footer. |
| `InsertPicturesInWord` | Insert images into a DOCX document. |
| `IterateTables` | Iterate through tables in a DOCX document and inspect contents. |
| `MapObjectToTable` | Map object/data structures into a DOCX table (templated table population). |
| `ReplaceTexts` | Find/replace text in a DOCX document. |
| `ResizeEmbededPicture` | Resize an embedded picture in DOCX. |
| `SimpleDocument` | Create a simple DOCX with basic paragraphs/runs. |
| `SimpleTable` | Create a simple table in DOCX. |
| `UpdateEmbeddedDoc` | Update an embedded document/object inside a DOCX container. |

---

## HSSF (Excel .xls)

| Example | Description |
|---|---|
| `AddHyperlinkInXls` | Add hyperlinks to cells in an `.xls` workbook. |
| `ApplyFontInXls` | Create/apply font styles in `.xls`. |
| `AutoSizeColumnInXls` | Auto-size columns based on content in `.xls`. |
| `ChangeSheetTabColorInXls` | Change worksheet tab color in `.xls`. |
| `ConditionalFormattingInXls` | Apply conditional formatting rules in `.xls`. |
| `ConvertExcelToHtml` | Convert an `.xls` sheet/workbook to HTML. |
| `CopyRowsAndCellsInXls` | Copy rows/cells (values + styles) within `.xls`. |
| `CopySheet` | Copy/clone a sheet in an `.xls` workbook. |
| `CreateDropDownListCellInXls` | Create data validation drop-down lists in `.xls`. |
| `CreateEmptyExcelFile` | Create an empty `.xls` workbook. |
| `CreateHeaderFooterInXls` | Create headers/footers in `.xls`. |
| `CustomColorInXls` | Define/use custom palette colors in `.xls`. |
| `DisplayGridlinesInXls` | Control gridline display settings in `.xls`. |
| `DrawingInXls` | Draw shapes/graphics in `.xls`. |
| `EnableAutoFilterInXls` | Enable auto-filter on a cell range in `.xls`. |
| `ExportXlsToDownload` | Write an `.xls` to a stream/response for download (web scenario). |
| `ExtractPicturesFromXls` | Extract embedded images from `.xls`. |
| `ExtractStringsFromXls` | Extract shared strings/text content from `.xls`. |
| `FillBackgroundInXls` | Set background fill color/pattern in `.xls`. |
| `GenerateXlsFromXlsTemplate` | Fill an `.xls` template to generate output. |
| `GroupRowAndColumnInXls` | Group/outline rows and columns in `.xls`. |
| `HideColumnAndRowInXls` | Hide rows/columns in `.xls`. |
| `ImportXlsToDataTable` | Import `.xls` data into a `DataTable`. |
| `InsertPicturesInXls` | Insert images into an `.xls` sheet. |
| `NumberFormatInXls` | Apply numeric/date formats in `.xls`. |
| `ProtectSheetInXls` | Protect a sheet with a password/permissions in `.xls`. |
| `RepeatingRowsAndColumns` | Configure repeating title rows/columns for printing. |
| `RotateTextInXls` | Rotate cell text in `.xls`. |
| `SetActiveCellRangeInXls` | Set active cell and selection range in `.xls`. |
| `SetAlignmentInXls` | Set horizontal/vertical alignment in `.xls`. |
| `SetBorderStyleInXls` | Apply border styles to cells in `.xls`. |
| `SetBordersOfRegion` | Apply borders around a region (cell range). |
| `SetCellCommentInXls` | Add cell comments/notes in `.xls`. |
| `SetDateCellInXls` | Write dates to cells with date formatting in `.xls`. |
| `SetPrintAreaInXls` | Set print area for `.xls` worksheets. |
| `SetPrintSettingsInXls` | Configure print settings (paper size, orientation, etc.) in `.xls`. |
| `SetWidthAndHeightInXls` | Set row heights and column widths in `.xls`. |
| `ShrinkToFitColumnInXls` | Shrink-to-fit text in a column/cell. |
| `SplitAndFreezePanes` | Split and/or freeze panes in `.xls`. |
| `UseNewlinesInCellsInXls` | Insert newlines within a cell and configure wrapping. |
| `ZoomSheet` | Set sheet zoom level in `.xls`. |

---

## xssf (Excel .xlsx) — partial listing (API truncated)

| Example | Description |
|---|---|
| `AddHyperlinkInXlsx` | Add hyperlinks to cells in an `.xlsx` workbook. |
| `ApplyFontInXlsx` | Create/apply font styles in `.xlsx`. |
| `AreaChart` | Create an area chart in `.xlsx`. |
| `BarChart` | Create a bar chart in `.xlsx`. |
| `BigGridTest` | Stress/large-grid test for `.xlsx` creation/handling. |
| `BorderStylesInXlsx` | Demonstrate border styles in `.xlsx`. |
| `ConditionalFormats` | Apply conditional formatting in `.xlsx`. |
| `CopySheet` | Copy/clone a sheet in `.xlsx`. |
| `CreateCommentInXlsx` | Add cell comments in `.xlsx`. |
| `CreateCustomProperties` | Write custom document properties into an OOXML `.xlsx` package. |
| `CreateEmptyWorkbook` | Create an empty `.xlsx` workbook. |
| `CreateHeaderFooterInXlsx` | Create headers/footers in `.xlsx`. |
| `CreateTableInXlsx` | Create an Excel “Table” (structured table) in `.xlsx`. |
| `CreateWorkbookFromDataSet` | Generate a workbook from a .NET `DataSet`. |
| `DataFormatsInXlsx` | Apply number/date data formats in `.xlsx`. |
| `DownloadXlsx` | Write an `.xlsx` to stream/response for download (web scenario). |
| `ExtractPicturesFromXlsx` | Extract embedded images from `.xlsx`. |
| `ExtractTextFromXlsx` | Extract text content from `.xlsx`. |
| `FillBackgroundInXlsx` | Set background fill color/pattern in `.xlsx`. |
| `FillDateOnlyValue` | Write date-only values with appropriate cell type/format. |
| `HideColumnAndRowInXlsx` | Hide rows/columns in `.xlsx`. |
| `InsertPicturesInXlsx` | Insert images into `.xlsx`. |
| `InsertRowInExistingSheet` | Insert a row into an existing `.xlsx` sheet (shift rows). |
| `LineChart` | Create a line chart in `.xlsx`. |
| `PageSetupInXlsx` | Configure page setup settings (margins, orientation, etc.) in `.xlsx`. |
| `PieChart` | Create a pie chart in `.xlsx`. |
| `PrintSetupInXlsx` | Configure print setup options in `.xlsx`. |
| `ProtectSheetInXlsx` | Protect a sheet with password/permissions in `.xlsx`. |
| `ReadSheetsByXSSFReader` | Read sheets using `XSSFReader` (streaming/event-style reading). |
| `ScatterChart` | Create a scatter chart in `.xlsx`. |
| `SetIsRightToLeftInXlsx` | Set worksheet layout to right-to-left in `.xlsx`. |
| `SetRowStyle` | Apply styling at the row level in `.xlsx`. |
| `SetWidthAndHeightInXlsx` | Set row heights and column widths in `.xlsx`. |
| `SplitAndFreezePanesInXlsx` | Split and/or freeze panes in `.xlsx`. |
| `WritePerformanceTest` | Performance benchmark for writing `.xlsx`. |
