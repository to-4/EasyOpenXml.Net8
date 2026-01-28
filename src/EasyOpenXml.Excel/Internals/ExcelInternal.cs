using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyOpenXml.Excel.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class ExcelInternal : IDisposable
    {
        private SpreadsheetDocument _document;
        private SheetManager _sheetManager;
        private bool _opened;
        private CellSnapshot _clipboard;

        // 列スタイルのキャッシュ（Min/Max を展開した辞書）
        private Dictionary<int, uint?>? _colStyleCache;
        internal int OpenBook(string strFileName, string strOverlay)
        {
            try
            {
                // 1. Open Excel file
                _document = SpreadsheetDocument.Open(strFileName, true);

                // 2. Initialize sheet manager
                _sheetManager = new SheetManager(_document);

                _opened = true;
                return 0;
            }
            catch
            {
                return -1;
            }
        }

        internal void CloseBook(bool mode)
        {
            if (!_opened) return;

            if (mode)
            {
                _document.Save();
            }

            Dispose();
        }

        internal int SheetNo
        {
            set => _sheetManager.SelectByIndex(value);
        }

        internal IReadOnlyList<string> SheetNames
            => _sheetManager.GetSheetNames();

        internal Pos Pos(int sx, int sy)
            => Pos(sx, sy, sx, sy);

        internal Pos Pos(int sx, int sy, int ex, int ey)
        {
            var proxy = new PosProxy(
                _document,
                _sheetManager.CurrentWorksheetPart,
                sx, sy, ex, ey);

            return new Pos(proxy);
        }

        internal CellWrapper Cell(string cell)
            => Cell(cell, 0, 0);

        internal CellWrapper Cell(string cell, int cx, int cy)
        {
            var proxy = new CellWrapperProxy(
                _document,
                _sheetManager.CurrentWorksheetPart,
                cell,
                cx,
                cy);

            return new CellWrapper(proxy);
        }

        internal void SetClipboard(CellSnapshot snapshot)
        {
            _clipboard = snapshot;
        }

        internal CellSnapshot GetClipboard()
        {
            return _clipboard;
        }

        internal void PrintArea(int sx, int sy, int ex, int ey)
        {
            var wb = _document.WorkbookPart.Workbook;

            // 1. DefinedNames は「既存があればそれを使う」
            var definedNames = wb.DefinedNames;
            if (definedNames == null)
            {
                definedNames = new DefinedNames();
                wb.InsertAfter(definedNames, wb.Sheets);
            }

            // 2. 現在シートの LocalSheetId（0-based）を求める
            //    LocalSheetId：「定義名（DefinedName）が、どのシートに属するか」を示す番号
            var sheets = wb.Sheets!.Elements<Sheet>().ToList();

            var currentName = _sheetManager.CurrentSheetName;
            var sheetIndex = sheets.FindIndex(s => string.Equals(s.Name?.Value, currentName, StringComparison.Ordinal));

            if (sheetIndex < 0)
                throw new InvalidOperationException($"Sheet not found: {currentName}");

            uint localSheetId = (uint)sheetIndex;

            // 3. 既存の Print_Area（同一シート分のみ）を削除
            var old = definedNames.Elements<DefinedName>()
                .Where(d => d.Name?.Value == "_xlnm.Print_Area"
                         && d.LocalSheetId != null
                         && d.LocalSheetId.Value == localSheetId)
                .ToList();

            foreach (var d in old)
                d.Remove();

            // 4. 正しい式を作る
            var sheetName = _sheetManager.CurrentSheetName.Replace("'", "''");
            var area = $"{AddressConverter.ToAbsoluteA1(sx, sy)}:{AddressConverter.ToAbsoluteA1(ex, ey)}";
            var formula = $"'{sheetName}'!{area}";

            // 5. シートスコープで追加（LocalSheetId が重要）
            definedNames.Append(new DefinedName
            {
                Name = "_xlnm.Print_Area",
                LocalSheetId = localSheetId,
                Text = formula
            });

            wb.Save();
        }

        internal void PrintArea(string a1Range)
        {
            if (string.IsNullOrWhiteSpace(a1Range))
                throw new ArgumentException("PrintArea A1 range is required.", nameof(a1Range));

            if (!AddressConverter.TryParseA1Range(a1Range, out var sx, out var sy, out var ex, out var ey))
                throw new ArgumentException("Invalid A1 range format. Example: \"A1:D20\".", nameof(a1Range));

            PrintArea(sx, sy, ex, ey); // 既存の int 版へ委譲
        }

        private string GetCurrentSheetName()
        {
            // SheetManager から現在インデックスを取らず、
            // Workbook -> Sheets から CurrentWorksheetPart を逆引き
            var sheets = _document.WorkbookPart.Workbook.Sheets
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>();

            var currentPart = _sheetManager.CurrentWorksheetPart;

            var sheet = sheets.FirstOrDefault(s =>
                _document.WorkbookPart.GetPartById(s.Id) == currentPart);

            if (sheet == null)
                throw new InvalidOperationException("Current sheet name could not be resolved.");

            return sheet.Name;
        }

        internal void RowDelete(int sy, int count)
        {
            Guards.EnsureOpened(_opened);

            if (sy < 0)
                throw new ArgumentOutOfRangeException(nameof(sy));
            if (count <= 0)
                throw new ArgumentOutOfRangeException(nameof(count));

            var wsPart = _sheetManager.CurrentWorksheetPart;
            var worksheet = wsPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            if (sheetData == null) return;

            // OpenXML RowIndex is 1-based
            uint startRow = (uint)(sy + 1);
            uint endRow = startRow + (uint)count - 1;

            // 1. Remove shared formulas in target rows (defensive)
            foreach (var cell in sheetData
                .Descendants<Cell>()
                .Where(c => c.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared))
            {
                // Shared formulas are fragile when rows are deleted.
                // We remove shared attributes and keep formula text if any.
                var f = cell.CellFormula;
                if (f != null)
                {
                    f.SharedIndex = null;
                    f.Reference = null;
                    f.FormulaType = null;
                }
            }

            // 2. Remove rows in [startRow, endRow]
            var rowsToRemove = sheetData.Elements<Row>()
                .Where(r => r.RowIndex != null &&
                            r.RowIndex.Value >= startRow &&
                            r.RowIndex.Value <= endRow)
                .ToList();

            foreach (var row in rowsToRemove)
                row.Remove();

            // 3. Shift rows below upward
            foreach (var row in sheetData.Elements<Row>())
            {
                if (row.RowIndex == null) continue;
                if (row.RowIndex.Value <= endRow) continue;

                var newIndex = row.RowIndex.Value - (uint)count;
                row.RowIndex.Value = newIndex;

                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference == null) continue;

                    AddressConverter.TryParseA1(
                        cell.CellReference.Value,
                        out var col,
                        out var _);

                    cell.CellReference.Value =
                        AddressConverter.ToA1(col, (int)newIndex);
                }
            }

            // 4. Remove calcChain (safe & recommended)
            RemoveCalcChain();

            worksheet.Save();

            // 5. Ensure recalculation on open
            EnsureRecalcOnOpen();
        }
        internal void SetCalculationMode(CalculationMode mode)
        {
            Guards.EnsureOpened(_opened);
            Guards.EnsureWorkbookPart(_document);

            var workbook = _document.WorkbookPart.Workbook;

            // 1. Ensure CalculationProperties exists
            var calcPr = workbook.CalculationProperties;
            if (calcPr == null)
            {
                calcPr = new CalculationProperties();
                workbook.Append(calcPr);
            }

            // 2. Apply mode
            switch (mode)
            {
                case CalculationMode.Manual:
                    calcPr.CalculationMode = CalculateModeValues.Manual;
                    calcPr.FullCalculationOnLoad = false;
                    break;

                case CalculationMode.Automatic:
                default:
                    calcPr.CalculationMode = CalculateModeValues.Auto;
                    calcPr.FullCalculationOnLoad = false;
                    break;
            }

            workbook.Save();
        }
        public void Dispose()
        {
            _document?.Dispose();
            _document = null;
            _opened = false;
        }

        internal void ExportSharedFormulasCsv(string outputPath)
        {
            Guards.EnsureOpened(_opened);

            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("outputPath is required.", nameof(outputPath));

            var sb = new StringBuilder();

            // Header
            sb.AppendLine("SheetName,Cell,Row,Col,SharedIndex,Formula,Reference");

            var wb = _document.WorkbookPart.Workbook;
            var sheets = wb.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();

            foreach (var sheet in sheets)
            {
                if (sheet.Id == null) continue;

                var worksheetPart = (WorksheetPart)_document.WorkbookPart.GetPartById(sheet.Id);
                var worksheet = worksheetPart.Worksheet;
                var sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) continue;

                var sheetName = sheet.Name?.Value ?? string.Empty;

                var sharedFormulaCells = sheetData
                    .Descendants<Cell>()
                    .Where(c => c.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared);

                foreach (var cell in sharedFormulaCells)
                {
                    var cellRef = cell.CellReference?.Value ?? string.Empty;

                    // parse column/row (A1 -> col, row)
                    int col = 0, row = 0;
                    if (!string.IsNullOrEmpty(cellRef))
                    {
                        AddressConverter.TryParseA1(cellRef, out col, out row);
                    }

                    var formulaText = cell.CellFormula?.Text ?? string.Empty;
                    var sharedIndexText = cell.CellFormula?.SharedIndex != null
                        ? cell.CellFormula.SharedIndex.Value.ToString()
                        : string.Empty;
                    var referenceText = cell.CellFormula?.Reference?.Value ?? string.Empty;

                    // CSV escape fields
                    sb.Append(EscapeCsv(sheetName)); sb.Append(',');
                    sb.Append(EscapeCsv(cellRef)); sb.Append(',');
                    sb.Append(row.ToString()); sb.Append(',');
                    sb.Append(col.ToString()); sb.Append(',');
                    sb.Append(EscapeCsv(sharedIndexText)); sb.Append(',');
                    sb.Append(EscapeCsv(formulaText)); sb.Append(',');
                    sb.AppendLine(EscapeCsv(referenceText));
                }
            }

            // Ensure directory exists
            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            // Write UTF-8 without BOM (consumer friendly)
            File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(false));
        }

        private void EnsureRecalcOnOpen()
        {
            var workbook = _document.WorkbookPart.Workbook;

            var calcPr = workbook.CalculationProperties ?? new CalculationProperties();

            // Excel を開いたとき必ず再計算
            calcPr.CalculationMode = CalculateModeValues.Auto;
            calcPr.FullCalculationOnLoad = true;

            workbook.CalculationProperties = calcPr;
            workbook.Save();
        }

        private void RemoveCalcChain()
        {
            var wbPart = _document.WorkbookPart;
            var calcChainPart = wbPart.CalculationChainPart;

            if (calcChainPart != null)
            {
                wbPart.DeletePart(calcChainPart);
            }
        }

        private Row GetOrCreateRow(SheetData sheetData, uint rowIndex)
        {
            // 1. 既存行を探す（RowIndex は r 属性）
            var row = sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == rowIndex);

            if (row != null)
                return row;

            // 2. なければ新規作成
            row = new Row { RowIndex = rowIndex };

            // 3. 正しい順序で挿入（昇順を保証）
            //    対象行より行番号が大きく、最も近い行を取得
            var refRow = sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value > rowIndex);

            if (refRow != null)
                sheetData.InsertBefore(row, refRow); // 近い行の前に挿入
            else
                sheetData.Append(row); // 対象行より大きい番号の行がないので、末尾に追加

            return row;
        }

        private static string EscapeCsv(string value)
        {
            if (value == null) return string.Empty;
            var needsQuote = value.Contains(',') || value.Contains('"') || value.Contains('\r') || value.Contains('\n');
            if (!needsQuote) return value;
            var escaped = value.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }


    }
}