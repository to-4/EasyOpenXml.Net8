using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyOpenXml.Excel.Models;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class ExcelInternal : IDisposable
    {
        private SpreadsheetDocument _document;
        private SheetManager _sheetManager;
        private bool _opened;
        private CellSnapshot _clipboard;

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

            // 1. Remove rows in [startRow, endRow]
            var rowsToRemove = sheetData.Elements<Row>()
                .Where(r => r.RowIndex != null &&
                            r.RowIndex.Value >= startRow &&
                            r.RowIndex.Value <= endRow)
                .ToList();

            foreach (var row in rowsToRemove)
                row.Remove();

            // 2. Shift rows below upward
            foreach (var row in sheetData.Elements<Row>())
            {
                if (row.RowIndex == null) continue;
                if (row.RowIndex.Value <= endRow) continue;

                var newIndex = row.RowIndex.Value - (uint)count;
                row.RowIndex.Value = newIndex;

                // 3. Update CellReference for each cell in the row
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

            worksheet.Save();
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
    }
}