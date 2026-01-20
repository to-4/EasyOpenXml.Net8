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

            // 2. 既存の Print_Area はすべて削除（安全第一）
            var old = definedNames.Elements<DefinedName>()
                .Where(d => d.Name?.Value == "_xlnm.Print_Area")
                .ToList();

            foreach (var d in old)
                d.Remove();

            // 3. 正しい式を作る
            var sheetName = _sheetManager.CurrentSheetName.Replace("'", "''");
            var area = $"{AddressConverter.ToAbsoluteA1(sx, sy)}:{AddressConverter.ToAbsoluteA1(ex, ey)}";
            var formula = $"'{sheetName}'!{area}";

            // 4. 新規に 1 件だけ追加（LocalSheetId なし）
            definedNames.Append(new DefinedName
            {
                Name = "_xlnm.Print_Area",
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

        public void Dispose()
        {
            _document?.Dispose();
            _document = null;
            _opened = false;
        }
    }
}