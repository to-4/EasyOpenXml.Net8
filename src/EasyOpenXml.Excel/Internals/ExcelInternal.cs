using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using EasyOpenXml.Excel.Models;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class ExcelInternal : IDisposable
    {
        private SpreadsheetDocument _document;
        private SheetManager _sheetManager;
        private bool _opened;

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

        public void Dispose()
        {
            _document?.Dispose();
            _document = null;
            _opened = false;
        }
    }
}