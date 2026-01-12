using System;
using System.Collections.Generic;
using EasyOpenXml.Excel.Internals;

namespace EasyOpenXml.Excel
{
    public sealed class ExcelDocument : IDisposable
    {
        private ExcelInternal _internal = new ExcelInternal();
        private bool _disposed;

        public void InitializeFile(string path1, string path2 = "")
        {
            _internal = new ExcelInternal();

            var result = _internal.OpenBook(path1, path2);
            if (result < 0)
            {
                throw new ExcelDocumentException("Failed to open Excel file.");
            }
        }

        public void SheetSelect(int sheetId)
        {
            EnsureNotDisposed();
            _internal.SheetNo = sheetId;
        }

        public IReadOnlyList<string> SheetNames
        {
            get
            {
                EnsureNotDisposed();
                return _internal.SheetNames;
            }
        }

        public void SetValue(int sx, int sy, object value)
        {
            SetValue(sx, sy, sx, sy, value);
        }
        public void SetValue(int sx, int sy, int ex, int ey, object value)
        {
            EnsureNotDisposed();
            var pos = _internal.Pos(sx, sy, ex, ey);
            pos.Value = value;
        }
        public void SetValue(string cell, object value)
        {
            EnsureNotDisposed();
            var c = _internal.Cell(cell);
            c.Value = value;
        }
        public void SetValue(string cell, int cx, int cy, object value)
        {
            EnsureNotDisposed();
            var c = _internal.Cell(cell, cx, cy);
            c.Value = value;
        }

        public object GetValue(int sx, int sy)
        {
            EnsureNotDisposed();

            // 1. Use internal Pos API
            var pos = _internal.Pos(sx, sy);

            // 2. Delegate to PosProxy.GetValue()
            return pos.Value;
        }
        public object GetValue(string cell)
        {
            EnsureNotDisposed();
            var c = _internal.Cell(cell);
            return c.Value;
        }

        public void FinalizeFile(bool save = true)
        {
            EnsureNotDisposed();
            _internal.CloseBook(save);
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ExcelDocument));
        }

        public void Dispose()
        {
            if (_disposed) return;
            _internal?.Dispose();
            _disposed = true;
        }
    }
}
