using System;
using DocumentFormat.OpenXml.Packaging;
using EasyOpenXml.Excel.Models;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class CellWrapperProxy
    {
        private readonly PosProxy _posProxy;

        internal CellWrapperProxy(
            SpreadsheetDocument document,
            WorksheetPart worksheetPart,
            string cell,
            int cx,
            int cy)
        {
            if (!AddressConverter.TryParseA1(cell, out var col, out var row))
                throw new ArgumentException("Invalid A1 cell reference.", nameof(cell));

            // cx, cy are offsets
            var ex = col + cx;
            var ey = row + cy;

            _posProxy = new PosProxy(document, worksheetPart, col, row, ex, ey);
        }

        internal object GetValue()
        {
            return _posProxy.GetValue();
        }

        internal void SetValue(object value, bool isString)
        {
            _posProxy.SetValue(value, isString);
        }
    }
}
