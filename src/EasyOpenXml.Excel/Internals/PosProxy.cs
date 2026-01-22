using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyOpenXml.Excel.Models;
using System;
using System.Globalization;
using System.Linq;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class PosProxy
    {
        private readonly SpreadsheetDocument _document;
        private readonly WorksheetPart _worksheetPart;
        private readonly int _sx;
        private readonly int _sy;
        private readonly int _ex;
        private readonly int _ey;

        private readonly SharedStringManager _sharedStrings;

        internal PosProxy(
            SpreadsheetDocument document,
            WorksheetPart worksheetPart,
            int sx, int sy, int ex, int ey)
        {
            // 1. Validate arguments (assume 1-based coordinates: (1,1) = A1)
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (worksheetPart == null) throw new ArgumentNullException(nameof(worksheetPart));
            if (sx <= 0 || sy <= 0 || ex <= 0 || ey <= 0)
                throw new ArgumentOutOfRangeException("Coordinates must be 1-based and positive.");

            // 2. Normalize range
            _sx = Math.Min(sx, ex);
            _sy = Math.Min(sy, ey);
            _ex = Math.Max(sx, ex);
            _ey = Math.Max(sy, ey);

            _document = document;
            _worksheetPart = worksheetPart;

            _sharedStrings = new SharedStringManager(_document);
        }

        internal SpreadsheetDocument Document => _document;

        internal object GetValue()
        {
            // MVP: read only the top-left cell of the range
            var cell = GetOrCreateCell(_sx, _sy, create: false);
            if (cell == null) return null;

            // 1. Shared string
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (cell.CellValue == null) return string.Empty;

                if (int.TryParse(cell.CellValue.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sstIndex))
                {
                    return _sharedStrings.GetStringByIndexOrEmpty(sstIndex);
                }
                return string.Empty;
            }

            // 2. Boolean
            if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                return cell.CellValue?.Text == "1";
            }

            // 3. Number / DateTime (date format is not reliably detectable without style parsing)
            //    Here we return double if parsable; otherwise raw text.
            var raw = cell.CellValue?.Text;
            if (raw == null) return null;

            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                return d;

            return raw;
        }

        internal void SetValue(object value, bool isString)
        {
            // 1. Write to each cell in the range
            for (int row = _sy; row <= _ey; row++)
            {
                for (int col = _sx; col <= _ex; col++)
                {
                    var cell = GetOrCreateCell(col, row, create: true);
                    WriteCellValue(cell, value, isString);
                }
            }

            // 2. Save worksheet part (document save is handled by CloseBook(save:true))
            _worksheetPart.Worksheet.Save();
        }

        private void WriteCellValue(Cell cell, object value, bool isString)
        {
            // 1. Null clears the cell
            if (value == null)
            {
                cell.CellValue = null;
                cell.DataType = null;
                return;
            }

            // 2. Force string when requested OR when actual value is string
            if (isString || value is string)
            {
                var text = value?.ToString() ?? string.Empty;

                var index = _sharedStrings.GetOrAddString(text);

                cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                return;
            }

            // 3. Boolean
            if (value is bool b)
            {
                cell.CellValue = new CellValue(b ? "1" : "0");
                cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                return;
            }

            // 4. DateTime (Excel stores date as OADate double)
            if (value is DateTime dt)
            {
                // NOTE:
                // 1. Excel stores dates as OADate (double).
                // 2. To show it as a date in Excel UI, you must apply a date NumberFormat via Styles.
                // 3. MVP: write numeric OADate only.
                var oa = dt.ToOADate();
                cell.CellValue = new CellValue(oa.ToString(CultureInfo.InvariantCulture));
                cell.DataType = null; // numeric
                return;
            }

            // 5. Numeric types (int/long/float/double/decimal etc.)
            if (value is byte || value is sbyte ||
                value is short || value is ushort ||
                value is int || value is uint ||
                value is long || value is ulong ||
                value is float || value is double ||
                value is decimal)
            {
                var text = Convert.ToString(value, CultureInfo.InvariantCulture);
                cell.CellValue = new CellValue(text);
                cell.DataType = null; // numeric
                return;
            }

            // 6. Fallback: write as string (safe default)
            var fallback = value.ToString() ?? string.Empty;
            var fallbackIndex = _sharedStrings.GetOrAddString(fallback);

            cell.CellValue = new CellValue(fallbackIndex.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        private Cell GetOrCreateCell(int col, int row, bool create)
        {
            // 1. Prepare SheetData
            var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                if (!create) return null;
                sheetData = _worksheetPart.Worksheet.AppendChild(new SheetData());
            }

            // 2. Find or create Row
            var rowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)row);
            if (rowElement == null)
            {
                if (!create) return null;

                rowElement = new Row { RowIndex = (uint)row };

                // Insert row keeping order by RowIndex (avoids Excel repair)
                var refRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value > (uint)row);
                if (refRow != null) sheetData.InsertBefore(rowElement, refRow);
                else sheetData.AppendChild(rowElement);
            }

            // 3. Find or create Cell by CellReference (e.g., "A1")
            var cellRef = AddressConverter.ToA1(col, row);
            var cell = rowElement.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));

            if (cell == null)
            {
                if (!create) return null;

                cell = new Cell { CellReference = cellRef };

                // Insert cell keeping order by column (avoids Excel repair)
                var refCell = rowElement.Elements<Cell>()
                    .FirstOrDefault(c => CompareCellReference(c.CellReference?.Value, cellRef) > 0);

                if (refCell != null) rowElement.InsertBefore(cell, refCell);
                else rowElement.AppendChild(cell);
            }

            return cell;
        }

        private static int CompareCellReference(string a, string b)
        {
            // 1. Compare by column index first, then row index
            if (string.IsNullOrEmpty(a)) return -1;
            if (string.IsNullOrEmpty(b)) return 1;

            AddressConverter.TryParseA1(a, out var aCol, out var aRow);
            AddressConverter.TryParseA1(b, out var bCol, out var bRow);

            var c = aCol.CompareTo(bCol);
            return c != 0 ? c : aRow.CompareTo(bRow);
        }

        internal void ApplyStyle(uint styleIndex)
        {
            for (int row = _sy; row <= _ey; row++)
            {
                for (int col = _sx; col <= _ex; col++)
                {
                    var cell = GetOrCreateCell(col, row, create: true);
                    cell.StyleIndex = styleIndex;
                }
            }

            _worksheetPart.Worksheet.Save();
        }

        internal CellSnapshot CaptureSnapshot()
        {
            var cell = GetOrCreateCell(_sx, _sy, create: false);

            if (cell == null)
            {
                return new CellSnapshot
                {
                    Value = null,
                    IsString = false,
                    StyleIndex = 0
                };
            }

            return new CellSnapshot
            {
                Value = GetValue(),
                IsString = cell.DataType?.Value == CellValues.SharedString,
                StyleIndex = cell.StyleIndex?.Value ?? 0
            };
        }

        internal void ApplySnapshot(CellSnapshot snapshot)
        {
            // 1. Apply value
            SetValue(snapshot.Value, snapshot.IsString);

            // 2. Apply style
            if (snapshot.StyleIndex != 0)
            {
                ApplyStyle(snapshot.StyleIndex);
            }
        }

        internal void Merge()
        {
            // 1. Single cell range does not need merge
            if (_sx == _ex && _sy == _ey)
                return;

            var worksheet = _worksheetPart.Worksheet;

            // 2. Ensure MergeCells element exists
            var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();

                // Insert MergeCells after SheetData (Excel repair safe position)
                var sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData != null)
                    worksheet.InsertAfter(mergeCells, sheetData);
                else
                    worksheet.Append(mergeCells);
            }

            // 3. Build merge range reference (e.g. "A1:C3")
            var from = AddressConverter.ToA1(_sx, _sy);
            var to = AddressConverter.ToA1(_ex, _ey);
            var refText = $"{from}:{to}";

            // 4. Avoid duplicate merge entries (MVP: simple check)
            var exists = mergeCells.Elements<MergeCell>()
                .Any(m => string.Equals(m.Reference?.Value, refText, StringComparison.OrdinalIgnoreCase));

            if (!exists)
            {
                mergeCells.Append(new MergeCell { Reference = refText });
            }

            // 5. Save worksheet
            worksheet.Save();
        }

        internal void Merge(
            HorizontalAlign horizontal,
            VerticalAlign vertical,
            bool wrapText)
        {
            // 1. Merge cells (existing behavior)
            Merge();

            // 2. Apply alignment via StyleManager
            var styleManager = new StyleManager(_document);

            // 左上セルの既存 StyleIndex をベースにする
            var baseCell = GetOrCreateCell(_sx, _sy, create: true);
            var baseStyleIndex = baseCell.StyleIndex?.Value ?? 0U;

            var newStyleIndex = styleManager.GetOrCreateAlignmentStyle(
                baseStyleIndex,
                horizontal,
                vertical,
                wrapText);

            // 3. Apply style to merged range
            ApplyStyle(newStyleIndex);
        }

    }
}
