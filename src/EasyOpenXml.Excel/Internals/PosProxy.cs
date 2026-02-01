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

        /// <summary>
        /// 列のスタイルインデックステーブル
        /// <para>コンストラクタ時に呼び出し</para>
        /// <para>ベーススタイルID取得時の使用</para>
        /// </summary>
        private readonly ColumnStyleIndexMap _columnStyleMap;

        private readonly int _sx;
        private readonly int _sy;
        private readonly int _ex;
        private readonly int _ey;

        private readonly SharedStringManager _sharedStrings;

        private uint _templateStyleIndex = 0; // 最後の逃げ（Normal など）。必要なら差し替え。

        // 列スタイルのキャッシュ（Min/Max を展開した辞書）
        private Dictionary<int, uint?>? _colStyleCache;

        internal PosProxy(
            SpreadsheetDocument document,
            WorksheetPart worksheetPart,
            int sx, int sy, int ex, int ey)
        {
            // 1. 引数を検証
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (worksheetPart == null) throw new ArgumentNullException(nameof(worksheetPart));
            if (sx <= 0 || sy <= 0 || ex <= 0 || ey <= 0)
                throw new ArgumentOutOfRangeException("Coordinates must be 1-based and positive.");

            // 2. 範囲指定
            _sx = Math.Min(sx, ex);
            _sy = Math.Min(sy, ey);
            _ex = Math.Max(sx, ex);
            _ey = Math.Max(sy, ey);

            _document = document;
            _worksheetPart = worksheetPart;
            _sharedStrings = new SharedStringManager(_document);

            // 列のスタイルIDテーブルを設定
            _columnStyleMap = new ColumnStyleIndexMap(_worksheetPart.Worksheet);
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

        /// <summary>
        /// セルを取得。なければ作成。
        /// <para>新規作成時のみ StyleIndex を継承（行→列→左→上→テンプレ）</para>
        /// <para>既存セルの StyleIndex は変更しない</para>
        /// </summary>
        internal Cell GetOrCreateCellWithStyle(int rowIndex, int columnIndex, bool create = true)
        {

            var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                if (!create) return null;
                sheetData = _worksheetPart.Worksheet.AppendChild(new SheetData());
            }

            if (sheetData == null) throw new ArgumentNullException(nameof(sheetData));
            if (rowIndex < 1) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            if (columnIndex < 1) throw new ArgumentOutOfRangeException(nameof(columnIndex));

            uint r = (uint)rowIndex;
            string cellRef = AddressConverter.ToA1(columnIndex, rowIndex);

            // 1) Row を取得/作成
            var row = GetOrCreateRow(sheetData, r);

            // 2) 既存セルがあればそれを返す（StyleIndex は一切触らない）
            var existing = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
            if (existing != null) return existing;

            // 3) 新規セル作成
            var newCell = new Cell { CellReference = cellRef };

            // 4) 新規セルの StyleIndex を推定（行→列→左→上→テンプレ）
            uint? style = InferStyleIndexForNewCell(sheetData, row, r, columnIndex);
            if (style.HasValue)
                newCell.StyleIndex = style.Value;

            // 5) セルを列順で挿入（A, B, C... の順。安全）
            InsertCellInOrder(row, newCell);

            return newCell;
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

            // #########################
            // pending 2026/02/01
            //var newStyleIndex = styleManager.GetOrCreateAlignmentStyle(
            //    baseStyleIndex,
            //    horizontal,
            //    vertical,
            //    wrapText);

            //// 3. Apply style to merged range
            //ApplyStyle(newStyleIndex);
            // #########################
        }


        private Row GetOrCreateRow(SheetData sheetData, uint rowIndex)
        {
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            if (row != null) return row;

            row = new Row { RowIndex = rowIndex };

            // RowIndex 昇順に挿入（整っていた方が安全）
            var refRow = sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value > rowIndex);

            if (refRow != null) sheetData.InsertBefore(row, refRow);
            else sheetData.AppendChild(row);

            return row;
        }

        private uint? InferStyleIndexForNewCell(SheetData sheetData, Row row, uint rowIndex, int colIndex)
        {
            // ① 行スタイル
            if (row.StyleIndex != null)
                return row.StyleIndex.Value;

            // ② 列スタイル
            var colStyle = GetColumnStyleIndex(colIndex);
            if (colStyle.HasValue)
                return colStyle.Value;

            // ③ 左隣セル（同じ行の colIndex-1）
            var left = FindCellInRowByColumnIndex(row, colIndex - 1);
            if (left?.StyleIndex != null)
                return left.StyleIndex.Value;

            // ④ 上セル（rowIndex-1 の同じ列）
            if (rowIndex > 1)
            {
                var upperRow = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex?.Value == rowIndex - 1);
                if (upperRow != null)
                {
                    var upper = FindCellInRowByColumnIndex(upperRow, colIndex);
                    if (upper?.StyleIndex != null)
                        return upper.StyleIndex.Value;

                    // 上の行に Row.StyleIndex があるなら、それを使う案もある（好み）
                    // if (upperRow.StyleIndex != null) return upperRow.StyleIndex.Value;
                }
            }

            // ⑤ テンプレート（既定）スタイル
            // null を返すと StyleIndex 未指定＝標準（0）になるので、
            // 「必ずテンプレ値を付けたい」なら _templateStyleIndex を返す。
            return _templateStyleIndex;
        }

        private uint? GetColumnStyleIndex(int colIndex1Based)
        {
            _colStyleCache ??= BuildColumnStyleCache();
            return _colStyleCache.TryGetValue(colIndex1Based, out var s) ? s : null;
        }

        private Dictionary<int, uint?> BuildColumnStyleCache()
        {
            var dic = new Dictionary<int, uint?>();

            // Columns は Worksheet 直下にある（存在しないテンプレもある）

            var cols = _worksheetPart.Worksheet.Elements<Columns>().FirstOrDefault();
            if (cols == null) return dic;

            foreach (var col in cols.Elements<Column>())
            {
                if (col.Min == null || col.Max == null) continue;
                uint? style = col.Style?.Value;

                for (uint c = col.Min.Value; c <= col.Max.Value; c++)
                    dic[(int)c] = style;
            }

            return dic;
        }

        private static Cell? FindCellInRowByColumnIndex(Row row, int colIndex1Based)
        {
            if (row == null) return null;
            if (colIndex1Based < 1) return null;

            string colName = AddressConverter.ToColumnName(colIndex1Based);

            foreach (var c in row.Elements<Cell>())
            {
                var r = c.CellReference?.Value;
                if (string.IsNullOrEmpty(r)) continue;

                // "BC12" → "BC"
                string existingCol = new string(r.TakeWhile(char.IsLetter).ToArray());
                if (string.Equals(existingCol, colName, StringComparison.Ordinal))
                    return c;
            }

            return null;
        }

        private static void InsertCellInOrder(Row row, Cell newCell)
        {
            // 列順で挿入したいので、次に大きい列参照の直前へ
            var newRef = newCell.CellReference?.Value;
            if (string.IsNullOrEmpty(newRef))
            {
                row.AppendChild(newCell);
                return;
            }

            var refCell = row.Elements<Cell>()
                .FirstOrDefault(c =>
                {
                    var r = c.CellReference?.Value;
                    if (string.IsNullOrEmpty(r)) return false;
                    return CompareCellRefByColumn(r, newRef) > 0;
                });

            if (refCell != null) row.InsertBefore(newCell, refCell);
            else row.AppendChild(newCell);
        }

        private static int CompareCellRefByColumn(string aRef, string bRef)
        {
            // "AA10" の "AA" だけを取り出して比較（行番号は無視）
            string aCol = new string(aRef.TakeWhile(char.IsLetter).ToArray());
            string bCol = new string(bRef.TakeWhile(char.IsLetter).ToArray());
            return string.Compare(aCol, bCol, StringComparison.Ordinal);
        }


        /// <summary>
        /// カラーオブジェクトを、ARGB 16進（例: "FFFF0000"）に変換します
        /// </summary>
        /// <param name="color">カラーオブジェクト</param>
        /// <returns>ARGB 16進 (例: @"FFFF0000") 文字列を返します</returns>
        private static string ToArgbHex(System.Drawing.Color color)
        {
            // AARRGGBB
            return color.A.ToString("X2", CultureInfo.InvariantCulture)
                 + color.R.ToString("X2", CultureInfo.InvariantCulture)
                 + color.G.ToString("X2", CultureInfo.InvariantCulture)
                 + color.B.ToString("X2", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// PosProxy で管理している範囲の、各セルデータ（Cell オブジェクト）を一つずつ返します
        /// <para>列挙メソッド</para>
        /// </summary>
        /// <param name="create">作成フラグ（true の場合、null であれば新規作成）</param>
        /// <returns>セルデータ（Cell オブジェクト）をひとつずつ返す</returns>
        internal System.Collections.Generic.IEnumerable<Cell> EnumerateTargetCells(bool create)
        {
            // 1. SheetData を準備
            var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                if (!create)
                {
                    yield break; // 列挙メソッドを終了
                }

                sheetData = _worksheetPart.Worksheet.AppendChild(new SheetData());
            }

            // 2. 範囲内の各セルデータ（Cell オブジェクト）を返す
            for (int row = _sy; row <= _ey; row++)
            {
                for (int col = _sx; col <= _ex; col++)
                {
                    // セルオブジェクトを取得
                    var cell = GetOrCreateCell(col, row, create);
                    if (cell != null)
                    {
                        // セルオブジェクトを返し、次に呼ばれるまでストップ
                        yield return cell;
                    }
                }
            }
        }

        /// <summary>
        /// 指定されたセルオブジェクトを基に、セル書式用のスタイルIDを取得します
        /// <para>下記の優先順位で取得</para>
        /// <para>1. セルが持つ書式情報のスタイルID</para>
        /// <para>2. 列が持つ書式情報のスタイルID</para>
        /// <para>3. 行が持つ書式情報のスタイルID</para>
        /// <para>4. ブックがテンプレートの書式情報のスタイルID</para>
        /// </summary>
        /// <param name="cell">セルオブジェクト</param>
        /// <returns>セル書式用のスタイルID</returns>
        internal uint ResolveBaseStyleIndex(Cell cell)
        {
            if (cell == null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            // 1. セルが持つ書式情報のスタイルID
            if (cell.StyleIndex != null)
            {
                return cell.StyleIndex.Value;
            }

            // 2. セル番地（行番号、列番号）を取得
            var a1 = cell.CellReference?.Value;
            if (string.IsNullOrEmpty(a1))
            {
                return 0U; // セル番地無し
            }
            if (!AddressConverter.TryParseA1(a1, out var col, out var row))
            {
                return 0U; // セル番地が不正
            }

            // 3. Column style (Column.Style)
            var colStyle = _columnStyleMap.TryGetStyleIndex(col);
            if (colStyle.HasValue)
                return colStyle.Value;

            // 3) Row style (Row.StyleIndex)
            var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData != null)
            {
                var rowElem = sheetData.Elements<Row>()
                    .FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)row);

                if (rowElem != null && rowElem.StyleIndex != null)
                    return rowElem.StyleIndex.Value;
            }


            // 5) Default style
            return 0U;
        }

        /// <summary>
        /// 列情報スタイルIDテーブルクラス
        /// <para>セルのベーススタイルIDの取得時に使用</para>
        /// </summary>
        private sealed class ColumnStyleIndexMap
        {
            private readonly Span[] _spans;

            internal ColumnStyleIndexMap(Worksheet ws)
            {
                var columns = ws.Elements<Columns>().FirstOrDefault();
                if (columns == null)
                {
                    _spans = Array.Empty<Span>();
                    return;
                }

                // Columns は重複/上書きがあり得るので、定義順を保持する（後勝ちにする）
                var list = new System.Collections.Generic.List<Span>();

                foreach (var c in columns.Elements<Column>())
                {
                    if (c.Style == null) continue;

                    int min = (int)(c.Min?.Value ?? 1);
                    int max = (int)(c.Max?.Value ?? (uint)min);
                    uint style = (uint)c.Style.Value;

                    list.Add(new Span(min, max, style));
                }

                _spans = list.ToArray();
            }

            internal uint? TryGetStyleIndex(int col)
            {
                // 後勝ち（後から定義された Column を優先）
                for (int i = _spans.Length - 1; i >= 0; i--)
                {
                    var s = _spans[i];
                    if (s.Min <= col && col <= s.Max)
                        return s.Style;
                }
                return null;
            }

            private readonly struct Span
            {
                public readonly int Min;
                public readonly int Max;
                public readonly uint Style;

                public Span(int min, int max, uint style)
                {
                    Min = min;
                    Max = max;
                    Style = style;
                }
            }
        }

        /// <summary>
        /// 指定された背景色を各セルの書式に設定します
        /// </summary>
        /// <param name="styleManager">スタイル管理オブジェクト</param>
        /// <param name="color">背景色</param>
        internal void SetBackColor(StyleManager styleManager, System.Drawing.Color color)
        {

            // 1. 背景色を取得
            var rgb = ToArgbHex(color);

            // 2. 塗りつぶし仕様データを取得
            var fill = new FillSpec(rgb);

            // 局所的なスタイルテーブル初期化
            var localCache = new Dictionary<uint, uint>();

            // 範囲内のセルを反復
            foreach (var cell in EnumerateTargetCells(create: true))
            {
                // スタイルインデックスを取得
                uint baseStyleIndex = ResolveBaseStyleIndex(cell);

                if (!localCache.TryGetValue(baseStyleIndex, out var newStyleIndex))
                {
                    // 局所的なスタイルテーブルには無し
                    // 新しいスタイルID を取得 ※ 引数の背景色を反映
                    newStyleIndex = styleManager.GetOrCreateStyle(
                        baseStyleIndex: baseStyleIndex,
                        fill: fill);

                    localCache.Add(baseStyleIndex, newStyleIndex);
                }

                // スタイルIDにセット
                cell.StyleIndex = newStyleIndex;
            }

            _worksheetPart.Worksheet.Save();
        }


    }
}
