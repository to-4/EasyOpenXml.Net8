using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class StyleManager
    {

        /// <summary>
        /// ブック全体の管理
        /// <para>SpreadSheetDocument のプロパティ</para>
        /// </summary>
        private readonly WorkbookPart _workbookPart;

        /// <summary>
        /// ブック全体の“スタイル定義”を保持する Part 
        /// <para>SpreadSheetDocument - WorkbookPart の WorkbookStylesPart プロパティ</para>
        /// </summary>
        private readonly WorkbookStylesPart _stylesPart;

        /// <summary>
        /// ブック全体で共有されるスタイル定義
        /// <para>WorkbookStylesPart のプロパティ</para>
        /// <para>※ 実際のファイルは、xl\styles.xml</para>
        /// </summary>
        private readonly Stylesheet _ss;

        /// <summary>
        /// FontSpec の ID テーブル (key: FontSpec, value: fontId)
        /// <para>※ fontId は、xl/styles.xml の ＜fonts＞ 要素内の ＜border＞ 要素の並び順（0 から始まる）</para>
        /// </summary>
        private readonly Dictionary<FontSpec,   uint> _fontCache   = new Dictionary<FontSpec, uint>();
        /// <summary>
        /// FillSpec の ID テーブル (key: FillSpec, value: fillId)
        /// <para>※ fillId は、xl/styles.xml の ＜fills＞ 要素内の <fill> 要素の並び順（0 から始まる）</para>
        /// </summary>
        private readonly Dictionary<FillSpec,   uint> _fillCache   = new Dictionary<FillSpec, uint>();
        /// <summary>
        /// BorderSpec の ID テーブル (key: BorderSpec, value: borderId)
        /// <para>※ borderId は、xl/styles.xml の ＜borders＞ 要素内の ＜border＞ 要素の並び順（0 から始まる）</para>
        /// </summary>
        private readonly Dictionary<BorderSpec, uint> _borderCache = new Dictionary<BorderSpec, uint>();
        /// <summary>
        /// StyleKey の ID テーブル (key: StyleKey, value: styleId)
        /// <para>※1. styleId は、xl/styles.xml の ＜cellXfs＞要素（cellXfs: Cell eXtended Formats）内の ＜xf＞ 要素の上からの順番（0 から始まる）</para>
        /// <para>※2. sheet1.xml などの各セルの　s　属性の値と styleId が紐づきます</para>
        /// <para>※3. xl/styles.xml の ＜cellStyleXfs＞ 要素（名前付きスタイルの元）が名称が似ていますが、当テーブルとは異なります</para>
        /// </summary>
        private readonly Dictionary<StyleKey,   uint> _styleCache  = new Dictionary<StyleKey, uint>();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="document">スプレッドシート ドキュメント</param>
        internal StyleManager(SpreadsheetDocument document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            _workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null.");

            // 2. Ensure styles part exists
            _stylesPart = _workbookPart.WorkbookStylesPart ?? _workbookPart.AddNewPart<WorkbookStylesPart>();
            if (_stylesPart.Stylesheet == null)
            {
                _stylesPart.Stylesheet = CreateDefaultStylesheet();
            }

            _ss = _stylesPart.Stylesheet;

            // 3. Ensure required collections exist
            EnsureCollections(_ss);

            // 4. Optional: warm up caches by scanning existing items (safe; can be omitted)
            WarmUpCaches();
        }

        /// <summary>
        /// スタイルIDを取得します。または作成します。
        /// <para>スタイルIDは、CellFormats に含まれる CellFormat を識別</para>
        /// </summary>
        /// <param name="baseStyleIndex">
        /// 基本的なスタイルID（指定セルの元のスタイルID）
        /// <para>※ 引数の font 等の情報が反映される前のスタイルID</para>
        /// </param>
        /// <param name="font">フォント仕様情報</param>
        /// <param name="fill">塗りつぶし仕様情報</param>
        /// <param name="border">罫線仕様情報</param>
        /// <param name="alignment">配置仕様情報</param>
        /// <returns>スタイルID</returns>
        internal uint GetOrCreateStyle(
            uint baseStyleIndex,
            FontSpec?   font = null,
            FillSpec?   fill = null,
            BorderSpec? border = null,
            AlignmentSpec? alignment = null)
        {
            // 1. スタイルテーブルから取得
            // - スタイルキー情報（フォント、網掛け、罫線、配置情報を含む）を取得
            var key = new StyleKey(baseStyleIndex, font, fill, border, alignment);
            uint cached;
            if (_styleCache.TryGetValue(key, out cached))
                return cached;

            // 2. 元のセルスタイル
            var cellFormats = _ss.CellFormats;
            var baseCf = (CellFormat)cellFormats.ChildElements[(int)baseStyleIndex];

            // 3. 元のセルスタイルからクローンを取得
            var newCf = (CellFormat)baseCf.CloneNode(true);

            // 4. フォントID を設定
            if (font.HasValue)
            {
                var fontId = GetOrCreateFontId(font.Value);
                newCf.FontId = fontId;
                newCf.ApplyFont = true;
            }

            // 5. 塗りつぶしID を設定
            if (fill.HasValue)
            {
                var fillId = GetOrCreateFillId(fill.Value);
                newCf.FillId = fillId;
                newCf.ApplyFill = true;
            }

            // 6. 罫線ID を設定
            if (border.HasValue)
            {
                var borderId = GetOrCreateBorderId(border.Value);
                newCf.BorderId = borderId;
                newCf.ApplyBorder = true;
            }

            // 7. 配置仕様を設定
            if (alignment.HasValue)
            {
                if (newCf.Alignment == null)
                    newCf.Alignment = new Alignment();

                var a = alignment.Value;

                if (a.H.HasValue)
                    newCf.Alignment.Horizontal = a.H.Value;

                if (a.V.HasValue)
                    newCf.Alignment.Vertical = a.V.Value;

                // 折り返し設定
                newCf.Alignment.WrapText = a.WrapText;

                newCf.ApplyAlignment = true;
            }

            // 8. セル書式リストの末尾に追加
            cellFormats.AppendChild(newCf);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            // ブック全体で共有されるスタイル定義を保存
            _ss.Save();

            // 9. セル書式リストの末尾の要素番号を取得
            var newIndex = (uint)(cellFormats.ChildElements.Count - 1);
            _styleCache[key] = newIndex;

            return newIndex;
        }

        // -------------------------
        // Fonts
        // -------------------------
        private uint GetOrCreateFontId(FontSpec spec)
        {
            uint id;
            if (_fontCache.TryGetValue(spec, out id))
                return id;

            var fonts = _ss.Fonts;

            // 1. Build Font
            var f = new Font();

            if (!string.IsNullOrEmpty(spec.Name))
                f.Append(new FontName { Val = spec.Name });

            if (spec.Size > 0)
                f.Append(new FontSize { Val = spec.Size });

            if (spec.Bold) f.Append(new Bold());
            if (spec.Italic) f.Append(new Italic());
            if (spec.Underline) f.Append(new Underline());

            // 2. Color
            if (!string.IsNullOrEmpty(spec.Rgb) || spec.Theme.HasValue || spec.Indexed.HasValue)
            {
                var c = new Color();
                if (!string.IsNullOrEmpty(spec.Rgb))
                    c.Rgb = new HexBinaryValue(spec.Rgb);

                if (spec.Theme.HasValue)
                    c.Theme = (uint)spec.Theme.Value;

                if (spec.Indexed.HasValue)
                    c.Indexed = (uint)spec.Indexed.Value;

                f.Append(c);
            }

            // 3. Append
            fonts.AppendChild(f);
            fonts.Count = (uint)fonts.ChildElements.Count;

            _ss.Save();

            id = (uint)(fonts.ChildElements.Count - 1);
            _fontCache[spec] = id;
            return id;
        }

        // -------------------------
        // Fills
        // -------------------------
        private uint GetOrCreateFillId(FillSpec spec)
        {
            uint id;
            if (_fillCache.TryGetValue(spec, out id))
                return id;

            var fills = _ss.Fills;

            // 1. "No fill" / default handling
            if (string.IsNullOrEmpty(spec.Rgb))
            {
                // Use existing "none" fill (usually index 0 or 1 depending on template)
                // We do NOT assume fixed; instead we create a PatternFill(None) if needed.
            }

            // 2. Build solid fill
            var fill = new Fill(
                new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = new HexBinaryValue(spec.Rgb) },
                    // BackgroundColor is typically not required for solid
                });

            // 3. Append
            fills.AppendChild(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            _ss.Save();

            id = (uint)(fills.ChildElements.Count - 1);
            _fillCache[spec] = id;
            return id;
        }

        // -------------------------
        // Borders
        // -------------------------
        private uint GetOrCreateBorderId(BorderSpec spec)
        {
            uint id;
            if (_borderCache.TryGetValue(spec, out id))
                return id;

            var borders = _ss.Borders;

            // 1. Build Border
            var b = new Border
            {
                LeftBorder     = BuildEdgeLeft(spec.Left),
                RightBorder    = BuildEdgeRight(spec.Right),
                TopBorder      = BuildEdgeTop(spec.Top),
                BottomBorder   = BuildEdgeBottom(spec.Bottom),
                DiagonalBorder = new DiagonalBorder()
            };

            // 2. Append
            borders.AppendChild(b);
            borders.Count = (uint)borders.ChildElements.Count;

            _ss.Save();

            id = (uint)(borders.ChildElements.Count - 1);
            _borderCache[spec] = id;
            return id;
        }

        private static LeftBorder BuildEdgeLeft(BorderEdgeSpec e)
        {
            // 1. Map our enum to OpenXml BorderStyleValues
            BorderStyleValues? style = ToBorderStyle(e.Line);

            var edge = new LeftBorder();
            if (style.HasValue)
                edge.Style = style.Value;

            // 2. Color
            if (!string.IsNullOrEmpty(e.Rgb))
                edge.Append(new Color { Rgb = new HexBinaryValue(e.Rgb) });

            return edge;
        }

        // Overloads for other edges
        private static RightBorder BuildEdgeRight(BorderEdgeSpec e)
        {
            BorderStyleValues? style = ToBorderStyle(e.Line);

            var edge = new RightBorder();
            if (style.HasValue)
                edge.Style = style.Value;

            if (!string.IsNullOrEmpty(e.Rgb))
                edge.Append(new Color { Rgb = new HexBinaryValue(e.Rgb) });

            return edge;
        }

        private static TopBorder BuildEdgeTop(BorderEdgeSpec e)
        {
            BorderStyleValues? style = ToBorderStyle(e.Line);

            var edge = new TopBorder();
            if (style.HasValue)
                edge.Style = style.Value;

            if (!string.IsNullOrEmpty(e.Rgb))
                edge.Append(new Color { Rgb = new HexBinaryValue(e.Rgb) });

            return edge;
        }

        private static BottomBorder BuildEdgeBottom(BorderEdgeSpec e)
        {
            BorderStyleValues? style = ToBorderStyle(e.Line);

            var edge = new BottomBorder();
            if (style.HasValue)
                edge.Style = style.Value;

            if (!string.IsNullOrEmpty(e.Rgb))
                edge.Append(new Color { Rgb = new HexBinaryValue(e.Rgb) });

            return edge;
        }

        private static BorderStyleValues? ToBorderStyle(BorderLine line)
        {
            switch (line)
            {
                case BorderLine.None: return null;
                case BorderLine.Thin: return BorderStyleValues.Thin;
                case BorderLine.Medium: return BorderStyleValues.Medium;
                case BorderLine.Thick: return BorderStyleValues.Thick;
                case BorderLine.Dashed: return BorderStyleValues.Dashed;
                case BorderLine.Dotted: return BorderStyleValues.Dotted;
                case BorderLine.Double: return BorderStyleValues.Double;
                default: return null;
            }
        }

        // -------------------------
        // Helpers
        // -------------------------
        private static Stylesheet CreateDefaultStylesheet()
        {
            // 1. Minimal stylesheet with required collections
            var ss = new Stylesheet
            {
                Fonts = new Fonts { Count = 1U, KnownFonts = true },
                Fills = new Fills { Count = 2U },
                Borders = new Borders { Count = 1U },
                CellStyleFormats = new CellStyleFormats { Count = 1U },
                CellFormats = new CellFormats { Count = 1U }
            };

            // 2. Default font
            ss.Fonts.AppendChild(new Font(
                new FontSize { Val = 11D },
                new FontName { Val = "Calibri" }));

            // 3. Default fills (required by Excel)
            ss.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.None }));
            ss.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));

            // 4. Default border
            ss.Borders.AppendChild(new Border(
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder(),
                new DiagonalBorder()));

            // 5. Default formats
            ss.CellStyleFormats.AppendChild(new CellFormat());
            ss.CellFormats.AppendChild(new CellFormat());

            return ss;
        }

        private static void EnsureCollections(Stylesheet ss)
        {
            if (ss.Fonts == null) ss.Fonts = new Fonts();
            if (ss.Fills == null) ss.Fills = new Fills();
            if (ss.Borders == null) ss.Borders = new Borders();
            if (ss.CellStyleFormats == null) ss.CellStyleFormats = new CellStyleFormats();
            if (ss.CellFormats == null) ss.CellFormats = new CellFormats();

            if (ss.Fonts.Count == null) ss.Fonts.Count = (uint)ss.Fonts.ChildElements.Count;
            if (ss.Fills.Count == null) ss.Fills.Count = (uint)ss.Fills.ChildElements.Count;
            if (ss.Borders.Count == null) ss.Borders.Count = (uint)ss.Borders.ChildElements.Count;
            if (ss.CellStyleFormats.Count == null) ss.CellStyleFormats.Count = (uint)ss.CellStyleFormats.ChildElements.Count;
            if (ss.CellFormats.Count == null) ss.CellFormats.Count = (uint)ss.CellFormats.ChildElements.Count;

            // Ensure required default fills exist
            if (ss.Fills.ChildElements.Count < 2)
            {
                ss.Fills.RemoveAllChildren();
                ss.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.None }));
                ss.Fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
                ss.Fills.Count = 2U;
            }

            if (ss.CellFormats.ChildElements.Count == 0)
            {
                ss.CellFormats.AppendChild(new CellFormat());
                ss.CellFormats.Count = 1U;
            }
        }

        private void WarmUpCaches()
        {
            // 1. Optional: scan existing fonts/fills/borders and register them to caches
            //    This avoids duplicates when template already contains desired specs.
            //    Minimal version: do nothing.
        }

        private static string ToArgbHex(System.Drawing.Color color)
        {
            // AARRGGBB
            return color.A.ToString("X2", CultureInfo.InvariantCulture)
                 + color.R.ToString("X2", CultureInfo.InvariantCulture)
                 + color.G.ToString("X2", CultureInfo.InvariantCulture)
                 + color.B.ToString("X2", CultureInfo.InvariantCulture);
        }
    }
}
