using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyOpenXml.Excel.Models;
using System.Globalization;
using OXmlFont = DocumentFormat.OpenXml.Spreadsheet.Font;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class StyleManager_old
    {
        private readonly WorkbookPart _workbookPart;
        private readonly Stylesheet _ss;
        private readonly Dictionary<string, uint> _styleCache = new Dictionary<string, uint>(StringComparer.Ordinal);


        internal StyleManager_old(SpreadsheetDocument document)
        {
            Guards.NotNull(document, nameof(document));
            _workbookPart = document.WorkbookPart
                ?? throw new InvalidOperationException("WorkbookPart is missing.");

            EnsureStylesheet();

            _ss = EnsureStylesheet(_workbookPart);
        }

        internal uint GetOrCreateNumberFormat(string format)
        {
            // - Return default style index (0)
            // - Real implementation will add NumberingFormats + CellFormats
            return 0;
        }

        internal uint GetOrCreateFontColor(System.Drawing.Color color)
        {
            // Excel expects ARGB hex (AARRGGBB)
            var argb = ToArgbHex(color);
            var key = "font:" + argb;

            if (_styleCache.TryGetValue(key, out var cached))
                return cached;

            // 1) Ensure required collections exist
            var fonts = _ss.Fonts ?? (_ss.Fonts = new Fonts());
            var fills = _ss.Fills ?? (_ss.Fills = CreateDefaultFills());      // keep minimum 2 fills
            var borders = _ss.Borders ?? (_ss.Borders = CreateDefaultBorders());
            var cellFormats = _ss.CellFormats ?? (_ss.CellFormats = CreateDefaultCellFormats());

            // 2) Add new Font with color
            var font = new Font(
                new Color { Rgb = argb } // <color rgb="FFFF0000"/>
            );
            fonts.Append(font);

            // Update count
            fonts.Count = (uint)fonts.ChildElements.Count;

            // 3) Add new CellFormat referring to the new FontId
            var fontId = (uint)(fonts.ChildElements.Count - 1);

            var cf = new CellFormat
            {
                FontId = fontId,
                FillId = 0,     // keep default fill
                BorderId = 0,
                ApplyFont = true
            };

            cellFormats.Append(cf);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            // 4) Save stylesheet and return styleIndex (= cell format index)
            _workbookPart.WorkbookStylesPart.Stylesheet.Save();

            var styleIndex = (uint)(cellFormats.ChildElements.Count - 1);
            _styleCache[key] = styleIndex;

            return styleIndex;
        }

        internal uint GetOrCreateFillColor(System.Drawing.Color color)
        {
            var argb = ToArgbHex(color);
            var key = "fill:" + argb;

            if (_styleCache.TryGetValue(key, out var cached))
                return cached;

            var fonts = _ss.Fonts ?? (_ss.Fonts = new Fonts(new Font()));
            var fills = _ss.Fills ?? (_ss.Fills = CreateDefaultFills());
            var borders = _ss.Borders ?? (_ss.Borders = CreateDefaultBorders());
            var cellFormats = _ss.CellFormats ?? (_ss.CellFormats = CreateDefaultCellFormats());

            // Add Fill (solid)
            var fill = new Fill(
                new PatternFill(
                    new ForegroundColor { Rgb = argb },
                    new BackgroundColor { Indexed = 64U }
                )
                { PatternType = PatternValues.Solid }
            );
            fills.Append(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            var fillId = (uint)(fills.ChildElements.Count - 1);

            var cf = new CellFormat
            {
                FontId = 0,
                FillId = fillId,
                BorderId = 0,
                ApplyFill = true
            };

            cellFormats.Append(cf);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            _workbookPart.WorkbookStylesPart.Stylesheet.Save();

            var styleIndex = (uint)(cellFormats.ChildElements.Count - 1);
            _styleCache[key] = styleIndex;

            return styleIndex;
        }

        private void EnsureStylesheet()
        {
            if (_workbookPart.WorkbookStylesPart != null) return;

            var part = _workbookPart.AddNewPart<WorkbookStylesPart>();
            part.Stylesheet = new Stylesheet(
                new Fonts(new OXmlFont()),
                new Fills(new Fill()),
                new Borders(new Border()),
                new CellFormats(new CellFormat())
            );
            part.Stylesheet.Save();
        }


        private static Stylesheet EnsureStylesheet(WorkbookPart workbookPart)
        {
            if (workbookPart.WorkbookStylesPart == null)
            {
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                stylesPart.Stylesheet.Save();
            }

            var ss = workbookPart.WorkbookStylesPart.Stylesheet ?? (workbookPart.WorkbookStylesPart.Stylesheet = new Stylesheet());

            // Ensure defaults exist (Excel likes these)
            ss.Fonts ??= new Fonts(new Font()) { Count = 1U };
            ss.Fills ??= CreateDefaultFills();
            ss.Borders ??= CreateDefaultBorders();
            ss.CellFormats ??= CreateDefaultCellFormats();

            workbookPart.WorkbookStylesPart.Stylesheet.Save();
            return ss;
        }

        private static Fills CreateDefaultFills()
        {
            // Required minimum: 2 fills (None and Gray125)
            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            )
            { Count = 2U };
            return fills;
        }

        private static Borders CreateDefaultBorders()
        {
            var borders = new Borders(new Border()) { Count = 1U };
            return borders;
        }

        private static CellFormats CreateDefaultCellFormats()
        {
            var cfs = new CellFormats(
                new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0
                }
            )
            { Count = 1U };
            return cfs;
        }

        private static string ToArgbHex(System.Drawing.Color color)
        {
            // AARRGGBB
            return color.A.ToString("X2", CultureInfo.InvariantCulture)
                 + color.R.ToString("X2", CultureInfo.InvariantCulture)
                 + color.G.ToString("X2", CultureInfo.InvariantCulture)
                 + color.B.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static uint GetNextCustomNumFmtId(NumberingFormats nfs)
        {
            uint max = 163; // custom should start at 164
            foreach (var nf in nfs.Elements<NumberingFormat>())
            {
                if (nf.NumberFormatId != null && nf.NumberFormatId.Value > max)
                    max = nf.NumberFormatId.Value;
            }
            return max + 1;
        }

        internal uint GetOrCreateAlignmentStyle(
            uint baseStyleIndex,
            HorizontalAlign hAlign,
            VerticalAlign vAlign,
            bool wrapText)
        {
            var key = $"align:{baseStyleIndex}:{hAlign}:{vAlign}:{wrapText}";
            if (_styleCache.TryGetValue(key, out var cached))
                return cached;

            var ss = _workbookPart.WorkbookStylesPart.Stylesheet;
            var cellFormats = ss.CellFormats;

            var baseCf = (CellFormat)cellFormats.ChildElements[(int)baseStyleIndex];
            var newCf = (CellFormat)baseCf.CloneNode(true);

            var alignment = newCf.Alignment ?? (newCf.Alignment = new Alignment());

            alignment.Horizontal = ToHorizontal(hAlign);
            alignment.Vertical = ToVertical(vAlign);
            alignment.WrapText = wrapText;

            newCf.ApplyAlignment = true;

            cellFormats.Append(newCf);
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            ss.Save();

            var newIndex = (uint)(cellFormats.ChildElements.Count - 1);
            _styleCache[key] = newIndex;
            return newIndex;
        }

        private static EnumValue<HorizontalAlignmentValues>? ToHorizontal(HorizontalAlign h)
        {
            switch (h)
            {
                case HorizontalAlign.Left: return HorizontalAlignmentValues.Left;
                case HorizontalAlign.Center: return HorizontalAlignmentValues.Center;
                case HorizontalAlign.Right: return HorizontalAlignmentValues.Right;
                default: return HorizontalAlignmentValues.General;
            }
        }

        private static EnumValue<VerticalAlignmentValues>? ToVertical(VerticalAlign v)
        {
            switch (v)
            {
                case VerticalAlign.Top: return VerticalAlignmentValues.Top;
                case VerticalAlign.Center: return VerticalAlignmentValues.Center;
                default: return VerticalAlignmentValues.Bottom;
            }
        }

    }
}
