using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class SheetManager
    {
        private readonly SpreadsheetDocument _document;
        private readonly WorkbookPart _workbookPart;

        private readonly List<SheetInfo> _sheets = new List<SheetInfo>();

        private int _currentIndex = -1;
        private WorksheetPart _currentWorksheetPart;

        internal int CurrentIndex => _currentIndex;
        internal string CurrentSheetName => _sheets[_currentIndex].Name;

        internal SheetManager(SpreadsheetDocument document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));

            // 1. Cache core parts
            _document = document;
            _workbookPart = _document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");

            // 2. Load sheet list
            LoadSheets();

            // 3. Default selection: leftmost sheet (index 0)
            if (_sheets.Count > 0)
            {
                SelectByIndex(0);
            }
        }

        internal WorksheetPart CurrentWorksheetPart
        {
            get
            {
                if (_currentWorksheetPart == null)
                    throw new InvalidOperationException("Worksheet is not selected. Set SheetNo first.");
                return _currentWorksheetPart;
            }
        }

        internal IReadOnlyList<string> GetSheetNames()
        {
            // Return a copy to avoid external mutation
            return _sheets.Select(s => s.Name).ToList();
        }

        internal void SelectByIndex(int index)
        {
            // 1. Validate
            if (_sheets.Count == 0)
                throw new InvalidOperationException("This workbook has no sheets.");

            if (index < 0 || index >= _sheets.Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Sheet index must be in 0..{_sheets.Count - 1}.");

            // 2. Resolve WorksheetPart
            var info = _sheets[index];
            var part = _workbookPart.GetPartById(info.RelationshipId) as WorksheetPart;

            if (part == null)
                throw new InvalidOperationException("WorksheetPart could not be resolved by relationship id.");

            // 3. Set current
            _currentIndex = index;
            _currentWorksheetPart = part;
        }

        internal void SelectByName(string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("Sheet name is required.", nameof(name));

            var index = _sheets.FindIndex(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new ArgumentOutOfRangeException(nameof(name), "The specified sheet name was not found.");

            SelectByIndex(index);
        }

        private void LoadSheets()
        {
            _sheets.Clear();

            // 1. Ensure workbook structure exists
            var workbook = _workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
            var sheets = workbook.Sheets ?? throw new InvalidOperationException("Sheets collection is missing.");

            // 2. Preserve the order as it appears in the workbook (left-to-right tabs)
            foreach (var sheet in sheets.Elements<Sheet>())
            {
                if (sheet == null) continue;

                var relId = sheet.Id?.Value;
                var name = sheet.Name?.Value ?? string.Empty;

                if (string.IsNullOrEmpty(relId))
                {
                    // If relationship id is missing, Excel may repair anyway.
                    // We treat it as invalid and skip.
                    continue;
                }

                _sheets.Add(new SheetInfo(name, relId));
            }

            if (_sheets.Count == 0)
                throw new InvalidOperationException("No valid sheets were found in the workbook.");
        }

        private sealed class SheetInfo
        {
            internal SheetInfo(string name, string relationshipId)
            {
                Name = name ?? string.Empty;
                RelationshipId = relationshipId ?? throw new ArgumentNullException(nameof(relationshipId));
            }

            internal string Name { get; }
            internal string RelationshipId { get; }
        }
    }
}
