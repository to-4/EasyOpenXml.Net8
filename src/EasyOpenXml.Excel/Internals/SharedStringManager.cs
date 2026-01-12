using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class SharedStringManager
    {
        private readonly WorkbookPart _workbookPart;
        private SharedStringTablePart _sstPart;

        internal SharedStringManager(SpreadsheetDocument document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            _workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
        }

        internal int GetOrAddString(string text)
        {
            // 1. Ensure SharedStringTablePart exists
            EnsureSharedStringTablePart();

            // 2. Find existing item (MVP: linear search; OK for small/medium use)
            //    If you expect huge strings, add a dictionary cache.
            var sst = _sstPart.SharedStringTable;
            var items = sst.Elements<SharedStringItem>().ToList();

            for (int i = 0; i < items.Count; i++)
            {
                var existing = items[i].InnerText ?? string.Empty;
                if (string.Equals(existing, text ?? string.Empty, StringComparison.Ordinal))
                    return i;
            }

            // 3. Add new item
            var newItem = new SharedStringItem(new Text(text ?? string.Empty));
            sst.AppendChild(newItem);

            // 4. Update counts (Excel sometimes repairs files if counts are missing/wrong)
            sst.Count = (uint)sst.Count();
            sst.UniqueCount = (uint)sst.Elements<SharedStringItem>().Count();

            _sstPart.SharedStringTable.Save();

            return items.Count; // new index
        }

        internal string GetStringByIndexOrEmpty(int index)
        {
            EnsureSharedStringTablePart();
            var sst = _sstPart.SharedStringTable;

            if (index < 0) return string.Empty;

            var item = sst.Elements<SharedStringItem>().ElementAtOrDefault(index);
            return item?.InnerText ?? string.Empty;
        }

        private void EnsureSharedStringTablePart()
        {
            // 1. Create if missing
            _sstPart ??= _workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            if (_sstPart == null)
            {
                _sstPart = _workbookPart.AddNewPart<SharedStringTablePart>();
                _sstPart.SharedStringTable = new SharedStringTable();
                _sstPart.SharedStringTable.Save();
            }
            else
            {
                // 2. Ensure table object exists
                _sstPart.SharedStringTable ??= new SharedStringTable();
            }
        }
    }
}
