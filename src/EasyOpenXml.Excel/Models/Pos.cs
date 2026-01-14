using DocumentFormat.OpenXml.Vml.Office;
using System;
using System.Collections.Generic;
using System.Text;
using EasyOpenXml.Excel.Internals;

namespace EasyOpenXml.Excel.Models
{
    public sealed class Pos
    {
        private readonly Internals.PosProxy _proxy;
        private PosAttr _attr;

        internal Pos(Internals.PosProxy proxy)
        {
            _proxy = proxy;
        }

        public object Value
        {
            get => _proxy.GetValue();
            set => _proxy.SetValue(value, isString: false);
        }

        public object Str
        {
            get => _proxy.GetValue();
            set => _proxy.SetValue(value, isString: true);
        }

        public PosAttr Attr
        {
            get
            {
                if (_attr == null)
                    _attr = new PosAttr(_proxy);
                return _attr;
            }
        }

        // Models/Pos.cs（追加）
        public void Copy()
        {
            var snapshot = _proxy.CaptureSnapshot();

            // ExcelInternal は PosProxy 経由では直接触れないため、
            // Clipboard は PosProxy.Document から辿る設計にする
            ExcelInternalAccessor.SetClipboard(_proxy.Document, snapshot);
        }

        public void Copy(string r, string c)
        {
            // MVP: ignore r,c and behave same as Copy()
            Copy();
        }

        public void Paste()
        {
            var snapshot = ExcelInternalAccessor.GetClipboard(_proxy.Document);
            if (snapshot == null)
                return; // or throw, depending on policy

            _proxy.ApplySnapshot(snapshot);
        }
    }
}

