using System;
using System.Drawing;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class PosAttrProxy
    {
        private readonly PosProxy _posProxy;
        private readonly StyleManager _styleManager;

        internal PosAttrProxy(PosProxy posProxy)
        {
            Guards.NotNull(posProxy, nameof(posProxy));
            _posProxy = posProxy;
            _styleManager = new StyleManager(posProxy.Document);
        }

        internal void SetFormat(string format)
        {
            if (string.IsNullOrEmpty(format)) return;

            var styleIndex = _styleManager.GetOrCreateNumberFormat(format);
            _posProxy.ApplyStyle(styleIndex);
        }

        internal void SetFontColor(Color color)
        {
            var styleIndex = _styleManager.GetOrCreateFontColor(color);
            _posProxy.ApplyStyle(styleIndex);
        }

        internal void SetBackColor(Color color)
        {
            var styleIndex = _styleManager.GetOrCreateFillColor(color);
            _posProxy.ApplyStyle(styleIndex);
        }
    }
}
