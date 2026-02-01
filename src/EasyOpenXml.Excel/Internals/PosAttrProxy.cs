using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Drawing;
using System.Globalization;

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

            //var styleIndex = _styleManager.GetOrCreateNumberFormat(format);
            //_posProxy.ApplyStyle(styleIndex);
        }

        internal void SetFontColor(System.Drawing.Color color)
        {
            //var styleIndex = _styleManager.GetOrCreateFontColor(color);
            //_posProxy.ApplyStyle(styleIndex);
        }

        /// <summary>
        /// 指定された色オブジェクトを基に、塗りつぶし（背景色）スタイルを設定します
        /// </summary>
        /// <param name="color">色オブジェクト</param>
        internal void SetBackColor(System.Drawing.Color color)
        {
            // 背景色をセット
            _posProxy.SetBackColor(_styleManager, color);
        }

    }
}
