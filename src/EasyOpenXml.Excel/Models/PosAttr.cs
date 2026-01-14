using System.Drawing;
using EasyOpenXml.Excel.Internals;

namespace EasyOpenXml.Excel.Models
{
    public sealed class PosAttr
    {
        private readonly PosAttrProxy _proxy;

        internal PosAttr(PosProxy posProxy)
        {
            _proxy = new PosAttrProxy(posProxy);
        }

        // Display format (NumberFormat)
        public string Format
        {
            set => _proxy.SetFormat(value);
        }

        // Font color
        public Color FontColor
        {
            set => _proxy.SetFontColor(value);
        }

        // Background color
        public Color BackColor
        {
            set => _proxy.SetBackColor(value);
        }
    }
}
