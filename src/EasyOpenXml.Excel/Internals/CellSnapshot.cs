using System.Drawing;

namespace EasyOpenXml.Excel.Internals
{
    internal sealed class CellSnapshot
    {
        internal object Value { get; init; }
        internal bool IsString { get; init; }

        internal uint StyleIndex { get; init; }
    }
}
