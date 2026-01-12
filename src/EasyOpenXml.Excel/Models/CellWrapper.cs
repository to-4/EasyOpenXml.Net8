// src/EasyOpenXml.Excel/Models/CellWrapper.cs
using EasyOpenXml.Excel.Internals;

namespace EasyOpenXml.Excel.Models
{
    public sealed class CellWrapper
    {
        private readonly CellWrapperProxy _proxy;

        internal CellWrapper(CellWrapperProxy proxy)
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
    }
}
