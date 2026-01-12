using System;
using System.Collections.Generic;
using System.Text;


namespace EasyOpenXml.Excel.Models
{
    public sealed class Pos
    {
        private readonly Internals.PosProxy _proxy;

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
    }
}

