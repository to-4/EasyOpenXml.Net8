// src/EasyOpenXml.Excel/ExcelDocumentException.cs
using System;
using System.Runtime.Serialization;

namespace EasyOpenXml.Excel
{
    /// <summary>
    /// Public-facing exception for EasyOpenXml.Excel.
    /// Note: OpenXml SDK exceptions must not leak through public API.
    /// </summary>
    [Serializable]
    public class ExcelDocumentException : Exception
    {
        public ExcelDocumentException()
        {
        }

        public ExcelDocumentException(string message)
            : base(message)
        {
        }

        public ExcelDocumentException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        // For serialization support (.NET Framework)
        protected ExcelDocumentException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        /// <summary>
        /// Optional error code for compatibility with return-code style APIs (0 / -1 etc.).
        /// </summary>
        public int ErrorCode { get; set; } = 0;
    }
}
