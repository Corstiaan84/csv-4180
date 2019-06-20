using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    [Serializable]
    public class ActiveWorksheetException : Exception
    {
        internal ActiveWorksheetException()
            : base("The active sheet is not the worksheet.")
        { }

        protected ActiveWorksheetException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}
