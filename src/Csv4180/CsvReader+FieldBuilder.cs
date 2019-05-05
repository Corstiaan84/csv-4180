using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace DocInOffice.Csv4180
{
    public partial class CsvReader
    {
        private class FieldBuilder
        {
            private readonly StringBuilder _stringBuilder = new StringBuilder();
            private readonly List<Char> _separatorBuffer = new List<Char>();

            public FieldBuilder()
            {

            }

            public Boolean HasValue
            {
                get { return _stringBuilder.Length > 0; }
            }

            public void Clear()
            {
                if (_stringBuilder.Length > 0)
                    _stringBuilder.Remove(0, _stringBuilder.Length);

                _separatorBuffer.Clear();
            }

            public override string ToString()
            {
                return _stringBuilder.ToString();
            }

            public void Append(Char ch)
            {
                this.FlushSeparatorBuffer();
                _stringBuilder.Append(ch);
            }

            public void AppendEndLine()
            {
                this.FlushSeparatorBuffer();
                _stringBuilder.AppendLine();
            }

            public void PutSeparator(Char ch)
            {
                Debug.Assert(ch != '\r');
                Debug.Assert(ch != '\n');
                Debug.Assert(Char.IsSeparator(ch));

                _separatorBuffer.Add(ch);
            }

            private void FlushSeparatorBuffer()
            {
                if (_separatorBuffer.Count > 0)
                {
                    _stringBuilder.Append(_separatorBuffer.ToArray());
                    _separatorBuffer.Clear();
                }
            }
        }
    }
}
