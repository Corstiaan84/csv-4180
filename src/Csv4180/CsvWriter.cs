using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;

namespace DocInOffice.Csv4180
{
    public partial class CsvWriter : ICsvWriter, IDisposable
    {
        private readonly TextWriter _textWriter;
        private readonly FieldBuilder _fieldBuilder = new FieldBuilder();

        public CsvWriter(TextWriter textWriter)
        {
            if (textWriter == null) throw new ArgumentNullException("textWriter");

            _textWriter = textWriter;
        }

    #region Disposing implementation
        public void Dispose()
        {
            this.FlushField();
        }
    #endregion

        public void Write(String value)
        {
            if (String.IsNullOrEmpty(value)) return;

            _fieldBuilder.Append(value);
        }

        public void EndField()
        {
            this.FlushField();
            _textWriter.Write(',');
        }

        public void EndRecord()
        {
            this.FlushField();
            _textWriter.Write('\r');
            _textWriter.Write('\n');
        }

        public void Flush()
        {
            this.FlushField();
        }

        private void FlushField()
        {
            Char previousChar = Char.MinValue;

            foreach (Char currentChar in _fieldBuilder)
            {
                if (currentChar == '\n')
                {
                    if (previousChar == '\r')
                    {
                        _textWriter.Write('\n');
                    }
                    else
                    {
                        _textWriter.Write('\r');
                        _textWriter.Write('\n');
                    }
                }
                else
                {
                    _textWriter.Write(currentChar);
                }

                previousChar = currentChar;
            }

            _fieldBuilder.Reset();
        }
    }
}
