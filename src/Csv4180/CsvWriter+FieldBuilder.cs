// Copyright © Leonid S. Maliutin <mals@live.ru> 2014

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocInOffice.Csv4180
{
    public partial class CsvWriter
    {
        private class FieldBuilder
        {
            private readonly List<Char> _charList = new List<Char>();
            private Boolean _quoteMode;

            public FieldBuilder()
            {

            }

            public void Append(String value)
            {
                for (Int32 i = 0; i < value.Length; i++)
                {
                    Char currentChar = value[i];

                    if (currentChar == ',' || currentChar == '\r' || currentChar == '\n')
                    {
                        _quoteMode = true;
                        _charList.Add(currentChar);
                    }
                    else if (currentChar == '"')
                    {
                        _quoteMode = true;
                        _charList.Add('"');
                        _charList.Add('"');
                    }
                    else
                    {
                        _charList.Add(currentChar);
                    }
                }
            }

            public void Reset()
            {
                _quoteMode = false;
                _charList.Clear();
            }

            public FieldCharEnumerator GetEnumerator()
            {
                return new FieldCharEnumerator(_charList, _quoteMode);
            }
        }
    }
}
