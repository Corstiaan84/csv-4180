//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2014
//

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;

namespace DocInOffice.Csv4180
{
    public partial class CsvReader
    {
        private class CharReader
        {
            private readonly TextReader _textReader;

            public CharReader(TextReader textReader)
            {
                if (textReader == null) throw new ArgumentNullException("textReader");

                _textReader = textReader;
            }

            public Boolean ReadChars(out Char currentChar, out Char nextChar)
            {
                Int32 c = _textReader.Read();

                if (c == -1)
                {
                    currentChar = Char.MinValue;
                    nextChar = Char.MinValue;

                    return false;
                }
                else
                {
                    currentChar = (Char)c;

                    Int32 c1 = _textReader.Peek();

                    nextChar = c1 == -1 ? Char.MinValue : (Char)c1;

                    return true;
                }
            }

            public void SkipNextChar()
            {
                _textReader.Read();
            }
        }
    }
}
