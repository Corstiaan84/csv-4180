// Copyright © Leonid S. Maliutin <mals@live.ru> 2014

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocInOffice.Csv4180
{
    public partial class CsvWriter
    {
        private struct FieldCharEnumerator
        {
            private readonly List<Char> _charList;
            private readonly Boolean _quoteMode;
            private Int32 _position;
            private Char _currentChar;


            public FieldCharEnumerator(List<Char> charList, Boolean quoteMode)
            {
                _charList = charList;
                _quoteMode = quoteMode;
                _position = quoteMode ? -2 : -1;
                _currentChar = default(Char);
            }

            public char Current
            {
                get 
                {
                    return _currentChar;
                }
            }

            public bool MoveNext()
            {
                _position += 1;

                if (_position == -1 && _quoteMode)
                {
                    _currentChar = '"';
                    return true;
                }
                else if(_position < _charList.Count)
                {
                    _currentChar = _charList[_position];
                    return true;
                }
                else if (_position == _charList.Count && _quoteMode)
                {
                    _currentChar = '"';
                    return true;
                }
                else
                {
                    _currentChar = default(Char);
                    return false;
                }
            }
        }
    }
}
