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
    public partial class CsvReader : ICsvReader
    {
        private readonly CharReader _charReader;
        private readonly List<FieldBuilder> _fieldBuilderList = new List<FieldBuilder>();

        public CsvReader(TextReader textReader)
        {
            _charReader = new CharReader(textReader);
        }

        public String[] ReadFields()
        {
            Char previewChar = Char.MinValue;
            Char currentChar = Char.MinValue;
            Char nextChar = Char.MinValue;

            if (_charReader.ReadChars(out currentChar, out nextChar) == false)
                return null;

            Int32 currentFieldNumber = -1;
            FieldBuilder currentFieldBuilder;

            this.PrepareNextField(ref currentFieldNumber, out currentFieldBuilder);

            Boolean quoteMode = false;

            do
            {
                if (currentChar == '"')
                {
                    if (quoteMode)
                    {
                        if (nextChar == '"')
                        {
                            currentFieldBuilder.Append('"');
                            _charReader.SkipNextChar();
                        }
                        else if (nextChar == Char.MinValue // конец файла
                               || nextChar == ','           // конец поля
                               || nextChar == '\r'          // конец записи
                               )
                        {
                            quoteMode = false;
                        }
                        else
                        {
                            currentFieldBuilder.Append('"');
                        }
                    }
                    else
                    {
                        if (previewChar == Char.MinValue // начало файла
                           || previewChar == ','           // начало поля
                           || previewChar == '\n'          // начало записи
                           )
                        {
                            quoteMode = true;
                        }
                        else
                        {
                            currentFieldBuilder.Append('"');
                        }
                    }
                }
                else if (currentChar == ',')
                {
                    if (quoteMode)
                    {
                        currentFieldBuilder.Append(currentChar);
                    }
                    else
                    {
                        this.PrepareNextField(ref currentFieldNumber, out currentFieldBuilder);
                    }
                }
                else if (currentChar == '\r')
                {
                    if (nextChar != '\n')
                    {
                        currentFieldBuilder.AppendEndLine();
                    }
                }
                else if (currentChar == '\n')
                {
                    if (quoteMode == false && previewChar == '\r')
                    {
                        break;
                    }
                    else
                    {
                        currentFieldBuilder.AppendEndLine();
                    }
                }
                else
                {
                    currentFieldBuilder.Append(currentChar);
                }

                previewChar = currentChar;
            } while (_charReader.ReadChars(out currentChar, out nextChar));

            String[] fields = new String[currentFieldNumber + 1];

            for(Int32 i = 0; i < fields.Length; i += 1)
            {
                FieldBuilder fieldBuilder = _fieldBuilderList[i];
                fields[i] = fieldBuilder.ToString();
            }

            return fields;
        }

        private void PrepareNextField(ref Int32 currentFieldNumber, out FieldBuilder currentFieldBuilder)
        {
            currentFieldNumber += 1;

            if (currentFieldNumber > _fieldBuilderList.Count - 1)
            {
                currentFieldBuilder = new FieldBuilder();
                _fieldBuilderList.Add(currentFieldBuilder);
            }
            else
            {
                currentFieldBuilder = _fieldBuilderList[currentFieldNumber];

                currentFieldBuilder.Clear();
            }
        }
    }
}
