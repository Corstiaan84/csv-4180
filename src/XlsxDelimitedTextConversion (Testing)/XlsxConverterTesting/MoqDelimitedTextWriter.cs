﻿//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2015
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocInOffice.Csv4180;
using DocInOffice.XlsxDelimitedTextConversion;

namespace DocInOffice.XlsxDelimitedTextConversion.XlsxConverterTesting
{
    class MoqDelimitedTextWriter : IDelimitedTextWriter
    {
        private readonly CsvWriter _csvWriter;

        public MoqDelimitedTextWriter(CsvWriter csvWriter)
        {
            if (csvWriter == null) throw new ArgumentNullException("csvWriter");

            _csvWriter = csvWriter;
        }

        public void Write(string value)
        {
            _csvWriter.Write(value);
        }

        public void EndField()
        {
            _csvWriter.EndField();
        }

        public void EndRecord()
        {
            _csvWriter.EndRecord();
        }

        public void Flush()
        {
            _csvWriter.Flush();
        }
    }
}