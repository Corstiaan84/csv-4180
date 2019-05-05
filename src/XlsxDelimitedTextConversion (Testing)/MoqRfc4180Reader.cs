//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2015
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using DocInOffice.Csv4180;
using DocInOffice.XlsxDelimitedTextConversion;

namespace DocInOffice.XlsxDelimitedTextConversion.XlsxConverterTesting
{
    internal class ProgressCsvReader : IDelimitedTextReader
    {
        private readonly StreamReader _streamReader;
        private readonly CsvReader _csvReader;

        private UInt32 _progressValue = 0;

        public ProgressCsvReader(StreamReader streamReader, CsvReader csvReader)
        {
            if (streamReader == null) new ArgumentNullException("streamReader");
            if (csvReader == null) new ArgumentNullException("csvReader");

            _streamReader = streamReader;
            _csvReader = csvReader;
        }

        public string[] ReadFields()
        {
            String[] fiedls = _csvReader.ReadFields();
            return fiedls;
        }
    }
}
