using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using DocInOffice.Csv4180;
using DocInOffice.XlsxDelimitedTextConversion;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    internal class MoqDelimitedTextReader : IDelimitedTextReader
    {
        private readonly StreamReader _streamReader;
        private readonly CsvReader _csvReader;

        private UInt32 _progressValue = 0;

        public MoqDelimitedTextReader(CsvReader csvReader)
        {
            if (csvReader == null) new ArgumentNullException("csvReader");

            _csvReader = csvReader;
        }

        public string[] ReadFields()
        {
            String[] fiedls = _csvReader.ReadFields();
            return fiedls;
        }
    }
}
