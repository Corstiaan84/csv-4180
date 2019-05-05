//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2014
//

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using DocInOffice.XlsxDelimitedTextConversion;

namespace DocInOffice.Csv4180.FileExtension
{
    internal class ProgressCsvReader : IDelimitedTextReader
    {
        private readonly StreamReader _streamReader;
        private readonly CsvReader _csvReader;
        private readonly Action<UInt32> _showProgress;

        private UInt32 _progressValue = 0;

        public ProgressCsvReader(StreamReader streamReader, CsvReader csvReader, Action<UInt32> showProgress)
        {
            if (streamReader == null) new ArgumentNullException("streamReader");
            if (csvReader == null) new ArgumentNullException("csvReader");
            if (showProgress == null) new ArgumentNullException("showProgress");

            _streamReader = streamReader;
            _csvReader = csvReader;
            _showProgress = showProgress;
        }

        public string[] ReadFields()
        {
            String[] fiedls = _csvReader.ReadFields();
            this._processProgress();
            return fiedls;
        }

        private void _processProgress()
        {
            Stream stream = this._streamReader.BaseStream;
            UInt32 currentValue = (UInt32)Math.Floor(100d * stream.Position / stream.Length);

            if (this._progressValue < currentValue)
            {
                this._progressValue = currentValue;
                this._showProgress(currentValue);
            }
        }
    }
}
