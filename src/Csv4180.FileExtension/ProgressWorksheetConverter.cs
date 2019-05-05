//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2015
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocInOffice.XlsxDelimitedTextConversion;

namespace DocInOffice.Csv4180.FileExtension
{
    class ProgressWorksheetConverter : WorksheetConverter
    {
        private readonly Action<UInt32> _reportProgress;
        private UInt32 _progressValue = 0;

        public ProgressWorksheetConverter(Action<UInt32> reportProgress)
        {
            if (reportProgress == null) throw new ArgumentNullException("reportProgress");

            _reportProgress = reportProgress;
        }

        protected override void OnRowConverted(Int32 rowIndex, Int32 rowCount)
        {
            UInt32 currentValue = (UInt32)Math.Floor(100d * rowIndex / rowCount);

            if (_progressValue < currentValue)
            {
                _progressValue = currentValue;
                _reportProgress(_progressValue);
            }
        }

    }
}
