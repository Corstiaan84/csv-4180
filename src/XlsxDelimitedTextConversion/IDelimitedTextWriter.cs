﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    public interface IDelimitedTextWriter
    {
        void Write(String value);

        void EndField();

        void EndRecord();

        void Flush();
    }
}