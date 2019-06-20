using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocInOffice.Csv4180
{
    public interface ICsvReader
    {
        String[] ReadFields();
    }
}
