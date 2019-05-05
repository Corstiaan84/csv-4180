//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2015
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    public interface IDelimitedTextReader
    {
        String[] ReadFields();
    }
}
