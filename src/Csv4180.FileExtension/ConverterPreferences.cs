using System;
using System.Runtime.InteropServices;

using DocInOffice.MSOfficeExtension;

namespace DocInOffice.Csv4180.FileExtension
{
    #region DEBUG
    [ComVisible(true)]
    [Guid("0829ED03-0E73-430C-831E-CB5DEBEDEB31")]
    #endregion
    public class ConverterPreferences : IConverterPreferences
    {
        public void HrCheckFormat(out int pFormat)
        {
            pFormat = 9;
        }

        public void HrGetLossySave(out int pfLossySave)
        {
            pfLossySave = Convert.ToInt32(true);
        }

        public void HrGetMacroEnabled(out int pfMacroEnabled)
        {
            pfMacroEnabled = Convert.ToInt32(false);
        }
    }
}
