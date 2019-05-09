using System;
using System.Runtime.InteropServices;

namespace DocInOffice.MSOfficeExtension
{
    [ComImport]
    [Guid("000C03D4-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [TypeLibType(TypeLibTypeFlags.FOleAutomation)]
    public interface IConverterPreferences
    {
        void HrGetMacroEnabled(out int pfMacroEnabled);

        void HrCheckFormat(out int pFormat);

        void HrGetLossySave(out int pfLossySave);
    }
}
