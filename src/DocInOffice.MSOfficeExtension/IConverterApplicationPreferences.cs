//
// Copyright © Leonid Maliutin <mals@live.ru> 2016
//

using System;
using System.Runtime.InteropServices;

namespace DocInOffice.MSOfficeExtension
{
    [ComImport]
    [Guid("000C03D5-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [TypeLibType(TypeLibTypeFlags.FOleAutomation)]
    public interface IConverterApplicationPreferences
    {
        void HrGetLcid(out uint plcid);

        void HrGetHwnd(out int phwnd);

        void HrGetApplication([MarshalAs(UnmanagedType.BStr)] out string pbstrApplication);

        void HrCheckFormat(out int pFormat);
    }
}
