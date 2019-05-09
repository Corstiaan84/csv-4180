using System;
using System.Runtime.InteropServices;

namespace DocInOffice.MSOfficeExtension
{
    [ComImport]
    [Guid("000C03D6-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [TypeLibType(TypeLibTypeFlags.FOleAutomation)]
    public interface IConverterUICallback
    {
        void HrReportProgress([In] uint uPercentComplete);

        void HrMessageBox([MarshalAs(UnmanagedType.BStr), In] string bstrText,
                          [MarshalAs(UnmanagedType.BStr), In] string bstrCaption,
                          [In] uint uType,
                          out int pidResult);

        void HrInputBox([MarshalAs(UnmanagedType.BStr), In] string bstrText,
                        [MarshalAs(UnmanagedType.BStr), In] string bstrCaption,
                        [MarshalAs(UnmanagedType.BStr)] out string pbstrInput,
                        [In] int fPassword);
    }
}
