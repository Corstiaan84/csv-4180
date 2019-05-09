using System;
using System.Runtime.InteropServices;

namespace DocInOffice.MSOfficeExtension
{
    [ComImport]
    [Guid("000C03D7-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [TypeLibType(TypeLibTypeFlags.FOleAutomation)]
    public interface IConverter
    {
        void HrInitConverter([MarshalAs(UnmanagedType.Interface), In] IConverterApplicationPreferences pcap, [MarshalAs(UnmanagedType.Interface)] out IConverterPreferences ppcp, [MarshalAs(UnmanagedType.Interface), In] IConverterUICallback pcuic);

        void HrUninitConverter([MarshalAs(UnmanagedType.Interface), In] IConverterUICallback pcuic);

        void HrImport([MarshalAs(UnmanagedType.BStr), In] string bstrSourcePath, [MarshalAs(UnmanagedType.BStr), In] string bstrDestPath, [MarshalAs(UnmanagedType.Interface), In] IConverterApplicationPreferences pcap, [MarshalAs(UnmanagedType.Interface)] out IConverterPreferences ppcp, [MarshalAs(UnmanagedType.Interface), In] IConverterUICallback pcuic);

        void HrExport([MarshalAs(UnmanagedType.BStr), In] string bstrSourcePath, [MarshalAs(UnmanagedType.BStr), In] string bstrDestPath, [MarshalAs(UnmanagedType.BStr), In] string bstrClass, [MarshalAs(UnmanagedType.Interface), In] IConverterApplicationPreferences pcap, [MarshalAs(UnmanagedType.Interface)] out IConverterPreferences ppcp, [MarshalAs(UnmanagedType.Interface), In] IConverterUICallback pcuic);

        void HrGetFormat([MarshalAs(UnmanagedType.BStr), In] string bstrPath, [MarshalAs(UnmanagedType.BStr)] out string pbstrClass, [MarshalAs(UnmanagedType.Interface), In] IConverterApplicationPreferences pcap, [MarshalAs(UnmanagedType.Interface)] out IConverterPreferences ppcp, [MarshalAs(UnmanagedType.Interface), In] IConverterUICallback pcuic);

        void HrGetErrorString([In] int hrErr, [MarshalAs(UnmanagedType.BStr)] out string pbstrErrorMsg, [MarshalAs(UnmanagedType.Interface), In] IConverterApplicationPreferences pcap);
    }
}
