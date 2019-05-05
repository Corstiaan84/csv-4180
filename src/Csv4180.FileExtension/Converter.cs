﻿//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2014
//

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Win32;

using Ude;

using DocInOffice.MSOfficeExtension;
using DocInOffice.XlsxDelimitedTextConversion;

namespace DocInOffice.Csv4180.FileExtension
{
    #region DEBUG
    [ComVisible(true)]
    [Guid("2B2A8095-CD42-4762-84E1-EEFD600C5CB3")]
    #endregion
    public class Converter : IConverter
    {
        private const String LogApplicationKeyName = "DocInOffice.Csv4180FileExtension";

        #region COM Registartion
        private const String Utf8ExportExtensionKeyName  = "CSV [RFC 4180] 1";
        private const String Utf16ExportExtensionKeyName = "CSV [RFC 4180] 2";
        private const String Utf32ExportExtensionKeyName = "CSV [RFC 4180] 3";
        private const String OemExportExtensionKeyName   = "CSV [RFC 4180] 4";

        private const String ImportExtensionKeyName = "CSV [RFC 4180]";

    #if DEBUG
        private static readonly String[] ExcelKeys = {
            @"SOFTWARE\Microsoft\Office\12.0\Excel",
            @"SOFTWARE\Microsoft\Office\14.0\Excel",
            @"SOFTWARE\Microsoft\Office\15.0\Excel",
            @"SOFTWARE\Microsoft\Office\16.0\Excel",
        };

        [ComRegisterFunctionAttribute]
        private static void OnRegistered(Type type)
        {
            String clsid = type.GUID.ToString("B").ToUpper();
            Console.WriteLine(clsid);

            try
            {
                using (RegistryKey clsidKey = Registry.ClassesRoot.OpenSubKey(Path.Combine("CLSID", clsid), true))
                {
                    Console.WriteLine(clsidKey.Name);
                    clsidKey.SetValue("AppID", clsid);
                    clsidKey.Flush();
                }

                using (RegistryKey appidClsidKey = Registry.ClassesRoot.CreateSubKey(Path.Combine("AppID", clsid)))
                {
                    Console.WriteLine(appidClsidKey.Name);
                    appidClsidKey.SetValue("", type.FullName);
                    appidClsidKey.SetValue("DllSurrogate", "");
                    appidClsidKey.Flush();
                }

                Action<String, String, String> addKey = (String parentKeyPath, String extensionKeyName, String extensionName) =>
                {
                    using (RegistryKey exportKey = Registry.LocalMachine.CreateSubKey(Path.Combine(parentKeyPath, extensionKeyName)))
                    {
                        Console.WriteLine(exportKey.Name);
                        exportKey.SetValue("Extensions", "csv", RegistryValueKind.String);
                        exportKey.SetValue("Name", extensionName, RegistryValueKind.String);
                        exportKey.SetValue("CLSID", clsid, RegistryValueKind.String);
                        exportKey.Flush();
                    }
                };

                foreach (String excelKey in ExcelKeys)
                {
                    String exportKeyPath = Path.Combine(excelKey, @"Text Converters\OOXML Converters\Export");

                    addKey(exportKeyPath, Utf8ExportExtensionKeyName,  "CSV [RFC 4180] UTF-8" );
                    addKey(exportKeyPath, Utf16ExportExtensionKeyName, "CSV [RFC 4180] UTF-16");
                    addKey(exportKeyPath, Utf32ExportExtensionKeyName, "CSV [RFC 4180] UTF-32");
                    addKey(exportKeyPath, OemExportExtensionKeyName,   "CSV [RFC 4180] OEM" );

                    String importKeyPath = Path.Combine(excelKey, @"Text Converters\OOXML Converters\Import");

                    addKey(importKeyPath, ImportExtensionKeyName,    "CSV [RFC 4180]");

                    String defaultKeyPath = Path.Combine(excelKey, @"Text Converters\Defaults");

                    using (RegistryKey defaultKey = Registry.LocalMachine.CreateSubKey(defaultKeyPath))
                    {
                        Console.WriteLine(defaultKey.Name);
                        defaultKey.SetValue("csv", ImportExtensionKeyName, RegistryValueKind.String);
                    }
                }

                using (RegistryKey eventLogApplicationKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\eventlog\Application", true))
                {
                    RegistryKey xlsxRfc4180EventLogKey = eventLogApplicationKey.OpenSubKey(LogApplicationKeyName);
                    try
                    {
                        if (xlsxRfc4180EventLogKey == null)
                            xlsxRfc4180EventLogKey = eventLogApplicationKey.CreateSubKey(LogApplicationKeyName);

                        Console.WriteLine(xlsxRfc4180EventLogKey.Name);
                        xlsxRfc4180EventLogKey.SetValue("EventMessageFile", @"%windir%\Microsoft.NET\Framework\v2.0.50727\EventLogMessages.dll");
                        xlsxRfc4180EventLogKey.Flush();
                    }
                    finally
                    {
                        if (xlsxRfc4180EventLogKey != null)
                            xlsxRfc4180EventLogKey.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                throw;
            }
        }

        [ComUnregisterFunctionAttribute]
        private static void OnUnregistered(Type type)
        {
            String clsid = type.GUID.ToString("B").ToUpper();

            Console.WriteLine(clsid);
            Console.WriteLine();

            Action<RegistryKey, String> deleteKey = (RegistryKey parentKey, String extensionKeyName) =>
                {
                    RegistryKey extensionKey = parentKey.OpenSubKey(extensionKeyName);

                    if (extensionKey == null)
                    {
                        Console.WriteLine("{0} has not been found.", extensionKeyName);
                    }
                    else
                    {
                        extensionKey.Close();
                        parentKey.DeleteSubKeyTree(extensionKeyName);
                        Console.WriteLine("{0} has been deleted.", extensionKeyName);
                    }
                };

            try
            {
                foreach (String excelKey in ExcelKeys)
                {
                    Console.WriteLine();
                    Console.WriteLine(excelKey);
                    String exportKeyPath = Path.Combine(excelKey, @"Text Converters\OOXML Converters\Export");
                    RegistryKey exportKey = Registry.LocalMachine.OpenSubKey(exportKeyPath, true);

                    if (exportKey != null)
                    {
                        Console.WriteLine(exportKey.Name);
                        deleteKey(exportKey, Utf8ExportExtensionKeyName);
                        deleteKey(exportKey, Utf16ExportExtensionKeyName);
                        deleteKey(exportKey, Utf32ExportExtensionKeyName);
                        deleteKey(exportKey, OemExportExtensionKeyName);

                        exportKey.Close();
                    }

                    String importKeyPath = Path.Combine(excelKey, @"Text Converters\OOXML Converters\Import");
                    RegistryKey importKey = Registry.LocalMachine.OpenSubKey(importKeyPath, true);

                    if (importKey != null)
                    {
                        Console.WriteLine(importKey.Name);

                        deleteKey(importKey, ImportExtensionKeyName);

                        importKey.Close();
                    }

                    String defaultKeyPath = Path.Combine(excelKey, @"Text Converters\Defaults");
                    RegistryKey defaultKey = Registry.LocalMachine.OpenSubKey(defaultKeyPath, true);

                    if (defaultKey == null)
                    {
                        Console.WriteLine("{0} has not been found.", defaultKeyPath);
                    }
                    else
                    {
                        Console.WriteLine(defaultKey.Name);

                        Object value = defaultKey.GetValue("csv");

                        if (ImportExtensionKeyName.Equals(value))
                        {
                            defaultKey.DeleteValue("csv");
                            defaultKey.Flush();
                        }

                        defaultKey.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                throw;
            }

            try
            {
                RegistryKey appidKey = Registry.ClassesRoot.OpenSubKey("AppID", true);

                if (appidKey != null)
                {
                    RegistryKey clsidKey = appidKey.OpenSubKey(clsid);

                    if (clsidKey != null)
                    {
                        clsidKey.Close();
                        appidKey.DeleteSubKeyTree(clsid);
                    }

                    appidKey.Close();
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                throw;
            }

            try
            {
                using (RegistryKey eventLogApplicationKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\eventlog\Application", true))
                {
                    Console.WriteLine();
                    Console.WriteLine(eventLogApplicationKey.Name);

                    deleteKey(eventLogApplicationKey, LogApplicationKeyName);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                throw;
            }
        }
    #endif
    #endregion

        private readonly TraceSource _traceSource;
        private Encoding _encoding;
        private string _errorMessage = null;

        public Converter()
        {
#if DEBUG
            _traceSource = new TraceSource(LogApplicationKeyName, SourceLevels.All);
#else
            _traceSource = new TraceSource(LogApplicationKeyName, SourceLevels.All);
#endif
            _traceSource.Listeners.Add(new EventLogTraceListener(LogApplicationKeyName));
        }

    #region Init/Uninit
        public void HrInitConverter
            (IConverterApplicationPreferences pcap
            , out IConverterPreferences ppcp
            , IConverterUICallback pcuic)
        {
            //int pidResult = 0;
            //pcuic.HrMessageBox("HrInitConverter", "FB Converter Server", 0, out pidResult);

            _traceSource.TraceInformation("Converter.HrInitConverter");
            
            ppcp = new ConverterPreferences();
        }

        public void HrUninitConverter(IConverterUICallback pcuic)
        {
            _traceSource.TraceInformation("Converter.HrUninitConverter");
        }
    #endregion

        public void HrGetErrorString
            ( int hrErr
            , out string pbstrErrorMsg
            , IConverterApplicationPreferences pcap)
        {
            _traceSource.TraceInformation("Converter.HrGetErrorString");
            pbstrErrorMsg = _errorMessage;
        }

        public void HrGetFormat 
            ( string bstrPath
            , out string pbstrClass
            , IConverterApplicationPreferences pcap
            , out IConverterPreferences ppcp
            , IConverterUICallback pcuic)
        {
#if DEBUG
            UInt32 plcid;
            pcap.HrGetLcid(out plcid);
            String application;
            pcap.HrGetApplication(out application);
            Int32 pFormat;
            pcap.HrCheckFormat(out pFormat);

            // Int32 pid0;
            // pcuic.HrMessageBox(String.Format("bstrPath: \"{0}\"\nplcid: {1}, application: \"{2}\", pFormat: {3}", bstrPath, plcid, application, pFormat), "HrGetFormat", 0, out pid0);

            _traceSource.TraceInformation("Converter.HrGetFormat\nbstrPath: \"{0}\"\nplcid: {1}, application: \"{2}\", pFormat: {3}", bstrPath, plcid, application, pFormat);
#endif
            
            ResolveEncoding(bstrPath, out pbstrClass, out _encoding);
            ppcp = new ConverterPreferences();
        }

        public void HrImport
            (string bstrSourcePath
            , string bstrDestPath
            , IConverterApplicationPreferences pcap
            , out IConverterPreferences ppcp
            , IConverterUICallback pcuic)
        {
            ppcp = new ConverterPreferences();

            try
            {
// #if DEBUG
                UInt32 plcid;
                pcap.HrGetLcid(out plcid);
                String application;
                pcap.HrGetApplication(out application);
                Int32 pFormat;
                pcap.HrCheckFormat(out pFormat);

                //                Int32 pid0;
                //                pcuic.HrMessageBox(String.Format("bstrSourcePath: \"{0}\"\nbstrDestPath \"{1}\"\nplcid: {2}, application: \"{3}\", pFormat: {4}", bstrSourcePath, bstrDestPath, plcid, application, pFormat), "HrImport", 0, out pid0);
// #endif
                _traceSource.TraceInformation("Converter.HrImport\nbstrSourcePath: \"{0}\"\nbstrDestPath \"{1}\"\nplcid: {2}, application: \"{3}\", pFormat: {4}", bstrSourcePath, bstrDestPath, plcid, application, pFormat);



                //using (CsvConverter csvConverter = new CsvConverter(bstrSourcePath, bstrDestPath))
                //{
                //if (fbDocument.ContentStatus != FB.ContentStatus.Correct)
                //{
                //    const UInt32 MB_YESNO = 0x00000004;
                //    const UInt32 MB_ICONWARNING = 0x00000030;
                //    const UInt32 type = MB_YESNO | MB_ICONWARNING;
                //    const Int32 IDYES = 6;


                //    Int32 pid;
                //    pcuic.HrMessageBox("Открываемый файл имеет неверный формат или повреждён.\nДалее будет осуществлена попытка его открыть, но при этом не гарантируется, что все данные будут отображены корректно.\n\nПродолжить?", "Внимание!", type, out pid);

                //    if (pid != IDYES)
                //        throw new ApplicationException();
                //}

                //DocInWords.Office.FBConversion.FBConverter fbConverter = new DocInWords.Office.FBConversion.FBConverter(bstrDestPath);
                //csvConverter.Convert();
                //}

                pcuic.HrReportProgress(0u);

                using (StreamReader sourceStream = new StreamReader(bstrSourcePath, _encoding))
                {
                    CsvReader csvReader = new CsvReader(sourceStream);
                    ProgressCsvReader progressCsvReader = new ProgressCsvReader(sourceStream, csvReader, pcuic.HrReportProgress);

                    using (FileStream destStream = File.OpenWrite(bstrDestPath))
                    {
                        CsvConverter.Convert(progressCsvReader, destStream);

                        destStream.Flush();
                    }
                }
            }
            catch (Exception ex)
            {
                _errorMessage = "Невозможно открыть файл. Возможно, он имеет неверный формат или повреждён.";

                _traceSource.TraceEvent(TraceEventType.Error, 1024, ex.ToString());
                Debug.Fail(ex.ToString());
                throw;
            }
            finally
            {
                pcuic.HrReportProgress(100u);
            }
        }

        public void HrExport
            (string bstrSourcePath
            , string bstrDestPath
            , string bstrClass
            , IConverterApplicationPreferences pcap
            , out IConverterPreferences ppcp
            , IConverterUICallback pcuic)
        {
//#if DEBUG
            UInt32 plcid;
            pcap.HrGetLcid(out plcid);
            String application;
            pcap.HrGetApplication(out application);
            Int32 pFormat;
            pcap.HrCheckFormat(out pFormat);

            //Int32 pid0;
            //pcuic.HrMessageBox(String.Format("bstrSourcePath: \"{0}\"\nbstrDestPath \"{1}\"\nbstrClass: \"{2}\"\nplcid: {3}, application: \"{4}\", pFormat: {5}", bstrSourcePath, bstrDestPath, bstrClass, plcid, application, pFormat), "HrExport", 0, out pid0);

            _traceSource.TraceInformation("Converter.HrExport\nbstrSourcePath: \"{0}\"\nbstrDestPath \"{1}\"\nbstrClass: \"{2}\"\nplcid: {3}, application: \"{4}\", pFormat: {5}", bstrSourcePath, bstrDestPath, bstrClass, plcid, application, pFormat);
// #endif

            ppcp = new ConverterPreferences();

            try
            {
                Encoding encoding = ComputeEncodeing(bstrClass);

                using (StreamWriter streamWriter = new StreamWriter(bstrDestPath, false, encoding))
                using (CsvWriter csvWriter = new CsvWriter(streamWriter))
                {
                    Rfc4180FileWriter fileWriter = new Rfc4180FileWriter(csvWriter);
                    WorksheetConverter worksheetConverter = new ProgressWorksheetConverter(pcuic.HrReportProgress);
                    XlsxConverter.Convert(bstrSourcePath, fileWriter, worksheetConverter);
                    fileWriter.Flush();
                    streamWriter.Flush();
                }
            }
            catch (ActiveWorksheetException)
            {
                _errorMessage = "В книге активной листом является не таблица.\nПожалуйста, выберите лист с таблицей и попробуйте сохранить снова.";
                throw;
            }
            catch (Exception ex)
            {
                _errorMessage = "Возникла внутренняя ошибка. Невозможно сохранить файл.";

                _traceSource.TraceEvent(TraceEventType.Error, 1025, ex.ToString());
                Debug.Fail(ex.ToString());
                throw;
            }
        }

        private static void ResolveEncoding(String filePath, out String exportExtensionKeyName, out Encoding encoding)
        {
            using (FileStream fs = File.OpenRead(filePath))
            {
                ICharsetDetector cdet = new CharsetDetector();
                
                byte[] buffer = new byte[1024];
                Boolean done = false;
                int count;

                while ((count = fs.Read(buffer, 0, buffer.Length)) > 0 && !done)
                {
                    cdet.Feed(buffer, 0, count);
                    done = cdet.IsDone();
                }
                cdet.DataEnd();

                if (String.Equals(cdet.Charset, Ude.Charsets.UTF8, StringComparison.Ordinal))
                {
                    exportExtensionKeyName = Utf8ExportExtensionKeyName;
                    encoding = Encoding.UTF8;
                }
                else if (String.Equals(cdet.Charset, Ude.Charsets.UTF16_LE, StringComparison.Ordinal))
                {
                    exportExtensionKeyName = Utf16ExportExtensionKeyName;
                    encoding = Encoding.Unicode;
                }
                else if (String.Equals(cdet.Charset, Ude.Charsets.UTF32_LE, StringComparison.Ordinal))
                {
                    exportExtensionKeyName = Utf32ExportExtensionKeyName;
                    encoding = Encoding.UTF32;
                }
                else
                {
                    exportExtensionKeyName = OemExportExtensionKeyName;
                    encoding = Encoding.Default;
                }
            }
        }

        private static Encoding ComputeEncodeing(String exportExtensionKeyName)
        {
            if (String.Equals(exportExtensionKeyName, Utf8ExportExtensionKeyName, StringComparison.Ordinal))
            {
                return Encoding.UTF8;
            }
            else if (String.Equals(exportExtensionKeyName, Utf16ExportExtensionKeyName, StringComparison.Ordinal))
            {
                return Encoding.Unicode;
            }
            else if (String.Equals(exportExtensionKeyName, Utf32ExportExtensionKeyName, StringComparison.Ordinal))
            {
                return Encoding.UTF32;
            }
            else if (String.Equals(exportExtensionKeyName, OemExportExtensionKeyName, StringComparison.Ordinal))
            {
                return Encoding.Default;
            }
            else
            {
                throw new Exception("Undefined export extension key name.");
            }
        }
    }
}