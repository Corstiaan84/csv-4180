using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Xml;
using System.Text;

using Ionic.Zip;

namespace DocInOffice.XlsxDelimitedTextConversion
{
    public static class CsvConverter
    {
        public static void Convert(IDelimitedTextReader csvReader, Stream xlsxStream)
        {
            WriteDelegate writeContetnt =
                (en, s) =>
                {
                    GenerateWorksheetPartContent(csvReader, s);
                };

            using (Stream templateStream = PrepareTemplateStream())
            {
                Save(templateStream, writeContetnt, xlsxStream);
            }
        }

        private static Stream PrepareTemplateStream()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();

            return assembly.GetManifestResourceStream("DocInOffice.XlsxDelimitedTextConversion.template.xlsx");
        }

        private static void GenerateWorksheetPartContent(IDelimitedTextReader csvReader, Stream contentStream)
        {   
            XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
            xmlWriterSettings.Encoding = Encoding.Unicode;

            XmlWriter xmlWriter = XmlWriter.Create(contentStream, xmlWriterSettings);

            GenerateWorksheetPartContent(csvReader, xmlWriter);

            xmlWriter.Flush();
        }

        private static void GenerateWorksheetPartContent(IDelimitedTextReader csvReader, XmlWriter xmlWriter)
        {
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            xmlWriter.WriteStartElement("sheetData", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            List<String> listValue = new List<String>();
            Int32 rowNumber = 0;
            Int32 columnNumber;

            for (String[] fields = csvReader.ReadFields();
                 fields != null;
                 fields = csvReader.ReadFields())
            {
                rowNumber += 1;
                columnNumber = 0;

                xmlWriter.WriteStartElement("row", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                foreach (String field in fields)
                {
                    xmlWriter.WriteStartElement("c", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    xmlWriter.WriteAttributeString("s", "1");
                    xmlWriter.WriteAttributeString("t", "inlineStr");
                    xmlWriter.WriteStartElement("is", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    xmlWriter.WriteStartElement("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    xmlWriter.WriteAttributeString("xml", "space", null, "preserve");
                    xmlWriter.WriteValue(field);
                    xmlWriter.WriteEndElement(); // t
                    xmlWriter.WriteEndElement(); // is
                    xmlWriter.WriteEndElement(); // c

                    columnNumber += 1;

                    listValue.Add(GetExcelColumnName(columnNumber) + rowNumber.ToString(CultureInfo.InvariantCulture));
                }

                xmlWriter.WriteEndElement(); // row
            }

            xmlWriter.WriteEndElement(); // sheetData

            if (listValue.Count > 0)
            {
                xmlWriter.WriteStartElement("ignoredErrors", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                xmlWriter.WriteStartElement("ignoredError", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                xmlWriter.WriteAttributeString("sqref", String.Join(" ", listValue.ToArray()));
                xmlWriter.WriteAttributeString("numberStoredAsText", "1");
                xmlWriter.WriteEndElement(); // ignoredErrors
                xmlWriter.WriteEndElement(); // ignoredErrors
            }

            xmlWriter.WriteEndElement(); // worksheet
        }

        private static String GetExcelColumnName(int columnNumber) // http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = System.Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private static void Save(Stream templateStream, WriteDelegate writeContetnt, Stream xlsxStream)
        {
            using (ZipFile zipFile = ZipFile.Read(templateStream))
            {
                ZipEntry zipEntry = zipFile.AddEntry(@"xl\worksheets\sheet.xml", writeContetnt);
                zipFile.Save(xlsxStream);
            }
        }
    }
}
