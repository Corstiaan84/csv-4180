using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

using NUnit.Framework;

namespace DocInOffice.Csv4180.CsvWriterTesting
{
    public class Tester
    {
        [TestCase]
        public void Test01()
        {
            StringBuilder stringBuilder = new StringBuilder();

            using (StringWriter stringWriter = new StringWriter(stringBuilder))
            using (CsvWriter csvWrite = new CsvWriter(stringWriter))
            {
                csvWrite.Write("1");
                csvWrite.EndField();
                csvWrite.Write("Привет");
                csvWrite.Write(" ");
                csvWrite.Write("мир");
                csvWrite.Write("!");
                csvWrite.EndField();
                csvWrite.Write("Hello, World!");
                csvWrite.EndRecord();
                csvWrite.Write("2");
                csvWrite.EndField();
                csvWrite.Write("Здрасте!");
                csvWrite.EndField();
                csvWrite.Write("Hi!");
                csvWrite.EndRecord();
            }

            String resultString = stringBuilder.ToString();

            Console.WriteLine("\""+ resultString + "\"");

            String templateString =
@"1,Привет мир!,""Hello, World!""
2,Здрасте!,Hi!
";

            Assert.That(resultString, Is.EqualTo(templateString));
        }

        [TestCase]
        public void Test02()
        {
            StringBuilder stringBuilder = new StringBuilder();

            using (StringWriter stringWriter = new StringWriter(stringBuilder))
            using (CsvWriter csvWrite = new CsvWriter(stringWriter))
            {
                csvWrite.Write("1");
                csvWrite.EndField();
                csvWrite.Write("Привет");
                csvWrite.Write("\r\n");
                csvWrite.Write("мир");
                csvWrite.Write("!");
                csvWrite.EndRecord();
                csvWrite.Write("2");
                csvWrite.EndField();
                csvWrite.Write("Hi!");
                csvWrite.EndRecord();
            }

            String resultString = stringBuilder.ToString();

            Console.WriteLine("\"" + resultString + "\"");

            String templateString =
@"1,""Привет
мир!""
2,Hi!
";

            Assert.That(resultString, Is.EqualTo(templateString));
        }

        [TestCase]
        public void Test03()
        {
            StringBuilder stringBuilder = new StringBuilder();

            using (StringWriter stringWriter = new StringWriter(stringBuilder))
            using (CsvWriter csvWrite = new CsvWriter(stringWriter))
            {
                csvWrite.Write("1");
                csvWrite.EndField();
                csvWrite.Write("Привет \"мир\"!");
                csvWrite.EndRecord();
                csvWrite.Write("2");
                csvWrite.EndField();
                csvWrite.Write("Hi!");
                csvWrite.EndRecord();
            }

            String resultString = stringBuilder.ToString();

            Console.WriteLine("\"" + resultString + "\"");

            String templateString =
@"1,""Привет """"мир""""!""
2,Hi!
";

            Assert.That(resultString, Is.EqualTo(templateString));
        }

        [TestCase]
        public void Test04()
        {
            StringBuilder stringBuilder = new StringBuilder();

            using (StringWriter stringWriter = new StringWriter(stringBuilder))
            using (CsvWriter csvWrite = new CsvWriter(stringWriter))
            {
                csvWrite.Write("1");
                csvWrite.EndField();
                csvWrite.Write("Привет, мир!");
                csvWrite.EndRecord();
                csvWrite.Write("2");
                csvWrite.EndField();
                csvWrite.Write("Hi!");
                csvWrite.EndRecord();
            }

            String resultString = stringBuilder.ToString();

            Console.WriteLine("\"" + resultString + "\"");

            String templateString =
@"1,""Привет, мир!""
2,Hi!
";

            Assert.That(resultString, Is.EqualTo(templateString));
        }
    }
}