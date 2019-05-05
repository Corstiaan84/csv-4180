//
// Copyright © Leonid S. Maliutin <mals@live.ru> 2014
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

using NUnit.Framework;

namespace DocInOffice.Csv4180.CsvReaderTesting
{
    public class Tester
    {
        private String _testFolderPath = @"..\..\CsvReaderTesterTesting";

        [SetUp]
        public void SetUp()
        {
            _testFolderPath = Path.GetFullPath(Path.Combine(TestContext.CurrentContext.TestDirectory, @"..\..\CsvReaderTesting"));
        }

        [TestCase("01_total.csv")]
        public void Test01(String fileName)
        {
            using (StreamReader streamReader = new StreamReader(Path.Combine(_testFolderPath, fileName)))
            {
                CsvReader reader = new CsvReader(streamReader);
                
                String[] fields1 = reader.ReadFields();

                Assert.That(fields1, Is.Not.Null);
                Assert.That(fields1, Has.Length.EqualTo(2));
                Assert.That(fields1[0], Is.EqualTo("1"));
                Assert.That(fields1[1], Is.EqualTo("один"));

                String[] fields2 = reader.ReadFields();

                Assert.That(fields2, Is.Not.Null);
                Assert.That(fields2, Has.Length.EqualTo(2));
                Assert.That(fields2[0], Is.EqualTo("2"));
                Assert.That(fields2[1], Is.EqualTo("два"));

                String[] fields3 = reader.ReadFields();

                Assert.That(fields3, Is.Not.Null);
                Assert.That(fields3, Has.Length.EqualTo(2));
                Assert.That(fields3[0], Is.EqualTo("3"));
                Assert.That(fields3[1], Is.EqualTo("строка с пробелами"));


                String[] fields4 = reader.ReadFields();

                Assert.That(fields4, Is.Not.Null);
                Assert.That(fields4, Has.Length.EqualTo(2));
                Assert.That(fields4[0], Is.EqualTo("4"));
                Assert.That(fields4[1], Is.EqualTo(" строка с пробелами и конечными пробелами "));

                String[] fields5 = reader.ReadFields();

                Assert.That(fields5, Is.Not.Null);
                Assert.That(fields5, Has.Length.EqualTo(2));
                Assert.That(fields5[0], Is.EqualTo("5"));
                Assert.That(fields5[1], Is.EqualTo(" строка с пробелами и конечными пробелами в ковычках "));

                String[] fields6 = reader.ReadFields();

                Assert.That(fields6, Is.Not.Null);
                Assert.That(fields6, Has.Length.EqualTo(2));
                Assert.That(fields6[0], Is.EqualTo("6"));
                Assert.That(fields6[1], Is.EqualTo(@"строка ""с"" ковычками"));

                String[] fields7 = reader.ReadFields();

                Assert.That(fields7, Is.Not.Null);
                Assert.That(fields7, Has.Length.EqualTo(2));
                Assert.That(fields7[0], Is.EqualTo("7"));
                Assert.That(fields7[1], Is.EqualTo(@"строка ""с"" ковычками"));

                String[] fields8 = reader.ReadFields();

                Assert.That(fields8, Is.Not.Null);
                Assert.That(fields8, Has.Length.EqualTo(2));
                Assert.That(fields8[0], Is.EqualTo("8"));
                Assert.That(fields8[1], Is.EqualTo(@"первая стока значения
вторя строка значения"));

                Assert.That(reader.ReadFields(), Is.Null);

            }
        }

        [TestCase("03_UTF_without_BOM.csv")]
        public void Test03(String fileName)
        {
            using (StreamReader streamReader = new StreamReader(Path.Combine(_testFolderPath, fileName)))
            {
                CsvReader reader = new CsvReader(streamReader);

                String[] fields1 = reader.ReadFields();

                Assert.That(fields1, Is.Not.Null);
                Assert.That(fields1, Has.Length.EqualTo(2));
                Assert.That(fields1[0], Is.EqualTo("1"));
                Assert.That(fields1[1], Is.EqualTo("один"));

                String[] fields2 = reader.ReadFields();

                Assert.That(fields2, Is.Not.Null);
                Assert.That(fields2, Has.Length.EqualTo(2));
                Assert.That(fields2[0], Is.EqualTo("2"));
                Assert.That(fields2[1], Is.EqualTo("два"));

                String[] fields3 = reader.ReadFields();

                Assert.That(fields3, Is.Not.Null);
                Assert.That(fields3, Has.Length.EqualTo(2));
                Assert.That(fields3[0], Is.EqualTo("3"));
                Assert.That(fields3[1], Is.EqualTo("строка с пробелами"));

                String[] fields4 = reader.ReadFields();

                Assert.That(fields4, Is.Not.Null);
                Assert.That(fields4, Has.Length.EqualTo(2));
                Assert.That(fields4[0], Is.EqualTo("4"));
                Assert.That(fields4[1], Is.EqualTo(" строка с пробелами и конечными пробелами "));

                String[] fields5 = reader.ReadFields();

                Assert.That(fields5, Is.Not.Null);
                Assert.That(fields5, Has.Length.EqualTo(2));
                Assert.That(fields5[0], Is.EqualTo("5"));
                Assert.That(fields5[1], Is.EqualTo(" строка с пробелами и конечными пробелами в ковычках "));

                String[] fields6 = reader.ReadFields();

                Assert.That(fields6, Is.Not.Null);
                Assert.That(fields6, Has.Length.EqualTo(2));
                Assert.That(fields6[0], Is.EqualTo("6"));
                Assert.That(fields6[1], Is.EqualTo(@"строка ""с"" ковычками"));

                String[] fields7 = reader.ReadFields();

                Assert.That(fields7, Is.Not.Null);
                Assert.That(fields7, Has.Length.EqualTo(2));
                Assert.That(fields7[0], Is.EqualTo("7"));
                Assert.That(fields7[1], Is.EqualTo(@"первая стока значения
вторя строка значения"));

                Assert.That(reader.ReadFields(), Is.Null);

            }
        }

        [TestCase("04_last_CRLF.csv", TestName="The file has more then one columns and last CRLF")]
        [TestCase("04_without_last_CRLF.csv", TestName = "The file has more then one columns and it does not have last CRLF")]
        public void Test04(String fileName)
        {
            using (StreamReader streamReader = new StreamReader(Path.Combine(_testFolderPath, fileName)))
            {
                CsvReader reader = new CsvReader(streamReader);

                String[] fields1 = reader.ReadFields();

                Assert.That(fields1, Is.Not.Null);
                Assert.That(fields1, Has.Length.EqualTo(2));
                Assert.That(fields1[0], Is.EqualTo("1"));
                Assert.That(fields1[1], Is.EqualTo("один"));

                String[] fields2 = reader.ReadFields();

                Assert.That(fields2, Is.Not.Null);
                Assert.That(fields2, Has.Length.EqualTo(2));
                Assert.That(fields2[0], Is.EqualTo("2"));
                Assert.That(fields2[1], Is.EqualTo("два"));

                Assert.That(reader.ReadFields(), Is.Null);
            }
        }

        [TestCase("05_1_column_1_row.csv", TestName = "Colunms: 1, Rows: 1")]
        public void Test05(String fileName)
        {
            using (StreamReader streamReader = new StreamReader(Path.Combine(_testFolderPath, fileName)))
            {
                CsvReader reader = new CsvReader(streamReader);

                String[] fields1 = reader.ReadFields();

                Assert.That(fields1, Is.Not.Null);
                Assert.That(fields1, Has.Length.EqualTo(1));
                Assert.That(fields1[0], Is.EqualTo("1"));

                Assert.That(reader.ReadFields(), Is.Null);

            }
        }

        [TestCase("06_1_column_2_rows.csv", TestName = "Colunms: 1, Rows: 2")]
        public void Test06(String fileName)
        {
            using (StreamReader streamReader = new StreamReader(Path.Combine(_testFolderPath, fileName)))
            {
                CsvReader reader = new CsvReader(streamReader);

                String[] fields1 = reader.ReadFields();

                Assert.That(fields1, Is.Not.Null);
                Assert.That(fields1, Has.Length.EqualTo(1));
                Assert.That(fields1[0], Is.EqualTo("1"));

                String[] fields2 = reader.ReadFields();

                Assert.That(fields2, Is.Not.Null);
                Assert.That(fields2, Has.Length.EqualTo(1));
                Assert.That(fields2[0], Is.EqualTo("2"));

                Assert.That(reader.ReadFields(), Is.Null);
            }
        }

        [TestCase("07_1_column_3_rows_last_empty.csv", TestName = "Colunms: 1, Rows: 2, Last row is empty")]
        public void Test07(String fileName)
        {
            using (StreamReader streamReader = new StreamReader(Path.Combine(_testFolderPath, fileName)))
            {
                CsvReader reader = new CsvReader(streamReader);

                String[] fields1 = reader.ReadFields();

                Assert.That(fields1, Is.Not.Null);
                Assert.That(fields1, Has.Length.EqualTo(1));
                Assert.That(fields1[0], Is.EqualTo("1"));

                String[] fields2 = reader.ReadFields();

                Assert.That(fields2, Is.Not.Null);
                Assert.That(fields2, Has.Length.EqualTo(1));
                Assert.That(fields2[0], Is.EqualTo("2"));

                Assert.That(reader.ReadFields(), Is.Null);
            }
        }
    }
}
