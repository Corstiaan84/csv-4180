using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;

using NUnit.Framework;

namespace DocInOffice.XlsxDelimitedTextConversion.XlsxConverterTesting
{
    namespace TypeTesting
    {
        public class Tester
        {
            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "Value", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test_01(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }
        
            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "String", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test_02(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }
        
            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "InlineString", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test_03(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }

            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "SharedString", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test_04(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }

            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "Number", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }

            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "Date", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test_05(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }

            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "Boolean", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test_06(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }

            [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { "Error", @"..\..\XlsxConverterTesting\TypeTester+Files" })]
            public void Test_07(String sourceFolderName, String sourceFolderPath)
            {
                TesterEngine.Test(sourceFolderName, sourceFolderPath);
            }
        }
    }
}
