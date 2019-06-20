using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

using NUnit.Framework;

namespace DocInOffice.XlsxDelimitedTextConversion.XlsxConverterTesting
{
    public class TemplateTester
    {
        public TemplateTester()
        {
        }

        [TestCaseSource(typeof(TesterEngine), "GetTestSource", new object[] { ".", @"..\..\XlsxConverterTesting\TemplateTester+Files" })]
        public void TestGeneralIssues(String tempFolderName, String sourceSubfolderPath)
        {
            TesterEngine.Test("General", sourceSubfolderPath);
        }
    }
}
