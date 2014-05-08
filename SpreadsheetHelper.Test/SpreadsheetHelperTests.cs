using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SpreadsheetHelper.Test
{
    [TestClass]
    public class SpreadsheetHelper_Tests
    {
        [TestMethod]
        public void SpreadsheetHelper()
        {
            Spreadsheet doc = new Spreadsheet();
            Assert.IsNotNull(doc);
        }
    }
}
