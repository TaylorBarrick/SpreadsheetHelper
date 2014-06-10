using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using SpreadsheetLight;
using System.ComponentModel;

namespace SpreadsheetHelper.Test
{
    [TestClass]
    public class SpreadsheetHelper_Tests
    {
        List<TestClass> tcs;

        [TestInitialize]
        public void Setup()
        {
            TestClass tc1 = new TestClass() { TestString = "Test" };
            TestClass tc2 = new TestClass() { TestString = "Test" };
            tcs = new List<TestClass>() { tc1, tc2 };
            foreach (string fn in Directory.GetFiles(".", "*testfile.xlsx"))
                File.Delete(fn);
        }

        [TestMethod]
        public void SpreadsheetHelper()
        {
            Spreadsheet doc = new Spreadsheet();
            Assert.IsNotNull(doc);
        }

        [TestMethod]
        public void SpreadsheetHelper_Save()
        {
            Spreadsheet doc = new Spreadsheet();
            doc.CreateAndAppendWorksheet<TestClass>(tcs);
            string filename = string.Format("{0:yyyyMMdd_hhmmss}testfile1.xlsx", DateTime.Now);
            doc.Save(filename);
            Assert.IsTrue(File.Exists(filename));
        }

        [TestMethod]
        public void SpreadsheetHelper_SaveStream()
        {
            Spreadsheet doc = new Spreadsheet();
            doc.CreateAndAppendWorksheet<TestClass>(tcs);
            string filename = string.Format("{0:yyyyMMdd_hhmmss}testfile2.xlsx", DateTime.Now);
            FileStream stream = new FileStream(filename, FileMode.CreateNew);
            doc.Save(stream);
            Assert.IsTrue(File.Exists(filename));
        }

        [TestMethod]
        public void SpreadsheetHelper_OrderProperties()
        {
            Spreadsheet doc = new Spreadsheet();
            List<PropertyInfo> props = doc.OrderProperties(typeof(TestClassWithOrder));
            Assert.AreEqual(props[0].Name, "TestString2");
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void SpreadsheetHelper_Save_MissingPath()
        {
            Spreadsheet doc = new Spreadsheet();
            doc.CreateAndAppendWorksheet<TestClass>(tcs); 
            doc.Save("");
        }

        [TestMethod]
        public void SpreadsheetHelper_Doc_CreateHeader()
        {
            Spreadsheet doc = new Spreadsheet();
            doc.CreateAndAppendWorksheet<TestClass>(tcs);
            PropertyInfo[] props = typeof(TestClass).GetProperties();
            int columnIndex = 1;
            foreach (PropertyInfo prop in typeof(TestClass).GetProperties())
            {
                Assert.AreEqual(prop.Name,doc.doc.GetCellValueAsString(1,columnIndex));
                columnIndex++;
            }
        }

        [TestMethod]
        public void SpreadsheetHelper_DisplayWidth()
        {
            TestClassWithDisplayWidth tc1 = new TestClassWithDisplayWidth();
            Spreadsheet doc = new Spreadsheet();
            doc.CreateAndAppendWorksheet<TestClassWithDisplayWidth>(new List<TestClassWithDisplayWidth>() { tc1 });
            PropertyInfo[] props = typeof(TestClassWithDisplayWidth).GetProperties();
            int columnIndex = 1;
            foreach (PropertyInfo prop in typeof(TestClassWithDisplayWidth).GetProperties())
            {
                foreach (DisplayWidth dw in prop.GetCustomAttributes(typeof(DisplayWidth)))
                {
                    Assert.AreEqual(dw.Width, doc.doc.GetColumnWidth(columnIndex));
                    break;
                }
                columnIndex++;
            }
        }
    }

    public class TestClass
    {
        public string TestString { get; set; }
        public string TestString2 { get; set; }
    }

    public class TestClassWithHide
    {
        [DisplayHide]
        public string TestString { get; set; }
        public string TestString2 { get; set; }
    }

    public class TestClassWithOrder
    {
        public string TestString { get; set; }
        [Display(Order = 1)]
        public string TestString2 { get; set; }
    }

    public class TestClassWithDisplayName
    {
        [DisplayName("NewName1")]
        public string TestString { get; set; }
        [DisplayName("NewName2")]
        public string TestString2 { get; set; }
    }

    public class TestClassWithDisplayWidth
    {
        [DisplayWidth(10)]
        public string TestString { get; set; }
        [DisplayWidth(10)]
        public string TestString2 { get; set; }
    }
}
