using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DSSExcelPlugin;
using System.Linq;
using System.Collections.Generic;

namespace DSSExcelTests
{
    [TestClass]
    public class BasicFunctionsTests
    {
        [TestMethod]
        public void GetTimeSeriesTableFromExcel()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\test.xlsx");
            var table = de.ExcelToTimeSeriesTable("sheet1");
            List<object> headers = table.Rows[0].ItemArray.ToList();
            var t = headers[0].GetType();
            var h = new List<object>() { "h1", "y1", "x2", "y2" };
            Assert.AreEqual(t, typeof(string));
            Assert.IsTrue(headers.SequenceEqual(h));
            Assert.AreEqual(table.Columns.Count, 4);
            Assert.AreEqual(table.Rows.Count, 4);
        }
    }
}
