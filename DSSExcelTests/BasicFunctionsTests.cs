﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DSSExcelPlugin;
using System.Linq;
using System.Collections.Generic;
using System.Data;

namespace DSSExcelTests
{
    [TestClass]
    public class BasicFunctionsTests
    {
        [TestMethod]
        public void GetTimeSeriesTableFromExcel()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\test.xlsx");
            var table = de.ExcelToDataTable("sheet1");
            List<object> headers = table.Rows[0].ItemArray.ToList();
            var t = headers[0].GetType();
            var h = new List<object>() { "h1", "y1", "x2", "y2" };
            foreach (DataRow item in table.Rows)
            {
                System.Diagnostics.Debug.WriteLine("{0} {1} {2} {3}", item[0].ToString(), item[1].ToString(), item[2].ToString(), 
                    item[3].ToString());
            }
            Assert.AreEqual(t, typeof(string));
            Assert.IsTrue(headers.SequenceEqual(h));
            Assert.AreEqual(table.Columns.Count, 4);
            Assert.AreEqual(table.Rows.Count, 4);
        }
    }
}