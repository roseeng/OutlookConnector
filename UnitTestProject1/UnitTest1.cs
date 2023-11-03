using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OutlookConnector;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            MailConnector.OpenEmail(@"C:\temp\test.msg", "goran@roseen.se", @"C:\temp\sql2.sql");
        }
    }
}
