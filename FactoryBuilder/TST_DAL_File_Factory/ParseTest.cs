using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TST_DAL_File_Factory
{
    [TestClass]
    public class ParseTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            string sFileName = @"C:\Users\David\Desktop\ForWork\DAL_TOP_AM\DAL_TOP_AM\Factory\OMNI\FileFactory.cs";
            string s = DAL_SQL_Server.Factory.CREATE_ENTITY.ParseText.ParseFactory.ParseFile(sFileName);

        }
    }
}
