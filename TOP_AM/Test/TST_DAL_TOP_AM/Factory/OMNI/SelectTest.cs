using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DAL_TOP_AM.Entities;

namespace TST_DAL_TOP_AM.Factory.OMNI
{
    [TestClass]
    public class SelectTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            List<Trade_OMNI> list = null;
            list = DAL_TOP_AM.Factory.OMNI.FileFactory.Select();
        }

        [TestMethod]
        public void TestMethod22()
        {
            List<Trade_OMNI> list = null;
            list = DAL_TOP_AM.Factory.OMNI.FileFactory.Select();


        }


    }
}
