using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TST_BLL_Disclosure_Forms
{
    [TestClass]
    public class CreateTest
    {
        [TestMethod]
        public void RunningCreateTest()
        {
            try
            {
                BLL_Disclosure_Forms.CreateDouments.Run();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        [TestMethod]
        public void _8point7_()
        {
            try
            {
                BLL_Disclosure_Forms._8point7_.OMNI.NewDisclosureForm.Run();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}