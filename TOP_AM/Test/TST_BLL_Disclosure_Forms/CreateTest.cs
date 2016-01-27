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
        public void CreateDisclosureFormsTest()
        {
            try
            {
                BLL_Disclosure_Forms.CreateDisclosureForms.Run();                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}