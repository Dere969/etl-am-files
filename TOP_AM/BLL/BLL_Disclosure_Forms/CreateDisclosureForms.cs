using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BLL_Disclosure_Forms
{
    public class CreateDisclosureForms
    {
        public static void Run()
        {
            try
            {
                _8point7_.CreateDisclosureForms.Run();
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }
    }
}
