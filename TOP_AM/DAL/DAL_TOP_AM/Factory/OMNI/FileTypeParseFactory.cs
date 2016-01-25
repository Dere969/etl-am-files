using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAL_TOP_AM.Factory.OMNI
{
    public class FileTypeParseFactory
    {
        public static Int32 NumOfSecuities(string textIn)
        {
            Int32 response = 0;
            try
            {
                response = Int32.Parse(textIn);
            }
            catch (Exception ex)
            {
                //throw ex;
            }
            return response;
        }

    }
}
