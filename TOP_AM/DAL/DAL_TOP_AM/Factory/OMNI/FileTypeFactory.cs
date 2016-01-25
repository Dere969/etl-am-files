using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DAL_TOP_AM.Entities;

namespace DAL_TOP_AM.Factory.OMNI
{
    public class FileTypeFactory
    {
        #region Constants

        public const string cFileName                   = @"C:\SVN_Workspace\etl-files-am\trunk\OMNI_Trades.txt";
        public const bool   cBlnFirstRowContainsHeader  = true;
        public const string cDelimiter                  = "|";

        public const Int16 cPurchaseOrSale              = 0;        
        public const Int16 cNumOfSecurities             = 1;
        public const Int16 cPrice                       = 2;
        public const Int16 cCCY                         = 3;

        #endregion

        #region Select

        public static Trade_OMNI Select(string lineIn)
        {
            Trade_OMNI response = new Trade_OMNI();
            try
            {                
                List<string> data = null;
                data = lineIn.Split(cDelimiter[0]).ToList();

                response.PurchaseOrSale     = data[cPurchaseOrSale];
                response.NumOfSecurities    = FileTypeParseFactory.NumOfSecuities(data[cNumOfSecurities]);
                response.Price              = ParsePrice(data[cPrice]);
                response.CCY                = data[cCCY];

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return response;
        }

        #endregion

        #region Methods.Parse

        public static Int32 ParseNumOfSecuities(string textIn)
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

        public static decimal ParsePrice(string textIn)
        {
            decimal response = 0;
            try
            {
                response = decimal.Parse(textIn);
            }
            catch (Exception ex)
            {
                //throw ex;
            }
            return response;
        }

        #endregion
    }
}