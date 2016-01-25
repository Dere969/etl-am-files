using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAL_TOP_AM.Entities
{
    public class Trade_OMNI
    {
        #region Attributes

        public string   PurchaseOrSale;
        public int      NumOfSecurities;
        public decimal  Price;
        public string   CCY;
        public string   Name;
        public string   TradeDate;
        public string   TransactionType;
        
        #endregion
    }
}