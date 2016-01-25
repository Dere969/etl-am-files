using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAL_TOP_AM.Entities
{
    public class DisclosureTrade
    {
        public string   PurchaseOrSale;
        public int      NumOfSecurities;
        public decimal  Price;
        public string   CCY;
    }
}
