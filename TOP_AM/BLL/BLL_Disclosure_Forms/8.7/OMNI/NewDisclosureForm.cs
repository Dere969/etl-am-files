using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace BLL_Disclosure_Forms._8point7_.OMNI
{
    public class NewDisclosureForm
    {
        private static List<DAL_TOP_AM.Entities.Trade_OMNI> list = null;
        public static void Run()
        {
            try
            {

                list = DAL_TOP_AM.Factory.OMNI.FileFactory.Select().OrderBy(x => x.PurchaseOrSale).OrderBy(x => x.Price).ToList();

                Run(list);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        public static void UpdateFields(ref Text TextIn)
        {
            try
            {
                string hash = TextIn.Text.ToLower();
                hash        = hash.Replace("#", "");
                List<DAL_TOP_AM.Entities.Trade_OMNI> listPurchased = list.Where(x => x.PurchaseOrSale.ToLower() == "purchase").ToList().OrderBy(x=>x.Price).ToList();
                List<DAL_TOP_AM.Entities.Trade_OMNI> listSales = list.Where(x => x.PurchaseOrSale.ToLower() == "sale").ToList().OrderBy(x => x.Price).ToList(); 
                          
                switch (hash)
                {
                    case "efm":
                        TextIn.Text = "JPMorgan International Bank";
                        break;
                    //case "class":
                    //    string sClass = string.Empty;
                    //    for (Int32 i = 0; i < list.Count; i++)
                    //    {
                    //        if (i > 0)
                    //            sClass += Environment.NewLine;
                    //        sClass += "Ordinary Share";
                    //    }
                    //    TextIn.Text = sClass;
                    //    break;
                    //case "purchaseorsale":
                    //    string sPurchaseOrSale = string.Empty;
                    //    for (Int32 i = 0; i < list.Count; i++)
                    //    {
                    //        if (i > 0)
                    //            sPurchaseOrSale += Environment.NewLine;
                    //        sPurchaseOrSale += list[i].PurchaseOrSale;
                    //    }
                    //    TextIn.Text = sPurchaseOrSale;
                    //    break;
                    //case "numofsecurities":
                    //    string sNumOfSecurities = string.Empty;
                    //    for (Int32 i = 0; i < list.Count; i++)
                    //    {
                    //        if (i > 0)
                    //            sNumOfSecurities += Environment.NewLine;
                    //        sNumOfSecurities += list[i].NumOfSecurities;
                    //    }
                    //    TextIn.Text = sNumOfSecurities;
                    //    break;
                    //case "price":
                    //    string sPrice = string.Empty;
                    //    for (Int32 i = 0; i < listPurchased.Count; i++)
                    //    {
                    //        if (sPrice.Length > 0)
                    //            sPrice += Environment.NewLine;
                    //        sPrice += listPurchased[i].Price.ToString("0.0000");
                    //    }
                    //    for (Int32 i = 0; i < listSales.Count; i++)
                    //    {
                    //        if (sPrice.Length > 0)
                    //            sPrice += Environment.NewLine;
                    //        sPrice += listSales[i].Price.ToString("0.0000");
                    //    }
                    //    TextIn.Text = sPrice;
                    //    break;
                    //case "totalclass":                        
                    //    TextIn.Text = "Ordinary Share";
                    //    break;
                    //case "totalpurchased":
                    //    TextIn.Text = listPurchased.Count.ToString();
                    //    break;
                    //case "totalsold":
                    //     TextIn.Text = listSales.Count.ToString();
                    //    break;
                    case "dateofdisclosure":
                        TextIn.Text = DateTime.Now.Date.ToString("dd/mm/yyyy");
                        break;
                    default:
                        //TextIn.Text = string.Empty;
                        break;
                }

                TextIn.Text = TextIn.Text.Replace(Environment.NewLine, "<br/>");
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        #region Methods.Helper

        private static void Run(List<DAL_TOP_AM.Entities.Trade_OMNI> ListIn)
        {
            try
            {
                string FileName = string.Format("{0}\\{1}", Constants.cRootDisclosurePath, Constants.cFileNameTemplateForm8point7);
                
                byte[] byteArray = File.ReadAllBytes(FileName);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                    {
                        foreach (Text textsection in doc.MainDocumentPart.Document.Descendants<Text>())
                        {
                            if (!textsection.Text.Contains("#"))
                                continue;
                            Text refText = textsection;                            
                            UpdateFields(ref refText);                            
                        }
                    }
                    string NewFIleName = string.Format("{0}\\{1:yyyy_MM_dd_HH_mm_ss_fff}-{2}", Constants.cRootDisclosurePath, DateTime.Now, "8.7.docx");
                    File.WriteAllBytes(NewFIleName, stream.ToArray());
                }
            }
            catch (IOException ex)
            {
                Int64 err = System.Runtime.InteropServices.Marshal.GetExceptionCode();
                switch (err)
                {
                    case -532462766:
                        throw ex;
                        break;
                    default:
                        break;
                }
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        #endregion

    }
}
