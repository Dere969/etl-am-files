using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text.RegularExpressions;
using LIB_Shared_Classes.Extensions;
using Word = Microsoft.Office.Interop.Word;

namespace BLL_Disclosure_Forms._8point7_
{
    public class CreateDisclosureForms
    {

        #region Attributes

        private static List<DAL_TOP_AM.Entities.Trade_OMNI> list            = null;
        private static List<DAL_TOP_AM.Entities.Trade_OMNI> listPurchased   = null;
        private static List<DAL_TOP_AM.Entities.Trade_OMNI> listSales       = null;

        private static Word.Document                        Document        = null;
        private static string                               FileName        = null;        
        private static string                               NewFileName     = null;

        #endregion

        #region Methods.Public

        public static void Run()
        {
            try
            {

                list            = DAL_TOP_AM.Factory.OMNI.FileFactory.Select().OrderBy(x => x.PurchaseOrSale).ThenBy(x => x.Price).ToList();
                listPurchased   = list.Where(x => x.PurchaseOrSale.ToLower() == "purchase").ToList().OrderBy(x => x.Price).ToList();
                listSales       = list.Where(x => x.PurchaseOrSale.ToLower() == "sale").ToList().OrderBy(x => x.Price).ToList(); 
               
                FileName        = string.Format("{0}\\{1}", Constants.cRootDisclosurePath, Constants.cFileNameTemplateForm8point7);
                NewFileName     = string.Format("{0}\\{1:yyyy_MM_dd_HH_mm_ss_fff}-{2}", Constants.cRootDisclosurePath, DateTime.Now, "8.7.doc");

                MicrosoftWord.CloneDocument(FileName, NewFileName, new MicrosoftWord.ProcessDocument(UpdateFields));
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        #endregion

        #region Methods.Private

        private static void UpdateFields(ref Word.Document DocumentIn)
        {
            try
            {
                Document = DocumentIn;
                UpdateEFM();
                UpdateClass();
                UpdatePurchaseOrSale();
                UpdateNumOfSecurities();
                UpdatePrice();
                UpdateTotalClass();
                UpdateTotalPurchased();
                UpdateTotalSold();
                UpdateDateOfDisclosure();
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateEFM()
        {
            try
            {
                string tFindTag              = "#efm#";
                string tNewText              = "JPMorgan International Bank";
                Int32  tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateClass()
        {
            try
            {
                string tFindTag = "#Class#";
                string tNewText = string.Empty;
                for (Int32 i = 0; i < list.Count; i++)
                {
                    if (i > 0)
                        tNewText += "\v";
                    tNewText += "Ordinary Share";
                }

                Int32 tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdatePurchaseOrSale()
        {
            try
            {
                string tFindTag = "#PurchaseOrSale#";
                string tNewText = string.Empty;
                for (Int32 i = 0; i < list.Count; i++)
                {
                    if (i > 0)
                        tNewText += "\v";//Environment.NewLine;
                    tNewText += list[i].PurchaseOrSale;
                }

                Int32 tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateNumOfSecurities()
        {
            try
            {
                string tFindTag = "#NumOfSecurities#";
                string tNewText = string.Empty;
                for (Int32 i = 0; i < list.Count; i++)
                {
                    if (i > 0)
                        tNewText += "\v";//Environment.NewLine;
                    tNewText += list[i].NumOfSecurities.ToString("###,###,###,###");
                }

                Int32 tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdatePrice()
        {
            try
            {
                string tFindTag = "#Price#";
                string tNewText = string.Empty;
                for (Int32 i = 0; i < listPurchased.Count; i++)
                {
                    if (tNewText.Length > 0)
                        tNewText += "\v";//Environment.NewLine;
                    tNewText += listPurchased[i].Price.ToString("0.0000");
                    tNewText += " " + listPurchased[i].CCY;
                }
                for (Int32 i = 0; i < listSales.Count; i++)
                {
                    if (tNewText.Length > 0)
                        tNewText += "\v";//Environment.NewLine;
                    tNewText += listSales[i].Price.ToString("0.0000");
                    tNewText += " " + listSales[i].CCY;
                }

                Int32 tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateTotalPurchased()
        {
            try
            {
                Int32 tTotal = 0;
                for (Int32 i = 0; i < listPurchased.Count; i++)
                {
                    tTotal += listPurchased[i].NumOfSecurities;
                }

                string tFindTag              = "#TotalPurchased#";
                string tNewText              = tTotal.ToString("###,###,###,###");
                Int32  tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateTotalSold()
        {
            try
            {
                Int32 tTotal = 0;
                for (Int32 i = 0; i < listSales.Count; i++)
                {
                    tTotal += listSales[i].NumOfSecurities;
                }

                string tFindTag              = "#TotalSold#";
                string tNewText              = tTotal.ToString("###,###,###,###");
                Int32  tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateTotalClass()
        {
            try
            {
                string tFindTag              = "#TotalClass#";
                string tNewText              = "Ordinary Share";
                Int32  tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateDateOfDisclosure()
        {
            try
            {
                string tFindTag              = "#dateofdisclosure#";
                string tNewText              = DateTime.Now.Date.ToString("dd/MM/yyyy");
                Int32  tExpectedReplaceCount = 1;

                Document.ReplaceText(tFindTag, tNewText, tExpectedReplaceCount); 
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
