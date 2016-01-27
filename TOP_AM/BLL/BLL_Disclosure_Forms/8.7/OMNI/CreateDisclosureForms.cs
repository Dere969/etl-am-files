using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text.RegularExpressions;

using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Runtime.InteropServices;

namespace BLL_Disclosure_Forms
{
    public class CreateDisclosureForms
    {

        #region Attributes

        private static List<DAL_TOP_AM.Entities.Trade_OMNI> list            = null;
        private static List<DAL_TOP_AM.Entities.Trade_OMNI> listPurchased   = null;
        private static List<DAL_TOP_AM.Entities.Trade_OMNI> listSales       = null;

        #endregion

        public static void Run()
        {
            try
            {

                list            = DAL_TOP_AM.Factory.OMNI.FileFactory.Select().OrderBy(x => x.PurchaseOrSale).ThenBy(x => x.Price).ToList();
                listPurchased   = list.Where(x => x.PurchaseOrSale.ToLower() == "purchase").ToList().OrderBy(x => x.Price).ToList();
                listSales       = list.Where(x => x.PurchaseOrSale.ToLower() == "sale").ToList().OrderBy(x => x.Price).ToList(); 
               
                Run(list);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        public static void UpdateFields(ref string TextIn)
        {
            try
            {
                UpdateEFM(ref TextIn);
                UpdateClass(ref TextIn);
                UpdatePurchaseOrSale(ref TextIn);
                UpdateNumOfSecurities(ref TextIn);
                UpdatePrice(ref TextIn);
                UpdateTotalClass(ref TextIn);
                UpdateTotalPurchased(ref TextIn);
                UpdateTotalSold(ref TextIn);
                UpdateDateOfDisclosure(ref TextIn);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateEFM(ref string TextIn)
        {
            try
            {
                string tFindTag              = "#efm#";
                string tNewText              = "JPMorgan International Bank";
                Int32  tExpectedReplaceCount = 1;

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateClass(ref string TextIn)
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

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdatePurchaseOrSale(ref string TextIn)
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

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateNumOfSecurities(ref string TextIn)
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

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdatePrice(ref string TextIn)
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

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount);
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateTotalPurchased(ref string TextIn)
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

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateTotalSold(ref string TextIn)
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

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateTotalClass(ref string TextIn)
        {
            try
            {
                string tFindTag              = "#TotalClass#";
                string tNewText              = "Ordinary Share";
                Int32  tExpectedReplaceCount = 1;

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void UpdateDateOfDisclosure(ref string TextIn)
        {
            try
            {
                string tFindTag              = "#dateofdisclosure#";
                string tNewText              = DateTime.Now.Date.ToString("dd/MM/yyyy");
                Int32  tExpectedReplaceCount = 1;

                ReplaceText(ref TextIn, tFindTag, tNewText, tExpectedReplaceCount); 
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }
      
        public static void xxxupdatefileds(ref Text TextIn)
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
                    case "class":
                        string sClass = string.Empty;
                        for (Int32 i = 0; i < list.Count; i++)
                        {
                            if (i > 0)
                                sClass += Environment.NewLine;
                            sClass += "Ordinary Share";
                        }
                        TextIn.Text = sClass;
                        break;
                    case "purchaseorsale":
                        string sPurchaseOrSale = string.Empty;
                        for (Int32 i = 0; i < list.Count; i++)
                        {
                            if (i > 0)
                                sPurchaseOrSale += Environment.NewLine;
                            sPurchaseOrSale += list[i].PurchaseOrSale;
                        }
                        TextIn.Text = sPurchaseOrSale;
                        break;
                    case "numofsecurities":
                        string sNumOfSecurities = string.Empty;
                        for (Int32 i = 0; i < list.Count; i++)
                        {
                            if (i > 0)
                                sNumOfSecurities += Environment.NewLine;
                            sNumOfSecurities += list[i].NumOfSecurities;
                        }
                        TextIn.Text = sNumOfSecurities;
                        break;
                    case "price":
                        string sPrice = string.Empty;
                        for (Int32 i = 0; i < listPurchased.Count; i++)
                        {
                            if (sPrice.Length > 0)
                                sPrice += Environment.NewLine;
                            sPrice += listPurchased[i].Price.ToString("0.0000");
                        }
                        for (Int32 i = 0; i < listSales.Count; i++)
                        {
                            if (sPrice.Length > 0)
                                sPrice += Environment.NewLine;
                            sPrice += listSales[i].Price.ToString("0.0000");
                        }
                        TextIn.Text = sPrice;
                        break;
                    case "totalclass":
                        TextIn.Text = "Ordinary Share";
                        break;
                    case "totalpurchased":
                        TextIn.Text = listPurchased.Count.ToString();
                        break;
                    case "totalsold":
                        TextIn.Text = listSales.Count.ToString();
                        break;
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
        private static Body body = null;
        private static void Runxx(List<DAL_TOP_AM.Entities.Trade_OMNI> ListIn)
        {
            try
            {
                string FileName     = string.Format("{0}\\{1}", Constants.cRootDisclosurePath, Constants.cFileNameTemplateForm8point7);               
                string NewFIleName  = string.Format("{0}\\{1:yyyy_MM_dd_HH_mm_ss_fff}-{2}", Constants.cRootDisclosurePath, DateTime.Now, "8.7.doc");
                string docText = null;

                File.Copy(FileName, NewFIleName);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(NewFIleName, true))
                {                    
                    using (StreamReader sr = new StreamReader(doc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }
                    //Paragraph v= doc.MainDocumentPart.Document.Descendants<Paragraph>();

                    //WordTools.ReplaceText((Paragraph)body, "#TotalPurchased#", "3");
                    UpdateFields(ref docText);

                    //docText = docText.Replace(Environment.NewLine, "\v");

                    using (StreamWriter sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
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

        private static Word.Document _msdoc = null;
        private static object o = Missing.Value;
        private static object oFalse = false;
        private object oTrue = true;
         private static void Run(List<DAL_TOP_AM.Entities.Trade_OMNI> ListIn)
        {
            try
            {
                string FileName     = string.Format("{0}\\{1}", Constants.cRootDisclosurePath, Constants.cFileNameTemplateForm8point7);               
                string NewFIleName  = string.Format("{0}\\{1:yyyy_MM_dd_HH_mm_ss_fff}-{2}", Constants.cRootDisclosurePath, DateTime.Now, "8.7.doc");
                string docText = null;

                File.Copy(FileName, NewFIleName);
            
                object o = Missing.Value;
                object oFalse = false;
                object oTrue = true;

                Word._Application app = null;
                Word.Documents docs = null;
                Word.Document doc = null;

                object path = NewFIleName;

                try
                {
                    app = new Word.Application();
                    app.Visible = false;
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                    docs = app.Documents;
                    doc = docs.Open(ref path, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
                    doc.Activate();

                    _msdoc = doc;

                    string s = "";
                    UpdateFields(ref s);
                    
                    doc.Save();
                    ((Word._Document)doc).Close(ref o, ref o, ref o);
                    app.Quit(ref o, ref o, ref o);
                }
                finally
                {
                    if (doc != null)
                        Marshal.FinalReleaseComObject(doc);

                    if (docs != null)
                        Marshal.FinalReleaseComObject(docs);

                    if (app != null)
                        Marshal.FinalReleaseComObject(app);
                }
            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        private static void ReplaceText(ref string TextIn, string FindTextIn, string ReplaceTextIn, Nullable<Int32> ExpectedReplaceCount)
        {
            try
            {

                foreach (Word.Range range in _msdoc.StoryRanges)
                    {
                        Word.Find find = range.Find;
                        object findText = FindTextIn;
                        object replacText = ReplaceTextIn;
                        object replace = Word.WdReplace.wdReplaceAll;
                        object findWrap = Word.WdFindWrap.wdFindContinue;

                        find.Execute(ref findText, ref o, ref o, ref o, ref oFalse, ref o,
                            ref o, ref findWrap, ref o, ref replacText,
                            ref replace, ref o, ref o, ref o, ref o);

                        Marshal.FinalReleaseComObject(find);
                        Marshal.FinalReleaseComObject(range);
                    }


                //Regex regexText = new Regex(FindTextIn);
                //MatchCollection mCollection = regexText.Matches(TextIn);

                //if (ExpectedReplaceCount.HasValue && mCollection.Count != ExpectedReplaceCount.Value)
                //    throw new ApplicationException(string.Format("Error Cannot Find Text '{0}': Expected Count {1}. Actual {2}. (ReplaceTagText)", FindTextIn, ExpectedReplaceCount, mCollection.Count));

                //TextIn = regexText.Replace(TextIn, ReplaceTextIn);

            }
            catch (Exception ex)
            {
                DAL_TOP_AM.Factory.LogEntry.InsertFactory.Insert(ex.Message, ex.StackTrace);
                throw ex;
            }
        }

        #endregion

        #region Methods.Alternative

        private static void RunOld(List<DAL_TOP_AM.Entities.Trade_OMNI> ListIn)
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
                            //Text refText = textsection;
                            //UpdateFields(ref refText);
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
