using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace BLL_Disclosure_Forms
{
    public class CreateDouments
    {

        public static string cRootDisclosurePath            = @"C:\SVN_Workspace\etl-files-am\trunk\Disclosures";
        public static string cFileNameTemplateForm8point3   = @"YYMMDD-Form-8.7-SECURITYNAME-DISCLOSERNAME-Final-version.docx";

        public static void Run()
        {
            try
            {
                string FileName = string.Format("{0}\\{1}", cRootDisclosurePath, cFileNameTemplateForm8point3);
                
                byte[] byteArray = File.ReadAllBytes(FileName);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                    {
                        foreach (Text textsection in doc.MainDocumentPart.Document.Descendants<Text>())
                        {
                            // do something here
                            if (textsection.InnerText.Contains("#"))
                            {
                                textsection.Text = "replaced";
                            }

                        }
                    }
                    string NewFIleName = string.Format("{0}\\{1:yyyy_MM_dd_HH_mm_ss_fff}-{2}", cRootDisclosurePath, DateTime.Now, "8.7.docx");
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
        
        [Obsolete]
        public static void xxxRun()
        {
            try
            {

                bool IsEditable = true;
                WordprocessingDocument doc = WordprocessingDocument.Open(string.Format("{0}\\{1}", cRootDisclosurePath, cFileNameTemplateForm8point3), IsEditable);
               
                foreach (Text textsection in doc.MainDocumentPart.Document.Descendants<Text>())
                {
                    // do something here
                    if (textsection.InnerText.Contains("#"))
                    {
                        textsection.Text  = "replaced";
                    }

                }
                
                doc.Close();
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


    }
}
