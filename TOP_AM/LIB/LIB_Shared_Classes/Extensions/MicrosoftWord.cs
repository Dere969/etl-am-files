using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace LIB_Shared_Classes.Extensions
{
    public static class MicrosoftWord
    {

        #region Attributes

        public delegate void ProcessDocument(ref Word.Document DocumentIn);

        #endregion

        #region Methods.Public

        public static void CloneDocument(string FileName, string NewFileName, ProcessDocument ProcessDocumentIn)
        {
            try
            {

                File.Copy(FileName, NewFileName);

                object o = Missing.Value;
                object oFalse = false;
                object oTrue = true;

                Word._Application app = null;
                Word.Documents docs = null;
                Word.Document doc = null;

                object path = NewFileName;

                try
                {
                    app = new Word.Application();
                    app.Visible = false;
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                    docs = app.Documents;
                    doc = docs.Open(ref path, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
                    doc.Activate();

                    ProcessDocumentIn(ref doc);

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
                throw ex;
            }
        }

        public static bool ReplaceText(this Word.Document DocumentIn, string FindTextIn, string ReplaceTextIn, Nullable<Int32> ExpectedReplaceCount)
        {
            bool response = false;
            try
            {

                foreach (Word.Range range in DocumentIn.StoryRanges)
                {
                    Word.Find find = range.Find;
                    object findText = FindTextIn;
                    object replacText = ReplaceTextIn;
                    object replace = Word.WdReplace.wdReplaceAll;
                    object findWrap = Word.WdFindWrap.wdFindContinue;
                    object o = Missing.Value;
                    object oFalse = false;
                    object oTrue = true;

                    response = find.Execute(
                        ref findText, ref o, ref o, ref o, ref oFalse, ref o,
                        ref o, ref findWrap, ref o, ref replacText,
                        ref replace, ref o, ref o, ref o, ref o);

                    Marshal.FinalReleaseComObject(find);
                    Marshal.FinalReleaseComObject(range);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return response;
        }

        #endregion

    }
}
