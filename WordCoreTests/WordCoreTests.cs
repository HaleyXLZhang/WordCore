using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WordCoreTests.LetterGenerationTest;

namespace WordCore.Tests
{
    [TestClass()]
    public class WordCoreTests
    {
        [TestMethod()]
        public void GetWordTablesTest()
        {
            using (WordCore wordCore = new WordCore())
            {
                wordCore.OpenWord(@"C:\Users\Administrator\Desktop\20243177_90272171(L12).doc");
                IList<string> tables = wordCore.GetWordTables();
            }
        }
        [TestMethod]
        public void GetDropDownlistOldVersion()
        {
            using (WordCore wordCore = new WordCore())
            {
                wordCore.OpenWord(@"C:\Users\Administrator\Desktop\新建 Microsoft Word 97 - 2003 文档.doc");
                wordCore.Set_DropDownList_SelectedText("d3", "2");
            }
        }
        [TestMethod]
        public void PastTest()
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.ApplicationClass();
            app.Visible = false;
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(@"C:\Users\Administrator\Desktop\WorkFiles\Letter Automation\Template\NewTemplate\(chi) Cancellation of PW1.doc", false);

            doc.Activate();


            doc.Tables[2].Cell(2, 2).Range.FormFields[1].Result = "44444444";
            int count = doc.FormFields.Count;

            //doc.Tables[1].ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs, false);
            //doc.Tables[2].Cell(1, 3).Range.Text = "33";
            //doc.Tables[3].Rows.Add();
            //object unite = WdUnits.wdStory;
            //app.Selection.EndKey(ref unite, Type.Missing); //将光标移动到文档末尾
            //doc.Tables[3].Rows[1].Cells[1].Range.Paste();
            //doc.Tables[3].Rows.Add();
            //Clipboard.Clear();
            //doc.Protect(WdProtectionType.wdAllowOnlyFormFields, true, Type.Missing, Type.Missing, true);
            //doc.Save();
            doc.Save();
            doc.Close();
            app.Quit();
        }
        [TestMethod]
        public void CopyTest()
        {
            //string reasonFile = @"C:\Users\Administrator\Desktop\GenerateLetter\PW + ER Reason.doc";
            //string templateFile = @"C:\Users\Administrator\Desktop\GenerateLetter\(chi) ER but missing information.doc";

            //using (WordCore wordCore = new WordCore())
            //{
            //    wordCore.Copy(templateFile, @"C:\Users\Administrator\Desktop\GenerateLetter\(chi) ER but missing informationTest1.doc");
            //}

            string len = "\r\newrwer";


        }
        [TestMethod]
        public void LetterGenerationTest()
        {
            EmployeeInfo employee = new EmployeeInfo() { address = "陕西省西安市雁塔区天谷八路环普科技园1", eRID = "HR342389", language = "C", name = "Haley", title = "TestMessage" };
            EmployerInfo employer = new EmployerInfo() { schemeName = "计划名称", name = "中软国际", schemeNumber = "901213", schemeCode = "CHNSOFT", language = "C", eRID = "HR565", address = "陕西省西安市雁塔区天谷八路环普科技园2" };

            string reasonFile = @"C:\Users\Administrator\Desktop\GenerateLetter\PW + ER Reason.doc";
            string templateFile = @"C:\Users\Administrator\Desktop\GenerateLetter\(chi) ER but missing informationTest.doc";

            //if Language=="C" worLettertableIndex=3 Language=="E"  worLettertableIndex=4
            int worLettertableIndex = 3;
            ReasonInfo reasoninfo = new ReasonInfo();

            using (WordCore wordCore = new WordCore())
            {
                wordCore.OpenWord(reasonFile, true);
                IList<string> codes = wordCore.GetTable_Clolumn_ByColumnIndex(1, 1);
                IList<string> shortCodes = wordCore.GetTable_Clolumn_ByColumnIndex(1, 2);
                for (int i = 0; i < codes.Count; i++)
                {
                    reasoninfo.Reasons.Add(new SelectReasonItem
                    {
                        Code = codes[i],
                        ShortCode = shortCodes[i],
                        Row = i + 2,
                        CopyColumn = 3
                    });
                }
                foreach (SelectReasonItem item in reasoninfo.Reasons)
                {
                    wordCore.CopyTable_Cell_ByRowIndexAndColumnIndex(1, item.Row, item.CopyColumn);
                    wordCore.OpenWord(templateFile);
                    wordCore.AppendPasteContentToTable(worLettertableIndex);
                    wordCore.OpenWord(reasonFile, true);
                }
                wordCore.ProtectDocument(templateFile);
            }
        }

        [TestMethod]
        public void FindTextInWord() {
            string strKeyEng = "Should you have any queries or require any assistance";
            string strKeyChi = "如有任何查詢或需要協助，請致電";
            string strAppendEng = "According to our record, we have not received your “Employee Application Form”.  Therefore your address of the above account is not updated.  In order to enable us to set up a complete member record, please download the “Employee Application Form” from our website, complete and return together with all necessary documents to us for the above account.\r\n";
            string strAppendChi = "根據我們的紀錄，我們尚未收到閣下的「僱員申請表」。因此，閣下於上述賬戶的地址並未更新。請閣下於我們的網址下載有關「僱員申請表」，填妥及連同所需文件遞交予我們，以便我們為閣下設立完整的成員紀錄。\r\n";
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.ApplicationClass();
            app.Visible = false;
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(@"C:\Users\Administrator\Desktop\LetterTool-DIS batch20181005\[Test](eng) Letter for DIS TI Account_20170405.docx");
            doc.Activate();     
            int i = 0, iCount = 0;
            Microsoft.Office.Interop.Word.Find wfnd;
            if (doc.Paragraphs != null && doc.Paragraphs.Count > 0)
            {
                iCount = doc.Paragraphs.Count;
                for (i = 1; i <= iCount; i++)
                {
                    wfnd = doc.Paragraphs[i].Range.Find;
                    wfnd.ClearFormatting();
                    wfnd.Text = strKeyEng;
                    if (wfnd.Execute(Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing))
                    {
                        doc.Paragraphs[i-2].Range.Text = strAppendEng;
                        break;
                    }
                }
            }
            doc.Save();
            doc.Close();
            app.Quit();
        }


        [TestMethod]
        public void SerachTextTest()
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.ApplicationClass();
            app.Visible = false;
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(@"C:\Users\Administrator\Desktop\Solution Design Documnets\Letter Generation Automation Tool User Requirements v1.0.docx", false);
            doc.Activate();
            
            object unite = WdUnits.wdStory;
            //   app.Selection.EndKey(ref unite, Type.Missing); //将光标移动到文档末尾
            app.Selection.WholeStory();
            app.Selection.Find.Forward = true;
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.MatchWholeWord = true;
            app.Selection.Find.MatchCase = false;
            app.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
            bool isFind = app.Selection.Find.Execute("Log文件夹");
            int start = app.Selection.Range.Start;
            int end = app.Selection.Range.End;
            int length = app.Selection.Range.StoryLength;
            Microsoft.Office.Interop.Word.Range range = app.Selection.Range;
            object p = range.Information[WdInformation.wdActiveEndPageNumber];

            range.SetRange(end - 1, app.ActiveDocument.Content.End);
            
            int MoveStartWhileCount = range.MoveStartUntil("\r", WdConstants.wdBackward);


       


            int getStart = range.Start;
            int getEnd = range.End;
            range.Select();
            int paragraphsCount = range.Paragraphs.Count;
            range.Find.Forward = true;
            range.Find.ClearFormatting();
            range.Find.MatchWholeWord = true;
            range.Find.MatchCase = false;
            range.Find.Wrap = WdFindWrap.wdFindContinue;
            bool isFind2 = app.Selection.Range.Find.Execute("Log文件夹");

            //doc.Protect(WdProtectionType.wdAllowOnlyFormFields, true, Type.Missing, Type.Missing, true);

            doc.Save();
            doc.Close();
            app.Quit();
        }





    }
}