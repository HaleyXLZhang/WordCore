using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
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
        public void PastTest() {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.ApplicationClass();
            app.Visible = false;
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(@"C:\Users\Administrator\Desktop\GenerateLetter\PW + ER Reason.doc",true);

            // dynamic range=doc.Tables[1].Cell(5, 3).Range.Borders.OutsideLineStyle=WdLineStyle.wdLineStyleNone;

            Range range = doc.Tables[1].Cell(5, 3).Range;

            doc.Tables[1].Cell(1, 1).Range.AutoFormat();
           // doc.Tables[1].Cell(1, 1).Range.al


            doc.Close();
            app.Quit();
        }






        [TestMethod]
        public void LetterGenerationTest()
        {
            EmployeeInfo employee = new EmployeeInfo() { address = "陕西省西安市雁塔区天谷八路环普科技园1", eRID = "HR342389", language = "C", name = "Haley", title = "TestMessage" };
            EmployerInfo employer = new EmployerInfo() { schemeName = "计划名称", name = "中软国际", schemeNumber = "901213", schemeCode = "CHNSOFT", language = "C", eRID = "HR565", address = "陕西省西安市雁塔区天谷八路环普科技园2" };
            string reasonFile = @"C:\Users\Administrator\Desktop\GenerateLetter\PW + ER Reason.doc";
            string templateFile = @"C:\Users\Administrator\Desktop\GenerateLetter\(chi) ER but missing information.doc";

          

            using (WordCore wordCore = new WordCore())
            {
                List<ReasonInfo> reasons = new List<ReasonInfo>();
                wordCore.OpenWord(@"C:\Users\Administrator\Desktop\GenerateLetter\PW + ER Reason.doc",true);
                IList<string> codes = wordCore.GetTable_Clolumn_ByColumnIndex(1, 1);
                IList<string> shortCodes = wordCore.GetTable_Clolumn_ByColumnIndex(1, 2);
                for (int i = 0; i < codes.Count; i++)
                {
                    reasons.Add(new ReasonInfo() { Code = codes[i], ShortCode = shortCodes[i] });
                }
                wordCore.CopyTable_Cell_ByRowIndexAndColumnIndex(1, 5, 3);

                wordCore.OpenWord(@"C:\Users\Administrator\Desktop\GenerateLetter\(chi) ER but missing information1.doc");
                wordCore.PastToBookmark("reason");

            }

        }

    }
}