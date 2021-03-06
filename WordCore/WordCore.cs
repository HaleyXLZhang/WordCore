﻿using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using WordCore.Common;
using WordCore.Interface;

namespace WordCore
{
    public class WordCore : IWord
    {
        /// <summary>
        /// Application 对象
        /// </summary>
        dynamic wordApp = null;
        /// <summary>
        /// Document 对象
        /// </summary>
        dynamic wordDoc = null;
        private string openFileName = string.Empty;
        public WordCore()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
            wordApp.Visible = false;
        }
        public void CreateWord(string directoryAndFileName, EmunSet.WdSaveFormat format= EmunSet.WdSaveFormat.wdFormatDocument)
        {   
            wordDoc = wordApp.Documents.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            wordDoc.SaveAs(ref directoryAndFileName, ref format, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
           
            wordDoc.Close(Missing.Value, Missing.Value, Missing.Value);
             
            wordApp.Quit(Missing.Value, Missing.Value, Missing.Value);

            wordDoc = null;

            wordApp = null;
        }
        public void OpenWord(string fileName)
        {
            openFileName = fileName;
            wordDoc = wordApp.Documents.Open(fileName,
                Missing.Value,
                false,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                false,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value);
            wordDoc.Activate();
        }
        public void OpenWord(string fileName, bool isReadOnly)
        {
            openFileName = fileName;
            wordDoc = wordApp.Documents.Open(fileName,
                Missing.Value,
                isReadOnly,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                false,
                Missing.Value,
                Missing.Value,
                Missing.Value,
                Missing.Value);
            wordDoc.Activate();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableIndex">Start with one</param>
        /// <param name="columnIndex">Start with one</param>
        /// <returns></returns>
        public IList<string> GetTable_Clolumn_ByColumnIndex(int tableIndex, int columnIndex)
        {
            List<string> columnRows = new List<string>();
            dynamic nowTable = wordDoc.Tables.Item(tableIndex);

            for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)
            {
                columnRows.Add(nowTable.Cell(rowPos, columnIndex).Range.Text);
            }
            return columnRows;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableIndex">Start with one</param>
        /// <param name="rowIndex">Start with one</param>
        /// <param name="columnIndex">Start with one</param>
        /// <returns></returns>
        public void CopyTable_Cell_ByRowIndexAndColumnIndex(int tableIndex, int rowIndex, int columnIndex)
        {
            dynamic nowTable = wordDoc.Tables.Item(tableIndex);
            dynamic cell = nowTable.Cell(rowIndex, columnIndex).Range;
            cell.Copy();
        }

        public void AppendPasteContentToTable(int tableIndex)
        {
            // object unite = WdUnits.wdStory;
            //  wordApp.Selection.EndKey(ref unite, Type.Missing); //将光标移动到文档末尾
            //  wordDoc.Paragraphs.Last.Range.Paste();
            if (wordDoc.Tables[tableIndex].Rows.Count >= 1)
            {

                string obj = wordDoc.Tables[tableIndex].Rows[wordDoc.Tables[tableIndex].Rows.Count].Cells[1].Range.Text;

                if (string.IsNullOrEmpty(obj) || string.IsNullOrEmpty(obj.Replace("\r\a", "")))
                {
                    wordDoc.Tables[tableIndex].Rows[wordDoc.Tables[tableIndex].Rows.Count].Cells[1].Range.Paste();
                }
                else
                {
                    wordDoc.Tables[tableIndex].Rows.Add();
                    wordDoc.Tables[tableIndex].Rows[wordDoc.Tables[tableIndex].Rows.Count].Cells[1].Range.Paste();
                }
            }
            //  dynamic tables = wordDoc.Tables;
            // TableConvertToText(tables);
            Clipboard.Clear();
           // 
            wordDoc.Save();
        }
        public void ProtectDocument(string DocumentName) {
            OpenWord(DocumentName);
            wordDoc.Protect(WdProtectionType.wdAllowOnlyFormFields, true, Type.Missing, Type.Missing, true);

            wordDoc.Save();
        }
        public void AppendContentToFirstParagraphs(string text)
        {
            wordDoc.Paragraphs[1].Range.Text = wordDoc.Paragraphs[1].Range.Text + text;
        }


        public int SearchActiveDocumentParagraphIndex(string strKeyWords)
        { 
            int i = -1, iCount = -1;
            dynamic wfnd;
            if (wordDoc.Paragraphs != null && wordDoc.Paragraphs.Count > 0)
            {
                iCount = wordDoc.Paragraphs.Count;
                for (i = 1; i <= iCount; i++)
                {
                    wfnd = wordDoc.Paragraphs[i].Range.Find;
                    wfnd.ClearFormatting();
                    wfnd.Text = strKeyWords;
                    if (wfnd.Execute(Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing,
                        Type.Missing))
                    {  
                        break;
                    }
                }
            }
            return i;
        }

        public void PasteToBookmark(string bookMarkName)
        {
            GotoBookMark(bookMarkName);
            // wordApp.Selection.Paste();
            int i = 1;
            for (; i <= wordDoc.Bookmarks.Count; i++)
            {
                if (wordDoc.Bookmarks[i].Name == bookMarkName)
                {
                    wordDoc.Bookmarks[i].Range.Paste();
                    break;
                }
            }
            dynamic tables = wordDoc.Tables;
            TableConvertToText(tables);
            Clipboard.Clear();
            wordDoc.Save();
        }

        public void SetTableCellValue(int tableIndex, int rowIndex, int columnIndex, string value)
        {
            dynamic tables = wordDoc.Tables;
            tables[tableIndex].Cell(rowIndex, columnIndex).Range.Text = value;
            wordDoc.Save();
        }

        private void TableConvertToText(dynamic tables)
        {
            int i;
            int tableCount = tables.Count;
            for (i = 1; i <= tableCount; i++)
            {
                if (i > 3)
                {
                    tables[i].ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs, false);
                }
            }


        }

        public IList<string> GetWordTables()
        {
            List<string> tables = new List<string>();
            for (int tablePos = 1; tablePos <= wordDoc.Tables.Count; tablePos++)
            {
                dynamic nowTable = wordDoc.Tables.Item(tablePos);
                string tableMessage = string.Format("第{0}/{1}个表:\n", tablePos, wordDoc.Tables.Count);

                for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)
                {
                    for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)
                    {
                        tableMessage += nowTable.Cell(rowPos, columPos).Range.Text;
                        tableMessage = tableMessage.Remove(tableMessage.Length - 2, 2);//remove \r\a
                        tableMessage += "\t";
                    }
                    tableMessage += "\n";
                }
                tables.Add(tableMessage);
            }
            return tables;
        }
        /// <summary>
        /// Copy full content from one word document to another
        /// </summary>
        public void Copy(string sourceWordFile, string destinationWordFile)
        {
            if (string.IsNullOrWhiteSpace(sourceWordFile) || string.IsNullOrWhiteSpace(destinationWordFile))
            {
                return;
            }
            OpenWord(sourceWordFile, true);
            wordApp.Selection.WholeStory();
            wordApp.Selection.Copy();
            CreateWord(destinationWordFile);
            wordApp.Selection.Paste();
            SaveAs(destinationWordFile);


        }
        public void SaveAs(string strFileName)
        {
            object fileName = strFileName;
            object missing = Missing.Value;
            wordDoc.SaveAs(fileName, missing, missing, missing, missing, missing, missing,
                               missing, missing, missing, missing, missing, missing, missing, missing, missing);
        }
        public void Set_ComboBox_SelectedText(string comboBoxTitle, string selectedText)
        {
            for (int i = 1; i < wordDoc.ContentControls.Count; i++)
            {
                if (wordDoc.ContentControls[i].Title == comboBoxTitle)
                {
                    for (int j = 1; j < wordDoc.ContentControls[i].DropdownListEntries.Count; j++)
                    {
                        string itemText = wordDoc.ContentControls[i].DropdownListEntries[j].Text;
                        if (itemText == selectedText)
                        {
                            wordDoc.ContentControls[i].DropdownListEntries[j].Select();
                        }
                    }
                }
            }
        }
        public void Set_DropDownList_SelectedText(string bookMark, string selectedText)
        {
            for (int i = 1; i <= wordDoc.Bookmarks.Count; i++)
            {
                string name = wordDoc.Bookmarks[i].Name;
                if (name == bookMark)
                {
                    for (int j = 1; j <= wordDoc.FormFields[bookMark].DropDown.ListEntries.Count; j++)
                    {
                        if (wordDoc.FormFields[bookMark].DropDown.ListEntries[j].Name.Contains(selectedText))
                        {
                            wordDoc.FormFields[bookMark].DropDown.Value = j;
                            break;
                        }
                    }
                    break;
                }
            }
        }
        public void GotoBookMark(string strBookMarkName)
        {
            int i = 1;
            for (; i <= wordDoc.Bookmarks.Count; i++)
            {
                if (wordDoc.Bookmarks[i].Name == strBookMarkName)
                {
                    break;
                }
            }

            if (i <= wordDoc.Bookmarks.Count)
            {
                object bookmark = (int)Common.WdGoToItem.wdGoToBookmark;
                object bookMarkName = strBookMarkName;
                wordDoc.GoTo(ref bookmark, Missing.Value, Missing.Value, ref bookMarkName);
            }
        }
        public void InsertText(string strBookMarkName, string text)
        {
            dynamic bks = wordDoc.Bookmarks;
            for (int i = 1; i <= bks.Count; i++)
            {
                if (bks[i].Name == strBookMarkName)
                {
                    bks[i].Range.Text = text;
                }
            }
        }
        public void Save()
        {
            wordDoc.Save();
        }
        public void Quit()
        {

            if (wordDoc != null) { wordDoc.Close(Type.Missing, Type.Missing, Type.Missing); }
            if (wordApp != null) { wordApp.Quit(); }

        }
        public void Dispose()
        {
            Quit();
        }

       
    }
}
