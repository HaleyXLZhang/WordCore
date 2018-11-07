using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordCore.CommClass;

namespace WordCore.Interface
{
    public interface IWord:IDisposable
    {
        void CreateWord(string directoryAndFileName, EmunSet.WdSaveFormat format = EmunSet.WdSaveFormat.wdFormatDocument);
        void OpenWord(string fileName);
        void AppendContentToFirstParagraphs(string text);
        void OpenWord(string fileName, bool isReadOnly);
        IList<string> GetTable_Clolumn_ByColumnIndex(int tableIndex, int columnIndex);
        void CopyTable_Cell_ByRowIndexAndColumnIndex(int tableIndex, int rowIndex, int columnIndex);
        void AppendPasteContentToTable(int tableIndex);
        void PasteToBookmark(string bookMarkName);
        void SetTableCellValue(int tableIndex, int rowIndex, int columnIndex, string value);
        IList<string> GetWordTables();
        void Copy(string sourceWordFile, string destinationWordFile);
        void SaveAs(string strFileName);
        int SearchActiveDocumentParagraphIndex(string strKeyWords);
        void Set_ComboBox_SelectedText(string comboBoxTitle, string selectedText);
        void Set_DropDownList_SelectedText(string bookMark, string selectedText);
        void GotoBookMark(string strBookMarkName);
        void InsertText(string strBookMarkName, string text);
        void ProtectDocument(string DocumentName);
        void Save();
        void Quit();
    }
}
