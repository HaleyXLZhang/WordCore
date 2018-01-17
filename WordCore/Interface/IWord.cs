using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordCore.Interface
{
    public interface IWord:IDisposable
    {
        void CreateWord(string savePath);
        void OpenWord(string fileName);
        void OpenWord(string fileName, bool isReadOnly);
        IList<string> GetTable_Clolumn_ByColumnIndex(int tableIndex, int columnIndex);
        void CopyTable_Cell_ByRowIndexAndColumnIndex(int tableIndex, int rowIndex, int columnIndex);
        void AppendPasteContentToTable(int tableIndex);
        void PasteToBookmark(string bookMarkName);
        void SetTableCellValue(int tableIndex, int rowIndex, int columnIndex, string value);
        IList<string> GetWordTables();
        void Copy(string sourceWordFile, string destinationWordFile);
        void SaveAs(string strFileName);
        void Set_ComboBox_SelectedText(string comboBoxTitle, string selectedText);
        void Set_DropDownList_SelectedText(string bookMark, string selectedText);
        void GotoBookMark(string strBookMarkName);
        void InsertText(string strBookMarkName, string text);
        void Save();
        void Quit();
    }
}
