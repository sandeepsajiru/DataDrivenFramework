using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beingzero.Framework.Datadriven.ExcelHelper
{
    public interface IExcelHelper
    {
        int GetRowCount(String sheetName);
        int GetColumnCount(String sheetName, int rowNumber);

        //int GetColumnNumber(String sheetName, String colName);
        //int GetColumnNumber(String sheetName, String colName, int startRowIdx);
        //void Save();

        String GetCellData(String sheetName, String colName, int rowNumber);
        String GetCellData(String sheetName, int colNumber, int rowNumber);
        bool SetCellData(String sheetName, String colName, int rowNumber, String data);
        bool SetCellData(String sheetName, int colNumber, int rowNumber, String data);
        void AddSheet(String sheetname);
        bool RemoveSheet(String sheetName);
        void AddColumn(String sheetName, String colName);
        bool SheetExists(String sheetName);
        void RemoveColumn(String sheetName, int colNum);
        int GetRowNumber(String sheetName, String colName, String value);
    }
}
