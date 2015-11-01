using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beingzero.Framework.Datadriven.ExcelHelper
{
    public class ExcelHelper : IExcelHelper
    {
        private string filePath;
        XLWorkbook workbook;
        IXLWorksheet worksheet;
        
        //var wb = new XLWorkbook(northwinddataXlsx);

        public ExcelHelper(string filePath)
        {
            this.filePath = filePath;
            workbook = new XLWorkbook(filePath);
            if(workbook.Worksheets.Count>0)
                worksheet = workbook.Worksheet(1);
        }

        private void Save()
        {
            try {
                workbook.Save();
                workbook = new XLWorkbook(filePath);
            }catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        public void AddColumn(string sheetName, string colName)
        {
            worksheet = workbook.Worksheet(sheetName);
            if(worksheet.LastColumnUsed()!=null)
                worksheet.LastColumnUsed().ColumnRight().Cell(1).Value = colName;
            else
                worksheet.Cell("A1").Value = colName;
            Save();
        }

        public void AddSheet(string sheetname)
        {
            workbook.Worksheets.Add(sheetname);
            Save();
        }

        public string GetCellData(string sheetName, int colNumber, int rowNumber)
        {
            return workbook.Worksheet(sheetName).Cell(rowNumber, colNumber).Value.ToString();
        }

        private int GetColumnNumber(String sheetName, String colName)
        {
            worksheet = workbook.Worksheet(sheetName);
            IXLRow r =  worksheet.Row(1);
            while (r != worksheet.LastRowUsed())
            {

                foreach(IXLCell c in r.CellsUsed())
                {
                    if (c.Value.Equals(colName))
                        return c.Address.ColumnNumber; // Starts with 0
                }

                r = r.RowBelow();
            }
            return -1;
        }
        public string GetCellData(string sheetName, string colName, int rowNumber)
        {
            worksheet = workbook.Worksheet(sheetName);
            int colNumber = GetColumnNumber(sheetName, colName);
            return GetCellData(sheetName, colNumber, rowNumber);
        }
        
        public int GetColumnCount(string sheetName, int rowNumber)
        {
            worksheet = workbook.Worksheet(sheetName);
            return worksheet.Row(rowNumber).CellsUsed().Count();
        }

        public int GetRowCount(string sheetName)
        {
            return worksheet.RangeUsed().RowCount();
        }

        public void RemoveColumn(string sheetName, int colNum)
        {
            worksheet = workbook.Worksheet(sheetName);
            worksheet.Column(colNum).Delete();
            Save();
        }

        public bool RemoveSheet(string sheetName)
        {
            workbook.Worksheets.Delete(sheetName);
            Save();
            return true;
        }

        public bool SetCellData(string sheetName, int colNumber, int rowNumber, string data)
        {
            workbook.Worksheet(sheetName).Cell(rowNumber, colNumber).Value = data;
            Save();
            return true;
        }

        public bool SetCellData(string sheetName, string colName, int rowNumber, string data)
        {
            int colNumber = GetColumnNumber(sheetName, colName);
            return SetCellData(sheetName, colNumber, rowNumber, data);
        }

        public bool SheetExists(string sheetName)
        {
            return workbook.TryGetWorksheet(sheetName, out worksheet);
        }

        public int GetRowNumber(String sheetName, String colName, String value)
        {
            int totalRows = GetRowCount(sheetName);
            for (int i = 1; i <= totalRows; i++)
            {
                if (GetCellData(sheetName, colName, i).Equals(value))
                    return i;
            }
            throw new Exception(String.Format("Value being Searched '{0}' in Sheet '{1}' under Column '{2}' Couldn't be found", value, sheetName, colName));
        }
    }
}
