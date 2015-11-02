using Beingzero.Framework.Datadriven.ExcelHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = DocumentFormat.OpenXml.Drawing.Charts.DataTable;

namespace TestOracle
{
    public class SelectionHelper
    {
        public static bool IsSuiteSelected(String suiteName, String filePath)
        {
            ExcelHelper eh = new ExcelHelper(filePath);
            int rowId = eh.GetRowNumber("Suite", "SuiteName", suiteName);

            if (eh.GetCellData("Suite", "Runmode", rowId).Equals("Y", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        public static bool IsTestSelected(String testName, String suiteFilePath)
        {
            ExcelHelper eh = new ExcelHelper(suiteFilePath);
            int rowId = eh.GetRowNumber("TestCases", "TestCases", testName);
            if (eh.GetCellData("TestCases", "Runmode", rowId).Equals("Y", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }


        public static Object[][] GetTestData(String testName, String suiteFilePath)
        {
            string[][] resultData;
            DataTable dt = new DataTable();
            ExcelHelper eh = new ExcelHelper(suiteFilePath);
            int rowId = eh.GetRowNumber("Data", 1, testName);
            int colHeaderRowNumber = rowId + 1;
            int dataStartRowNumber = rowId + 2;

            int i = 0;
            while(!eh.GetCellData("Data", 1, i+dataStartRowNumber).Equals(""))
            {
                i++;
            }

            int dataEndRowNumber = dataStartRowNumber + i-1;

            int j = 1;
            while(!eh.GetCellData("Data", j, colHeaderRowNumber).Equals(""))
            {
                j++;
            }

            resultData = new string[i][];
            Object[][] newResultData = new Object[i][];

            int r=0, c=0;
            int dataEndColNumber = j-1;
            for(i=dataStartRowNumber;i<=dataEndRowNumber;i++, r++)
            {
                c = 0;
                resultData[r] = new string[j - 1];
                newResultData[r] = new Object[1];
                Dictionary<String, String> dict = new Dictionary<string, string>();
                for (j = 1; j <= dataEndColNumber; j++, c++)
                {
                    resultData[r][c] = eh.GetCellData("Data", j, i);
                    dict.Add(eh.GetCellData("Data", j, colHeaderRowNumber), eh.GetCellData("Data", j, i));
                    newResultData[r][0] = dict;
                }
            }

            //return resultData;
            return newResultData;
        }
    }
}
