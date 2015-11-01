using Beingzero.Framework.Datadriven.ExcelHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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


        public static void ReadTestData(String testName, String suiteFilePath)
        {
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

            int dataEndColNumber = j-1;
            String dataValue;
            for(i=dataStartRowNumber;i<=dataEndRowNumber;i++)
            {
                for(j=1;j<= dataEndColNumber; j++)
                    dataValue = eh.GetCellData("Data", j, i);
            }
        }
    }
}
