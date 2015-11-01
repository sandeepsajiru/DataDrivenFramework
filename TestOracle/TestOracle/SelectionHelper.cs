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
    }
}
