using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestOracle;

namespace NUnitAutomatedTests
{
    public class NUnitTestBase
    {
        public void ValidateRunModes(String suiteName, String testName, String dataRunMode)
        {
            bool isSuiteSelected = SelectionHelper.IsSuiteSelected(suiteName, "Data\\Suite.xlsx");
            bool isTestSelected = SelectionHelper.IsTestSelected(testName, "Data\\SuiteA.xlsx");
            bool isRunModeYes = dataRunMode.Equals("Y", StringComparison.OrdinalIgnoreCase);
            if (!(isSuiteSelected && isTestSelected && isRunModeYes))
            {
                throw new Exception(String.Format(
                    "Skipping Test {0}(RunMode: {1}) in Suite {2}(RunMode:{3}) with DataRunMode {4}",
                    testName, isTestSelected, suiteName, isSuiteSelected,dataRunMode));
            }
        }
    }
}
