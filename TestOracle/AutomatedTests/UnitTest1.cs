using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestOracle;

namespace AutomatedTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            bool selected = SelectionHelper.IsSuiteSelected("SuiteA", "Data\\Suite.xlsx");
            selected = SelectionHelper.IsSuiteSelected("SuiteB", "Data\\Suite.xlsx");
            selected = SelectionHelper.IsSuiteSelected("SuiteC", "Data\\Suite.xlsx");

            selected = SelectionHelper.IsTestSelected("Test1", "Data\\SuiteA.xlsx");
            selected = SelectionHelper.IsTestSelected("Test2", "Data\\SuiteA.xlsx");
            selected = SelectionHelper.IsTestSelected("Test3", "Data\\SuiteA.xlsx");
            selected = SelectionHelper.IsTestSelected("Test4", "Data\\SuiteA.xlsx");
            selected = SelectionHelper.IsTestSelected("Test5", "Data\\SuiteA.xlsx");


            SelectionHelper.ReadTestData("Test2", "Data\\SuiteA.xlsx");
        }
    }
}
