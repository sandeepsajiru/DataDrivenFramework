using System;
using System.Collections.Generic;
using NUnit.Framework;
using TestOracle;

namespace NUnitAutomatedTests
{
    [TestFixture]
    public class NUnitTest1
    {
        [Test]
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

            Console.WriteLine(selected);

            SelectionHelper.GetTestData("Test2", "Data\\SuiteA.xlsx");
        }

        private String test2Name = "Test2";
        [Test, TestCaseSource("SecondData")]
        public void DataDrivenNUnitTest2(Dictionary<String, String> table)
        {
            Console.WriteLine("");   
        }

        IEnumerable<Object[]>  SecondData()
        {
            Object[][] data = SelectionHelper.GetTestData(test2Name, "Data\\SuiteA.xlsx");
            for (int i = 0; i < data.Length; i++)
            {
                    yield return data[i];
            }
                
        }


        private String test1Name = "Test1";
        [Test, TestCaseSource("FirstData")]
        public void DataDrivenNUnitTest1(Dictionary<String, String> table)
        {
            Console.WriteLine("");
        }

        IEnumerable<Object[]> FirstData()
        {
            Object[][] data = SelectionHelper.GetTestData(test1Name, "Data\\SuiteA.xlsx");
            for (int i = 0; i < data.Length; i++)
            {
                yield return data[i];
            }

        }

    }
}