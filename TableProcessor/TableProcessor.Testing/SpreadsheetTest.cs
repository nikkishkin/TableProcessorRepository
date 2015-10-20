using System.Linq;
using NUnit.Framework;

namespace TableProcessor.Testing
{
    [TestFixture]
    public class SpreadsheetTest
    {
        /// <summary>
        /// Test where there are no errors in initial cells
        /// </summary>
        [Test]
        public void CalculatePositiveTest()
        {
            string[] initialRows = new string[3];
            initialRows[0] = "12\t=C2\t3\t'Sample";
            initialRows[1] = "=A1+B1*C1/5\t=A2*B1\t=B3-C3\t'Spread";
            initialRows[2] = "'Test\t=4-3\t5\t'Sheet";

            Spreadsheet spreadsheet = new Spreadsheet{Data = GetMatrix(initialRows)};
            spreadsheet.Calculate();

            string[] resultRows = new string[3];
            resultRows[0] = "12\t-4\t3\tSample";
            resultRows[1] = "4\t-16\t-4\tSpread";
            resultRows[2] = "Test\t1\t5\tSheet";

            string[,] resultTable = GetMatrix(resultRows);

            Assert.IsTrue(MatrixesAreEqual(spreadsheet.Data, resultTable));
        }

        /// <summary>
        /// Test where some cells in initial table contain loop references
        /// </summary>
        [Test]
        public void CalculateLoopReferencesTest()
        {
            string[] initialRows = new string[3];
            initialRows[0] = "12\t=C2\t=A2*7\t'Sample"; // Third element contains loop reference
            initialRows[1] = "=A1+B1*C1/5\t=A2*B1\t=B3-C3\t'Spread"; // First and second elements contains loop references
            initialRows[2] = "'Test\t=4-3\t5\t'Sheet";

            Spreadsheet spreadsheet = new Spreadsheet { Data = GetMatrix(initialRows) };
            spreadsheet.Calculate();

            string[] resultRows = new string[3];
            resultRows[0] = "12\t-4\t#The cell contains loop reference\tSample";
            resultRows[1] = "#The cell contains loop reference\t#The cell contains loop reference\t-4\tSpread";
            resultRows[2] = "Test\t1\t5\tSheet";

            string[,] resultTable = GetMatrix(resultRows);

            Assert.IsTrue(MatrixesAreEqual(spreadsheet.Data, resultTable));
        }

        /// <summary>
        /// Test where some cells in initial table contain references to non-existent cells
        /// </summary>
        [Test]
        public void CalculateNonExistentReferencesTest()
        {
            string[] initialRows = new string[2];
            initialRows[0] = "3\t7";
            initialRows[1] = "=A1+C1\t=B1+1"; // Cell C1 doesn't exist

            Spreadsheet spreadsheet = new Spreadsheet { Data = GetMatrix(initialRows) };
            spreadsheet.Calculate();

            string[] resultRows = new string[2];
            resultRows[0] = "3\t7";
            resultRows[1] = "#Referenced cell doesn't exist\t8";

            string[,] resultTable = GetMatrix(resultRows);

            Assert.IsTrue(MatrixesAreEqual(spreadsheet.Data, resultTable));
        }

        /// <summary>
        /// Test where some cells in initial table reference to cells which cannot be part of expression
        /// </summary>
        [Test]
        public void CalculateInvalidReferencedCellTest()
        {
            string[] initialRows = new string[2];
            initialRows[0] = "3\t'Hello";
            initialRows[1] = "=A1+B1\t=A1+1"; // Cell B1 contains text line which cannot be used as a part of expression

            Spreadsheet spreadsheet = new Spreadsheet { Data = GetMatrix(initialRows) };
            spreadsheet.Calculate();

            string[] resultRows = new string[2];
            resultRows[0] = "3\tHello";
            resultRows[1] = "#Referenced cell cannot be part of expression\t4";

            string[,] resultTable = GetMatrix(resultRows);

            Assert.IsTrue(MatrixesAreEqual(spreadsheet.Data, resultTable));
        }

        /// <summary>
        /// Get matrix from array of rows, where each row's elements are separeted by tabulation char
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
        private string[,] GetMatrix(string[] rows)
        {
            int columnsNumber = rows[0].Split('\t').Length;
            string[,] table = new string[rows.Length,columnsNumber];

            for (int i = 0; i < rows.Length; i++)
            {
                for (int j = 0; j < columnsNumber; j++)
                {
                    table[i, j] = rows[i].Split('\t')[j];
                }
            }

            return table;
        }

        private bool MatrixesAreEqual(string[,] m1, string[,] m2)
        {
            for (int i = 0; i < m1.GetLength(0); i++)
            {
                for (int j = 0; j < m1.GetLength(1); j++)
                {
                    if (m1[i,j] != m2[i,j])
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}
