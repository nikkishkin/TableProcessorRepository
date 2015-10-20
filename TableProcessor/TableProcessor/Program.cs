using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TableProcessor
{
    class Program
    {
        private static int FIRST_CELL_OFFSET = 2;

        static void Main(string[] args)
        {
            // The last element of current line and the first argument of next line always belong to the same arg
            // We should split such arg
            args = args.SelectMany(arg => Regex.Split(arg, Environment.NewLine)).ToArray();

            int rowsNumber = Convert.ToInt32(args[0]);
            int colsNumber = Convert.ToInt32(args[1]);

            // Fill the table with data received from command line
            string[,] table = new string[rowsNumber,colsNumber];
            for (int i = 0; i < rowsNumber; i++)
            {
                for (int j = 0; j < colsNumber; j++)
                {
                    table[i, j] = args[FIRST_CELL_OFFSET + i * colsNumber + j];
                }
            }

            Spreadsheet spreadsheet = new Spreadsheet {Data = table};
            spreadsheet.Calculate();

            for (int i = 0; i < rowsNumber; i++)
            {
                for (int j = 0; j < colsNumber; j++)
                {
                    Console.Write(spreadsheet.Data[i, j] + "\t");
                }
                Console.WriteLine();
            }

            Console.ReadLine();
        }

    }
}
