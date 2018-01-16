using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using SpreadsheetLight;

namespace ImportExcelLight
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = @"E:\Planilhas\Planilha de teste de contrato de cliente.xlsx";
            var sheetNames = GetSheetNames(fileName);

            foreach (var sheetName in sheetNames)
            {
                var table = CreateDataTable(fileName, sheetName);
                Console.WriteLine($"Planilha {sheetName} - Total de registros: {table.Rows.Count}");
            }
            
            Console.ReadLine();
        }

        public static List<string> GetSheetNames(string fileName)
        {
            using (var sl = new SLDocument(fileName))
                return sl.GetSheetNames();
        }

        public static DataTable CreateDataTable(string fileName, string sheetName)
        {
            using (var sl = new SLDocument(fileName, sheetName))
            {
                var stats = sl.GetWorksheetStatistics();
                var tableName = sl.GetCurrentWorksheetName();
                var table = new DataTable(tableName);

                var firstLine = true;
                for (var row = stats.StartRowIndex + 1; row <= stats.EndRowIndex; row++)
                {
                    var newRow = table.NewRow();

                    for (var column = stats.StartColumnIndex; column <= stats.EndColumnIndex; column++)
                    {
                        var columnName = sl.GetCellValueAsString(1, column).Trim();
                        var value = sl.GetCellValueAsString(row, column);

                        if (firstLine && !table.Columns.Contains(columnName))
                            table.Columns.Add(columnName, typeof(string));

                        newRow[columnName] = value;
                    }

                    table.Rows.Add(newRow);
                    firstLine = false;
                }

                return table;
            }
        }
    }
}