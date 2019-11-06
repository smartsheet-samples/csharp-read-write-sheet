using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

// Add nuget reference to smartsheet-csharp-sdk (https://www.nuget.org/packages/smartsheet-csharp-sdk/)
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;

namespace csharp_read_write_sheet
{
    class Program
    {
        // The API identifies columns by Id, but it's more convenient to refer to column names
        static Dictionary<string, long> columnMap = new Dictionary<string, long>(); // Map from friendly column name to column Id 

        static void Main(string[] args)
        {
            // Initialize client. Uses API access token from environment variable SMARTSHEET_ACCESS_TOKEN
            SmartsheetClient smartsheet = new SmartsheetBuilder().Build();

            Sheet sheet = smartsheet.SheetResources.ImportXlsSheet("../../../Sample Sheet.xlsx", null, 0, null);
            
            // Load the entire sheet
            sheet = smartsheet.SheetResources.GetSheet(sheet.Id.Value, null, null, null, null, null, null, null);
            Console.WriteLine("Loaded " + sheet.Rows.Count + " rows from sheet: " + sheet.Name);

            // Build column map for later reference
            foreach (Column column in sheet.Columns)
                columnMap.Add(column.Title, (long)column.Id);

            // Accumulate rows needing update here
            List<Row> rowsToUpdate = new List<Row>();

            foreach (Row row in sheet.Rows)
            {
                Row rowToUpdate = evaluateRowAndBuildUpdates(row);
                if (rowToUpdate != null)
                    rowsToUpdate.Add(rowToUpdate);
            }

            // Finally, write all updated cells back to Smartsheet 
            Console.WriteLine("Writing " + rowsToUpdate.Count + " rows back to sheet id " + sheet.Id);
            smartsheet.SheetResources.RowResources.UpdateRows(sheet.Id.Value, rowsToUpdate);
            Console.WriteLine("Done (Hit enter)");
            Console.ReadLine();
        }


        /*
         * TODO: Replace the body of this loop with your code
         * This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
         * 
         * Return a new Row with updated cell values, else null to leave unchanged
         */
        static Row evaluateRowAndBuildUpdates(Row sourceRow)
        {
            Row rowToUpdate = null;

            // Find cell we want to examine
            Cell statusCell = getCellByColumnName(sourceRow, "Status");
            if (statusCell.DisplayValue == "Complete")
            {
                Cell remainingCell = getCellByColumnName(sourceRow, "Remaining");
                if (remainingCell.DisplayValue != "0")                  // Skip if "Remaining" is already zero
                {
                    Console.WriteLine("Need to update row # " + sourceRow.RowNumber);

                    var cellToUpdate = new Cell
                    {
                        ColumnId = columnMap["Remaining"],
                        Value = 0
                    };

                    var cellsToUpdate = new List<Cell>();
                    cellsToUpdate.Add(cellToUpdate);

                    rowToUpdate = new Row
                    {
                        Id = sourceRow.Id,
                        Cells = cellsToUpdate
                    };
                }
            }
            return rowToUpdate;
        }


        // Helper function to find cell in a row
        static Cell getCellByColumnName(Row row, string columnName)
        {
            return row.Cells.First(cell => cell.ColumnId == columnMap[columnName]);
        }
    }
}
