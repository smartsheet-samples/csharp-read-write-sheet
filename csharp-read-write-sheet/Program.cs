using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

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
            // Get API access token from App.config file or environment
            string accessToken = ConfigurationManager.AppSettings["AccessToken"];
            if (string.IsNullOrEmpty(accessToken))
                accessToken = Environment.GetEnvironmentVariable("SMARTSHEET_ACCESS_TOKEN");
            if (string.IsNullOrEmpty(accessToken))
                throw new Exception("Must set API access token in App.conf file");

            // Get sheet Id from App.config file 
            string sheetIdString = ConfigurationManager.AppSettings["SheetId"];
            long sheetId = long.Parse(sheetIdString);

            // Initialize client
            SmartsheetClient ss = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            // Load the entire sheet
            Sheet sheet = ss.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);
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
            ss.SheetResources.RowResources.UpdateRows(sheetId, rowsToUpdate);
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
                    // We need to update this row, so create a rowBuilder and list of cells to update
                    var rowBuilder = new Row.UpdateRowBuilder(sourceRow.Id);
                    var cellsToUpdate = new List<Cell>();

                    // Build each new cell value
                    var cellToUpdate = new Cell.UpdateCellBuilder(columnMap["Remaining"], 0).Build();   // Set value to 0

                    // Add to update list
                    cellsToUpdate.Add(cellToUpdate);
                    rowToUpdate = rowBuilder.SetCells(cellsToUpdate).Build();
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
