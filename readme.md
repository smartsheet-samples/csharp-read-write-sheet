## csharp-read-write-sheet

This is a minimal Smartsheet sample that demonstrates how to
* Import a sheet
* Load a sheet
* Loop through the rows
* Check for rows that meet a criteria
* Update cell values
* Write the results back to the original sheet


This sample scans a sheet for rows where the value of the "Status" column is "Complete" and sets the "Remaining" column to zero.
This is implemented in the `evaluateRowAndBuildUpdates()` method which you should modify to meet your needs.


## Setup
Update the App.config file with an API access token, obtained from the Smartsheet Account button, under Personal settings

Build and run the application.

The rows marked "Complete" will have the "Remaining" value set to 0. (Note that you will have to refresh in the desktop application to see the changes)

## See also
- https://github.com/smartsheet-platform/smartsheet-csharp-sdk
- http://smartsheet-platform.github.io/api-docs/
- https://www.smartsheet.com/
