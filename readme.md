# ⚠️ Notice: this repository has been deprecated ⚠️

This .NET 4 sample has been replaced with the new .NET Core Smartsheet C# SDK and can be accessed [here](https://github.com/smartsheet/smartsheet-csharp-sdk/tree/mainline/sdk-csharp-sample-net80)

## See also
- https://github.com/smartsheet/smartsheet-csharp-sdk
- https://smartsheet.redoc.ly
- https://www.smartsheet.com/



---
---

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
Set the system environment variable `SMARTSHEET_ACCESS_TOKEN` to the value of your access token, obtained from the Smartsheet Account button, under Personal settings.

Build and run the application.

The rows marked "Complete" will have the "Remaining" value set to 0. (Note that you will have to refresh in the desktop application to see the changes)

