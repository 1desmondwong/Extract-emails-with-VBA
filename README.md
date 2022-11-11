# VBA-Extractor

This is my Visual Basic for Applications (VBA) code to extract emails into a workbook for text mining in R or Python.

To use the code, go into the workbook's VBA code using Alt-F11. I've put in notes on how to update your file paths, and how the code works.

The Extract1 macro pulls emails out of your your Outlook inbox or sub-folder into a spreadsheet, bccs included. (You can also do this with Power Query: https://www.xelplus.com/import-outlook-to-excel-with-power-query/) 

Assuming you've run Extract1 for a few people, the Extract2 macro pulls their spreadsheets into separate tabs on the workbook. You just need to place their spreadsheets into a folder on your computer beforehand.
