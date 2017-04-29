# trelloToXL
Facility to import a Trello JSON file into an excel spreadsheet

Written by Kevin Harper, 16-Mar-17.

The spreadsheet trelloJsonToXl.xlsm uses the JSON export from Trello to import a rich content into Excel as a list of cards or a list of actions. 
Either download the spreadsheet and use directly, or import the TrelloImport.bas VBA module into your own sheet.  If you use this approach, you will also need to import the VBA code for Tim Hall's JSON parser, referenced at the link below.

No Chrome or other browser extensions are required for the export/import. The import scripts use the JSON parser capability developed by Tim Hall, available at:  
https://github.com/VBA-tools/VBA-JSON
  The VBA-JSON parser is already included in the spreadsheet file.

From within the spreadsheet, there are two import options
1. The "ImportedCards" sheet runs vbscript ImportMyTrello - this creates a list of all the cards on your Trello board,
excluding those that have been archived.
2. The "ImportedActions" sheet runs vbscript ImportActionsFromTrello - this creates a list of
all the actions that have been carried out on your Trello board -
by applying Excel filters, you can create a list of specific actions;  eg, such as a record of who and when moved cards into a particular column.

To export a board from Trello into Excel using the spreadsheet, carry out the following steps:

(1) In Trello, for your chosen board and using the right-hand side menu options, select:
       More / Print and Export / Export to JSON

o  In Chrome, this will display the JSON code in the open tab.  Save this as a local file on your computer,
by right clicking and selecting "Save as..."
o  In Internet Explorer, this will download the JSON export as a local file on your computer.

(2) Using the spreadsheet, click the "Import" button on either worksheet and select your downloaded JSON file
....and the import should proceed.

