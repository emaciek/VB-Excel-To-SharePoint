# VB-Excel-To-SharePoint
Transfer Excel rows into SharePoint list.
## Idea
Excel is very common format for many forms which are sent to some business unit for processing.
Excel is very flexible in terms of automation, controlling data, getting data from different sources and so on.

If business unit get this kind of form for processing, there is a need for "central database" for incoming data and tracking progress, filtering, making reports. Typical scenario is one big Excel file to paste all data into it. There are meny problems with this kind of files.
So let's put **Excel rows into SharePoint list automatically!**
To acomplish this you have an excel file with Visual Basic macro, which is using SharePoint web service to create new list items based on selected rows in Excel.

## The file
>Order form.xlsm
This is the example file, which you can use to acomplish our goal of automatic creation of list items.
Some information about it:
* All worksheets should be [protected] (https://support.office.com/en-ca/article/Password-protect-worksheet-or-workbook-elements-dbf706e0-ba22-4a08-84d8-552db16eef11) in order to prevent "outside" users from changing its structure.
* Columns A,B,C and D shoud be hidden. You need to unhide them to use export macro.
* In the file there is implemented automatic cells coloring and data validations.
## The SharePoint list compatible with Excel file
To **create SharePoint list** you need to import ForListCreation.xlsx as stated in [article](https://support.office.com/en-ca/article/Create-a-list-based-on-a-spreadsheet-380cfeb5-6e14-438e-988a-c2b9bea574fa). The outcome will be sharepoint list with columns "compatible" with our excel columns from A to K.