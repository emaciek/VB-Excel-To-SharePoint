# VB-Excel-To-SharePoint
Transfer Excel rows into SharePoint list.
# Idea
Excel is very common format for many forms which are sent to some business unit for processing.
Excel is very flexible in terms of automation, controlling data, getting data from different sources and so on.

If business unit get this kind of form for processing, there is a need for "central database" for incoming data and tracking progress, filtering, making reports. Typical scenario is one big Excel file to paste all data into it. There are meny problems with this kind of files.
So let's put **Excel rows into SharePoint list automatically!**
To acomplish this you have an excel file with Visual Basic macro, which is using SharePoint web service to create new list items based on selected rows in Excel.

# The file
>Order form.xlsm
This is the example file, which you can use to acomplish our goal of automatic creation of list items.
Some information about it:
* There are two macros: UpdateSharepoint and Footer. Footer is here only to put form version in the footer when you print it.
* All worksheets should be [protected] (https://support.office.com/en-ca/article/Password-protect-worksheet-or-workbook-elements-dbf706e0-ba22-4a08-84d8-552db16eef11) in order to prevent "outside" users from changing its structure.
* Columns A,B,C and D shoud be hidden. You need to unhide them to use export macro.
* In the file there is implemented automatic cells coloring and data validations.

#The SharePoint list compatible with Excel file

To **create SharePoint list** you need to import ForListCreation.xlsx as stated in [article](https://support.office.com/en-ca/article/Create-a-list-based-on-a-spreadsheet-380cfeb5-6e14-438e-988a-c2b9bea574fa). The outcome will be sharepoint list with columns "compatible" with our excel columns from A to K.

#Order form.xlsm code changes for your environment

##How to find out SharePoint list ID?
SharePoint lists have ID's like {9C3BA4B4-8960-467A-8500-D9C911C94C74}

* By going to list settings throug web interface. In address bar there will be address:

http://intranet/internal/_layouts/15/listedit.aspx?List={9C3BA4B4-8960-467A-8500-D9C911C94C74}
which consists of list ID at the end.

* By opening list in SharePoint Designer. Lists and libraries -> List name -> List information -> List ID

## How to find out SharePoint list column names?
You need to visit address like below. In this example "internal" is site name. Of course after "List=" you need to put your list ID.

```
(http://intranet/internal/_vti_bin/owssvr.dll?Cmd=Display&List={9C3BA4B4-8960-467A-8500-D9C911C94C74}&XMLDATA=TRUE)
```

You will get:
```xml
<xml xmlns:s="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:rs="urn:schemas-microsoft-com:rowset" xmlns:z="#RowsetSchema">
<s:Schema id="RowsetSchema">
   <s:ElementType name="row" content="eltOnly" rs:CommandTimeout="30">
      <s:AttributeType name="ows_LinkTitle" rs:name="Order link" rs:number="1">
         <s:datatype dt:type="string" dt:maxLength="512"/>
      </s:AttributeType>
      <s:AttributeType name="ows_Order_x0020_number" rs:name="Order number" rs:number="2">
         <s:datatype dt:type="float" dt:maxLength="8"/>
      </s:AttributeType>
      <s:AttributeType name="ows_Order_x0020_date" rs:name="Order date" rs:number="3">
         <s:datatype dt:type="float" dt:maxLength="8"/>
      </s:AttributeType>
      <s:AttributeType name="ows_Org_x0020_unit" rs:name="Org unit" rs:number="4">
         <s:datatype dt:type="string" dt:maxLength="512"/>
      </s:AttributeType>
	  ...
```
**Important!** Column names are "frozen" on time of creation. Later if you change column names through web interface it will not make any effect on "internal" column names above.

## Changing variables
* listid (your list ID)
* listname (name of your list, in our example "Orders")
* listURL (URL to the web services) for example: http://intranet/internal/_vti_bin/Lists.asmx
## Fields
You have XML with column names like *ows_Order_x0020_number* you need to place names in "Fields" variable but without "ows_" in front of it. It should be: "<Field Name='Order_x0020_number'>". 
You have to do it for every field you need.
arglist.Item(2) means that you are putting conents of the second column from your excel.
##How it works
- You need to mark cells that you want to transfer to SharePoint list. So in our example from A9 to K12. Then you have to push "SharePoint button. (You can change the script to get the same range every time.)
- Script is taking selected range, and for every row, it is creating key-value store for the data.
- Then we are calling updateSharePointList function, giving it key-value as argument. We are doing it for every row.
- The updateSharePointList function is putting arglist.Item into XML.
- In XML Cmd='New' means we are creating new records. If you change it to Cmd='Update' ,and you will have ID of list element in your excel, it will update it instead of creating new one. But it is out of scope of this manual.
- Next we are taking that XML and calling SharePoint web service.
- Web service is giving us response and we are showing popup with ID of the list element created.
- If there is any error we are showing popup with xml response.
- You can see more debug information in VBA editor in "Immediate" window.
- After items creation in SharePoint, we are opening this list in browser and filtering it through order number. In that way user will be able to see, what was added, and maybe change something in SharePoint editing.

#License
Creative commons attribution license
https://creativecommons.org/licenses/by/3.0/
