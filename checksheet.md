
# C#: Check the Excel SheetName
## Using Microsft.ACE.OlEDB


If your data pipeline reads a user-uploaded excel file into a database, then a common issue you run into is your ETL process failing spectacularly :fearful: because the 
the sheet names do not match the default set-up.

Here's a base code for figuring outsheetnames before trying to load to DB. 

You can build on it to allow your ETL to fail more gracefully by catching user errors and sending appropriate messages with SMTP tasks in SSIS.

### Pre-requisites
- Visual Studio

- Access Database Engine to facilate the transfer of data between Excel and VS. 



### Add the following namespaces to your script
```C#
using System.IO;
using System.Data.OleDb;
using System.Data;
```

### In Main function
1. Set up the filepath variable as a string. 
2. Ascertain that the filepath exists 
3. Set up a connection string variable with the Microsoft.ACE.OLEDB driver.
4. If your file is .xls then use **Extended Properties = Excel 8.0**. For .xlsx,  then use **Extended Properties = Excel 12.0**
5. Open the connection.

```C#
 string filepath = @"C:\filepath";
 
 if (File.Exists(filepath))
 {
     string  connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties=Excel 12.0";
     var conn = new OleDbConnection(connstring);
      
 ```
 
 
 6. Open the connection.
 7. Get meta-data for the 'Tables' in the excelfile using GetSchema("Tables"). *The driver interprets the sheets in a excel file as tables*
 8. A DataTable is returned to our variable table.
 9. To traverse the structure of the DataTable and get the info we want, we must access Data Table properties of Rows.
 
  
```C#
      conn.Open();
      var table = conn.GetSchema("Tables");
      var rows = table.Rows;      

```
10. A DataRowCollection is returned to our variable rows.
11. Loop through each row in the collection.
12. Access the index position which has the name of a sheet row["TABLE_NAME"].

```C#
  foreach (DataRow row in rows)
  {
    Console.WriteLine("Table Name:" + row["TABLE_NAME"].ToString());
    
  }
  conn.Close();
         
```

See fullcode below.



 
 
