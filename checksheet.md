
# C#: Check the Excel SheetName
## Using Microsoft.ACE.OLEDB


When data pipeline requires a user-uploaded excel file to be read to a database,  a common breakpoint is the ETL process failing spectacularly :fearful: because the 
the excel sheetnames do not match the default set-up.

Here's a base code for programmatically checking that sheetnames  are correct before trying to load to DB. 


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
4. For the connection string, if your file is .xls then use **Extended Properties = Excel 8.0**. For .xlsx,  then use **Extended Properties = Excel 12.0**


```C#
 string filepath = @"C:\filepath";
 string sheetname =  "somename$"; // Variable that has the sheetname we going to compare against; usually you can grab thisfrom Script task Dts variables
 bool sheet_name_matches; //writable Dts variable stores the outcome of our check
 if (File.Exists(filepath))
 {
     string  connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties=Excel 12.0";
     var conn = new OleDbConnection(connstring);
     .....
 ```
 
 
 5. Open the connection.
 6. Get meta-data for the 'Tables' in the excelfile using GetSchema("Tables"). *The driver interprets the sheets in a excel file as tables*
 7. A DataTable is returned to our variable table.
 8. To traverse the structure of the DataTable and get the info we want, we must access Data Table properties of Rows.
 
  
```C#
      conn.Open();
      var table = conn.GetSchema("Tables");
      var rows = table.Rows;  
      ...

```
10. A DataRowCollection is returned to our variable rows.
11. Loop through each row in the collection.
12. Access the index position which has the name of a sheet with row["TABLE_NAME"].
13. Compare the table/sheet name at that index to the expected name we have stored in out sheetname variable.
14. If we find a match, break out.

```C#
    foreach (DataRow row in rows)
    {
        sheet_name_matches = row["TABLE_NAME"].Equals(sheetname);
        if (sheet_name_matches)
        {
           //Console.WriteLine("Success!"); 
           break;
         }

     }
     ...
         
```

See fullcode below.


```C#
 static void Main(string[] args)
 {
           string filepath = @"C:\filepath";
           string sheetname =  "somesheetname$"; // Variable that has the sheetname we want to find; IRL, we can put this in a read only Dts variables
           bool sheet_name_matches; // A boolean that stores the outcome of our script. IRL, we can put this in a read/write Dts variable
           
            if (File.Exists(filepath))
            {               
                connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties=Excel 12.0";
                var conn = new OleDbConnection(connstring);
                conn.Open();             

                var table = conn.GetSchema("Tables");
                var rows = table.Rows;

                foreach (DataRow row in rows)
                {
                    sheet_name_matches = row["TABLE_NAME"].Equals(sheetname);
                    if (sheet_name_matches)
                    {
                        //Console.WriteLine("Success!"); 
                        break;
                    }

                }

                conn.Close();

                //Dts.TaskResult = sheet_name_matches ? ScriptResults.Success   : ScriptResults.Failure
            }
            else
            {
                Console.WriteLine("Error. File does not exist at specified path");
            }
}
 

```



 
 
