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
