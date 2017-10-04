/*
* Description: A generic way of injecting a macro and executing it on a MS Excel Spreadsheet and then extracting a dataTable.
* Author: Asad Zia
* Version 1.0
*/

private static System.Data.DataTable InjectMacro(string wbPath, string Macro, string MacroName, string SheetName, string index)
{   
  string msg = "";                                            /* The output message. Can be used for debugging. Change return type to String*/
  Microsoft.Office.Interop.Excel.Application xl = null;       /* The excel application instance*/
  Microsoft.Office.Interop.Excel._Workbook wb = null;         /* The Excel workbook instance */
  Microsoft.Office.Interop.Excel._Worksheet sheet = null;     /* The Excel sheet instance */
  bool SaveChanges = false;                                   /* A flag for saving the changes made on the spreadsheet */
  Microsoft.Vbe.Interop.VBComponent module = null;
  System.Data.DataTable result = null;
  object rows = 0, columns = 0;

  try
  {
    GC.Collect(); // Garbage collection of any COM instances

    // Create a new instance of Excel from scratch
    xl = new Microsoft.Office.Interop.Excel.Application();
    xl.Visible = true;

    // Add one workbook to the instance of Excel
    wb = (Microsoft.Office.Interop.Excel._Workbook)(xl.Workbooks.Open(wbPath, Missing.Value, Missing.Value, Missing.Value, Missing.Value)); 

    // Check if the spreadsheet has a password
    if (!wb.HasPassword) {msg = msg + "No Password!";}

    // Get a reference to the one and only worksheet in our workbook
    sheet = (Microsoft.Office.Interop.Excel._Worksheet)wb.Worksheets[SheetName];
    sheet.Activate();

    try
    {
      // Dynamically create a code module and load it with the string we formatted
      // in the .GetMacro() method above.
      module = wb.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
      module.CodeModule.AddFromString(Macro);

      // Run the named VBA Sub that we just added.  In our sample, we named the Sub FormatSheet
      wb.Application.Run(MacroName, SheetName, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value);

    }
    catch (Exception err)
    {
      msg = "Error: ";
      msg = String.Concat(msg, err.Message);
      //return msg;
    }

    try
    {
      // Dynamically create a code module and load it with the Macro String
      module = wb.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
      module.CodeModule.AddFromString(GetRowMacro());

      // Run the named VBA Sub.
      rows = wb.Application.Run("LastRow", Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value);

      //msg =  msg + " " + Convert.ToString(rows) + "\r\n";
    }
    catch (Exception err)
    {
      msg = "Error: ";
      msg = String.Concat(msg, err.Message);
      //return msg;
    }

    try
    {
      // Dynamically create a code module and load it with the Macro String
      module = wb.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
      module.CodeModule.AddFromString(GetColumnMacro());

      // Run the named VBA Sub.
      columns = wb.Application.Run("LastColumn", Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
      Missing.Value, Missing.Value, Missing.Value);

      //msg =  msg + " " + Convert.ToString(columns) + "\r\n";
    }
    catch (Exception err)
    {
      msg = "Error: ";
      msg = String.Concat(msg, err.Message);
      //return msg;
    }

    // Get datatable
    result = (ReadExcelFile(wbPath,(Convert.ToInt32(rows)).ToString(), ExcelColumnFromNumber(Convert.ToInt32(columns)), sheet.Name, index)).Tables[0];

    // Let loose control of the Excel instance
    xl.Visible = false;
    xl.UserControl = false;

    // Set a flag saying that all is well and it is ok to save our changes to a file.
    SaveChanges = false;

    //  Save the file to disk
    /* wb.SaveAs(FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
    null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
    false, false, null, null, null);*/
    //return msg + "Success - Macro Nailed!";

  }
  catch (Exception err)
  {
    msg = "Error: ";
    msg = String.Concat(msg, err.Message);
    msg = String.Concat(msg, " Line: ");
    msg = String.Concat(msg, err.Source);
    //return msg;
  }
  finally
  {

    try
    {
      // Repeat xl.Visible and xl.UserControl releases just to be sure
      // we didn't error out ahead of time.
      xl.Visible = false;
      xl.UserControl = false;

      // Close the document and avoid user prompts to save if our
      // method failed.
      wb.Close(SaveChanges, null, null);
      xl.Workbooks.Close();
    }
    catch { }

    // Gracefully exit out and destroy all COM objects to avoid hanging instances
    // of Excel.exe whether our method failed or not.
    xl.Quit();

    if (module != null) { Marshal.ReleaseComObject(module); }
    if (sheet != null) { Marshal.ReleaseComObject(sheet); }
    if (wb != null) { Marshal.ReleaseComObject(wb); }
    if (xl != null) { Marshal.ReleaseComObject(xl); }

    module = null;
    sheet = null;
    wb = null;
    xl = null;
    GC.Collect();
  }
  return result;
}


/*
* Description: A way to convert a numerical column into its alphabetical equivalent.
* Author: Asad Zia
* Version 1.0
*/
public static string ExcelColumnFromNumber(int column)
{
  string columnString = "";
  decimal columnNumber = column;
 
  while (columnNumber > 0)
  {
    decimal currentLetterNumber = (columnNumber - 1) % 26;
    char currentLetter = (char)(currentLetterNumber + 65);
    columnString = currentLetter + columnString;
    columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
  }

  return columnString;
}
   
/*
* Description: Connection string generation function for OleDB.
* Author: Asad Zia
* Version 1.0
*/     
private static string GetConnectionString(string filePath)
{
    Dictionary<string, string> props = new Dictionary<string, string>();

    // XLSX - Excel 2007, 2010, 2012, 2013
    //props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
    //props["Extended Properties"] = "Excel 12.0 XML";
    //props["Data Source"] = filePath;

    // XLS - Excel 2003 and Older
    props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
    props["Extended Properties"] = "'Excel 8.0; HDR=Yes; IMEX=1;'";     //HDR is for considering first row as header and IMEX is for setting all value as text items
    props["Data Source"] = filePath;

    StringBuilder sb = new StringBuilder();

    foreach (KeyValuePair<string, string> prop in props)
    {
        sb.Append(prop.Key);
        sb.Append('=');
        sb.Append(prop.Value);
        sb.Append(';');
    }

    return sb.ToString();
}


/*
* Description: Reading excel datatable from a specific sheet.
* Author: Asad Zia
* Version 1.0
*/
private static DataSet ReadExcelFile(string filePath, string row, string column, string sheet, string index)
{
    // the string index parameter should be the cell reference for example, 'A1'.

    DataSet ds = new DataSet();

    string connectionString = GetConnectionString(filePath);

    using (OleDbConnection conn = new OleDbConnection(connectionString))
    {
        conn.Open();
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = conn;

        // Get all Sheets in Excel File
        System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

        // Loop through all Sheets to get data
        foreach (DataRow dr in dtSheet.Rows)
        {
            string sheetName = dr["TABLE_NAME"].ToString();

            if (!sheetName.EndsWith("$"))       // For sheet names with whitespace in their names, we get something like -->  'White Space$' ... An apostrophe at the END!
            {
              if (!sheetName.EndsWith("'"))
              {
                continue;
              }
              else
              {
                sheetName = sheetName.Replace("'", "");
              }
            }

            if (sheetName.Contains(sheet))
            {    
                // Get all rows from the Sheet
                cmd.CommandText = "SELECT * FROM [" + sheetName + index + ":" + column + row + "]" ;

                System.Data.DataTable dt = new System.Data.DataTable();
                dt.TableName = sheetName;

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);

                ds.Tables.Add(dt);
            }
        }

        cmd = null;
        conn.Close();
    }

    return ds;
}

/*
* Description: Writing excel datatable from a specific sheet.
* Author: Asad Zia
* Version 1.0
*/
private static void WriteExcelFile(string filePath)
{
    string connectionString = GetConnectionString(filePath);

    using (OleDbConnection conn = new OleDbConnection(connectionString))
    {
        conn.Open();
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = conn;

        cmd.CommandText = "CREATE TABLE [table1] (id INT, name VARCHAR, datecol DATE );";
        cmd.ExecuteNonQuery();

        cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(1,'AAAA','2014-01-01');";
        cmd.ExecuteNonQuery();

        cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(2, 'BBBB','2014-01-03');";
        cmd.ExecuteNonQuery();

        cmd.CommandText = "INSERT INTO [table1](id,name,datecol) VALUES(3, 'CCCC','2014-01-03');";
        cmd.ExecuteNonQuery();

        cmd.CommandText = "UPDATE [table1] SET name = 'DDDD' WHERE id = 3;";
        cmd.ExecuteNonQuery();

        conn.Close();
    }
}
