/*
         * Description: The method which invokes a macro within an MS Excel Spreadsheet.
         * Author: Asad Zia
         * Version 1.0
         */

        private static string RunMacroInExcel(string wbPath, string MacroName, string password)
        {
            /* Variable declarations */
            
            string msg = "";                                            /* The output message */
            Microsoft.Office.Interop.Excel.Application xl = null;       /* The excel application instance*/
            Microsoft.Office.Interop.Excel._Workbook wb = null;         /* The Excel workbook instance */
            Microsoft.Office.Interop.Excel._Worksheet sheet = null;     /* The Excel sheet instance */
            bool SaveChanges = false;                                   /* A flag for saving the changes made on the spreadsheet */
            
            try
            {
                // Create a new instance of Excel from scratch

                xl = new Microsoft.Office.Interop.Excel.Application();
                xl.Visible = true;
                
                /* Prevent popups */

                // xl.Interactive = false;
                // xl.ScreenUpdating = false;
                xl.DisplayAlerts = false;

                // Add one workbook to the instance of Excel
                // Open in read only mode (3rd argument is true) so as to avoid popup which asks for read or write access.
                
                wb = (Microsoft.Office.Interop.Excel._Workbook)(xl.Workbooks.Open(wbPath, Missing.Value, true, Missing.Value, password));
                
                // Get a reference to the one and only worksheet in our workbook

                sheet = (Microsoft.Office.Interop.Excel._Worksheet)wb.ActiveSheet;
                
                sheet.Unprotect(password);
                
                if  (wb.HasVBProject)
                {
                    try
                    {
                        // Run the named VBA Sub.

                        wb.Application.Run(MacroName, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value);
                    }
                    catch (COMException err)
                    {
                        msg = "Error: ";
                        msg = String.Concat(msg, err.Message);
                        return msg;
                        //return msg + "Macro Executed: No" + "\r\n";
                    }
                    catch (Exception e)
                    {
                        msg = "Error: ";
                        msg = String.Concat(msg, e.Message);
                        return msg;
                    }
                }

                // Let loose control of the Excel instance

                xl.Visible = false;
                xl.UserControl = false;
                 
                // Create a Customized Message to return to the object
                 
                msg = String.Concat(msg, MacroName + " executed successfully!");
                return msg;
                           
            }
            catch (Exception err)
            {
                msg = "Error: ";
                msg = String.Concat(msg, err.Message);

                return msg;
            }
              finally
              {
                  try
                  {
                      // Repeat xl.Visible and xl.UserControl releases just to be sure
                      // we didn't error out ahead of time.

                      xl.Visible = false;
                      xl.UserControl = false;

                      /*Important to save the document like this. This option can be removed if the document is not to be saved.*/
                        
                      wb.SaveAs(wbPath, wb.FileFormat, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                      wb.Close(SaveChanges, null, null);
                      xl.Workbooks.Close();
                      
                      /* Force garbage collection for ALL open Excel COM instances */
                      
                      GC.Collect();
                      GC.WaitForPendingFinalizers();
                  }
                  catch { }

                  // Gracefully exit out and destroy all COM objects to avoid hanging instances
                  // of Excel.exe whether our method failed or not.

                  xl.Quit();

                  if (sheet != null) { Marshal.ReleaseComObject(sheet); }
                  if (wb != null) { Marshal.ReleaseComObject(wb); }
                  if (xl != null) { Marshal.ReleaseComObject(xl); }

                  sheet = null;
                  wb = null;
                  xl = null;
              }

        }
