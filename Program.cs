/****************************** Module Header ******************************\
File Name:    TXT2XLSX.cs
Purpose:      Convert a text file (CSV, TDF) to an XLSX file
By:           Jos van Nijnatten
Date:         28-10-2012
Version:      1.0
\***************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Default.Window;
using Default.Utilities;
using Microsoft.Office;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace TXT2XLSX
{
    class Convert
    {
        #region Global variables
        Excel.Application excelApplication = null;
        Excel.Workbook    excelWorkBook    = null;
        private string    iniFileName  = null;
        private bool      overwriteIni = false;
        private bool      iniCreated   = false;
        private bool      inifile      = false;
        private bool      imported     = false;
        private bool      converted    = false;
        private bool      exported     = false;
        private bool      iniDeleted   = false;
        private Exception exception    = null;
        private DataSet   dataset      = new DataSet();
        public  int       terminated   = 0; // 0 is success
        #endregion
        #region Program constants
        public readonly static string ABOUT   = Std.AssemblyTitle() + " v" + Std.GetVersion() + "\n" + Std.AssemblyCopyright() + "\n" + Std.AssemblyDescription() + "\n Icons from iconset 'Tulliana' by M.Umut Pulat (http://12m3.deviantart.com/) (LGPL v2.1).";
        public readonly static string HELP    = Std.AssemblyTitle() + " v" + Std.GetVersion() + "\n" + Std.AssemblyCopyright() + "\n" + "Use this program to convert Text files to Excel XLSX files.\n'" + Path.GetFileName(Std.GetExeName()) + " inputFileName exportFileName [options]'\n\nOptions are:\n[overwriteIni [hasColumnNames [generateColumnNames [generateIdentity]]]]\nbool overwriteIni\n    If there is an INI file (containing the ODBC driver settings for the input file), overwrite it/add to it or don't and use this INI file as it is\nbool hasColumnNames\n    Specify if the input file has column names\nbool generateColumnNames\n    If there are no column names already ('hasColumnNames') then generate them above the columns (A, B, C, ...)\nbool generateIdentity\n    Generate identities in the first column (1, 2, 3, ...)";
        public readonly static string VERSION = Std.AssemblyTitle() + " v" + Std.GetVersion(); //"\n"
        #endregion


        #region Main external entry
        /// <summary>
        /// The first function to be called when the program starts.
        /// Use as "[Program name].exe inputFileName outputFileName [options...]"
        /// </summary>
        /// <param name="args">Command line arguments</param>
        /// <returns>Success (0 is success, anything else not so much)</returns>
        static int Main(string[] args)
        {
            int argn = args.Length;
            
            if (Array.IndexOf<string>(args, "--about") >= 0)
            {
                MessageBox.Show(ABOUT,   "About...", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            else if ((Array.IndexOf<string>(args, "--help") >= 0) ||
                     (argn == 0))
            {
                MessageBox.Show(HELP,    "Help...",  MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            else if (Array.IndexOf<string>(args, "--version") >= 0)
            {
                MessageBox.Show(VERSION, "Version...",  MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            else if (argn == 1)
            {
                string outFileName = Path.GetFullPath(args[0]) + ".xlsx";
                return new Convert(args[0], outFileName).terminated;
            }
            else if (argn == 2)
            {
                return new Convert(args[0], args[1]).terminated;
            }
            else if (argn == 3)
            {
                return new Convert(args[0], args[1], ToBool(args[2], false)).terminated;
            }
            else if (argn == 4)
            {
                return new Convert(args[0], args[1], ToBool(args[2], false), ToBool(args[3], true)).terminated;
            }
            else if (argn == 5)
            {
                return new Convert(args[0], args[1], ToBool(args[2], false), ToBool(args[3], true), ToBool(args[4], false)).terminated;
            }
            else if (argn == 6)
            {
                return new Convert(args[0], args[1], ToBool(args[2], false), ToBool(args[3], true), ToBool(args[4], false), ToBool(args[5], false)).terminated;
            }
            else if (argn == 7)
            {
                return new Convert(args[0], args[1], ToBool(args[2], false), ToBool(args[3], true), ToBool(args[4], false), ToBool(args[5], false), args[6]).terminated;
            }
            else
            {
                return 1; // error
            }
        }
        #endregion


        #region Constructors
        /// <summary>
        /// The main instance of the program; Converts the input file to an XLSX file using the MS ODBC driver to read the file and imports its data and then uses Excel to make the output XLSX file.
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <param name="exportFileName">Relative or absolute filename of the file to create</param>
        public Convert(string importFileName, string exportFileName) : this(importFileName, exportFileName, false, true, false, false, null) { }

        /// <summary>
        /// The main instance of the program; Converts the input file to an XLSX file using the MS ODBC driver to read the file and imports its data and then uses Excel to make the output XLSX file.
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <param name="exportFileName">Relative or absolute filename of the file to create</param>
        /// <param name="overwriteIni">If there is an INI file (containing the ODBC driver settings for the input file), overwrite it/add to it or don't and use this INI file as it is</param>
        public Convert(string importFileName, string exportFileName, bool overwriteIni) : this(importFileName, exportFileName, overwriteIni, true, false, false, null) { }

        /// <summary>
        /// The main instance of the program; Converts the input file to an XLSX file using the MS ODBC driver to read the file and imports its data and then uses Excel to make the output XLSX file.
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <param name="exportFileName">Relative or absolute filename of the file to create</param>
        /// <param name="overwriteIni">If there is an INI file (containing the ODBC driver settings for the input file), overwrite it/add to it or don't and use this INI file as it is</param>
        /// <param name="hasColumnNames">Specify if the input file has column names</param>
        public Convert(string importFileName, string exportFileName, bool overwriteIni, bool hasColumnNames) : this(importFileName, exportFileName, overwriteIni, hasColumnNames, false, false, null) { }

        /// <summary>
        /// The main instance of the program; Converts the input file to an XLSX file using the MS ODBC driver to read the file and imports its data and then uses Excel to make the output XLSX file.
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <param name="exportFileName">Relative or absolute filename of the file to create</param>
        /// <param name="overwriteIni">If there is an INI file (containing the ODBC driver settings for the input file), overwrite it/add to it or don't and use this INI file as it is</param>
        /// <param name="hasColumnNames">Specify if the input file has column names</param>
        /// <param name="generateColumnNames">If there are no column names already ('hasColumnNames') then generate them above the columns (A, B, C, ...)</param>
        public Convert(string importFileName, string exportFileName, bool overwriteIni, bool hasColumnNames, bool generateColumnNames) : this(importFileName, exportFileName, overwriteIni, hasColumnNames, generateColumnNames, false, null) { }

        /// <summary>
        /// The main instance of the program; Converts the input file to an XLSX file using the MS ODBC driver to read the file and imports its data and then uses Excel to make the output XLSX file.
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <param name="exportFileName">Relative or absolute filename of the file to create</param>
        /// <param name="overwriteIni">If there is an INI file (containing the ODBC driver settings for the input file), overwrite it/add to it or don't and use this INI file as it is</param>
        /// <param name="hasColumnNames">Specify if the input file has column names</param>
        /// <param name="generateColumnNames">If there are no column names already ('hasColumnNames') then generate them above the columns (A, B, C, ...)</param>
        public Convert(string importFileName, string exportFileName, bool overwriteIni, bool hasColumnNames, bool generateColumnNames, bool generateIdentity) : this(importFileName, exportFileName, overwriteIni, hasColumnNames, generateColumnNames, generateIdentity, null) { }


        /// <summary>
        /// The main instance of the program; Converts the input file to an XLSX file using the MS ODBC driver to read the file and imports its data and then uses Excel to make the output XLSX file.
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <param name="exportFileName">Relative or absolute filename of the file to create</param>
        /// <param name="overwriteIni">If there is an INI file (containing the ODBC driver settings for the input file), overwrite it/add to it or don't and use this INI file as it is</param>
        /// <param name="hasColumnNames">Specify if the input file has column names</param>
        /// <param name="generateColumnNames">If there are no column names already ('hasColumnNames') then generate them above the columns (A, B, C, ...)</param>
        /// <param name="generateIdentity">Generate identities in the first column (1, 2, 3, ...)</param>
        public Convert(string importFileName, string exportFileName, bool overwriteIni, bool hasColumnNames, bool generateColumnNames, bool generateIdentity, string sep)
        {
            this.iniFileName  = Path.GetDirectoryName(importFileName) + Path.DirectorySeparatorChar + "Schema.ini";
            this.overwriteIni = overwriteIni;

            importFileName   = Path.GetFullPath(importFileName.Trim());
            exportFileName   = Path.GetFullPath(exportFileName.Trim());

            #region Add region to Ini
            while (true) {
                this.terminated = 0;
                this.inifile = AddToIniFile(importFileName, hasColumnNames, sep);

                if (this.inifile) {
                    break;
                } else {
                    string message = "Error writing to ini file";

                    if (this.exception != null)
                    {
                        message += "\n" + this.exception.Message;
                    }

                    DialogResult btn = MessageBox.Show(message, "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

                    if (btn == DialogResult.Retry) {
                        // retry
                    } else if (btn == DialogResult.Cancel) {
                        break;
                    }
                }
            }
            #endregion

            #region When there is an Ini file, import the text file
            while (this.inifile) { // Only when there is na Ini file
                this.terminated = 0;
                this.imported = this.ImportText(importFileName);

                if (this.imported) {
                    break;
                } else {
                    string message = "Error importing the text file";

                    if (this.exception != null)
                    {
                        message += "\n" + this.exception.Message;
                    }

                    DialogResult btn = MessageBox.Show(message, "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

                    if (btn == DialogResult.Retry) {
                        // retry
                    } else if (btn == DialogResult.Cancel) {
                        break;
                    }
                }
            }
            #endregion

            #region Open workbook and Excel
            if (this.imported)
            {
                this.excelApplication = new Excel.Application();
                this.excelWorkBook = this.excelApplication.Workbooks.Add();
                this.excelApplication.Visible = false;
            }
            #endregion

            #region Convert the dataset to an Excel worksheet
            if (this.imported) // Only when the data is imported
            {
                this.terminated = 0;
                this.converted = this.ConvertTextToWorksheet(hasColumnNames, generateColumnNames, generateIdentity);

                if (!this.converted)
                {
                    string message = "Error importing the text file\n" + this.exception.Message;

                    MessageBox.Show(message, "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                }
            }
            #endregion

            #region Export to XLSX
            while (this.converted) // Only when the data is converted
            {
                this.terminated = 0;
                this.exported = this.ExportXLSX(exportFileName);

                if (this.exported) {
                    break;
                } else {
                    string message = "Error exporting to XLSX";

                    if (this.exception != null)
                    {
                        message += "\n" + this.exception.Message;
                    }

                    DialogResult btn = MessageBox.Show(message, "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

                    if (btn == DialogResult.Retry) {
                        // retry
                    } else if (btn == DialogResult.Cancel) {
                        break;
                    }
                }
            }
            #endregion

            #region Close workbook and Excel
            try
            {
                if (this.excelWorkBook != null)
                {
                    this.excelWorkBook.Close(
                        Missing.Value,
                        Missing.Value,
                        Missing.Value
                        );
                }
            }
            catch (Exception ex)
            {
                // this.terminated = 1;
                exception = ex;
                /* Thats all right... */
            }
            #endregion

            #region Delete created Ini file
            if (this.iniCreated) {
                while (!this.iniDeleted)
                {
                    this.terminated = 0;
                    for (int i = 0; i < 3; i++)
                    {
                        try {
                            File.Delete(this.iniFileName);
                            this.iniDeleted = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            this.terminated = 1;
                            this.exception = ex;
                            Thread.Sleep(100);
                        }
                    }

                    if (!this.iniDeleted)
                    {
                        string message = "Error deleting Ini file";

                        if (this.exception != null)
                        {
                            message += "\n" + this.exception.Message;
                        }

                        DialogResult btn = MessageBox.Show(message, "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);

                        if (btn == DialogResult.Retry) {
                            // retry
                        } else if (btn == DialogResult.Cancel) {
                            break;
                        }
                    }
                }
            }
            #endregion
        }
        #endregion


        #region Processing the input file to the output file
        /// <summary>
        /// Create the 'Schema.ini' file for the ODBC driver, if overwrite is on in the object
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <param name="hasColumnNames">Specify if the input file has column names</param>
        /// <returns>Success</returns>
        private bool AddToIniFile(string importFileName, bool hasColumnNames, string sep)
        {
            try
            {
                if (( File.Exists(this.iniFileName) && this.overwriteIni) || // File exists and you may overwrite
                    (!File.Exists(this.iniFileName)))                        // File does not exist
                {
                    this.iniCreated    = !File.Exists(this.iniFileName);
                    string sectionName = Path.GetFileName(importFileName);
                    Ini    ini         = new Ini(this.iniFileName);
                    string extension   = Path.GetExtension(importFileName).ToLower();

                    string format = "";
                    if (string.IsNullOrEmpty(sep))
                    {
                        switch (extension)
                        {
                            case ".csv":
                                format = "CSVDelimited";
                                break;
                            case ".tab":
                                format = "TabDelimited";
                                break;
                            default:
                                format = "CSVDelimited";
                                string delimiter = "";

                                while (true)
                                {
                                    DialogResult res = Default.Window.Dialog.Prompt("Input required", "Please enter a delimiter character for file:\n" + importFileName, ref delimiter);
                                    if (res == DialogResult.OK)
                                    {
                                        if (delimiter.Length == 1)
                                        {
                                            if (delimiter == ",")
                                            {
                                                format = "CSVDelimited";
                                            }
                                            else if ((delimiter ==  "\t") &&
                                                     (delimiter == "\\t"))
                                            {
                                                format = "TabDelimited";
                                            }
                                            else
                                            {
                                                format = "Delimited(" + delimiter + ")";
                                            }
                                            break;
                                        }
                                        else
                                        {
                                            MessageBox.Show("Error", "The delimiter can't be empty or a sequence of characters", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                break;
                        }
                    }
                    else
                    {
                        switch (sep)
                        {
                            case ",":
                                format = "CSVDelimited";
                                break;
                            case "\t":
                                format = "TabDelimited";
                                break;
                            case "\\t":
                                format = "TabDelimited";
                                break;
                            default:
                                format = "Delimited(" + sep + ")";

                                if (sep.Length > 1)
                                {
                                    do
                                    {
                                        MessageBox.Show("Error", "The delimiter can't be empty or a sequence of characters", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        DialogResult res = Default.Window.Dialog.Prompt("Input required", "Please enter a delimiter character for file:\n" + importFileName, ref sep);
                                        
                                        if (res == DialogResult.Cancel)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            switch (sep)
                                            {
                                                case ",":
                                                    format = "CSVDelimited";
                                                    break;
                                                case "\t":
                                                    format = "TabDelimited";
                                                    break;
                                                case "\\t":
                                                    format = "TabDelimited";
                                                    break;
                                                default:
                                                    format = "Delimited(" + sep + ")";
                                                    break;
                                            }
                                        }
                                    } while ((sep.Length != 1) && (sep != "\\t"));
                                }

                                break;
                        }
                    }

                    ini.WriteIniValue(sectionName, "ColNameHeader", (hasColumnNames ? "True" : "False"));
                    ini.WriteIniValue(sectionName, "Format",        format);
                    //ini.WriteIniValue(sectionName, "MaxScanRows",   "1024");
                    ini.WriteIniValue(sectionName, "CharacterSet",  "OEM");
                }

                return true;
            }
            catch (Exception ex)
            {
                this.terminated = 1;
                this.exception = ex;
                return false;
            }
        }


        /// <summary>
        /// Import the file to import to the programs memory, into a DataSet, using the ODBC driver
        /// </summary>
        /// <param name="importFileName">Relative or absolute filename of the file to convert</param>
        /// <returns>Success</returns>
        private bool ImportText(string importFileName) {
            string connectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + Path.GetDirectoryName(importFileName) + ";Extensions=asc,csv,tab,txt;Persist Security Info=False";
            OdbcConnection conn = new OdbcConnection();

            string worksheet = Path.GetFileName(importFileName);

            try
            {
                conn = new OdbcConnection(connectionString);
                conn.Open();

                if ((conn.State != ConnectionState.Closed)  &&
                    (conn.State != ConnectionState.Broken)) {

                    // Fill the dataset with data
                    string sql_select = "select * from [" + Path.GetFileName(importFileName) + "]";
                
                    OdbcDataAdapter dataAdapter = new OdbcDataAdapter(sql_select, conn);

                    dataAdapter.Fill(this.dataset, worksheet);

                        /*
                        if (hasColumnNames)
                        {
                            System.Data.DataTable dt = this.dataset.Tables[worksheet];
                        
                            if ((dt.Columns.Count > 0) && (dt.Rows.Count > 0))
                            {
                                Hashtable columns = new Hashtable();

                                // Get column names and remove them from the dataset
                                DataRow row = dt.Rows[0];
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    string cell = (row[i]).ToString();
                                    if (!string.IsNullOrEmpty(cell))
                                    {
                                        columns.Add(i, cell);
                                    }
                                    else
                                    {
                                        columns.Add(i, ColumnName(i));
                                    }
                                }
                                row.Delete();

                                // Add column names (first row)
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    dt.Columns[i].ColumnName = columns[i].ToString();
                                }

                                // make changes permanent
                                dt.AcceptChanges();
                            }
                        }
                        */
                    }
                else
                {
                    throw new Exception("Connection to text file was broken or closed.");
                }
            }
            catch (Exception ex)
            {
                this.terminated = 1;
                this.exception = ex;
                return false;
            }
            finally
            {
                conn.Close();
            }

            return true;
        }


        /// <summary>
        /// Convert the imported file (in the DataSet) to an XLSX file using MS Excel
        /// </summary>
        /// <param name="hasColumnNames">Specify if the input file has column names</param>
        /// <param name="generateColumnNames">If there are no column names already ('hasColumnNames') then generate them above the columns (A, B, C, ...)</param>
        /// <param name="generateIdentity">Generate identities in the first column (1, 2, 3, ...)</param>
        /// <returns>Success</returns>
        private bool ConvertTextToWorksheet(bool hasColumnNames, bool generateColumnNames, bool generateIdentity) {
            try
            {
                for (int k = 0; k < this.dataset.Tables.Count; k++)
                {
                    System.Data.DataTable dt = this.dataset.Tables[k];
                    Worksheet excelWorkSheet = (Worksheet)this.excelWorkBook.Worksheets.Add(
                        Missing.Value,
                        Missing.Value,
                        Missing.Value,
                        Missing.Value);
                    excelWorkSheet.Name = dt.TableName;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            int row = 1 + (generateColumnNames ? 1 : 0);
                            int col = 1 + (generateIdentity ? 1 : 0);

                            // Column Header
                            if ((i == 0) && (generateColumnNames || hasColumnNames))
                            {
                                if (hasColumnNames)
                                {
                                    excelWorkSheet.Cells[1, (col + j)] = dt.Columns[j].ColumnName;
                                }
                                else if (generateColumnNames)
                                {
                                    excelWorkSheet.Cells[1, (col + j)] = ColumnName(j + 1);
                                }
                            }

                            // Row Identity
                            if ((j == 0) && generateIdentity)
                            {
                                excelWorkSheet.Cells[(row + i), 1] = (1 + i).ToString();
                            }

                            row += i;
                            col += j;

                            // Cells
                            excelWorkSheet.Cells[row, col] = dt.Rows[i][j].ToString();
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                this.terminated = 1;
                this.exception = ex;
                return false;
            }
        }


        /// <summary>
        /// Export the in-memory XLSX file to the harddisk
        /// </summary>
        /// <param name="exportFileName">Relative or absolute filename of the file to create</param>
        /// <returns>Success</returns>
        private bool ExportXLSX(string exportFileName) {
            try
            {
                this.excelWorkBook.SaveAs(
                    exportFileName,
                    Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value,
                    Missing.Value,
                    false,
                    false,
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlUserResolution,
                    true,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value);

                return true;
            }
            catch (Exception ex)
            {
                /***
                 * Exceptions and solutions found:
                 *  - Message: Exception from HRESULT: 0x800A03EC
                 *    Solution 1:
                 *      Activate your version of Excel
                 *    Solution 2:
                 *      Create directory "C:/Windows/SysWOW64/config/systemprofile/Desktop" (64 bit Windows)
                 *                       "C:/Windows/System32/config/systemprofile/Desktop" (32 bit Windows)
                 *      Set Full Control permissions for directory 'Desktop'
                 *    Solution 3:
                 *      
                 **/

                this.terminated = 1;
                this.exception = ex;
                return false;
            }
        }
        #endregion


        #region Helper functions
        /// <summary>
        /// Generate the Nth column name, e.g. 1 is A, 26 is Z, etc. 
        /// </summary>
        /// <param name="n">Column number</param>
        /// <returns>Column name</returns>
        private static string ColumnName(int n)
        {
            StringBuilder columnName = new StringBuilder();
            while (true)
            {
                if (n > 0)
                {
                    char c = ((char)(((n -1) % 26) + 65));
                    columnName.Insert(0, c);
                    n = ((n - ((n - 1) % 26)) / 26);

                    if (n == 0)
                    {
                        break;
                    }
                }
                else
                {
                    break;
                }
            }
            return columnName.ToString();
        }


        /// <summary>
        /// Convert a string to a boolean: "true" and "1" are true, "false" and "0" are false
        /// </summary>
        /// <param name="input">String to convert to a boolean</param>
        /// <param name="dflt">Default output if input is not "true", "1", "false" or "0"</param>
        /// <returns></returns>
        private static bool ToBool(string input, bool dflt)
        {
            bool rtn = dflt;

            if (input.ToLower().Equals("true") || input.Equals("1"))
            {
                rtn = true;
            }
            else if (input.ToLower().Equals("false") || input.Equals("0"))
            {
                rtn = false;
            } else {
                MessageBox.Show("Could not convert string to boolean, using default ('" + dflt.ToString() + "');\n'" + input + "'", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return rtn;
        }
        #endregion
    }
}
