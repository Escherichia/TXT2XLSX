/****************************** Module Header ******************************\
File Name:    Synchron\Log.cs
Purpose:      General use in any C# project to write to a .Log-file.
By:           Jos van Nijnatten
Date:         29-09-2012
Version:      1.0
\***************************************************************************/

using System;
using System.IO;
using Default.Utilities;

namespace Default.Utilities
{
    public class Log
    {
        private string sFileName;
        private long   maxSize = 1024L * 1024L * 10L; // 10 Mb

        #region Constructor
        /// <summary>Creates a new <see cref="Ini"/> instance. Default logfile is the programname + '.log'</summary>
        public Log() : this(@"WorkData\" + Std.GetFileNameOnly(Std.GetExeName()) + ".log") { }

        /// <summary>Creates a new <see cref="Ini"/> instance.</summary>
        /// <param name="FileName">Path to the LOG file.</param>
        public Log(string FileName)
        {
            this.sFileName = Path.GetFullPath(FileName);

            #region Create file
            if (!File.Exists(this.sFileName))
            {
                CreateFile(this.sFileName);
            }
            #endregion
        }
        #endregion

        /// <summary>Adds a text to the log-file. Create file if not exist.</summary>
        /// <param name="sMessage">The text to add to the log file.</param>
        public void LogMessage(string sMessage)
        {
            if (File.Exists(this.sFileName))
            {
                sMessage = string.Format("{0:G}: {1}{2}", DateTime.Now, sMessage, Environment.NewLine);

                try
                {
                    File.AppendAllText(this.sFileName, sMessage);
                }
                catch
                {
                    Console.WriteLine("Error writing line(s) to log-file ('" + this.sFileName + "'). Message:\n" + sMessage);
                }

                fileBackup(this.sFileName, this.maxSize, false);
            }
        }

        #region Helper functions
        /// <summary>
        /// Removes the first few lines so that the last entries are present.
        /// </summary>
        /// <param name="fileName">Name of the file which you want to truncate</param>
        /// <param name="maxSize">Max size of the origional file.</param>
        /// <param name="stack">Stack the old backup files or throw them away.</param>
        private static void fileBackup(string fileName, long maxSize, bool stack)
        {
            if (File.Exists(fileName))
            {
                long fileSize = new FileInfo(fileName).Length;

                if (fileSize > maxSize)
                {
                    if (!stack)
                    {
                        try
                        {
                            // Delete old backup file and create a new one
                            if (File.Exists(fileName + ".bak"))
                            {
                                File.Delete(fileName + ".bak");
                                Debug.write("Removed old backup file '" + fileName + ".bak';");
                            }
                            File.Move(fileName, fileName + ".bak");
                            Debug.write("Created file '" + fileName + ".bak';");
                        }
                        catch (Exception e)
                        {
                            Debug.write("Can't backup file '" + fileName + "';");
                            Debug.write(e);
                        }
                    }
                    else
                    {
                        string   dir   = Path.GetDirectoryName(fileName);
                        string[] files = Directory.GetFiles(dir);
                        int i;

                        // Find amount of files that are backup files of 'fileName'
                        for (i = 0; i < files.Length; i++ )
                        {
                            string newFileName = fileName + "." + i + ".bak";
                            int    keyIndex    = Array.FindIndex<string>(files, item => item == newFileName);

                            if (keyIndex == -1)
                            {
                                break;
                            }
                        }

                        // Create new backup file
                        try
                        {
                            File.Move(fileName, fileName + "." + i + ".bak");
                            Debug.write("Created backup file '" + fileName + "." + i + ".bak';");
                        }
                        catch (Exception e)
                        {
                            Debug.write("Can't backup file '" + fileName + "';");
                            Debug.write(e);
                        }
                    }

                    // Create new log file.
                    CreateFile(fileName);
                }
            }
            else
            {
                Debug.write("File not found; " + fileName, Debug.trace);
            }
        }

        /// <summary>
        /// Create a file and the directories to the file.
        /// </summary>
        /// <param name="sFileName">The file to create</param>
        private static void CreateFile(string sFileName)
        {
            try
            {
                // Create Directory
                string dir = Path.GetDirectoryName(sFileName);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                // Create File
                if (!File.Exists(sFileName))
                {
                    FileStream fs = File.Create(sFileName);
                    fs.Close();
                }
                Debug.write("Created log-file '" + sFileName + "'...");
            }
            catch (Exception)
            {
                Debug.write("Error creating log-file '" + sFileName + "'...");
            }
        }
        #endregion
    }
}
