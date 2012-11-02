/****************************** Module Header ******************************\
File Name:    Default\Ini.cs
Purpose:      General use in any C# project to save variables in an ini file.
By:           Koen Surtel
Date:         
Version:      1.0
\***************************************************************************/

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using Default.Utilities;

namespace Default.Utilities
{
    public class Ini
    {
        #region External
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string value, string filePath);

        [DllImport("kernel32", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Auto)]
        private static extern int GetPrivateProfileString(string section, string key, string defaultValue, string sResultBuf, int nSize, string filePath);

        [DllImport("kernel32.dll", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Auto)]
        private static extern int GetPrivateProfileSection(string section, string sResultBuf, int nSize, string filePath);

        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileSectionNamesA")]
        private static extern int GetPrivateProfileSectionNames(byte[] lpszResultBuffer, int nSize, string filePath);
        #endregion

        private string sINIpath;
        private bool bNeedUpdate;

        /// <summary>Creates a new <see cref="Ini"/> instance.</summary>
        /// <param name="INIPath">Path to the INI file.</param>
        public Ini() : this(@"WorkData\" + Std.AssemblyProduct() + ".ini") { }
        public Ini(string INIPath)
        {
            sINIpath = Path.GetFullPath(INIPath);

            #region Create file
            if (!File.Exists(this.sINIpath))
            {
                try
                {
                    // Create Directory
                    string dir = Path.GetDirectoryName(this.sINIpath);
                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }
                    // Create File
                    if (!File.Exists(this.sINIpath))
                    {
                        FileStream fs = File.Create(this.sINIpath);
                        fs.Close();
                    }
                    Debug.write("Created ini-file '" + this.sINIpath + "'...");
                }
                catch (Exception)
                {
                    Debug.write("Error creating ini-file '" + this.sINIpath + "'...");
                }
            }
            #endregion

            UpdateIniFile();
        }

        /// <summary>Destructs this <see cref="Ini"/> instance.</summary>
        /// <param name="INIPath">Path to the INI file.</param>
        ~Ini()
        {
            if (bNeedUpdate)
                UpdateIniFile();
            sINIpath = null;
        }

        /// <summary>Write an entry in a section of the INI file. Use "fINI.UpdateIniFile();" to
        /// force the INI file to be written at the end of a series of write commands.</summary>
        /// <param name="Section">Section to write in.</param>
        /// <param name="Key">Name of key within the section.</param>
        /// <param name="Value">The value to write.</param> 
        public void WriteIniValue(string Section, string Key, string Value)
        {
            if ((Section != null) && (Key != null) && (Value != null) && (sINIpath != null))
            {
                try
                {
                    WritePrivateProfileString(Section, Key, Value, sINIpath);
                }
                catch { }
                bNeedUpdate = true;
            }
        }

        /// <summary>Read an entry in a section of the INI file. Use "fINI.UpdateIniFile();"</summary>
        /// <param name="Section">Section to write in.</param>
        /// <param name="Key">Name of key within the section.</param>
        /// <param name="Value">The value to write.</param> 
        public string ReadIniValue(string Section, string Key, string DefaultValue)
        {
            string sResult = new string('\0', 2048);
            int byteSize = 0;

            if ((Section != null) && (Key != null) && (sINIpath != null))
            {
                try
                {
                    byteSize = GetPrivateProfileString(Section, Key, DefaultValue, sResult, sResult.Length, sINIpath);
                }
                catch { }
            }
            if (byteSize <= 0)
                sResult = "";
            return sResult.Substring(0, byteSize);  // the last trailing char(0) is removed
        }

        /// <summary>Forces the cached data entries to be written to the INI file.</summary>
        public void UpdateIniFile()
        {
            WritePrivateProfileString(null, null, null, sINIpath);
            bNeedUpdate = false;
        }

        /// <summary>Reads an entry in a section of the INI file.</summary>
        /// <param name="Section">The section name where the key should be in.</param>
        /// <param name="Key">Name of key within the section.</param>
        /// <param name="DefaultValue">The default value in case the data is unavailable.</param>
        public string ReadString(string Section, string Key, string DefaultValue)
        {
            return ReadIniValue(Section, Key, DefaultValue);
        }

        public int ReadInteger(string Section, string Key, int DefaultValue)
        {
            int nNum = 0;
            string sResult = ReadIniValue(Section, Key, DefaultValue.ToString());

            try
            {
                nNum = Int32.Parse(sResult);
            }
            catch { }
            return nNum;
        }

        /// <summary>Reads a value in an INI file. Valid values are "true" or "false".</summary>
        public bool ReadBool(string Section, string Key)
        {
            string sResult = ReadIniValue(Section, Key, "false");
            return sResult.ToLower().Equals("true");
        }

        /// <summary>Write an entry in a section of the INI file. Use "fINI.UpdateIniFile();" to
        /// force the INI file to be written at the end of a series of write commands.</summary>
        /// <param name="Section">Section to write in.</param>
        /// <param name="Key">Name of key within the section.</param>
        /// <param name="Value">The value to write.</param>
        public void WriteString(string Section, string Key, string Value)
        {
            WriteIniValue(Section, Key, "\"" + Value + "\"");
        }

        /// <summary>Write an entry in a section of the INI file. Use "fINI.UpdateIniFile();" to
        /// force the INI file to be written at the end of a series of write commands.</summary>
        /// <param name="Section">Section to write in.</param>
        /// <param name="Key">Name of key within the section.</param>
        /// <param name="Value">The value to write.</param>
        public void WriteInteger(string Section, string Key, int Value)
        {
            WriteIniValue(Section, Key, Value.ToString());
        }

        /// <summary>Writes a value in an INI file. Boolean is converted to "true" or "false".</summary>
        public void WriteBool(string Section, string Key, bool Value)
        {
            if (Value)
                WriteIniValue(Section, Key, "true");
            else
                WriteIniValue(Section, Key, "false");
        }

        /// <summary>Reads the contents of a value in an INI file. If it is not the same it will update it.</summary>
        /// <param name="Section">Section to write in.</param>
        /// <param name="Key">Name of key within the section.</param>
        /// <param name="Value">The value to write.</param>
        public void UpdateString(string Section, string Key, string Value)
        {
            string sOldValue = ReadIniValue(Section, Key, "");
            if (sOldValue != Value)
            {
                WriteString(Section, Key, Value);
                UpdateIniFile();
            }
        }

        /// <summary>Deletes an entire section from an INI file.</summary>
        /// <param name="Section">Section to write in.</param>
        public void EraseSection(string Section)
        {
            if ((Section != null) && (sINIpath != null))
            {
                WritePrivateProfileString(Section, null, null, sINIpath);
                bNeedUpdate = true;
            }
        }

        /// <summary>Deletes a key and it's contents from a section in an INI file.</summary>
        /// <param name="Section">Section to write in.</param>
        /// <param name="Key">Name of key within the section.</param>
        public void DeleteKey(string Section, string Key)
        {
            if ((Section != null) && (Key != null) && (sINIpath != null))
            {
                WritePrivateProfileString(Section, Key, null, sINIpath);
                bNeedUpdate = true;
            }
        }

        /// <summary> Reads a whole section (keys and values) of the INI file.
        ///  example : ListBox1.Items.AddRange(fINI.ReadSection("Settings")); </summary>
        /// <param name="section">Section to read.</param>
        public string[] ReadSection(string Section)
        {
            string sResult = new string('\0', 2048);
            int byteSize = 0;

            try
            {
                byteSize = GetPrivateProfileSection(Section, sResult, sResult.Length, sINIpath);
            }
            catch { }
            if (byteSize <= 0)
            {
                sResult = "\0";
                byteSize = 1;
            }
            // the last trailing char(0) is removed
            return sResult.Substring(0, byteSize - 1).Split('\0');
        }

        /// <summary>
        /// Reads all keys of a section of the INI file.
        ///  example : comboBox1.Items.AddRange(fINI.ReadSectionKeys("Settings"));
        /// </summary>
        /// <param name="section">Section to read.</param>
        public string[] ReadSectionKeys(string Section)
        {
            String sResult = new String(' ', 2048);
            int byteSize = 0;

            if ((Section != null) && (sINIpath != null))
            {
                try
                {
                    byteSize = GetPrivateProfileString(Section, null, null, sResult, sResult.Length, sINIpath);
                }
                catch { }
            }
            if (byteSize <= 0)
            {
                sResult = "\0";
                byteSize = 1;
            }
            // the last trailing char(0) is removed
            return sResult.Substring(0, byteSize - 1).Split('\0');
        }

        /// <summary> Gets a list of available sections of the INI file.
        ///  example : comboBox1.Items.AddRange(fINI.ReadSections()); </summary>
        public string[] ReadSections()
        {
            byte[] buffer = new byte[2048];
            String sResult = "";

            try
            {
                int byteSize = GetPrivateProfileSectionNames(buffer, buffer.Length, sINIpath);
                if (byteSize > 1)
                {   // the last trailing char(0) is removed
                    sResult = System.Text.Encoding.Default.GetString(buffer, 0, byteSize - 1);
                }
            }
            catch { }
            return sResult.Split('\0');
        }

    }
}
