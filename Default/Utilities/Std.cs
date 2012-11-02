/****************************** Module Header ******************************\
File Name:    Default\Std.cs
Purpose:      General use in any C# project; various functions regarding the Environment
By:           
Date:         
Version:      
\***************************************************************************/

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;

namespace Default.Utilities
{
	public static class Std
	{
        /// <summary>Returns the base path of the location of this executable or dll.</summary>
        public static string GetBasePath()
        {
            string sPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
            if (sPath.StartsWith("file:\\"))
                sPath = sPath.Remove(0, 6);
            return sPath + "\\";
        }

        public static string GetExeName()
        {
            return Assembly.GetEntryAssembly().Location;
        }

        public static string GetFileNameOnly(string sFilePath)
        {
            return Path.GetFileNameWithoutExtension(sFilePath);
        }

        public static string GetVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        public static string AssemblyTitle()
        {
            // Get all Title attributes on this assembly
            object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
            // If there is at least one Title attribute
            if (attributes.Length > 0)
            {
                // Select the first one
                AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                // If it is not an empty string, return it
                if (titleAttribute.Title != "")
                    return titleAttribute.Title;
            }
            // If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
            return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
        }

        public static string AssemblyVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        public static string AssemblyDescription()
        {
            // Get all Description attributes on this assembly
            object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
            // If there aren't any Description attributes, return an empty string
            if (attributes.Length == 0)
                return "";
            // If there is a Description attribute, return its value
            return ((AssemblyDescriptionAttribute)attributes[0]).Description;
        }

        public static string AssemblyProduct()
        {
            // Get all Product attributes on this assembly
            object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
            // If there aren't any Product attributes, return an empty string
            if (attributes.Length == 0)
                return "";
            // If there is a Product attribute, return its value
            return ((AssemblyProductAttribute)attributes[0]).Product;
        }

        public static string AssemblyCopyright()
        {
            // Get all Copyright attributes on this assembly
            object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
            // If there aren't any Copyright attributes, return an empty string
            if (attributes.Length == 0)
                return "";
            // If there is a Copyright attribute, return its value
            return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
        }

        public static string AssemblyCompany()
        {
            // Get all Company attributes on this assembly
            object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
            // If there aren't any Company attributes, return an empty string
            if (attributes.Length == 0)
                return "";
            // If there is a Company attribute, return its value
            return ((AssemblyCompanyAttribute)attributes[0]).Company;
        }

        public static string AssemblyGUID()
        {
            Guid assemblyGuid = Guid.Empty;

            object[] assemblyObjects = System.Reflection.Assembly.GetEntryAssembly().GetCustomAttributes(
                                         typeof(System.Runtime.InteropServices.GuidAttribute), true);

            if (assemblyObjects.Length > 0)
                return ((System.Runtime.InteropServices.GuidAttribute)assemblyObjects[0]).Value;
            else
                return "";
        }


        /// <summary>Returns a part of a string which is separated by comma's.</summary>
        /// <param name="sLine">The string to extract the substring from.</param>
        /// <param name="nItem">The index of the part you want to extract. The first part is index 1.</param>
        /// <param name="cSeparator">The separation character.</param>
        public static string Get_MyItem(string sLine, int nItem, char cSeparator)
        {
            string[] sList = sLine.Split(cSeparator);

            if ((sList.Length >= nItem) && (nItem > 0))
                return sList[nItem - 1];
            else
                return "";
        }

        /// <summary>Returns a part of a string which is separated by comma's.</summary>
        /// <param name="sLine">The string to extract the substring from.</param>
        /// <param name="nItem">The index of the part you want to extract. The first part is index 1.</param>
        public static string Get_Item(string sLine, int nItem)
        {
            return Get_MyItem(sLine, nItem, ',');
        }

        /// <summary>Returns an integer by extracting a value from a string which is separated by comma's.</summary>
        /// <param name="sLine">The string to extract the value from.</param>
        /// <param name="nItem">The index of the part you want to convert. The first part is index 1.</param>
        public static int Get_IntItem(string sLine, int nItem)
        {
            int nNum = 0;
            string sRes = Get_MyItem(sLine, nItem, ',');

            try
            {
                nNum = Int32.Parse(sRes);
            }
            catch { }
            return nNum;
        }
	}
}
