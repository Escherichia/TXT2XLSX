/****************************** Module Header ******************************\
File Name:    Default\Debug.cs
Purpose:      General use in any C# project print debug information from a
              project to the terminal or to a debug-log file
By:           Jos van Nijnatten
Date:         29-09-2012
Version:      1.0
\***************************************************************************/

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Default.Utilities;

namespace Default.Utilities
{
    public partial class Debug
    {
        #region Variables
        public  enum   MODE       {NONE, CONSOLE, FILE};
        private static MODE       mode = MODE.NONE;
        private static Log        dbgLog;
        public  static MODE       debug
        {
            get
            {
                return mode;
            }
            set
            {
                if (mode == MODE.NONE)
                {
                    if (value == MODE.FILE)
                    {
                        dbgLog = new Log(@"Log\debug.log");
                    }
                    mode = value;
                }
                else
                {
                    throw new Exception("You should only set the debug mode at one place!");
                }
            }
        }
        public  static StackTrace trace
        {
            get
            {
                return new StackTrace();
            }
        }
        #endregion

        #region Static functions
        public static void write(string message) 
        {
            write(message, null);
        }

        public static void write(Exception e) {
            write(e.Message + "\n" + e.StackTrace);
        }

        public static void write(string message, StackTrace stacktrace)
        {
            // 
            if (stacktrace != null)
            {
                MethodBase method     = (stacktrace.GetFrame(1).GetMethod());
                string     filename   = (stacktrace.GetFrame(1).GetFileName());
                int        linenumber = (stacktrace.GetFrame(1).GetFileLineNumber());
                message += " (called from method '" + method.Name + "' in '" + method.Module + 
                    ((filename == null) ? null : 
                        (":" + filename + 
                            ((linenumber == 0) ? null : (":" + linenumber))
                        )
                    ) + "')";
            }
            
            // Do the stuff
            switch(mode)
            {
                case MODE.CONSOLE:
                    Console.WriteLine(message);
                    break;
                case MODE.FILE:
                    if (dbgLog != null)
                    {
                        dbgLog.LogMessage(message);
                    }
                    break;
                default:
                    break;
            }

        }
        #endregion
    }
}
