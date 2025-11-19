using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LChart_Comparison_Tool
{ 
    class Utility
    {
        static int myCnt = 0;
        public static string m_sTempPath = Path.GetTempPath() /*+ @"\bin\Debug\"*/;
        public static string m_sBinPath = System.IO.Directory.GetCurrentDirectory();

        public static string m_sToolName = "CleaningTicket";
        public static string m_sVersion = "1.2";
        public static string m_sDate = "19-Jan-24";
        public static void WriteErrorLog(string sWarning, string sStackTrace, string sMsg)
        {
            try
            {
                if (sWarning.Trim().Length > 0)
                    System.Windows.Forms.MessageBox.Show(sWarning.Trim());
                string sLogFilepath = m_sBinPath + @"\"+ "ErrorLog.dat";
                if (sLogFilepath != null && sLogFilepath.Length > 0)
                {
                    try
                    {
                        //Write the Error Log
                        StreamWriter sw;
                        if (myCnt == 0)//First Error
                        {
                            sw = System.IO.File.CreateText(sLogFilepath);
                            DateTime date = DateTime.Now;
                            sw.WriteLine("********** " + m_sToolName + " V" + m_sVersion + " *********");
                            sw.WriteLine("Version           : " + m_sVersion);
                            sw.WriteLine("Release Date      : " + m_sDate);
                            sw.WriteLine("Time              : " + date.ToString());
                            sw.WriteLine("**********************************************");
                        }
                        else//Next Errors Append
                        {
                            sw = System.IO.File.AppendText(sLogFilepath);
                        }
                        //----- This is for error
                        if (sWarning.Length > 0 || sStackTrace.Length > 0)
                        {
                            sw.WriteLine("-----------------");
                            if (sWarning.Length > 0)
                                sw.WriteLine("Warning     : " + sWarning);//Write the Error Message
                            if (sStackTrace.Length > 0)
                                sw.WriteLine("StackTrace  : " + sStackTrace);//Write the StackTrace
                            if (sStackTrace.Length > 0)
                                sw.WriteLine("Msg  : " + sMsg);//Write the StackTrace
                            sw.WriteLine("-----------------");
                        }
                        else if (sMsg.Length > 0)
                        {
                            sw.WriteLine(sMsg);
                        }
                        //close
                        sw.Close();
                        myCnt++;
                    }
                    catch
                    { }
                }
            }
            catch { }
        }
        //*******************************************************
        //Function  : WriteErrorLog
        //Purpose   : Write the error data to log
        //*******************************************************
        public static void WriteErrorLog(Exception ee)
        {
            try
            {
                if (ee != null)
                {
                    WriteErrorLog(ee.Message, ee.StackTrace, "");
                }
            }
            catch { }
        }

    }
    
}
