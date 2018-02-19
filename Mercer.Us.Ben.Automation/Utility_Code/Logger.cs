using System;

using System.IO;
using System.Text;
using Microsoft;

namespace Mercer.Us.Ben.Automation.Utility_Code
{

    public sealed class Logger
    {
        public static string logfilefolderpath = "";
        //public static StreamWriter aWriter;


        public static void createLogFile()
        {
            try
            {
                DateTime date = DateTime.Now;
                string datevalue = string.Format("{0:yyyy_MM_dd}", date);
                DateTime cal = DateTime.Now;
                string date_time = string.Format("{0:MM_dd_yyyy_HH_mm_ss}", cal);

                string currPrjDirpath = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory()));
                string projectName = "TestResults";
                string solutionPath = currPrjDirpath.Replace(projectName, "");
                string logfolder = solutionPath + "log" + Path.DirectorySeparatorChar + datevalue + Path.DirectorySeparatorChar;
                if (!Directory.Exists(logfolder))
                {
                    Directory.CreateDirectory(logfolder);
                    Console.WriteLine("Log folder path is created!!");
                }
                else
                {
                    Console.WriteLine("Log folder path is already created!!");
                }
                string logfilepath = logfolder + date_time + ".txt";



                if (!File.Exists(logfilepath))
                {
                    File.CreateText(logfilepath).Dispose();
                }

                logfilefolderpath = logfilepath;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.Write(e.StackTrace);
            }
        }

        public static void info(string mssg)
        {
            Console.WriteLine("info : " + mssg);
            log("info : " + mssg);
        }

        public static void warn(string mssg)
        {
            Console.WriteLine("warn : " + mssg);
            log("warn : " + mssg);
        }

        public static void error(string mssg)
        {
            Console.WriteLine("error : " + mssg);
            log("error : " + mssg);
        }

        public static void error(string mssg, Exception e)
        {
            Console.WriteLine("error : " + mssg + " " + e);
            log("error : " + mssg + " " + e);
        }

        public static void error(Exception e)
        {
            //		System.out.println("error : "+ e);
            log("error : " + e);
        }

        public static void fatal(string mssg, Exception e)
        {
            Console.WriteLine("fatal : " + mssg + " " + e);
            log("fatal : " + mssg + " " + e);
        }

        // Creates log 
        public static void log(string loddata)
        {
            try
            {
                string filpath = logfilefolderpath;
                //TimeZone tz = TimeZone.CurrentTimeZone;
                DateTime now = DateTime.Now;
                StreamWriter aWriter = new StreamWriter(filpath, true);
                string currentTime = string.Format("{0:yyyy.MM.dd hh:mm:ss }", now);
                try
                {
                    aWriter.WriteLine(currentTime + " " + loddata + "\n");
                    aWriter.Flush();

                }
                finally
                {
                    aWriter.Close();
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.Write(e.StackTrace);
                Console.WriteLine("Log | Exception " + e.StackTrace);
            }
        }
    }
}
