using System;
using System.IO;
using compare_lib;

namespace runner
{
    class Program
    {
        enum ExitCode : int {
            Succeed = 0,
            FailedToParsePlan = 1,
            FailedToCompare = 2,
            UnknownError = 3,
            InvalidArguments = 9
        }

        private static void PrintHelper(){
            Console.WriteLine("\r\nUsage :     runner.exe xml_file_path out_log_file_path");
            Console.WriteLine("\r\nExample :   runner.exe c:\\compareplan.xml c:\\out.log");
            Console.WriteLine("\r\nExit Code : Succeed = 0");
            Console.WriteLine("            FailedToParsePlan = 1");
            Console.WriteLine("            FailedToCompare = 2");
            Console.WriteLine("            UnknownError = 3");
            Console.WriteLine("            InvalidArguments = 9");
        }

        static int Main(string[] args)
        {
            Console.WriteLine("");
            if (args.Length != 2){
                Console.WriteLine("Error: Number of argument is invalid !");
                PrintHelper();
                return (int)ExitCode.InvalidArguments;
            }

            string xmlFIlePath = args[0];
            string logFIlePath = args[1];

            if(!File.Exists(xmlFIlePath)){
                Console.WriteLine("Error: Xml file argument doesn't exist !");
                PrintHelper();
                return (int)ExitCode.InvalidArguments;
            }

            if(!Directory.Exists(Path.GetDirectoryName(logFIlePath))){
                Console.WriteLine("Error: The directory of the log file argument doesn't exist !");
                PrintHelper();
                return (int)ExitCode.InvalidArguments;
            }

            Logger logger=null;
            try{
                try{
                    logger = new Logger(args[1]);
                }catch(Exception ex){
                    Console.WriteLine("Failed to create log file: " + ex.Message);
                    return (int)ExitCode.InvalidArguments;
                }
                RunnerPlan data;
                try{
                    logger.Log("Parse xml file to get the compare plan (" + xmlFIlePath + ")");
                    data = RunnerPlan.ParseFromXml(xmlFIlePath);
                    logger.Log("\tCleanRegEx : " + data.CleanRegEx);
                    logger.Log("\tCompShape : " + data.CompShape);
                    logger.Log("\tCompStyle : " + data.CompStyle);
                    logger.Log("\tCompValue : " + data.CompValue);
                }catch(Exception ex){
                    logger.Log("Failed to parse Xml file: " + ex.Message);
                    return (int)ExitCode.FailedToParsePlan;
                }

                Compare compare;
                try{
                    compare = new Compare();
                    compare.InfoEvent += new InfoUpdateEventHandler(logger.Log);
                    compare.CompareFiles(data.FilesA, data.FilesB, data.CleanRegEx, data.ReportFolder, data.CompValue, data.CompStyle, data.CompStyle, true);
                }catch(Exception ex){
                    logger.Log("Failed to compare files: " + ex.Message);
                    return (int)ExitCode.FailedToCompare;
                }
                logger.Log("End of comparison.");
                return (int)ExitCode.Succeed;
            }catch(Exception){
                return (int)ExitCode.UnknownError;
            }finally{
                logger.Close();
            }

        }

        private class Logger
        {
            FileStream fileStream;
            StreamWriter streamWriter;

            public Logger(string filename){

                fileStream = new FileStream(filename, FileMode.Create);
                streamWriter = new StreamWriter(fileStream);
            }

            public void LogAndClose(string s){
                Log(s);
                Close();
            }

            public void Log(string s){
                streamWriter.WriteLine(getTime() + "  " + s);
                Console.WriteLine(s);
            }

            public void Close(){
                streamWriter.Close();
                fileStream.Close();
            }

            private string getTime(){
                return DateTime.Now.ToString(@"yyyy/MM/dd HH:mm:ss");
            }
        }

    }
}
