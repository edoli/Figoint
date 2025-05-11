namespace EdoliAddIn
{
    class DebugUtil
    {
        public static System.IO.StreamWriter logFile = new System.IO.StreamWriter("log.txt");


        public static void WriteLine(object obj)
        {
            logFile.WriteLine(obj);
            logFile.Flush();
        }
    }
}
