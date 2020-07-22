using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace TrayGuard
{
    class Document
    {

        public static List<string> pathList = new List<string>()
        {
            //AppDomain.CurrentDomain.BaseDirectory + "log\\",
            AppDomain.CurrentDomain.BaseDirectory + "pqm\\"
        };

        public static void CreateDocument()
        {
            foreach (string path in pathList)
            {
                Directory.CreateDirectory(path);
            }
        }
    }

    class Log : Document
    {
        public static void WriteLog(string SN, string judge)
        {
            string path = Document.pathList[0] + "log" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            using (StreamWriter file = new StreamWriter(path, true))
            {
                string str = DateTime.Now.ToString("yyyy/MM/dd,HH:mm:ss,")
                    + SN + ","
                    + judge;
                file.WriteLine(str);// 直接追加文件末尾，换行
            }
        }
    }

    class Pqm : Document
    {
        static string type = TfSQL.readIni_static("pqm", "type", Environment.CurrentDirectory + @"\csv.ini");
        static string factory = TfSQL.readIni_static("pqm", "factory", Environment.CurrentDirectory + @"\csv.ini");
        static string building = TfSQL.readIni_static("pqm", "building", Environment.CurrentDirectory + @"\csv.ini");
        static string line = TfSQL.readIni_static("pqm", "line", Environment.CurrentDirectory + @"\csv.ini");
        static string process = TfSQL.readIni_static("pqm", "process", Environment.CurrentDirectory + @"\csv.ini");
        static string inspect = TfSQL.readIni_static("pqm", "inspect", Environment.CurrentDirectory + @"\csv.ini");
        static string machineName = TfSQL.readIni_static("pqm", "MachineName", Environment.CurrentDirectory + @"\csv.ini");

        public static void WriteCSV(string SN, string checkItem, string checkTotal)
        {
            string fileName = type + factory + building + line + process + DateTime.Now.ToString("yyyyMMddHHmmss") + SN;
            string path = Document.pathList[1] + fileName + ".csv";
            using (StreamWriter file = new StreamWriter(path, true))
            {
                string[] csvStr = new string[] { type, factory, building, line, process,
                    SN, "", "", DateTime.Now.ToString("yy,MM,dd,HH,mm,ss"), "1", inspect, "0.0",
                    checkItem, checkTotal, "1", "MACHINE",machineName };

                string str = String.Join(",", csvStr);

                file.WriteLine(str);// 直接追加文件末尾，换行 
            }
        }
    }
}
