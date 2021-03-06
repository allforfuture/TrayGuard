﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace TrayGuard
{
    class Pqm
    {
        readonly static string iniPath = Environment.CurrentDirectory + @"\config.ini";
        readonly static bool isNewFormat = TfSQL.readIni_static("CSV", "isNewFormat", iniPath) == "0" ? false : true;
        readonly static string type = TfSQL.readIni_static("CSV", "type", iniPath);
        readonly static string factory = TfSQL.readIni_static("CSV", "factory", iniPath);
        readonly static string building = TfSQL.readIni_static("CSV", "building", iniPath);
        readonly static string lineString = TfSQL.readIni_static("CSV", "lineString", iniPath);
        static string[] lineArr = lineString.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
        static Dictionary<string, string> lineDic = lineArr.ToDictionary(
            sKey => sKey.Split('|')[0],
            sElement => sElement.Split('|')[1]
            );
        readonly static string process = TfSQL.readIni_static("CSV", "process", iniPath);
        readonly static string inspect = TfSQL.readIni_static("CSV", "inspect", iniPath);
        readonly static string machineName = TfSQL.readIni_static("CSV", "MachineName", iniPath);
        readonly static string pqmPath = TfSQL.readIni_static("CSV", "path", iniPath);

        public static void CreateDocument()
        {
            Directory.CreateDirectory(pqmPath);
        }

        public static void WriteCSV(string SN, string checkItem, string checkTotal)
        {
            string line = lineDic[SN.Substring(7,1)];
            string fileName = type + factory + building + line + process + DateTime.Now.ToString("yyyyMMddHHmmss") + SN + ".csv";
            string path = Path.Combine(pqmPath, fileName);
            using (StreamWriter file = new StreamWriter(path, true))
            {
                string[] csvStr = new string[] { type, factory, building, line, process,
                    SN, "", "", DateTime.Now.ToString("yy,MM,dd,HH,mm,ss"), "1", inspect, "0.0",
                    checkItem, checkTotal, "1", "MACHINE",machineName };
                string[] csvStr_New = new string[] { type, factory, building, line, process,
                    SN, "", "", "", DateTime.Now.ToString("yy,MM,dd,HH,mm,ss"), "N/A", "N/A", "1", inspect, "0.0",
                    checkItem, checkTotal, "1", "MACHINE",machineName };

                string str = isNewFormat ? String.Join(",", csvStr_New) : String.Join(",", csvStr);

                file.WriteLine(str);// 直接追加文件末尾，换行 
            }
        }
    }
}
