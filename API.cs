using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;

namespace TrayGuard
{
    class API
    {
        static string API_Path = TfSQL.readIni_static("API", "PATH", Environment.CurrentDirectory + @"\config.ini");
        public static string Judge(string SN)
        {
            //POST参数param
            string POSTparam = "module_sn=" + SN;

            // param转换
            byte[] data = Encoding.ASCII.GetBytes(POSTparam);

            ServicePointManager.Expect100Continue = false; //HTTP错误（417）对应

            //创建请求
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(API_Path);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;

            //POST写入
            try
            {
                using (Stream reqStream = request.GetRequestStream())
                    reqStream.Write(data, 0, data.Length);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); Environment.Exit(0); }

            //取得响应
            WebResponse response = request.GetResponse();

            //读取结果
            string APIstr = "";
            using (Stream resStream = response.GetResponseStream())
            using (var reader = new StreamReader(resStream, Encoding.GetEncoding("UTF-8")))
                APIstr = reader.ReadToEnd();

            if (APIstr == "{}")
            {
                MessageBox.Show("接收到空值。\r\n请检查API的值是否正确。\r\n将会关闭程序。", "API接收", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }

            JObject JO = null;
            try { JO = JObject.Parse(APIstr); }
            catch { MessageBox.Show("返回的Json:\r\n" + APIstr, "解析Json失败", MessageBoxButtons.OK, MessageBoxIcon.Error); Environment.Exit(0); }
            //关闭
            string result;
            switch ((int)JO["result"])
            {
                case 0:
                    result = "OK";
                    break;
                case 1:
                    result = "NG";
                    break;
                case 2:
                    result = "KEEP";
                    break;
                case 3:
                    result = "DUPLICATE";
                    break;
                case 4:
                    result = "HOLD";
                    break;
                case 5:
                    result = "SCRAP";
                    break;
                case 6:
                    result = "RETEST";
                    break;
                case 7:
                    result = "RECHECK";
                    break;
                default:
                    result = (string)JO["result"];
                    break;
            }
            return result;
        }
    }
}
