using System;
using System.IO;
using System.Net;
using System.Text;

namespace Geo
{
    class HttpGetHelper
    {
        /// <summary>
        /// 高德地图解析函数
        /// </summary>
        /// <param name="strResult">返回结果</param>
        public string GaoDeAnalysis(string parameters)
        {
            string strResult = "";
            
            string url = string.Format("http://restapi.amap.com/v3/geocode/geo?{0}", parameters);
            try
            {
                HttpWebRequest req = WebRequest.Create(url) as HttpWebRequest;
                req.ContentType = "multipart/form-data";
                req.Accept = "*/*";

                req.UserAgent = "";
                req.Timeout = 30000;//30秒连接不成功就中断 
                req.Method = "GET";
                req.KeepAlive = true;

                HttpWebResponse response = req.GetResponse() as HttpWebResponse;
                using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    strResult = sr.ReadToEnd();
                }

                string[] strs = strResult.Split('"');

                int j = 0;
                bool isFinded = false;
                for (int i = 0; i < strs.Length; i++)
                {
                    if (strs[i] == "location")
                    {
                        j = i + 2;
                        isFinded = true;
                        break;
                    }
                }

                if (isFinded)
                {
                    strResult = strs[j];
                }
                else
                {
                    strResult = "NULL";
                }
            }
            catch (Exception ex)
            {
                strResult = "NULL";
            }

            return strResult;
        }
    }
}
