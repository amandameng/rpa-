using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Class1
    {

        const string appSecret = "5d107792";    //这里输入提供的app_secret   
        const string appKey = "cbf8162a5ab1d29f8b3abbf7068f72d9";   //这里输入提供的app_key  

        const string image_url = "http://fapiao.glority.cn/dist/img/sample.jpg";
        public static string CalculateMD5Hash(string input)
        {
            // step 1, calculate MD5 hash from input 
            MD5 md5 = System.Security.Cryptography.MD5.Create();
            byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
            byte[] hash = md5.ComputeHash(inputBytes);

            // step 2, convert byte array to hex string 
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                sb.Append(hash[i].ToString("X2"));
            }
            return sb.ToString().ToLower();
        }

        public static double ConvertToUnixTimestamp(DateTime date)
        {
            DateTime origin = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            TimeSpan diff = date.ToUniversalTime() - origin;
            return Math.Floor(diff.TotalSeconds);
        }

        static void Main(string[] args)
        {
            double timeStamp = ConvertToUnixTimestamp(DateTime.Now);

            string token = CalculateMD5Hash(appKey + '+' + timeStamp + '+' + appSecret);
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("app_key", appKey);
            dic.Add("timestamp", timeStamp.ToString());
            dic.Add("token", token);
            dic.Add("image_url", image_url);

            string result = "";
            string url = "http://fapiao.glority.cn/v1/item/get_item_info";
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";
            #region 添加Post 参数 
            StringBuilder builder = new StringBuilder();
            int i = 0;
            foreach (var item in dic)
            {
                if (i > 0)
                    builder.Append("&");
                builder.AppendFormat("{0}={1}", item.Key, item.Value);
                i++;
            }
            byte[] data = Encoding.UTF8.GetBytes(builder.ToString());
            req.ContentLength = data.Length;
            using (Stream reqStream = req.GetRequestStream())
            {
                reqStream.Write(data, 0, data.Length);
                reqStream.Close();
            }
            #endregion
            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            Stream stream = resp.GetResponseStream();
            //获取响应内容 
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                result = reader.ReadToEnd();
            }

            Console.WriteLine(result);

            Console.ReadLine();
        }
    }
}
