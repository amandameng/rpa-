using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {

        const string appSecret = "cbf8162a5ab1d29f8b3abbf7068f72d9";    //这里输入提供的app_secret   
        const string appKey = "5d107792";   //这里输入提供的app_key  

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
            Console.WriteLine(origin);
            Console.WriteLine(date.ToUniversalTime());
            TimeSpan diff = date.ToUniversalTime() - origin;
            return Math.Floor(diff.TotalSeconds);
        }
        public static string GetBase64FromImage2(string imagefile)
        {
            string strbaser64 = "";
            try
            {
                Console.WriteLine(imagefile);
            Bitmap bmp = new Bitmap(imagefile);
            MemoryStream ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            byte[] arr = new byte[ms.Length];
            ms.Position = 0;
            ms.Read(arr, 0, (int)ms.Length);
            ms.Close();
            strbaser64 = Convert.ToBase64String(arr);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw new Exception("Something wrong during convert!");
            }
            return strbaser64;
        }
        public static byte[] GetByteFromImage(string imagefile)
        {
            byte[] arr = new byte[] { };
            try
            {
                Console.WriteLine(imagefile);
                Bitmap bmp = new Bitmap(imagefile);
                MemoryStream ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                arr = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(arr, 0, (int)ms.Length);
                ms.Close();
               
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw new Exception("Something wrong during convert!");
            }
            return arr;
        }

        public static byte[] FileStreamReadToBytes(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return new byte[] { };
            }

            byte[] bytesArr = null;
            using (FileStream fs = new FileStream(fileName, FileMode.Open))
            {
                bytesArr = new byte[fs.Length];
                fs.Write(bytesArr, 0, bytesArr.Length);
                return bytesArr;
            }
        }


        public static byte[] ReadFileToBinaryBytes(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return new byte[] { };
            }

            byte[] bytesArr = null;
            using (FileStream fs = new FileStream(fileName, FileMode.Open))
            {
                using (BinaryReader binReader = new BinaryReader(fs, Encoding.UTF8))
                {
                    bytesArr = binReader.ReadBytes((int)fs.Length);
                    return bytesArr;
                }
            }
        }

        public static byte[] GetBase64FromImage(string imagefile)
        {
            //string strbaser64;
            try
            {
                byte[] imageArr;

                using (FileStream fs = new FileStream(imagefile, FileMode.Open))
                {
                    byte[] bytesArr = new byte[fs.Length];
                    fs.Read(bytesArr, 0, bytesArr.Length);
                    imageArr = bytesArr;
                    fs.Close();
                }
                return imageArr;
                //strbaser64 = Convert.ToString(imageArr);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw new Exception("Something wrong during convert!");
            }
            //return strbaser64;
        }

        static void Main(string[] args)
        {
            using (var pipeClient = new NamedPipeClientStream(".", "RobotPipe_a80f08d5-c2ca-45e8-99c2-a1cf494733e4", PipeDirection.In))
            {
                pipeClient.Connect();

                using (var sr = new StreamReader(pipeClient))
                {
                    string url;
                    while ((url = sr.ReadLine()) != null)
                    {
                        Console.WriteLine(url);
                    }
                }
            }

            double timeStamp = ConvertToUnixTimestamp(DateTime.Now);
            string imageFile = @"D:\联合银行\解放路支行发票需求\发票\发票\二手车发票2.jpg";
            byte[] imageArr = GetBase64FromImage(imageFile);
            //byte[] image_file = FileStreamReadToBytes(imageFile);

            string token = CalculateMD5Hash(appKey + '+' + timeStamp + '+' + appSecret);
            string result = "";
            string imgBase64 = Convert.ToBase64String(imageArr);
            using (HttpClient httpClient = new HttpClient())
            {
                MultipartFormDataContent form = new MultipartFormDataContent();
                form.Add(new StringContent(appKey), "app_key");
                form.Add(new StringContent(timeStamp.ToString()), "timestamp");
                form.Add(new StringContent(token), "token");
                form.Add(new StringContent(imgBase64), "image_data");

                HttpResponseMessage response = httpClient.PostAsync("http://fapiao.glority.cn/v1/item/get_item_info_with_validation", form).Result;
                response.EnsureSuccessStatusCode();
                result = response.Content.ReadAsStringAsync().Result;
            }

            if (string.IsNullOrEmpty(result))
            {
                throw new Exception("返回的票据信息结果为空");
            }
            Console.WriteLine(result);
            //var pxmInvoiceResponse = JsonConvert.DeserializeObject<PxmInvoiceResponse>(result);
            //if (pxmInvoiceResponse == null)
            //{
            //    throw new Exception("解析票据信息结果为空");
            //}
            //if (pxmInvoiceResponse.Result == 0)
            //{
            //    throw new Exception($"ErrorCode:{pxmInvoiceResponse.Error}, Message:{pxmInvoiceResponse.Message}");
            //}
           



            //    Dictionary<string, string> dic = new Dictionary<string, string>();
            //    dic.Add("app_key", appKey);
            //    dic.Add("timestamp", timeStamp.ToString());
            //    dic.Add("token", token);
            //    //dic.Add("image_url", image_file);
            //    //dic.Add("image_data", image_data);
            //    dic.Add("image_data", Convert.ToBase64String(image_data));
            //    Console.WriteLine(timeStamp); //1590416212
            //    Console.WriteLine(token); //9ab2e3d0f3cdf7fc0660bba164ae966f

            //    string result = "";
            //    string url = "http://fapiao.glority.cn/v1/item/get_item_info";

            //    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            //    req.Method = "POST";
            //    req.ContentType = "application/x-www-form-urlencoded";
            //    #region 添加Post 参数 
            //    StringBuilder builder = new StringBuilder();
            //    int i = 0;
            //    foreach (var item in dic)
            //    {
            //        if (i > 0)
            //            builder.Append("&");
            //        builder.AppendFormat("{0}={1}", item.Key, item.Value);
            //        i++;
            //    }
            //    byte[] data = Encoding.UTF8.GetBytes(builder.ToString());
            //    req.ContentLength = data.Length;
            //    using (Stream reqStream = req.GetRequestStream())
            //    {
            //        reqStream.Write(data, 0, data.Length);
            //        reqStream.Close();
            //    }
            //    #endregion
            //    HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            //    Stream stream = resp.GetResponseStream();
            //    //获取响应内容 
            //    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            //    {
            //        result = reader.ReadToEnd();
            //    }

            //    Console.WriteLine(result);

            //    Console.ReadLine();
        }
    }
}
