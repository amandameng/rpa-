using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ConsoleApp1.helpers
{
    class PiaoXiaoMiInvoiceParser
    {
        // 机动车购置发票识别
        static void ParseInvoiceJson()
        {
            //string file = @"D:\云扩科技\RPA\OCR\机动车购置.txt";

            string file = @"D:\云扩科技\RPA\OCR\票小秘-识别加查验.txt";
            string result = File.ReadAllText(file);

            JObject jObject = (JObject)JsonConvert.DeserializeObject(result);
            JValue resultValue = (JValue)jObject["result"];
            Console.WriteLine(resultValue.Value);

            JObject data = (JObject)jObject["response"]["data"];
            JArray identify_results = (JArray)data["identify_results"];

            //Console.WriteLine(data["identify_results"]);

            DataTable resultTable = new DataTable();

            List<JToken> items = identify_results.Children().ToList();
            JObject item = (JObject)items[0];
            //foreach (JObject item in items)
            // {
            JValue type = (JValue)item["type"];
            Console.WriteLine(type.Value.ToString());

            JObject validation = (JObject)item["validation"];
            string ValdateStatus = "";
            foreach (JProperty jp in (JToken)validation)
            {
                string name = jp.Name.ToString();
                string value = jp.Value.ToString();
                if ("code".Equals(name))
                {
                    Console.WriteLine("name: " + name + "---code:" + value);

                    switch (value)
                    {
                        case "10000":
                            ValdateStatus = "正常";
                            break;
                        case "10001":
                            ValdateStatus = "查无此票";
                            break;
                        case "10002":
                            ValdateStatus = "查验信息不一致";
                            break;
                        case "10003":
                            ValdateStatus = "验真次数超过限制，同一张票一天最多可以查验5次";
                            break;
                        case "10004":
                            ValdateStatus = "不支持验真发票类型";
                            break;
                        case "10005":
                            ValdateStatus = "无效参数";
                            break;
                        case "10006":
                            ValdateStatus = "其它错误";
                            break;
                    }

                }

                Console.WriteLine(ValdateStatus);
            }

            JObject details = (JObject)item["details"];

            DataTable dt = new DataTable();
            DataRow row = dt.NewRow();

            //foreach (JProperty jp in (JToken)details)
            //{
            //    string name = jp.Name.ToString();
            //    dt.Columns.Add(name);
            //}

            foreach (JProperty jp in (JToken)details)
            {
                string name = jp.Name.ToString();
                string value = jp.Value.ToString();
                switch (name)
                {
                    case "code":
                        Console.WriteLine("code:" + value);
                        break;
                    case "number":
                        Console.WriteLine("number:" + value);
                        break;
                    case "date":
                        DateTime tmpDate;
                        bool isDate = DateTime.TryParse(value, out tmpDate);
                        if (isDate)
                        {
                            Console.WriteLine(tmpDate.ToString("yyyy-MM-dd"));
                        }
                        break;
                    case "pretax_amount":
                        Console.WriteLine("pretax_amount:" + value);
                        break;
                    case "buyer":
                        Console.WriteLine("buyer:" + value);
                        break;
                    case "buyer_id":
                        Console.WriteLine("buyer_id:" + value);
                        break;
                }
            }


            dt.Rows.Add(row);
            resultTable = dt.Copy();

            //DataColumnCollection allSrcDataColumns = dt.Columns;

            foreach (DataColumn col in dt.Columns)
            {
                Console.WriteLine(col.ColumnName);
            }

            //}
        }
    }
}
