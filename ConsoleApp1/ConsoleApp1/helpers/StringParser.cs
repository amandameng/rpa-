using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleApp1.helpers
{
    class StringParser
    {
        public static void GetRiskRate()
        {
            string 产品全名称 = "“乐惠”2020年第33期（保6M）-PR3";
            string[] 产品名称数组 = 产品全名称.Split(new string[] { "（", "-" }, StringSplitOptions.RemoveEmptyEntries);
            string 风险等级 = string.Empty;
            string 产品名称 = 产品名称数组[0];

            Console.WriteLine(产品名称.Split(new string[] { "“", "”" }, StringSplitOptions.RemoveEmptyEntries)[0]);
            Console.WriteLine(产品名称.Split(new string[] { "“", "”" }, StringSplitOptions.RemoveEmptyEntries)[1]);


            if (产品名称数组.Length > 0)
            {
                风险等级 = 产品名称数组[产品名称数组.Length - 1];
            }
            Console.WriteLine(风险等级);
        }

        public static void StrCapture()
        {
            //string 金额 = "“乐惠”2020年第64期（3M-最美逆行者定向白名单）";
            string 金额 = "“乐惠月月赢”2020年第64期（3M-最美逆行者定向白名单）";

            Regex nameRxg2 = new Regex(@"(月月赢|双季赢|日日升|周周乐)");

            Console.WriteLine(nameRxg2.IsMatch(金额));
            MatchCollection matchedGroup2 = nameRxg2.Matches(金额);
            if (matchedGroup2.Count > 0 && matchedGroup2[0].Groups.Count > 1)
            {
                string projectName = matchedGroup2[0].Value.Trim();
                Console.WriteLine(projectName);
            }
            else
            {
                Console.WriteLine("No match price part1");
            }

        }
        public static void arrJoin()
        {
            string[] emails = "zhanglingling@zhlxsemicon.com;lijuan@zhlxsemicon.com;zhanglingling@zhlxsemicon.com".Split(';', StringSplitOptions.None).Distinct().ToArray();
            Console.WriteLine(String.Join(";", emails));
        }
    }
}
