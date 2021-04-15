using System;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.CSharp;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using System.Reflection;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.ComponentModel;
using Xceed.Wpf.Toolkit;
using System.Security.Cryptography;
using HtmlAgilityPack;
using System.Windows.Forms;
using System.Drawing.Imaging;
using ImageMagick;
using ConsoleApp1.helpers;
using System.Globalization;

public class Class1
{
    public static void SaveTxtFile(string filePath)
    {
        String[] arr1 = new String[] { @"D:\文件监听用\源\rpa-1.txt", @"D:\文件监听用\源\rpa-2.txt" };

        
        string text = string.Join("\n", arr1); //\r\n表示换一行
        try
        {
            using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write))
            {
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(text);
                sw.Flush();
                sw.Close();
                fs.Close();
            }
        }
        catch
        {
            //"保存文本文件出错！"
        }
    }

    // 去哪网抓飞机票价
    public static void parsePrice()
    {
        String priceHtml = "<b style=\"width:64px;left:-64px\"><i style='width: 16px;'>5</i><i style='width: 16px;'>9</i><i style='width: 16px;'>8</i><i style='width: 16px;'>0</i></b><b style='width: 16px;left:-48px'>0</b><b style='width: 16px;left:-32px'>7</b><b style='width: 16px;left:-64px'>1</b>";
        String[] priceArr = priceHtml.Split(new string[] { "</b>" }, StringSplitOptions.RemoveEmptyEntries);
        
        String pricePart1 = priceArr[0];  //"<b style='width:48px;left:-48px'><i style='width: 16px;'>4</i><i style='width: 16px;'>1</i><i style='width: 16px;'>7</i>";

        Regex widthRgx = new Regex(@"<b style=""width:(\d+)px");
        MatchCollection mc = widthRgx.Matches(pricePart1);
        Int32 priceDigitNum = 0;
        if (mc.Count > 0 && mc[0].Groups.Count > 1)
        {
            priceDigitNum = Convert.ToInt32(mc[0].Groups[1].Value.Trim())/16;
        }
        else
        {
            Console.WriteLine("No match price part1");
        }
        Regex priceShadowRgx = new Regex(@">(\d)<");
        MatchCollection shadowPriceMc = priceShadowRgx.Matches(pricePart1);



        Int32[] PriceShadowArr = new Int32[priceDigitNum];


        if (shadowPriceMc.Count > 0)
        {
            for(Int32 index = 0; index < shadowPriceMc.Count; index++)
            {
                PriceShadowArr[index] = Convert.ToInt32(shadowPriceMc[index].Groups[1].Value.Trim());
            }
        }
        else
        {
            Console.WriteLine("No match price");
        }

        Dictionary<Int32, Int32> positionMap = new Dictionary<Int32, Int32>();
       // positionMap.ContainsKey

        // get the Real price
        for (Int32 i = 1; i < priceArr.Length; i++)
        {
            Regex RealPriceRgx = new Regex(@"left:(-\d+)px'.*>(\d)");
            MatchCollection priceRealMc = RealPriceRgx.Matches(priceArr[i]);
            Int32 mapKey = 0;
            if (priceRealMc.Count > 0)
            {
                for (Int32 index = 0; index < priceRealMc[0].Groups.Count; index++)
                {
                    if (index == 1)
                    {
                        mapKey = Convert.ToInt32(priceRealMc[0].Groups[1].Value.Trim()) / 16  + priceDigitNum;
                    }
                    else if (index == 2)
                    {
                        positionMap[mapKey] = Convert.ToInt32(priceRealMc[0].Groups[2].Value.Trim());
                    }

                }
            }
            else
            {
                Console.WriteLine("No match in" + priceArr[i]);
            }
        }


        for(Int32 index=0; index < PriceShadowArr.Length; index++)
        {
            try
            {
                PriceShadowArr[index] = positionMap[index];
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        String RealPrice = "";
        foreach(Int32 i in PriceShadowArr)
        {
            RealPrice = RealPrice + i;
        }

    }

    public static string 拼音字头(string strText)
    {
        string file = @"D:\拼音对照.txt";
        string 拼音字头列表 = File.ReadAllText(file);

        if (strText == null || strText.Length == 0)
            return strText;
        System.Text.StringBuilder myStr = new System.Text.StringBuilder();
        foreach (char vChar in strText)
        {
            if ((int)vChar >= 19968 && (int)vChar <= 40869)
                myStr.Append(拼音字头列表[(int)vChar - 19968]);
            else
                myStr.Append(char.ToUpper(vChar));
        }
        return myStr.ToString();
    }

    public static string GenerateMD5(string txt)
    {
        using (MD5 mi = MD5.Create())
        {
            byte[] buffer = Encoding.Default.GetBytes(txt);
            //开始加密
            byte[] newBuffer = mi.ComputeHash(buffer);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < newBuffer.Length; i++)
            {
                sb.Append(newBuffer[i].ToString("x2"));
            }
            return sb.ToString();
        }
    }

    public static void htmltableToDataTable()
    {
        string file = @"D:\Bottime\Project1\tablehtml.txt";
        string htmlCode = File.ReadAllText(file);
        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
        doc.LoadHtml(htmlCode);
        var headers = doc.DocumentNode.SelectNodes("//tr/th");
        DataTable table = new DataTable();
        foreach (HtmlNode header in headers)
            table.Columns.Add(header.InnerText); // create columns from th
                                                 // select rows with td elements 
        foreach (HtmlNode row in doc.DocumentNode.SelectNodes("//tr[td]"))

            table.Rows.Add(row.SelectNodes("td").Select(td => td.InnerText).ToArray());

        foreach(DataRow dr in table.Rows)
        {
            Console.WriteLine(dr[0] + "," + dr[1] + "," + dr[2] + "," + dr[8]);
        }
    }

    public static void getCellRowCol(string index1)
    {
        string bef = string.Empty;//前面字母
        string after = string.Empty;//后面数字

        foreach (char c in index1)
        {
            int ee;
            bool isT = int.TryParse(c.ToString(), out ee);
            if (isT)
            {
                after += c;
            }
            else
            {
                bef += c;
            }
        }
        Console.WriteLine(after);
    }

    public static String ConvertToChinese(Decimal number)
    {
        var s = number.ToString("#L#E#D#C#K#E#D#C#J#E#D#C#I#E#D#C#H#E#D#C#G#E#D#C#F#E#D#C#.0B0A");
        var d = Regex.Replace(s, @"((?<=-|^)[^1-9]*)|((?'z'0)[0A-E]*((?=[1-9])|(?'-z'(?=[F-L\.]|$))))|((?'b'[F-L])(?'z'0)[0A-L]*((?=[1-9])|(?'-z'(?=[\.]|$))))", "${b}${z}");
        var r = Regex.Replace(d, ".", m => "负元空零壹贰叁肆伍陆柒捌玖空空空空空空空分角拾佰仟万亿兆京垓秭穰"[m.Value[0] - '-'].ToString());
        return r;
    }

    public static string MoneySmallToBig(string par)
    {
        String[] Scale = { "分", "角", "元", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿", "拾", "佰", "仟", "兆", "拾", "佰", "仟" };
        String[] Base = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
        String Temp = par;
        string result = null;
        int index = Temp.IndexOf(".", 0, Temp.Length);//判断是否有小数点
        if (index != -1)
        {
            Temp = Temp.Remove(Temp.IndexOf("."), 1);
            for (int i = Temp.Length; i > 0; i--)
            {
                int Data = Convert.ToInt16(Temp[Temp.Length - i]);
                result += Base[Data - 48];
                result += Scale[i - 1];
            }
        }
        else
        {
            for (int i = Temp.Length; i > 0; i--)
            {
                int Data = Convert.ToInt16(Temp[Temp.Length - i]);
                result += Base[Data - 48];
                result += Scale[i + 1];
            }
        }
        return result;
    }

   

    private static void fun10()
    {
        List<string> list = new List<string>();
        int i = list.Count;
        //string a = list[0];
        list.Add("111");
        int index = 1;

        string a = "2123123";
        int df = a.Length;

        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add("Column-4", System.Type.GetType("System.String"));
        dtTemp.Columns.Add("账户标识", System.Type.GetType("System.String"));
        //dtTemp.Columns.Add("cardtype", System.Type.GetType("System.String"));
        //dtTemp.Columns.Add("idcard", System.Type.GetType("System.String"));
        //dtTemp.Columns.Add("state", System.Type.GetType("System.String"));
        dtTemp.Rows.Add("81005167152", "杨燕");
        dtTemp.Rows.Add("81005167153", "杨燕");
       
        //mydt.Rows.Add("81005167152", "杨燕", "身份证", "3205241963110121");
        //dtTemp.Rows.Add("81005167152", "杨燕", "护照", "3205241963110121");

        DataTable dtAll = dtTemp;
        //dtAll.Columns["Column-4"].ColumnName = "Column4";

        foreach (DataRow dr in dtTemp.Rows)
        {
          
            if (dr["Column-4"].ToString() != "")
            {
                //取该行数据导入至dtAll
                DataRow[] drs = dtAll.Select("`Column-4` = '81005167152' and 账户标识='杨燕'");
                //string aaa = drs[0]["管理机构"].ToString();
                if (drs.Length == 1)
                {
                    int rowindex = dtAll.Rows.IndexOf(drs[0]);
                    dtAll.Rows[rowindex]["管理机构"] = dr["管理机构"].ToString();
                    dtAll.Rows[rowindex]["账户标识"] = dr["账户标识"].ToString();
                    //dtAll.Rows[rowindex]["开立日期"] = dr["开立日期"].ToString();
                    //dtAll.Rows[rowindex]["到期日期"] = dr["到期日期"].ToString();
                    //dtAll.Rows[rowindex]["借款金额"] = dr["借款金额"].ToString();
                    //dtAll.Rows[rowindex]["账户币种"] = dr["账户币种"].ToString();
                }
                else
                {
                    DataRow dr1 = dtAll.NewRow();
                    //dr1["REPORT_NO"] = "11";
                    //dr1["CREATE_TIME"] = DateTime.Now.ToString();
                    //dr1["Type"] = "账户" + (index + 1).ToString();

                    dr1["管理机构"] = dr["管理机构"].ToString();
                    dr1["账户标识"] = dr["账户标识"].ToString();
                    //dr1["开立日期"] = dr["开立日期"].ToString();
                    //dr1["到期日期"] = dr["到期日期"].ToString();
                    //dr1["借款金额"] = dr["借款金额"].ToString();
                    //dr1["账户币种"] = dr["账户币种"].ToString();
                    dtAll.Rows.Add(dr1);
                }
                break;
            }
        }
    }

    public static void parseStringToCSV()
    {
        string file = @"C:\Users\Lenovo\Desktop\test\test\result.txt";
        DataTable dt = OpenCSV(file, new string[] { " " });
        foreach (DataRow row in dt.Rows)
        {
            Console.WriteLine("={0}-{1}-{2}=", row["code"], row["number1"], row["number2"]);
        }
    }
    public class User
    {
        public int age { get; set; }
        public string name { get; set; }
    }
    public static void clipboardOp()
    {
        User userIn = new User();
        userIn.name = "Jack";
        userIn.age = 18;
        Clipboard.SetData("mydata", "Jack");
        string userOut = (String)Clipboard.GetData("mydata");
        Console.WriteLine(userOut);
    }
    public static void readLargeFile()
    {

        string filePath = @"D:\数字员工workspace\读取日志\logs\EzDataAccess.log";
        int bufferSize = 1024000; //每次读取的字节数
        byte[] buffer = new byte[bufferSize];
        System.IO.FileStream stream = null;
        try
        {
            stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open);
            long fileLength = stream.Length;//文件流的长度
            Console.WriteLine("fileLength:" + fileLength);
            int readCount = (int)Math.Ceiling((double)(bufferSize / fileLength)); //需要对文件读取的次数
            int tempCount = 0;//当前已经读取的次数

            do
            {
                stream.Read(buffer, tempCount * bufferSize, bufferSize); //分readCount次读取这个文件流，每次从上次读取的结束位置开始读取bufferSize个字节
                                                                         //这里加入接收和处理数据的逻辑
                Console.WriteLine(buffer.ToString());
                //
            }
            while (tempCount < readCount);
        }
        catch
        {

        }
        finally
        {
            if (stream != null)
                stream.Dispose();
        }
    }

    public static void readLines()
    {
        
        EncodingInfo[] ecds =System.Text.Encoding.GetEncodings();
        foreach(EncodingInfo enc in ecds)
        {
            Console.WriteLine(enc.Name);
        }

        string _filePath = @"D:\数字员工workspace\读取日志\logs\EzDataAccess.log";

        int _currentLine = 0;
        System.Text.Encoding _encoding = Encoding.GetEncoding("utf-8");
        var lines = File.ReadLines(_filePath, _encoding).Take(1);
        var iter = lines.GetEnumerator();
        int size = lines.Count<string>();
        throw (new Exception("Error"));
        while (iter.MoveNext())
        {
            string dataStr = iter.Current;
            Console.WriteLine(dataStr);
        }
    }

    public static string RegFetchString(string 项目名称)
    {
        //string 项目名称 = "关于变更国联玉如意年年发10号集合资产管理计划资产管理合同的公告";
        Regex nameRxg = new Regex(@"国联.*计划");
        MatchCollection matchedGroup = nameRxg.Matches(项目名称);
        int groupCount = matchedGroup[0].Groups.Count;
        string projectName = "";
        if (matchedGroup.Count > 0 && groupCount >= 1)
        {
            projectName = matchedGroup[0].Groups[groupCount - 1].Value.Trim();
            Console.WriteLine(projectName);
            // Console.WriteLine(拼音字头(projectName));
        }
        else
        {
            Console.WriteLine("No match price part1");
        }
        return projectName;
    }

    public static void filterFilesNew()
    {
        string prodDir = @"D:\RPAwork\资产管理业务综合报送平台\产品pdf\";
        string[] filesList = new String[] {
           prodDir + "国联定新2号定向资产管理计划2020年第2季度管理报告.pdf",
           prodDir + "国联汇鑫26号单一资产管理计划2020年第2季度管理报告.pdf",
           prodDir + "国联汇鑫26号单一资产管理计划2020年第2季度托管报告.pdf",
           prodDir + "国联玉如意6号集合资产管理计划2020年第2季度管理报告.pdf",
           prodDir + "国联玉如意6号集合资产管理计划2020年第2季度托管报告.pdf"
      };
        Dictionary<String, List<String>> filesMapper = new Dictionary<String, List<String>> { };
        int i = 0;

        foreach (string item in filesList)
        {
            string fileName = Path.GetFileNameWithoutExtension(item);
            fileName = RegFetchString(fileName);
            if (filesMapper.ContainsKey(fileName))
            {
                filesMapper[fileName].Add(item);
            }
            else
            {
                filesMapper.Add(fileName, new List<String> { item });
            }
            i++;
        }
        Console.WriteLine(filesMapper.Keys.Count);

        foreach (KeyValuePair<string, List<String>> entry in filesMapper)
        {
            bool isSeason = false;
            string 信息报告名称;
            string 报告类型;
            //string docFile = entry.Value.FirstOrDefault(it => Path.GetExtension(it).Contains(".doc"));
            string[] pdfFiles = entry.Value.ToArray();
            string prodName = entry.Key;
            string firstPdf = pdfFiles[0];
            if(prodName.Contains("定向") || prodName.Contains("单一"))
            {
                报告类型 = "单一定期报告";
            }
            else
            {
                报告类型 = "集合定期报告";
            }
            // 获取年开始
            Regex nameRxg = new Regex(@"(\d{4})年");
            MatchCollection matchedGroup = nameRxg.Matches(firstPdf);
            string year = "";
            string season = "";
            if (matchedGroup.Count > 0)
            {
                int groupCount = matchedGroup[0].Groups.Count;

                if (groupCount >= 1)
                {
                    year = matchedGroup[0].Groups[groupCount - 1].Value.Trim();
                    Console.WriteLine(year);;
                }
            }
            //获取年结束
            信息报告名称 = prodName + year + "年";
            if (firstPdf.Contains("季度"))
            {
                isSeason = true;
                // 季度需要
                season = firstPdf.Split(new String[] { "第", "季度" }, StringSplitOptions.RemoveEmptyEntries)[1];
                信息报告名称 = 信息报告名称 + "第" + season + "季度";
            }
            if (pdfFiles.Length > 1)
            {
                信息报告名称 = 信息报告名称 + "管理报告、托管报告";
            }
            else
            {
                信息报告名称 = 信息报告名称 + "管理报告";
            }
            Console.WriteLine("信息报告名称:" + 信息报告名称);
            Console.WriteLine("year:" + year);
            Console.WriteLine("season:" + season);
            //Console.WriteLine(docFile);
            //bool hasDoc = entry.Value.Exists(it => Path.GetExtension(it).Contains(".doc"));
            //if (hasDoc)
            //{
            //    开放时间 = "get from doc";
            //}
            //else
            //{
            //    开放时间 = "NOW";
            //}
            //Console.WriteLine(entry.Key);
            //Console.WriteLine(开放时间);
        }

    }

    public static void filterFiles()
    {
        string[] filesList = new String[] {
          @"D:\数字员工\RpaWork\无锡国联\关于变更国联玉如意年年发10号集合资产管理计划资产管理合同的公告.docx",
          @"D:\数字员工\RpaWork\无锡国联\关于变更国联玉如意年年发10号集合资产管理计划资产管理合同的公告.pdf",
          @"D:\数字员工\RpaWork\无锡国联\关于变更国联玉如意年年发10号集合资产管理计划资产管理合同的征询意见函（致投资者）.pdf",
          @"D:\数字员工\RpaWork\无锡国联\关于变更国联玉如意年年发10号集合资产管理计划资产管理合同的征询意见函（致托管人）.pdf",
          @"D:\数字员工\RpaWork\无锡国联\关于变更国联玉如意年年发11号集合资产管理计划资产管理合同的公告.pdf",
          @"D:\数字员工\RpaWork\无锡国联\关于变更国联玉如意年年发12号集合资产管理计划资产管理合同的公告.docx"
      };
        Dictionary<String, List<String>> filesMapper = new Dictionary<String, List<String>> { };
        int i = 0;
        
        foreach (string item in filesList) {
            string fileName = Path.GetFileNameWithoutExtension(item);
            fileName = RegFetchString(fileName);
          if (filesMapper.ContainsKey(fileName))
            {
                filesMapper[fileName].Add(item);
            }
            else
            {
                filesMapper.Add(fileName, new List<String> { item});
            }
            i++;
        }
        Console.WriteLine(filesMapper.Keys.Count);

        foreach(KeyValuePair<string, List<String>> entry in filesMapper)
        {

            string 开放时间 = string.Empty;
            string docFile = entry.Value.FirstOrDefault(it => Path.GetExtension(it).Contains(".doc"));
            string[] pdfFiles = entry.Value.Where(it => Path.GetExtension(it).Contains(".pdfs")).ToArray();
            foreach (string item in pdfFiles)
            {
                Console.WriteLine("PDF:" + item);
            }
            Console.WriteLine(docFile);
            bool hasDoc = entry.Value.Exists(it => Path.GetExtension(it).Contains(".doc"));
            if (hasDoc)
            {
                开放时间 = "get from doc";
            }
            else
            {
                开放时间 = "NOW";
            }
            Console.WriteLine(entry.Key);
            Console.WriteLine(开放时间);
        }

    }
    public static string fetchDate(string dateStr)
    {
        Regex nameRxg = new Regex(@"\d{4}年\d{1,2}月\d{1,2}日");
        MatchCollection matchedGroup = nameRxg.Matches(dateStr);
        string projectName = "";
        if (matchedGroup.Count > 0)
        {
            int groupCount = matchedGroup[0].Groups.Count;

            if(groupCount >= 1)
        {
                projectName = matchedGroup[0].Groups[groupCount - 1].Value.Trim();
                Console.WriteLine(projectName);
                // Console.WriteLine(拼音字头(projectName));
            }
        }
        else
        {
            Console.WriteLine("No match price part1");
        }
        return projectName;
    }
    public static void ParseComboxJson()
    {
        //Console.WriteLine(Convert.ToDateTime("2020-11-13", "yyy-mm-dd"));
        Dictionary<string, string> resultDi = new Dictionary<string, string>();

        string file = @"C:\Users\Lenovo\Desktop\text.txt";
        string result = File.ReadAllText(file);

        JObject jObject = JObject.Parse(result);

        JArray jItems = (JArray)jObject["Items"][0]["Items"][2]["Items"][0]["Items"][4]["Items"][0]["Items"];

        List<JToken> items = jItems.Children().ToList();
        foreach(JToken item in items)
        {
            JObject joitem = (JObject)item;
            if (joitem["ControlType"].ToString() == "AWT component")
            {
                string[] valueArr = joitem["Value"].ToString().Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);
                string lableValue = valueArr[valueArr.Length - 1];
                Console.WriteLine("lableValue-----" + lableValue);  // lableValue

                JArray lableItems = (JArray)joitem["Items"];
                List<JToken> labekItems = lableItems.Children().ToList();
                foreach (JToken labelItem in labekItems)
                {
                    JObject jolabelItem = (JObject)labelItem;
                    string controlTypeValue = jolabelItem["ControlType"].ToString();
                    if (controlTypeValue  == "text" || controlTypeValue == "combo box")
                    {
                        string finalValue = jolabelItem["Value"].ToString();
                        Console.WriteLine("Final Value-----" + finalValue);  // finalValue
                        resultDi.Add(lableValue, finalValue);
                    }
                }
            }
        }
        Console.WriteLine(resultDi.Values);

    }

    public static void getStrWordWrap()
    {
        string thestr = "this is ad sds sdfsdf qweqwe sdfsdf wdsdasd sdasda asdasdas";
        int position = 6;
        bool flag = false;

        while (!flag)
        {
            string item = thestr.Substring(position - 1, 1);
            byte itemByte = ASCIIEncoding.ASCII.GetBytes(item)[0];
            if (itemByte == 32)
            {
                flag = true;
            }
            else
            {
                position -= 1;
            }
        }
        Console.WriteLine(thestr.Substring(0, position - 1).Length);
        Console.WriteLine(thestr.Substring(0, position - 1));
    }

    public static string convertImg(string fileName, string filter)
    {
        Console.WriteLine(fileName);

        string fileBaseName = Path.GetFileNameWithoutExtension(fileName);
        string fileExtension = Path.GetExtension(fileName);

        string dir = @"C:\Users\Lenovo\Pictures\图片处理结果";
        string dstfilePath = Path.Combine(dir, string.Format("{0}.{1}", fileBaseName, filter.ToLower()));
        switch (filter.ToUpper())
        {
            case "JPG":
                if (fileExtension == ".ico")
                {
                    Image img = Image.FromFile(fileName);
                    img.Save(dstfilePath, ImageFormat.Jpeg);

                   
                }
                else
                {
                    using (MagickImage image = new MagickImage(fileName))
                    {
                        //将透明色更改成白色(这里不指定默认是黑色)
                        image.Opaque(MagickColors.Transparent, MagickColors.White);
                        // Save frame as jpg
                        image.Write(dstfilePath);
                    }
                }
               
                break;
            case "PNG":
                if (fileExtension == ".ico")
                {
                    Image img = Image.FromFile(fileName);
                    img.Save(dstfilePath, ImageFormat.Png);
                }
                else
                {
                    using (MagickImage image = new MagickImage(fileName))
                    {
                        // Save frame as jpg
                        image.Write(dstfilePath);
                    }
                }

                break;
            //case "GIF":
            //    using (MagickImage image = new MagickImage(fileName))
            //    {
            //        // Save frame as jpg
            //        image.Write(dstfilePath);
            //    }
            //    break;
            case "BMP":
                if (fileExtension == ".ico")
                {
                    Image img = Image.FromFile(fileName);
                    img.Save(dstfilePath, ImageFormat.Bmp);
                }
                else
                {
                    using (MagickImage image = new MagickImage(fileName))
                    {
                        // Save frame as jpg
                        image.Write(dstfilePath);
                    }
                }

                break;
            //case "ICO":
            //    using (MagickImage image = new MagickImage(fileName))
            //    {
            //        // Save frame as jpg
            //        image.Write(dstfilePath);
            //    }
            //    break;
            default: break;
        }

        return dstfilePath;

}

    public static string getEleValue(string result, string PONumberKeyWord)
    {
        int PONumIndex = result.IndexOf(PONumberKeyWord);
        int PONumStrEndIndex = result.IndexOf(";", PONumIndex);
        if(PONumStrEndIndex == -1)
        {
            PONumStrEndIndex = result.Length;

        }
        int POEndIndex = PONumIndex + PONumberKeyWord.Length;
        string PONum = result.Substring(POEndIndex, PONumStrEndIndex - POEndIndex).Trim();
        return PONum;
    }

    public static string RemoveInvalidCharForFile(string fileName)
    {
        // 非法字符有：:<>:?*"|\/
        //string fileName = "文件?名称<";
        //StringBuilder builder = new StringBuilder(fileName); //需要操作很长的字符串，或者要对字符串进行非常频繁的操作
     
        foreach (char invalidChar in Path.GetInvalidFileNameChars())
        {
            Console.WriteLine(invalidChar);
            fileName = fileName.Replace(invalidChar.ToString(), string.Empty);
            Console.WriteLine("fileName" + fileName);
            //builder.Replace(invalidChar.ToString(), string.Empty);
        }
        return fileName;
    }

    static void Main()
    {
        //Console.WriteLine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
        // Console.WriteLine(System.Guid.NewGuid().ToString("N"));
        //Console.WriteLine("bcd58243-dd2c-40f7-a11b-3ae94245c484".Length);
        // SortFiles();
        //TxtParser.getEmail();

        //Console.WriteLine(ProcessHelper.ProcessExist("Encoo.Android.Automation"));

        
        
 
        string 乙方 = "850元建议明昭商贸有限公司进行维修。";
        Console.WriteLine(乙方.Substring(0, 4));
        Regex 建议正则 = new Regex(@"建议由*([\u4e00-\u9fa5]+)");
        Match matchResult = 建议正则.Match(乙方);
        string 乙方公司 = matchResult.Value;
        Console.WriteLine(乙方公司.Split(new String[] { "建议由", "建议"}, StringSplitOptions.RemoveEmptyEntries)[0]);

        // Console.WriteLine(("\"文件 ?名称<").Replace("\"", string.Empty));
        Console.WriteLine(RemoveInvalidCharForFile("\"文件 ?名称<"));


        string 系统流水号1 = "LX-R313-2021030007_CZ_001";

        Console.WriteLine(系统流水号1.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries)[0]);

        double dbdata = 0.5009;
        string str1 = String.Format("{0:F}", dbdata);//默认为保留两位


        Console.WriteLine(Math.Round(0.50009, 2));
        string totalAmount = "999951.14";
        string amount2 = "182628.95";
        double 汇率 = Convert.ToDouble("6.475400");

        double 本币1 = Convert.ToDouble(totalAmount)/ 汇率;
        double 本币2 = Convert.ToDouble(amount2)/ 汇率;

        string asd = "￥999951.14";
        


        Console.WriteLine(asd.Substring(1));
        Console.WriteLine(本币2);
        Console.WriteLine(本币1 + 本币2);

        Regex 成交总价正则 = new Regex(@"成交总价：(\d+(\.\d+)?)");
        string 外币备注 = "外币汇率：6.4623;贸易方式：一般贸易;币种：USD;合同协议号：ZHLX&ATC21-14;成交方式：FOV;成交总价：53639.56";
        Console.WriteLine(成交总价正则.Match(外币备注).Value);
        Console.WriteLine(外币备注.Replace(成交总价正则.Match(外币备注).Value, "成交总价：" + String.Format("{0:F}", 本币2)));


        //备注.Split(new String[] { "外币汇率：", ";", "" }, StringSplitOptions.None);



        Regex nameRxg11 = new Regex(@"([\u4e00-\u9fa5]{2})铭牌");
        String 铭牌全名称 = "14183-00078:A:货物铭牌,TIFU:专用";

        Console.WriteLine(nameRxg11.Match(铭牌全名称).Value);

        string 系统流水号 = "YXLX-R310-2021030029_CZ_1";

        string 单号 = 系统流水号.Split(new String[] { "-", "_" }, StringSplitOptions.RemoveEmptyEntries)[2];

        Console.WriteLine(单号.Substring(单号.Length - 2, 2));

        String[] arr1 = (new List<string> { }).ToArray();


        string text = string.Join("\n", arr1); //\r\n表示换一行
        Console.WriteLine(text);
        //sendMail();
        //ImageToBase64.ConvertToBase64(@"C:\Users\Lenovo\Pictures\各种类型图片\flower.jpg");

        // TxtParser.mciPdf();

        //Console.WriteLine(ExcelConvert.ToIndex("BB"));
        //Console.WriteLine(ExcelConvert.ToName(ExcelConvert.ToIndex("BB")))
        //获取当前月前一月的月末;

        DateCasesHelper.preMonthLastDay();

        string 备注 = "外币汇率：6.4623;贸易方式：一般贸易;币种：USD;合同协议号：ZHLX&ATC21-14;成交方式：FOV;成交总价：53639.56";

        //备注.Split(new String[] { "外币汇率：", ";", "" }, StringSplitOptions.None);


        string 外币汇率索引 = getEleValue(备注, "成交总价：");


        Console.WriteLine(外币汇率索引);

        DateTime date2;
        bool isValidDate = DateTime.TryParse("2021/2/5 0:00:00", out date2);
        Console.WriteLine(date2.ToString());

        List<string[]> list2 = new List<string[]>() {new string[]{ "a" }, new string[]{ "b" } };
        Console.WriteLine(list2[0][0]);
        string dateStr = "27/01/2021 14:31:00";
        //string dateStr = "27/01/2021 14:31:00";
        Console.WriteLine(DateTime.Parse("2021/1/27 4:32:43"));
        IFormatProvider ifp = new CultureInfo("en-us", true);
        DateTime resultDateTime;
        Console.WriteLine(DateTime.ParseExact(dateStr, "dd/MM/yyyy HH:mm:ss", null));
        bool parsedCorrectly = DateTime.TryParseExact(dateStr, "dd/MM/yyyy HH:mm:ss", null, DateTimeStyles.None, out resultDateTime);
        Console.WriteLine(resultDateTime);
        Console.WriteLine(parsedCorrectly);
        Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

        //int month = DateTime.Now.Month;
        //Console.WriteLine(month);
        //Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd"));

        //TxtParser.FlexMNPo();
        Console.WriteLine(TxtParser.Asc("A"));
        
        Console.WriteLine(TxtParser.Chr(66));
        //TxtParser.FlexMNPoV2();

        Double avgQty = double.MinValue;

        if (avgQty == double.MinValue) {
            Console.WriteLine(avgQty + "--------");
        }

        DateTime out_dt;
        DateTime.TryParse("08-FEB-2021", out out_dt);
        Console.WriteLine(out_dt);
       // 
        //TxtParser.FlexRMAPo();

        //string resultFile = convertImg(@"C:\Users\Lenovo\Pictures\各种类型图片\delete_hover.png", "png");
        //string resultFile1 = convertImg(@"C:\Users\Lenovo\Pictures\各种类型图片\delete_hover.png", "bmp");
        string resultFile2 = convertImg(@"C:\Users\Lenovo\Pictures\各种类型图片\encoo2.ico", "png");



        string[] columns = new string[]{
"组别标识", "行号", "库区编码(订单行)", "批次", "计划数量", "物料编码", "来源行号",
"超量接收允差百分比", "单据类型", "库存组织编码", "仓库编码", "货主编码", "库区编码(订单头)",
"来源仓库编码", "来源货主编码", "来源库区编码", "供应商", "客户", "联系人", "手机号", "座机号",
"电子邮件", "传真", "邮编", "国家","省", "城市", "地区(乡镇)", "入库详细地址", "重量", "总重量",
"体积","来源编号", "来源单据类型", "备注说明", "行备注", "采购凭证号", "合同号", "合同名称",
"预算主体", "采购部门", "业务员", "移动类型", "计划项目", "采购员", "主项", "头attribute11",
"头attribute12", "头attribute13", "头attribute14", "头attribute15", "wbs编号", "请购人","请购部门"
};
        int colNums = columns.Length;
        string[] rowsData = new string[colNums] ;



        try
        {
            Object obj = null;
            Console.WriteLine(obj.ToString());
        }
        catch(Exception e)
        {
            Console.WriteLine(e);
        }

       // getStrWordWrap();
       //DateTime outDt;
       //DateTime.TryParse(fetchDate("司 2020年9月7日"), out outDt);

        //Console.WriteLine(outDt.ToString("yyyy-MM-dd"));
        // ParseComboxJson();


        //ParseCurrencyJson();
        //filterFilesNew();

        //readLines();
        string gapStr = String.Empty;
        for (int i = 0; i < 1; i++)
        {
            gapStr = gapStr + "0";
        }

        string code = "000735";
        int codeLength = code.Length;
        int nextCodeInt = Convert.ToInt32(code) + 1;
        int nextCodeLength = nextCodeInt.ToString().Length;

        StringBuilder sb = new StringBuilder();
        
        for (int i = 0; i < (codeLength - nextCodeLength); i++)
        {
            sb.Append("0");
        }
        sb.Append(nextCodeInt.ToString());
        Console.WriteLine(sb);

        //Dictionary<string, Dictionary<string, string>> dic1 = new Dictionarycolumns<string, Dictionary<string, string>>();
        //Dictionary<string, string> tmpDic1 = new Dictionary<string, string>();
        //tmpDic1.Add("值", "1");
        //tmpDic1.Add("缩写", "zh");
        //dic1.Add("口径", tmpDic1);
        //tmpDic1 = new Dictionary<string, string>();
        //tmpDic1.Add("值", "2");
        //if (tmpDic1.ContainsKey("sa")) Console.WriteLine(tmpDic1["sa"]);

        //DateTime dt1;
        //bool isDate1 = DateTime.TryParse("2020/5/18", out dt1);

        //Console.WriteLine(dt1.ToString("yyyy-MM-dd"));

        //string[] apo = new string[] { "a", "b" };
        //Console.WriteLine(string.Join("", apo));
        ////parseStringToCSV();

        //Console.WriteLine(DateTime.Now.ToString("yyyyMMdd"));

        //excelNumToColName(4);
        string fileName1 = "国联金如意6号第361期产品开放销售公告（新手专享）";
        if (fileName1.Contains("新手专享"))
        {
            Console.WriteLine("---");
        }

        //Regex nameRxg = new Regex(@"((\d{4})(年度|年))?(.{6})");
        //String 项目名称 = "内控非现场专项审计";

        Regex nameRxg = new Regex(@"第(\d+)期");
        
        String 项目名称 = "国联金如意6号第361期产品开放销售公告（新手专享）";
        MatchCollection matchedGroup = nameRxg.Matches(项目名称);
        int groupCount = matchedGroup[0].Groups.Count;
        if (matchedGroup.Count > 0 && groupCount > 1)
        {
            string projectName = matchedGroup[0].Groups[groupCount-1].Value.Trim();
            Console.WriteLine(projectName);
            // Console.WriteLine(拼音字头(projectName));
        }
        else
        {
            Console.WriteLine("No match price part1");
        }
        // fun10();

        Console.WriteLine("a" + 12);
        //Regex nameRxg1 = new Regex(@"第(\d+)期");
        string 产品名称1 = "“乐惠”2020年第33期";

        string tt = "(abc)";
        if(tt.Contains("("))
        {
            tt = tt.Replace("(", "（");
            //tt = tt.Replace("b", "B");

        };

        Console.WriteLine(tt);

        string[] 产品名称数组1 = 产品名称1.Split(new string[] { "第", "期" }, StringSplitOptions.RemoveEmptyEntries);


        Console.WriteLine(((3.90+3.95)/2).ToString("0.0000"));

        
        Console.WriteLine(string.Format("{0:N}", 5000 * 10000));
        //DateTime dt2;
        //bool isDate1 = DateTime.TryParse("2020年5月18日", out dt2);
        //Console.WriteLine(dt2.ToString("yyyyMMdd"));
        Regex nameRxg3 = new Regex(@"(\d+)");
        MatchCollection matchedGroup3 = nameRxg3.Matches("按5w和50w金额分档，参照老说明书");
        if (matchedGroup3.Count > 1 && matchedGroup3[1].Groups.Count > 1)
        {
            Double 预计募集金额 = Convert.ToDouble(matchedGroup3[1].Value.Trim()); // 预计募集金额, 0.5 => 5000，以亿为单位
            Console.WriteLine(预计募集金额.ToString("0.00"));
        }


        string 临时预计募集金额 = "3.90%-3.95%";
        Double result;
        bool ifDouble = Double.TryParse(临时预计募集金额, out result);
        Console.WriteLine(result);

        //Regex nameRxg3 = new Regex(@"(\d+.?\d*)");
        //MatchCollection matchedGroup3 = nameRxg3.Matches(临时预计募集金额);
        //if (matchedGroup3.Count > 0 && matchedGroup3[0].Groups.Count > 1)
        //{
        //    Double 预计募集金额 = Convert.ToDouble(matchedGroup3[0].Value.Trim()) * 10000; // 预计募集金额, 0.5 => 5000，以亿为单位
        //    Console.WriteLine(预计募集金额.ToString("0.00"));
        //}

        if (true) Console.WriteLine("111");
        getCellRowCol("3000W起成立");

        //Regex nameRxg1 = new Regex(@"(（.+）)");
        
        //MatchCollection matchedGroup1 = nameRxg1.Matches(产品全名称);
        //if (matchedGroup1.Count > 0 && matchedGroup1[0].Groups.Count > 1)
        //{
        //    string projectName = matchedGroup1[0].Value.Trim();
        //    Console.WriteLine(projectName);
        //}
        //else
        //{
        //    Console.WriteLine("No match price part1");
        //}

        Console.WriteLine("BWB".ToLower());
        string hoemfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Console.WriteLine(hoemfolder);

       

        DateTime dt = DateTime.Now;
        Console.WriteLine(dt.AddMonths(-1).AddDays(1 - dt.Day).AddMonths(1).AddDays(-1));
        String[] strArr2 = new String[] { "a", "b", "c", "d" };
        String[] strArr3 = new String[] { "e", "f" };

        List<String> strArr4 = new List<String> { };
        strArr4 = strArr4.Concat(strArr2).ToList();
        strArr4 = strArr4.Concat(strArr3).ToList();

        htmltableToDataTable();

       

        Dictionary<string, string> thisMonthMapper = new Dictionary<string, string> {
        {"1", "Jan"}, {"2", "Feb"}, {"3", "Mar"},
        {"4", "Apr"}, {"5", "May"}, {"6", "June"},
        {"7", "July"}, {"8", "Aug"}, {"9", "Sep"},
        {"10", "Oct"}, {"11", "Nov"},{"12", "Dec"}
    };
        
        thisMonthMapper.Keys.ToList().ElementAt(0);
        //ParseInvoiceVerificationJson();


        Regex reg = new Regex(@"[\u4e00-\u9fa5]");//正则表达式
        string str = "证券投资合计:";
        if (reg.IsMatch(str))
        {
            Console.WriteLine("有汉字答");
        }
        else
        {
            Console.WriteLine("没汉字");
        }


        Double dlvalue;
        bool isDouble;
        isDouble = Double.TryParse("102,232,122.232", out dlvalue);

        if (isDouble)
        {
            Console.WriteLine("Is Double" + dlvalue);
        }
        else
        {
            Console.WriteLine("Is not Double");
        }

        isDouble = Double.TryParse("ss", out dlvalue);

        if (isDouble)
        {
            Console.WriteLine("Is Double" + dlvalue);
        }
        else
        {
            Console.WriteLine("Is not Double");
        }

        string a = "asd";
        if("asd".Equals(a))
        {
            Console.WriteLine("1111");
        }
        //DateTime dt;
        //bool isDate = DateTime.TryParse("2020年5月18日", out dt);
        //Console.WriteLine(dt.AddMonths(-1).ToString("yyyy-MM"));

        string testStr = "杭州联合银行乐惠系列2020年第21期封闭净值型人民币理财产品20200424估值报表";
        //string testStr = "杭州联合银行乐惠联盈1号2019年第3期封闭净值型人民币理财产品20200424估值报表";
        //string testStr = "杭州联合银行乐惠系列-CT第1期封闭净值型人民币理财产品20200424估值报表";
        //string testStr = "“乐惠-双季赢”第1期估值表_20200511";
        string[] strArr = testStr.Split(new string[] { "杭州联合银行", "封闭净值型人民币理财产品", "估值报表", "估值表", "_" }, StringSplitOptions.RemoveEmptyEntries);

        string 产品名 = strArr[0];
        string yearReg = @"[\d{4}年]*第";
        string[] resultArr = Regex.Split(产品名, yearReg);
        
        Console.WriteLine(resultArr[0] + "_" + resultArr[1]);

        string 期号 = "21";
        Console.WriteLine(期号.Split("期")[0]);


        foreach (string item in strArr)
        {
            Console.WriteLine(item);
            DateTime tmpDate;
            bool isDate = DateTime.TryParseExact(item, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture,
                                   System.Globalization.DateTimeStyles.None, out tmpDate);
            if (isDate)
            {
                Console.WriteLine(tmpDate.ToString("yyyy-MM-dd"));
            }
        }
            //DataTable dst = OpenCSV(@"‪D:\Bottime\0512-readcsv-datatable\test.csv", new char { '   ' });
            //foreach (DataRow row in dst.Rows)
            //{
            //    Console.WriteLine("{0}, {1}, {2}", row[0], row[1], row[2]);
            //}

            //List<String> asd = new String[] { null, "122132", "" }.Where(s => !String.IsNullOrEmpty(s)).ToList();
            //String lll = string.Join(", ", asd);
            //Console.WriteLine(lll);
            // sendMail();
            //parsePrice();
            //parseJson2();
            //GetCharSpellCode("");
            //Console.WriteLine(拼音字头("梅晓瑜"));

            //Console.WriteLine(System.IO.Path.Combine(@"D:\联合银行\解放路支行发票查验", "发票"));
            //String[] strArr2 = new String[] { "a", "b", "c", "d" };

            //foreach (string item in strArr2.TakeLast(3))
            //{
            //    Console.WriteLine("=========" + item);
            //}

            //DateTime dt;
            //bool isDate = DateTime.TryParse("2019年10月30日", out dt);
            //Console.WriteLine(dt.ToString("yyyy-MM-dd"));
            Int32 totalImages = 1110;
        Console.WriteLine(totalImages / 100 + 1);


        //DateTime dt;
        //bool isDate = DateTime.TryParse("2020/5/2 0:00:00", out dt);

        //Console.WriteLine(dt.ToString("yyyy年MM月dd日"));

       //Regex nameRxg = new Regex(@"(\d{4})(年度|年)(.{6})");
       // String 项目名称 = "内控非现场专项审计";
       // MatchCollection matchedGroup = nameRxg.Matches(项目名称);
       // if (matchedGroup.Count > 0 && matchedGroup[0].Groups.Count > 1)
       // {
       //     string projectName = matchedGroup[0].Groups[3].Value.Trim();
       //     Console.WriteLine(拼音字头(projectName));
       // }
       // else
       // {
       //     Console.WriteLine("No match price part1");
       // }

        String illegal = "贷款资 : * ?金用于承兑汇票保证金 \"< > |";
        string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

        foreach (char c in invalid)
        {
            illegal = illegal.Replace(c.ToString(), "");
        }
        Console.WriteLine(illegal);
        String 支行反馈意见 = "责任人（签名）";
        Console.WriteLine(Convert.ToString(支行反馈意见).Split(new string[] { "责任人（签名）" }, StringSplitOptions.None)[0]);
        Console.WriteLine(Convert.ToString(支行反馈意见.Split("责任人（签名）")[0]).Trim());


        Console.WriteLine(Path.GetFileName(@"D:\联合银行\内审\员工事实确认处理中\宝善支行"));
        String bank = "联合银行四季青支行";

        Console.WriteLine(bank.Replace("支行", "").Replace("联合银行", ""));

        String 序号 = "";
        Int32 原始序号 = Convert.ToInt32("1090");
        if(原始序号 < 10)
        {
            序号 = "00" + 原始序号;
        }
        else if(原始序号 >= 10 && 原始序号 < 100)
        {
            序号 = "0" + 原始序号;
        }
        else
        {
            序号 = "" + 原始序号;
        }
        Console.WriteLine(序号);
        //Console.WriteLine(DateTime.Now.Month);
        //Console.WriteLine(@"D:\文件监听用\1.txt\r\n");
        //SaveTxtFile(@"D:\文件监听用\1.txt");

        //List<Dictionary<String, String>> listChildren = new List<Dictionary<String, String>>();

        //Dictionary<String, String> child1 = new Dictionary<String, String>();
        //child1.Add("age", "4岁");
        //child1.Add("name", "wendy");
        //listChildren.Add(child1);

        //Dictionary<String, String> child2 = new Dictionary<String, String>();
        //child2.Add("age", "3岁");
        //child2.Add("name", "lily");
        //child2.Add("name", "lily");
        //listChildren.Add(child2);

        //Dictionary<String, String> child3 = new Dictionary<String, String>();
        //child3.Add("age", "5岁");
        //child3.Add("name", "john");
        //listChildren.Add(child3);

        //String[] strArr = new String[] { @"D:\文件监听用\源\rpa - 1.txt", @"D:\文件监听用\源\rpa - 2.txt" }; // 旧
        //String[] strArr2 = new String[] { @"D:\文件监听用\源\rpa - 1.txt", @"D:\文件监听用\源\rpa - 2.txt" }; // 新

        //String[] addedItems = new String[strArr2.Length];
        //String[] removedItems = new String[strArr.Length];
        //Console.WriteLine(addedItems.Where(s => !string.IsNullOrEmpty(s)).Count());
        //Int32 itemIndex = 0;
        //String itemm;

        //foreach (string item in strArr2) // 新
        //{
        //    // Boolean itemNotExist = false;
        //    Int32 i = 0;
        //    foreach(string item2 in strArr) // 旧
        //    {
        //        Console.WriteLine(item == item2);
        //        if (item == item2)
        //        {
        //            break;
        //        }
        //        Console.WriteLine(item + item2 + "------------");
        //        i = i + 1;
        //    }
        //    if (i == strArr.Length)
        //    {
        //        addedItems[itemIndex] = item;
        //    }
        //    itemIndex = itemIndex + 1;
        //}

        //Int32 itemIndex2 = 0;
        //foreach (string item in strArr) // 旧
        //{
        //    // Boolean itemNotExist = false;
        //    Int32 i = 0;
        //    foreach (string item2 in strArr2) //新
        //    {
        //        if (item == item2)
        //        {
        //            break;
        //        }
        //        i = i + 1;
        //    }
        //    if (i == strArr2.Length)
        //    {
        //        removedItems[itemIndex2] = item;
        //    }
        //    itemIndex2 = itemIndex2 + 1;
        //}
        ////foreach (string item in strArr)
        ////{
        ////    if (!strArr2.Contains(item))
        ////    {
        ////        removedItems.Append(item);
        ////    }
        ////}


        ////Console.WriteLine(string.Join("\n", strArr));
        ////String[] oldFileStrArr = string.Join("\n", strArr).Split("\n");

        ////List<string> oldFileList = strArr.ToList();
        ////List<string> currentFilesList = strArr2.ToList();

        ////IEnumerable<string> addedFiles = currentFilesList.Except(oldFileList);
        ////IEnumerable<string> removedFiles = oldFileList.Except(currentFilesList);

        //foreach (string item in addedItems)
        //{
        //    Console.WriteLine(item);
        //}
        //Console.WriteLine("=======================");
        //foreach (string item in removedItems)
        //{
        //    Console.WriteLine(item);
        //}


        //    string testStr = "最近 六个月\n平均";

        //    string bbn = "1,2000";

        //    Console.WriteLine("\"" + bbn + "\"");

        //    Console.WriteLine(testStr.Replace("\n", "").Replace(" ", ""));

        //    string df = "a2";
        //    int col;
        //    bool isInt = int.TryParse(df, out col);

        //    Console.WriteLine(col);
        //    Console.WriteLine(DateTime.Now.ToString("yyyy") + "\\" + DateTime.Now.ToString("MM") + "\\" + DateTime.Now.ToString("dd") + "\\" + DateTime.Now.ToString("HHmmssfff"));

        // 获取文件的文件夹名称
        //    string fullPath = @"D:\云扩科技\RPA\测试\test.txt";
        //    DirectoryInfo dirInfo = new DirectoryInfo(fullPath);
        //    string currentDir = dirInfo.Parent.FullName;

        //    Console.WriteLine(currentDir);

        //    string filename = System.IO.Path.GetFileName(fullPath);//文件名  “Default.aspx”
        //    string extension = System.IO.Path.GetExtension(fullPath);//扩展名 “.aspx”
        //    string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(fullPath);// 没有扩展名的文件名 “Default”

        //    string beginCell = "B22";
        //    string endCell = "D25";
        //    string range = "";
        //    Regex rgx = new Regex(@"(\d+)");
        //    MatchCollection mc = rgx.Matches(beginCell);
        //    Int32 rowNumber = 0;
        //    Int32 beginRow = 0;
        //    Int32 endRow = 0;
        //    if (mc.Count > 0)
        //    {
        //        beginRow = Convert.ToInt32(mc[0].Value.Trim());
        //    }

        //    // 获取列号
        //    rgx = new Regex(@"(\d+)");
        //    mc = rgx.Matches(endCell);

        //    if (mc.Count > 0)
        //    {
        //        endRow = Convert.ToInt32(mc[0].Value.Trim());
        //    }
        //    if ((endRow - beginRow) > 2)
        //    {
        //        range = "$" + beginRow.ToString() + ":$" + endRow.ToString();
        //    }


        //Console.WriteLine(range);
        //ParseJson();

        //ParseIdJson();
        //Object[] a = new Object[] { null, 4 };

        //foreach(Object x in a)
        //{
        //    Console.WriteLine(x);
        //}
        // string a = "$";
        //Console.WriteLine(a);

        //DataTable dst = OpenCSV(@"D:\encootech\digi4th-test\BVTWorkFlows\Datable-get-cell-value\test.csv");
        //foreach (DataRow row in dst.Rows)
        //{
        //    Console.WriteLine("{0}, {1}, {2}, {3}",row[0], row[1],row[2], row[3]);
        //}
        //;

        //ReadCsv();
        //Console.WriteLine(DateTime.Now.ToString("%yy-%M-%d"));
        //if (System.IO.Directory.Exists(@"D:\云扩科技\RPA\测?"))
        //{
        //    Console.WriteLine("1111");

        //}
        //else
        //{
        //    Console.WriteLine("2222");
        //}
        //Object[] ccobj = new object[] { "a", "b" };
        //IEnumerable<Object> ieObj = ccobj.AsEnumerable<Object>();

        //Object[] objArr = new object[] { 1, "n", null, "发我", 3 };
        //foreach (Object item in objArr)
        //{
        //    Console.WriteLine(Convert.ToString(item));
        //}

        //object obj = new
        //{
        //    Name = "小明",
        //    SayHi = "你妹的！"
        //};

        //foreach (PropertyInfo p in obj.GetType().GetProperties())
        //{
        //    Console.WriteLine("{0},{1}", p.Name, p.GetValue(obj));
        //}

        //string[] arr = { "candy", "Amanda", "1qa", "!dd" , "孟凡玲", "banana", "?sd", "@ss", "~ss", "$", "%", "^", "*", "&", "+"};
        //arr = arr.OrderBy(p => p).ToArray();
        //Console.WriteLine("Amanda".StartsWith("A"));

        //Console.WriteLine(arr);
        //tryCatchTest();

        //TimeSpan ts = new TimeSpan(0, 0, 0, 0, -1);
        //Thread.Sleep(ts);
        // DelectDir(@"D:\ExcelFile\dst?\ras-s");
        //String[] ids = new String[] { "1111", "2222", "3333", "444", "555", "666", "777" };

        //Console.WriteLine(string.Join(",", ids));

        //List<Boolean> listResult = new List<Boolean> { };

        //listResult.Add(true);
        //listResult.Add(false);
        //listResult.Add(true);

        //Console.WriteLine(listResult.Contains(false));
        //Console.WriteLine(listResult.Count == 3 && !listResult.Contains(false));

        //Char a = 'a';
        //Console.WriteLine(a.GetType());
        //Console.WriteLine(a is System.Object);


        //List<String> objIds = ids.ToList();
        //foreach (string item in objIds)
        //{
        //    Console.WriteLine(item);
        //}
        //ForEachTest();
        //string p = "パスワード：123456";
        //string password = p.Split(new string[]{ "パスワード：" }, StringSplitOptions.RemoveEmptyEntries)[0];
        //Console.WriteLine(password);
        //string a = "12000000";
        //Console.WriteLine((Convert.ToInt32(a) * 1.1).ToString("C"));
        //string b = string.Format("{0:000100.000}", 123);
        //Console.WriteLine(b);
        //Object[] ids = new object[] { "1111", "2222", "3333", "444", "555", "666", "777" };
        //List<Object> objIds = ids.ToList();
        //Random r = new Random();
        //int n = r.Next(0, objIds.Count - 1);
        //Console.WriteLine(objIds[n]);
        //objIds.RemoveAt(n);
        //Object aa = 1;
        //string c = "AB21";

        //Int32 i = 10;
        //i.ToString();

        //Regex rgx1 = new Regex(@"(\D+)");
        //string match = "";
        //MatchCollection mc1 = rgx1.Matches(c);
        //if (mc1.Count > 0)
        //{
        //    match = mc1[0].Value.Trim();
        //}

        //Console.WriteLine(match);
        //for (int i = 0; i < mc.Count; i++) //在输入字符串中找到所有匹配
        //{
        //    //results[i] = mc[i].Value; //将匹配的字符串添在字符串数组中
        //    Console.WriteLine(mc[i].Value.Trim());
        //    //matchposition[i] = mc[i].Index; //记录匹配字符的位置
        //}


        // string[] strName = System.Text.RegularExpressions.Regex.Split(a, "\w", System.Text.RegularExpressions.RegexOptions.IgnoreCase);



    }


    public static void sendMail()
    {
        string sender = "mengfanling825@163.com";

        //SmtpClient smtp = new SmtpClient
        //{
        //    Host = "smtp.163.com",
        //    Port = 25,
        //    EnableSsl = true,
        //    DeliveryMethod = SmtpDeliveryMethod.Network,
        //    UseDefaultCredentials = false,
        //    Credentials = new NetworkCredential(sender, "GFCXJEXGGREVWSCY")
        //};

        SmtpClient smtp = new SmtpClient
        {
            Host = "smtp.partner.outlook.cn",
            Port = 587,
            EnableSsl = true,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential("mengfanling@encootech.com", "Amanda_yc_2019")
        };

        MailMessage mm = new MailMessage();
        mm.From = new MailAddress(sender, "Amanda Meng");
        mm.To.Add("1172261995@qq.com");

        mm.Subject = "带有图片的邮件";

        string plainTextBody = "如果你邮件客户端不支持HTML格式，或者你切换到“普通文本”视图，将看到此内容";
        mm.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(plainTextBody, null, "text/plain"));

        ////HTML格式邮件的内容   
        string htmlBodyContent = "如果你的看到<b>这个</b>， 说明你是在以 <span style=\"color:red\">HTML</span> 格式查看邮件<br><br>";
        htmlBodyContent += "<img src=\"cid:Email_Image_Test\">";   //注意此处嵌入的图片资源   
        AlternateView htmlBody = AlternateView.CreateAlternateViewFromString(htmlBodyContent, null, "text/html");


        LinkedResource lrImage = new LinkedResource(@"D:\chrome downloads\browser-market-shares-in.png", "image/png");
        lrImage.ContentId = "Email_Image_Test"; //此处的ContentId 对应 htmlBodyContent 内容中的 cid: ，如果设置不正确，请不会显示图片   
        htmlBody.LinkedResources.Add(lrImage);
        mm.AlternateViews.Add(htmlBody);
        mm.IsBodyHtml = true;
        ////要求回执的标志   
       // mm.Headers.Add("Disposition-Notification-To", sender);

        ////自定义邮件头   
        //mm.Headers.Add("X-Website", "https://www.encootech.com/");

        ////针对 LOTUS DOMINO SERVER，插入回执头   
        //mm.Headers.Add("ReturnReceipt", "1");

        mm.Priority = MailPriority.Normal; //优先级   
       // mm.ReplyTo = new MailAddress("test2@163.com", "我自己");

        ////如果发送失败，SMTP 服务器将发送 失败邮件告诉我   
        mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

        ////异步发送完成时的处理事件   
        // smtp.SendCompleted += new SendCompletedEventHandler(smtp_SendCompleted);

        ////开始异步发送   
        smtp.Send(mm);
    }

    public static void tryCatchTest()
    {
        string readString = Console.ReadLine();
        int readValue;
        try
        {
            readValue = int.Parse(readString);
            Console.WriteLine(readValue);
        }
        catch (OverflowException)
        {
            Console.WriteLine("err:转化的不是一个int型数据");
        }
        catch (FormatException)
        {
            Console.WriteLine("err:格式错误");
        }
        catch (ArgumentNullException)
        {
            Console.WriteLine("err:null");
        }
        Console.ReadLine();
    }


    public static void DelectDir(string srcPath)
    {
        try
        {
            DirectoryInfo dir = new DirectoryInfo(srcPath);
            FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
            foreach (FileSystemInfo i in fileinfo)
            {
                if (i is DirectoryInfo)            //判断是否文件夹
                {
                    DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                    subdir.Delete(true);          //删除子目录和文件
                }
                else
                {
                    //如果 使用了 streamreader 在删除前 必须先关闭流 ，否则无法删除 sr.close();
                    File.Delete(i.FullName);      //删除指定文件
                }
            }
        }
        catch (Exception e)
        {
            throw;
        }
    }

    static void ForEachTest()
    {
        Int32[] intAry = new Int32[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
        Object[] a = new Object[] { 4, new Int32[] { 1, 2, 3 } };
        List<Object> ls = new List<Object> { 1, "a", 2.5 };

        foreach ( int item in intAry.ToList())
        {
   
            if (item % 2 == 0)
            {
                if(item > 5)
                {
                    continue;
                    Console.WriteLine(item.ToString() + "能被2整除");
                }
                else
                {
                    Console.WriteLine(item.ToString() + "能被2整除");
                }
               
            }else
            {
                Console.WriteLine(item.ToString() + "不能被2整除");
            }
        }
    }

    static void ParseInvoiceVerificationJson()
    {
        string file = @"D:\云扩科技\RPA\OCR\机动车发票查验.txt";
        string result = File.ReadAllText(file);

        JObject jObject = (JObject)JsonConvert.DeserializeObject(result);
        JValue resultValue = (JValue)jObject["state"];
        Console.WriteLine(resultValue.Value);
    }

    static void ParseCurrencyJson()
    {
        //string file = @"D:\云扩科技\RPA\OCR\机动车购置.txt";

        string file = @"D:\云扩科技\RPA\OCR\currency.txt";
        string result = File.ReadAllText(file);

        JObject jObject = (JObject)JsonConvert.DeserializeObject(result);
        JObject resultValue = (JObject)jObject["result"];
        //Console.WriteLine(resultValue.Value);

        JObject list = (JObject)resultValue["list"];
        Console.WriteLine(list);
        foreach (JProperty jp in (JToken)list)
        {
            Console.WriteLine(jp.Name);
            JObject valueObj = (JObject)jp.Value;
            Console.WriteLine(valueObj["name"]);
            Console.WriteLine(valueObj["rate"]);
        }
    }

       
    static void ParseJson()
    {
        string file = @"D:\云扩科技\RPA\测试文件\华云gs.txt";
        string jsonstr = File.ReadAllText(file);

        JObject obj = (JObject)JsonConvert.DeserializeObject(jsonstr);

        List<JToken> data = obj["data"].Children().ToList();

        DataTable 出差行程 = new DataTable();

        foreach (JToken item in data)
        {
            Console.WriteLine(item["报销编号"]);
            Console.WriteLine(item["出差行程"].Children().Count());

            if (item["出差行程"].Children().Count() > 0)
            {
                JToken token = item["出差行程"][0];
                DataTable dt = new DataTable();
                foreach (JProperty jp in token)
                {
                    string name = jp.Name;
                    dt.Columns.Add(name);
                }
                for (int i = 0; i < item["出差行程"].Children().Count(); i++)
                {
                    DataRow row = dt.NewRow();
                    JObject ccobj = item["出差行程"].Children().ToList()[i] as JObject;
                    IEnumerable<JToken> ss = (JToken)ccobj.AsEnumerable<JToken>();
                    foreach (JToken jkon in ss)
                    {
                        string name = ((JProperty)jkon).Name;
                        string value = ((JProperty)jkon).Value.ToString();
                        row[name] = value;
                    }
                    dt.Rows.Add(row);
                }
                foreach (DataRow dr in dt.Rows)
                {
                    //Object[] drArr = dr.ItemArray;
                    dr["住宿税率"] = dr["住宿税率"].ToString() == "0.06" ? "百分之6" : dr["住宿税率"].ToString() == "0.03" ? "百分之3" : "百分之0";
                }
                出差行程 = dt.Copy();
                foreach (DataRow row in 出差行程.Rows)
                {
                    Console.WriteLine(row["启程地点"]);
                    Console.WriteLine(row["到达地点"]);
                }
            }
        }

    }

    static void ReadCsv()
    {
        try
        {
            using (var sr = new System.IO.StreamReader(@"C:\BVTResult\Current\Result_4564.csv"))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                   
                    string[] values = line.Split(',');
                    for (int i = 0; i < values.Length; i++)
                    {
                        System.Console.Write("{0} ", values[i]);
                    }
                    System.Console.WriteLine();
                }
            }
        }
        catch (System.Exception e)
        {
            System.Console.WriteLine(e.Message);
        }
    }


    static void parseString()
    {
        string str = @"14:58:28 1第77届金球奖红毯 456万
2澳山火烟雾至南美新 435万
3特朗普再警告伊朗 350万
4伊拉克首都遭袭击 337万
5苏州十全街塌陷 279万
6暴雪蓝色预警继续 273万
7袁泉看夏雨变魔术 228万
8金球奖亚裔影后 178万
9海底捞吃出烟头新 151万
10林更新活跃像粉头新 123万";

        DataTable dt = new DataTable();
        DataColumn title = new DataColumn("Title", typeof(String));
        dt.Columns.Add(title);
        DataColumn rank = new DataColumn("Rank", typeof(String));
        dt.Columns.Add(rank);

        string[] sArray = str.Split('\n');
        foreach (string line in sArray)
        {
            string[] row = line.Split(' ');
            dt.Rows.Add(row[0], row[1]);
        }

        foreach (DataRow row in dt.Rows)
        {
            Console.WriteLine(row["Title"]);
            Console.WriteLine(row["Rank"]);
        }
    }

    static void parseJson2()
    {
        string file = @"D:\云扩科技\RPA\天匠AI开放平台\员工.txt";
        string result = File.ReadAllText(file);
        //string result = "[{\"1\":\"6月\",\"2\":\"王紫\",\"3\":\"IT部\",\"4\":\"1232223\",\"5\":\"2015.06.07\",\"6\":\"全勤\",\"7\":\"15000\",\"8\":\"4\",\"9\":\"400\",\"10\":\"300\",\"11\":\"700\",\"12\":\"0\",\"13\":\"400\",\"14\":\"800\",\"15\":\"1000\",\"16\":\"12300\",\"17\":\"请在两个月内核对好工资账单，过期不予处理，谢谢！\"}]";
        JArray resultArr = (JArray)JsonConvert.DeserializeObject(result);
        List<JToken> items = resultArr.Children().ToList();

        DataTable resultTable = new DataTable();
        int[] a = new int[17];
        int index = 1;
        foreach (int i in a)
        {
            resultTable.Columns.Add(index.ToString());
            index = index + 1;
        }
        foreach (JToken item in items)
        {
            DataRow row = resultTable.NewRow();
            foreach(DataColumn col in resultTable.Columns)
            {
                row[col.ColumnName] = item[col.ColumnName];
            }
            resultTable.Rows.Add(row);
        }

        Console.WriteLine(resultTable.Rows[0]["1"]);
       
        foreach (DataRow r in resultTable.Rows)
        {
            Console.WriteLine(r["1"]);
            foreach (DataColumn c in resultTable.Columns)
            {
                Console.WriteLine(r[c.ColumnName]);
            }
            Console.WriteLine("===================");
        }
    }

    // 身份证识别
    static void ParseIdJson()
    {
        string file = @"D:\OCR\身份证识别结果.txt";
        string result = File.ReadAllText(file);

        JObject resultObj = (JObject)JsonConvert.DeserializeObject(result);

        string 识别状态 = Convert.ToString(resultObj["image_status"]);
        string 住址;
        string 出生日期;
        string 姓名;
        string 公民身份号码;
        string 性别;
        string 民族;

        if (识别状态 == "normal")
        {
            识别状态 = "成功";
            Console.WriteLine("--------识别成功！------");
        }
        else
        {
            识别状态 = "失败";
            Console.WriteLine("--------识别失败，部分信息不全！------");
        }
        Console.WriteLine(识别状态);
        List <JToken> wordsResult = resultObj["words_result"].Children().ToList();
        foreach (JProperty item in wordsResult)
        {
            string name = item.Name.ToString();
            if(name== "住址")
            {
                住址 = item.Value["words"].ToString();
                Console.WriteLine(住址);
            }

            if (name == "出生")
            {
                出生日期 = item.Value["words"].ToString();
                Console.WriteLine(出生日期);
            }

            if (name == "姓名")
            {
                姓名 = item.Value["words"].ToString();
                Console.WriteLine(姓名);
            }

            if (name == "公民身份号码")
            {
                公民身份号码 = item.Value["words"].ToString();
                Console.WriteLine(公民身份号码);
            }


            if (name == "性别")
            {
                性别 = item.Value["words"].ToString();
                Console.WriteLine(性别);
            }


            if (name == "民族")
            {
                民族 = item.Value["words"].ToString();
                Console.WriteLine(民族);
            }

        }


        //JArray identify_results = (JArray)imageStatus["identify_results"];

        ////Console.WriteLine(data["identify_results"]);

        //DataTable resultTable = new DataTable();

        //List<JToken> items = identify_results.Children().ToList();

        //foreach (JObject item in items)
        //{
        //    JObject details = (JObject)item["details"];


        //    DataTable dt = new DataTable();
        //    DataRow row = dt.NewRow();

        //    foreach (JProperty jp in (JToken)details)
        //    {
        //        string name = jp.Name.ToString();
        //        dt.Columns.Add(name);
        //    }

        //    foreach (JProperty jp in (JToken)details)
        //    {
        //        string name = jp.Name.ToString();
        //        string value = jp.Value.ToString();
        //        if (!string.IsNullOrEmpty(name) && name != "items")
        //        {
        //            row[name] = value;
        //        }
        //    }


        //    dt.Rows.Add(row);
        //    resultTable = dt.Copy();

        //    foreach (DataColumn col in dt.Columns)
        //    {
        //        Console.WriteLine(col);
        //    }

        //}
    }

    /// <summary>
    /// 将CSV文件的数据读取到DataTable中
    /// </summary>
    /// <param name="fileName">CSV文件路径</param>
    /// <returns>返回读取了CSV数据的DataTable</returns>
    public static DataTable OpenCSV(string filePath, string[] colSep)
    {
        Encoding encoding = Encoding.UTF8;
        DataTable dt = new DataTable();
        FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
        //StreamReader sr = new StreamReader(fs, Encoding.UTF8);
        StreamReader sr = new StreamReader(fs, encoding);
        //string fileContent = sr.ReadToEnd();
        //encoding = sr.CurrentEncoding;
        //记录每次读取的一行记录
        string strLine = "";
        //记录每行记录中的各字段内容
        string[] aryLine = null;
        string[] tableHead = null;
        //标示列数
        int columnCount = 0;
        //标示是否是读取的第一行
        bool IsFirst = true;
        //逐行读取CSV中的数据
        while ((strLine = sr.ReadLine()) != null)
        {
            //strLine = Common.ConvertStringUTF8(strLine, encoding);
            //strLine = Common.ConvertStringUTF8(strLine);

            if (IsFirst == true)
            {
                tableHead = strLine.Split(colSep, StringSplitOptions.RemoveEmptyEntries);
                IsFirst = false;
                columnCount = tableHead.Length;
                //创建列
                for (int i = 0; i < columnCount; i++)
                {
                    DataColumn dc = new DataColumn(tableHead[i]);
                    dt.Columns.Add(dc);
                }
            }
            else
            {
                aryLine = strLine.Split(colSep, StringSplitOptions.RemoveEmptyEntries);
                DataRow dr = dt.NewRow();
                for (int j = 0; j < columnCount; j++)
                {
                    dr[j] = aryLine[j];
                }
                dt.Rows.Add(dr);
            }
        }
        if (aryLine != null && aryLine.Length > 0)
        {
            dt.DefaultView.Sort = tableHead[0] + " " + "asc";
        }

        sr.Close();
        fs.Close();
        return dt;
    }
}
