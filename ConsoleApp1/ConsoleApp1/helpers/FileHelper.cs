using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConsoleApp1.helpers
{
    class FileHelper
    {
        public void Run()
        {
            //在这里编写您的代码
            string 下载文件夹 = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
            string 下载的文件 = MostCurrentFile(下载文件夹);
        }
        //在这里编写您的函数或者类

        public string MostCurrentFile(string 目标路径)
        {
            string fileFilter = "aap1*.xls";
            //string filePath = @"D:\chrome downloads";
            DirectoryInfo di = new DirectoryInfo(目标路径);

            FileInfo[] arrFi = di.GetFiles(fileFilter); // aqpq开头的文件
            SortAsFileCreationTime(ref arrFi);

            //Console.WriteLine(arrFi[0].CreationTime);
            //Console.WriteLine(DateTime.Now.AddMinutes(-10));
            //Console.WriteLine(arrFi[0].CreationTime.CompareTo(DateTime.Now.AddMinutes(-10)));
            if (arrFi.Length > 0 && arrFi[0].LastWriteTime.CompareTo(DateTime.Now.AddMinutes(-10)) > 0)
            {
                return arrFi[0].FullName;
            }
            else
            {
                return string.Empty;
            }

        }

        public void SortAsFileCreationTime(ref FileInfo[] arrFi)
        {
            Array.Sort(arrFi, delegate (FileInfo x, FileInfo y) { return y.LastWriteTime.CompareTo(x.LastWriteTime); });
        }
    }
}
