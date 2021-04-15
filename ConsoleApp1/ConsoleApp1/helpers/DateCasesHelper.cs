using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.helpers
{
    class DateCasesHelper
    {
        //  //获取前一月的月末;
        public static void preMonthLastDay()
        {
            DateTime dt = DateTime.Parse("2021/03/02");
            Console.WriteLine(dt.AddDays(-dt.Day));
        }
    }
}
