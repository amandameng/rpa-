using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 获得指定工作日
{
    [ToolboxBitmap(typeof(Class1), "time.jpg")]
    [DisplayName("获得指定工作日")]
    [Description("获得指定工作日")]
    public class Class1 : BaseActivity
    {
        public Class1()
        {
            DisplayName = "获得指定工作日";
        }

        [Category("输入")]
        [DisplayName("天数")]
        [Description("天数")]
        public InArgument<int> InName { get; set; }

        [Category("输出")]
        [DisplayName("返回值")]
        [Description("返回值")]
        public OutArgument<DateTime> Result { get; set; }

        protected override void Execute(NativeActivityContext context)
        {

            Result.Set(context, getNextDate(InName.Get(context)));
            Console.WriteLine(InName.Get(context));
        }

        static bool stopCycle = false;
        public static DateTime nextDate(DateTime dt, int addDay)
        {
            if (stopCycle) return dt;
            string date = dt.ToString("MMdd");
            List<string> holiday = new List<string>(new String[] { "1001", "1002", "1003", "1004", "1005", "1006", "1007", "1008" });
            List<string> work = new List<string>(new String[] { "0927", "1010" });

            if (holiday.IndexOf(date) >= 0)  //法定节假日
            {
                dt = dt.AddDays(addDay);
                dt = nextDate(dt, addDay);
            }
            int week = (int)dt.DayOfWeek;
            if (week == 0 || week == 6)
            {
                if (work.IndexOf(date) >= 0) // 周末补班
                {
                    return dt;
                }
                else
                {
                    dt = dt.AddDays(addDay);
                    dt = nextDate(dt, addDay);
                }

            }
            return dt;
        }
        public static DateTime getNextDate(int day)
        {
            int addDay = day >= 0 ? 1 : -1;
            DateTime currDate = DateTime.Now;
            DateTime nDate = currDate.AddDays(day);
            nDate = nextDate(nDate, addDay);
            return nDate;
        }
    }
}
