using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;

namespace ConsoleApp1.helpers
{
    class ProcessHelper
    {
        public static bool ProcessExist(string processName)
        {
            var processList = Process.GetProcessesByName(processName);
            return processList.Length > 0;
        }

    }
}
