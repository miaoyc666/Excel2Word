using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Timers;

namespace CreateReport 
{
    class Program 
    {
        static void Main(string[] args) 
        {
            ScanFloder.Scan(@"最终报表");
            ScanFloder.Scan(@"原始记录");
            DeleteChildFloders.delete(@"D:\浙江省质量检测研究院\serverData\通过的表单");
        }
    }
}
