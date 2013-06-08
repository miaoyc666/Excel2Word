using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CreateReport 
{
    class ChangeText 
    {

        public static String  replaceWords(String str)
        {
            String tmp_str = str.Replace("\\s", " ");
            tmp_str = tmp_str.Replace("\\n", "\n");
            
            Console.WriteLine(tmp_str);

            if (tmp_str.ToString().Equals("!!~")) 
            {
                return "";
            }
            else
            {
                return tmp_str;
            }
        }
    }
}
