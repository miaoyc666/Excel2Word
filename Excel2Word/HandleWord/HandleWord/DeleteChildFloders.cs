using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CreateReport 
{
    class DeleteChildFloders 
    {
        public static void delete(string fatherDir) 
        {
            string[] childs = null;
            try 
            {
                if (!Directory.Exists(fatherDir)) 
                {
                    Directory.CreateDirectory(fatherDir);
                }
                else 
                {
                    childs = Directory.GetDirectories(fatherDir);
                }
            } 
            catch { }

            foreach(string child in childs)
            {
                DeleteFolder(child);
            }
        }

        private static void DeleteFolder(string dir) 
        {
            // 循环文件夹里面的内容
            foreach (string f in Directory.GetFileSystemEntries(dir)) 
            {
                // 如果是文件存在
                if (File.Exists(f)) 
                {
                    FileInfo fi = new FileInfo(f);
                    if (fi.Attributes.ToString().IndexOf("Readonly") != 1) 
                    {
                        fi.Attributes = FileAttributes.Normal;
                    }
                    // 直接删除其中的文件
                    File.Delete(f);
                } 
                else 
                {
                    // 如果是文件夹存在
                    // 递归删除子文件夹
                    DeleteFolder(f);
                }
            }
            // 删除已空文件夹
            Directory.Delete(dir);
        }
    }
}
