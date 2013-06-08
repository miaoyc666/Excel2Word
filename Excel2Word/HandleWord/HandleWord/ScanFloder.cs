using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;

namespace CreateReport 
{
    class ScanFloder 
    {

        public static void Scan(string tempalte) 
        {
            string path = @"D:\浙江省质量检测研究院\serverData\通过的表单"; //存放pass表单的文件夹
            string templatePath = @"D:\浙江省质量检测研究院\表单模板\" + tempalte + @".doc";
            string baseSavePath = @"D:\浙江省质量检测研究院";
            string[] childs = null;

            try 
            {
                if (!Directory.Exists(path)) 
                {
                    Directory.CreateDirectory(path);
                }
                else 
                {
                    childs = Directory.GetDirectories(path);
                }
            }
            catch { }

            Hashtable allData = new Hashtable();
            
            foreach (String dataTable in childs) 
            {    
                //用来获得一个文件夹中的所有数据
                foreach(string dataFile in Directory.GetFiles(dataTable))
                {
                    string[] dataName = dataFile.Split('\\');
                    dataName = dataName[dataName.Length - 1].Split('.');//获取文件名的一个数组，用来区分邮箱附带的那些附件

                    if (dataName.Length == 4) 
                    {
                        StreamReader readData = new StreamReader(dataFile);

                        string tmp_str = readData.ReadToEnd().ToString();
                        Scanner handleText = new Scanner(tmp_str);
                        while(handleText.hasNext())
                        {
                            string _key = handleText.nextWord();
                            string _value = ChangeText.replaceWords(handleText.nextWord());
                            allData.Add(_key, _value);
                        }
                        readData.Close();
                    } 
                    else if (dataName[0].Equals("shenheDate") || dataName[0].Equals("zhujianDate")) 
                    {
                        StreamReader readData = new StreamReader(dataFile);
                        string tmp_str = readData.ReadToEnd().ToString();

                        Scanner handleText = new Scanner(tmp_str);
                        while(handleText.hasNext())
                        {
                            allData.Add(dataName[0].ToString(), handleText.nextWord());
                        }
                        readData.Close();
                    }
                }
                //获得主检跟审核人姓名
                string[] dataFolder = dataTable.Split('\\');
                dataFolder = dataFolder[dataFolder.Length - 1].Split('$');
                string shenhe = dataFolder[3];
                string zhujian = dataFolder[0];
                allData.Add("shenhe", shenhe);//通过username找到对应签名
                allData.Add("zhujian", zhujian);
                //开始对加载的一个文件夹的数据开始操作
                handleWord report = new handleWord();
                report.CreateNewDocument(templatePath);
                
                foreach(DictionaryEntry de in allData)
                {
                    Controller.Insert(de, report);
                }

                //可能需要一个计算路径的东西
                string[] tmp = dataTable.Split('\\');
                string savePath = baseSavePath + "\\" + tempalte + "\\"+ tmp[tmp.Length - 1] + @".doc";
                report.SaveDocument(savePath);
                //操作完成后，清除hashtable
                allData.Clear();
            }
        }
    }
}
