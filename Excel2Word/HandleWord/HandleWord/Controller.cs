using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;


namespace CreateReport {
    class Controller {

        //建立开关这个对象，存数据。
        class onOff {
            private string _wei;
            private string _name;
            private string _type;
            public onOff(string wei, string name, string type) {
                _wei = wei;
                _name = name;
                _type = type;
            }
            public string getWei() {
                return _wei;
            }
            public string getName() {
                return _name;
            }
            public string getType() {
                return _type;
            }
        }

        public static void Insert(DictionaryEntry de,handleWord report) {
            switch (Classify(de.Key.ToString())){
                case 0:
                    InsertProduceName(de, report);
                    break;
                case 1:
                    InsertNormally(de, report);
                    break;
                case 2:
                    InsertP4(de, report);
                    break;
                case 3:
                    InsertAppNum(de, report);
                    break;
                case 4:
                    InsertReportNum(de, report);
                    break;
                case 5:
                    InsertTypeName(de,report);
                    break;
                case 6:
                    InsertRatedVottage(de,report);
                    break;
                case 7:
                    InsertRatedCurrent(de, report);
                    break;
                case 8:
                    InsertLingJian(de,report);
                    break;
                case 9:
                    InsertSigniture(de, report);
                    break;
                default:
                    break;
            }
        }

        private static int Classify(string tmp_st) {
            if (tmp_st.Equals("p3_2")) {return 0;
            }else if(isP4(tmp_st)){
                return 2;
            }else if (tmp_st.Equals("p1_3_1")) { //申请编号
                return 3;
            } else if (tmp_st.Equals("p1_0_1")) { //报告编号
                return 4;
            } else if (tmp_st.Equals("p3_1_1_1")) { //型号
                return 5;
            } else if (tmp_st.Equals("p3_1_2_1")) { //电压
                return 6;
            } else if (tmp_st.Equals("p3_1_3_1")) { //电流
                return 7;
            } else if(tmp_st.Equals("lj2_0")){
                return 8;
            } else if(tmp_st.Equals("shenhe") || tmp_st.Equals("zhujian")){
                return 9;
            }
            else { return 1; }
        }

        private static void InsertProduceName(DictionaryEntry de, handleWord report) {

            string[] childData = de.Value.ToString().Split('\n');
            List<onOff> list = new List<onOff>();
            //整理p3_2产品名称部分。形成对象。
            for (int i = 0; i < childData.Length; i++) {
                string[] tmp = childData[i].Split('|');
                if (tmp.Length == 3) {
                    onOff tmpONOFF = new onOff(tmp[0], tmp[1], tmp[2]);
                    list.Add(tmpONOFF);
                }
            }

            List<string> on_off_name = new List<string>();//存放不重复的开关名称
            List<ArrayList> on_off_type = new List<ArrayList>();//存放分类好了的开关类型，如果输入时，没有顺序输入，也没问题。
            
            foreach (onOff ONOFF in list) {
                if (on_off_name.Contains(ONOFF.getName())) {
                } else {
                    on_off_name.Add(ONOFF.getName());
                }
            }

            for (int i = 0; i < on_off_name.Count; i++) {
                ArrayList tmpArr = new ArrayList();
                foreach (onOff ONOFF in list) {
                    if (ONOFF.getName().Equals(on_off_name.ElementAt(i))) {
                        tmpArr.Add(ONOFF.getType());
                    } else { }
                }
                on_off_type.Add(tmpArr);
            }

            //整理数据成为string
            string p0_produce_name = null;
            foreach(string name in on_off_name){
                p0_produce_name = p0_produce_name + name + ";" + "\n";
            }
            string p0_produce_type = null;
            foreach(ArrayList arr in on_off_type){
                foreach(string tmp_str in arr){
                    p0_produce_type = p0_produce_type + tmp_str + "\n";
                }
                p0_produce_type = p0_produce_type.TrimEnd("\n".ToCharArray()) + ";\n";
            }
            string all_type = null;
            foreach(onOff ONOFF in list){
                all_type = all_type + ONOFF.getType() + "、";
            }
            if(all_type != null){
                all_type = all_type.TrimEnd("、".ToCharArray());
            }
            
            //开始插入数据
            report.InsertValue("p0_produce_name", p0_produce_name);
            report.InsertValue("p0_produce_type", p0_produce_type);
            report.InsertValue("p3_2", de.Value.ToString().Replace("|", ""));
            report.InsertValue("c8_1_5_2", all_type);
            report.InsertValue("p1_cover_type", all_type);

        }

        private static void InsertNormally(DictionaryEntry de, handleWord report) {
            report.InsertValue(de.Key.ToString(),de.Value.ToString());
        }

        private static void InsertP4(DictionaryEntry de, handleWord report) {
            int position = isContain(de.Value.ToString());
            if (position == -1) {
                InsertNormally(de, report);
            } else {
                string bookmark = "p4_c_" + position.ToString();
                report.InsertSymbol(bookmark, -3976, "Wingdings");
            }
            
        }

        private static void InsertAppNum(DictionaryEntry de, handleWord report) {
            report.InsertValue("p1_3_1_position_0", de.Value.ToString());
            report.InsertValue("p1_3_1_position_1", de.Value.ToString());
        }

        private static void InsertReportNum(DictionaryEntry de, handleWord report) {
            //report.InsertValue("p1_0_1_position_0", de.Value.ToString());
            //report.InsertValue("p1_0_1_position_1", de.Value.ToString());
            //report.InsertValue("p1_0_1_position_2", de.Value.ToString());
            //report.InsertValue("p1_0_1_position_3", de.Value.ToString());
        }
        private static void InsertRatedCurrent(DictionaryEntry de, handleWord report) {
            report.InsertValue("p3_1_3_1_position_0", de.Value.ToString());
            report.InsertValue("p3_1_3_1_position_1", de.Value.ToString());
            report.InsertValue("p3_1_3_1_position_2", de.Value.ToString());
        }
        private static void InsertRatedVottage(DictionaryEntry de, handleWord report) {
            report.InsertValue("p3_1_2_1_position_0", de.Value.ToString());
            report.InsertValue("p3_1_2_1_position_1", de.Value.ToString());
            report.InsertValue("p3_1_2_1_position_2", de.Value.ToString());
        }
        private static void InsertTypeName(DictionaryEntry de, handleWord report) {
            report.InsertValue("p3_1_1_1_position_0", de.Value.ToString());
            report.InsertValue("p3_1_1_1_position_1", de.Value.ToString());
            report.InsertValue("p3_1_1_1_position_2", de.Value.ToString());
        }

        private static void InsertSigniture(DictionaryEntry de, handleWord report) {
            string picPath = @"D:\浙江省质量检测研究院\serverData\签名\";
            report.InsertPicture(de.Key.ToString(),picPath+de.Value.ToString()+".PNG",100,45);
        }

        private static void InsertLingJian(DictionaryEntry de, handleWord report) {
            string[] allLingJian = de.Value.ToString().Split('|');
            int index = 0;
            string markTitle = "lj_";
            foreach(string value in allLingJian){
                string bookmark = markTitle + index;
                report.InsertValue(bookmark, value);
                index++;
            }
        }
        
        //判断id是否是属于page4的
        private static bool isP4(string tmpStr) {
            string[] childChapter = tmpStr.Split('_');
            if (childChapter[0].Equals("p4")) {
                return true;
            } else {
                return false;
            }
        }

        private static int isContain(string checkString) {
            string[] allWords = { "1", "2", "3", "03", "4", "5", "6", "6/2", "7", "正常间隙结构", "小间隙结构",
                                    "微间隙结构", "旋转开关", "倒扳开关", "跷板开关", "按钮开关", "拉线开关", "明装式开关", 
                                    "暗装式开关", "半暗装式开关", "面板式开关", "框缘式开关", "无需移动导线便可拆卸盖或盖板的固定式开关（结构A）", 
                                    "不移动导线便不能拆卸盖或盖板的固定式开关（结构B）", "螺钉端子", "螺栓端子", "柱型端子", "鞍型端子", "罩式端子", 
                                    "仅适于连接硬导线的无螺纹型端子", "适于连接硬导线和软导线的无螺纹端子", "有", "无",
                                    "卡扣式", "螺钉式", "整体式", "其他"
                                };
            checkString = checkString.Trim();
            for (int i = 0; i < allWords.Length; i++) {
                if (checkString.Equals(allWords[i]))
                    return i;
            }
            return -1;
        }
        
    }
}
