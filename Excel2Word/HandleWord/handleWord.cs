using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace CreateReport {
    class handleWord {
        private _Application wordApp = null;
        private _Document wordDoc = null;
        public _Application Application {
            get {
                return wordApp;
            }
            set {
                wordApp = value;
            }
        }
        public _Document Document {
            get {
                return wordDoc;
            }
            set {
                wordDoc = value;
            }
        }

        //通过模板创建新文档
        public void CreateNewDocument(string filePath) {
            //killWinWordProcess();
            wordApp = new ApplicationClass();
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = filePath;
            wordDoc = wordApp.Documents.Open(ref templateName, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);
        }

        //保存新文件
        public void SaveDocument(string filePath) {
            object fileName = filePath;
            object format = WdSaveFormat.wdFormatDocument;//保存格式
            object miss = System.Reflection.Missing.Value;
            wordDoc.SaveAs(ref fileName, ref format, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss);
            //关闭wordDoc，wordApp对象
            object SaveChanges = WdSaveOptions.wdSaveChanges;
            object OriginalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
            object RouteDocument = false;
            wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            wordApp.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
        }

        //在书签处插入值
        public bool InsertValue(string bookmark, string value) {
            object bkObj = bookmark;
            if (wordApp.ActiveDocument.Bookmarks.Exists(bookmark)) {
                wordApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
                wordApp.Selection.TypeText(value);
                return true;
            }
            return false;
        }

        //插入符号
        public bool InsertSymbol(string bookmark, int charNum, string Font) {
            object bkObj = bookmark;
            if (wordApp.ActiveDocument.Bookmarks.Exists(bookmark)) {

                wordApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();

                wordApp.Selection.InsertSymbol(charNum, Font, true);
                return true;
            }
            return false;
        }

        //插入一段文字,text为文字内容
        public void InsertText(string bookmark, string text) {
            object oStart = bookmark;
            object range = wordDoc.Bookmarks.get_Item(ref oStart).Range;
            Paragraph wp = wordDoc.Content.Paragraphs.Add(ref range);
            wp.Format.SpaceBefore = 6;
            wp.Range.Text = text;
            wp.Format.SpaceAfter = 24;
            wp.Range.InsertParagraphAfter();
            wordDoc.Paragraphs.Last.Range.Text = "\n";
        }

        //插入图片
        public void InsertPicture(string bookmark, string picturePath, float width, float hight) {
            object bkObj = bookmark;
            if (wordApp.ActiveDocument.Bookmarks.Exists(bookmark)) {
                object miss = System.Reflection.Missing.Value;
                object oStart = bookmark;
                Object linkToFile = false;       //图片是否为外部链接
                Object saveWithDocument = true;  //图片是否随文档一起保存
                object range = wordDoc.Bookmarks.get_Item(ref oStart).Range;//图片插入位置
                wordDoc.InlineShapes.AddPicture(picturePath, ref linkToFile, ref saveWithDocument, ref range);
                wordDoc.Application.ActiveDocument.InlineShapes[1].Width = width;   //设置图片宽度
                wordDoc.Application.ActiveDocument.InlineShapes[1].Height = hight;  //设置图片高度
            }
        }

        // 杀掉winword.exe进程
        public void killWinWordProcess() {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes) {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "") {
                    process.Kill();
                }
            }
        }
    }
}
