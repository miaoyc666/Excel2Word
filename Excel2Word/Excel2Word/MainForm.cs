using System;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace Excel2Word
{
    public partial class MainForm : Form
    {
        private _Application wordApp = null;
        private _Document wordDoc = null;

        public MainForm()
        {
            InitializeComponent();
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            DataSet excelData = ImportFromExcel();
            if (0 == excelData.Tables.Count)
            {
                return;
            }

            int nRowCount = excelData.Tables[0].Rows.Count;
            int nColCount = excelData.Tables[0].Columns.Count;

            for (int nRowIndex = 0; nRowIndex < nRowCount; ++nRowIndex)
            {
                List<string> lstColsData = new List<string>();
                for (int nColIndex = 0; nColIndex < nColCount; ++nColIndex)
                {
                    lstColsData.Add(excelData.Tables[0].Rows[nRowIndex][nColIndex].ToString());
                }
            }

            string path = @"D:\浙江省质量检测研究院\serverData\通过的表单\test.docx"; //存放pass表单的文件夹
            CreateNewDocument(path);
        }

        /// 从选择的Excel文件导入
        public DataSet ImportFromExcel()
        {
            DataSet dataSet = new DataSet();
            if (openFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                dataSet = doImport(openFileDlg.FileName);
            }
            return dataSet;
        }

        /// 从指定的Excel文件导入
        public DataSet ImportFromExcel(string strFileName_)
        {
            DataSet dataSet = new DataSet();
            dataSet = doImport(strFileName_);
            return dataSet;
        }

        /// 执行导入
        private DataSet doImport(string strFileName_)
        {
            if (strFileName_ == "")
            {
                return null;
            }

            string ExcelTableName = "";
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" +
                "Data Source=" + strFileName_ + ";" +
                "Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";

            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            OleDbDataAdapter myCommand;
            // 获取文件中TABLE类型的表

            // TODO:
            // 此处可以改为读取多个DataTable, 现阶段为只读取第一个Sheet中的数据
            //

            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            DataSet dsExcel = new DataSet();
            try
            {
                DataRow dataRow = schemaTable.Rows[0];
                ExcelTableName = dataRow["TABLE_NAME"].ToString().Trim();
                //从对应Excel内容的表中获取数据
                string strExcel = "select * from [" + ExcelTableName + "]";
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(dsExcel);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
            }
            return dsExcel;
        }

        /// 执行导出
        private bool doExport(string strFileName_)
        {
            return true;
        }

        //通过模板创建新文档
        public void CreateNewDocument(string filePath)
        {
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
    }
}
