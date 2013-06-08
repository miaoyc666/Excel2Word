using System;
using System.IO;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;

namespace WindowsFormsApplication1
{
    public partial class TestForm : Form
    {
        public TestForm()
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
            
            int asd = excelData.Tables[0].Rows.Count;       // 取第一个表中的数据
            List<int> test = new List<int>();

            
            for ( int index = 0; index < 10; ++index )
            {
                test.Insert(index, index);
            }
            for ( int index = 0; index < test.Count; ++index )
            {
            }
        }



        #region 从Excel文件导入到DataSet
        
        //        /// <summary>
        //        /// 从Excel导入文件
        //        /// </summary>
        //        /// <param name="strExcelFileName">Excel文件名</param>
        //        /// <returns>返回DataSet</returns>
        //        public DataSet ImportFromExcel(string strExcelFileName)
        //        {
        //            return doImport(strExcelFileName);
        //        }
        /// <summary>
        /// 从选择的Excel文件导入
        /// </summary>
        /// <returns>DataSet</returns>
        public DataSet ImportFromExcel()
        {
            DataSet dataSet = new DataSet();
            if (openFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                dataSet = doImport(openFileDlg.FileName);
            }
            return dataSet;
        }

        /// <summary>
        /// 从指定的Excel文件导入
        /// </summary>
        /// <param name="strFileName">Excel文件名</param>
        /// <returns>DataSet</returns>
        public DataSet ImportFromExcel(string strFileName)
        {
            DataSet dataSet = new DataSet();
            dataSet = doImport(strFileName);
            return dataSet;
        }

        /// <summary>
        /// 执行导入
        /// </summary>
        /// <param name="strFileName">文件名</param>
        /// <returns>DataSet</returns>
        private DataSet doImport(string strFileName)
        {
            if (strFileName == "")
            {
                return null;
            }

            string ExcelTableName = "";
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" +
                "Data Source=" + strFileName + ";" +
                "Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";

            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            OleDbDataAdapter myCommand;
            //获取文件中TABLE类型的表
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            DataSet dsExcel = new DataSet();
            try
            {
                DataRow dr = schemaTable.Rows[0];
                ExcelTableName = dr["TABLE_NAME"].ToString().Trim();
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
        #endregion
    }
}
