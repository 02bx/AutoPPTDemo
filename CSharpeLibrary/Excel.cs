using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;

namespace CSharpeLibrary
{
    public class Excel
    {
        Range r = null;
        //读取用例excel
        public void ReadExcel(string excelFilePath)
        {
            Microsoft.Office.Interop.Excel.Application ExcelObj = null;
            Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;
            try
            {
                //.net2.0
                //ExcelObj = new Microsoft.Office.Interop.Excel.ApplicationClass();
                //.net4.0
                ExcelObj = new Microsoft.Office.Interop.Excel.Application();
                ExcelObj.Visible = false;
                theWorkbook = ExcelObj.Workbooks.Open(excelFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
                Microsoft.Office.Interop.Excel.Worksheet xsheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);
                string strFirstSheetName = xsheet.Name;
                int RowCount = xsheet.UsedRange.Cells.Rows.Count;
                int ColumnCount = xsheet.UsedRange.Cells.Columns.Count;
                int obj = 0;
                int method = 0;
                int value = 0;
                int expect = 0;
                int actual = 0;
                int check = 0;
                for (int i = 0; i < ColumnCount; i++)
                {
                    r = (Range)xsheet.Cells[1, i];
                    if (r.Text.ToString() == "对象")
                    {
                        obj = i;
                    }
                    else if (r.Text.ToString() == "属性")
                    {
                        method = i;
                    }
                    else if (r.Text.ToString() == "值")
                    {
                        value = i;
                    }
                    else if (r.Text.ToString() == "期望值")
                    {
                        expect = i;
                    }
                    else if (r.Text.ToString() == "实际值")
                    {
                        actual = i;
                    }
                    else if (r.Text.ToString() == "检查点")
                    {
                        check = i;
                    }
                }
                //循环每一行
                for (int i = 2; i <= RowCount; i++)
                {
                    string str = "";
                    for (int j = 1; j <= ColumnCount; j++)
                    {
                        r = (Range)xsheet.Cells[i, j];
                        if (j == obj)
                        {
                            str += r.Text;
                        }
                        else if (j == method)
                        {
                            str += r.Text;
                        }
                        else if (j == value)
                        {
                            str += r.Text;
                        }
                        else if (j == expect)
                        {
                            str += r.Text;
                        }
                        else if (j == actual)
                        {
                            str += r.Text;
                        }
                        else if (j == check)
                        {
                            str += r.Text;
                        }
                        //MessageBox.Show(r.Text.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("导入出错：" + ex, "错误信息");
            }
            finally
            {
                theWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                theWorkbook = null;
                ExcelObj.Quit();
                ExcelObj = null;
            }
        }

        //读取对象库
        public DataSet ReadObject(string path, string tableHeader)
        {
            string con = "";
            con = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "; Extended Properties=\"Excel 8.0;IMEX=1;HDR=" + tableHeader + "\"";
            OleDbConnection olecon = new OleDbConnection(con);
            OleDbDataAdapter myda = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", con);
            DataSet myds = new DataSet();
            myda.Fill(myds);
            return myds;
        }

    }
}
