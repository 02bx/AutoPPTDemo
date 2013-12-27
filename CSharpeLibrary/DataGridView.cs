using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace CSharpeLibrary
{
    public class DataGridView : System.Windows.Forms.DataGridView
    {
        public DataGridView()
        { 
        
        }

        /// <summary>
        /// 行头显示行号（放入"RowPostPaint"事件）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <param name="dgv"></param>
        public static void dgv_RowPostPaint(System.Windows.Forms.DataGridView dgv, DataGridViewRowPostPaintEventArgs e)
        {
            SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor);
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), dgv.DefaultCellStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4);
        }

        /// <summary>
        /// 禁止列头排序(放入"ColumnHeaderMouseClick"事件)
        /// </summary>
        /// <param name="dgv"></param>
        public static void DoNotDataGridViewColumnSort(System.Windows.Forms.DataGridView dgv)
        {
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        /// <summary>
        /// 自动适应列的尺寸
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="dgv"></param>
        public static void DataGridViewAutoSizeColumn(DataTable dt, System.Windows.Forms.DataGridView dgv)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dgv.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        /// <summary>
        /// 增加DataTable列名
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="colName"></param>
        /// <returns></returns>
        public static DataTable AddDataTableColumn(DataTable dt, string[] colName)
        {
            for (int i = 0; i < colName.Length; i++)
            {
                dt.Columns.Add(new DataColumn(colName[i]));
            }
            return dt;
        }

        /// <summary>
        /// 复制行
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="dgv"></param>
        public static void CopyRow(DataTable dt, System.Windows.Forms.DataGridView dgv)
        {
            if (dgv.CurrentRow != null)
            {
                int index = dgv.CurrentRow.Index;
                DataRow dr = dt.NewRow();
                dr.ItemArray = dt.Rows[index].ItemArray;
                dt.Rows.Add(dr);
            }
        }

        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="dgv"></param>
        public static void DelRow(DataTable dt, System.Windows.Forms.DataGridView dgv)
        {
            if (dgv.CurrentRow != null)
            {
                int index = dgv.CurrentRow.Index;
                dt.Rows.RemoveAt(index);
            }
        }

        /// <summary>
        /// 上移行
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="dgv"></param>
        public static void UpMoveRow(DataTable dt, System.Windows.Forms.DataGridView dgv)
        {
            int index = dgv.CurrentRow.Index;
            if (dgv.CurrentRow.Index <= 0)
            {
                return;
            }
            else
            {
                DataRow tempdr = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    tempdr[i] = dt.Rows[index][i];
                }
                dt.Rows.InsertAt(tempdr, index - 1);
                dt.Rows.RemoveAt(index + 1);
                dgv.ClearSelection();
                dgv.Rows[index - 1].Selected = true;
                dgv.CurrentCell = dgv.Rows[index - 1].Cells[0];
            }
        }

        /// <summary>
        /// 下移行
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="dgv"></param>
        public static void DownMoveRow(DataTable dt, System.Windows.Forms.DataGridView dgv)
        {
            int index = dgv.CurrentRow.Index;
            if (index == dt.Rows.Count - 1)
            {
                return;
            }
            else if (index == -1)
            {
                return;
            }
            else
            {
                DataRow tempdr = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    tempdr[i] = dt.Rows[index][i];
                }
                dt.Rows.InsertAt(tempdr, index + 2);
                dt.Rows.RemoveAt(index);
                dgv.ClearSelection();
                dgv.Rows[index + 1].Selected = true;
                dgv.CurrentCell = dgv.Rows[index + 1].Cells[0];
            }
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="dgv"></param>
        /// <returns></returns>
        public static string DataGridViewToExcel(string path, System.Windows.Forms.DataGridView dgv)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                return "无法创建Excel对象，可能您的机子未安装Excel";
            }

            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1

            //写入标题
            for (int i = 0; i < dgv.ColumnCount; i++)
            {
                worksheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;
            }
            //写入数值
            for (int r = 0; r < dgv.Rows.Count; r++)
            {
                for (int i = 0; i < dgv.ColumnCount; i++)
                {
                    worksheet.Cells[r + 2, i + 1] = dgv.Rows[r].Cells[i].Value;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            worksheet.Columns.EntireColumn.AutoFit();//列宽自适应

            if (path != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(path);
                }
                catch (Exception ex)
                {
                    return "导出文件时出错,文件可能正被打开！\n" + ex.Message;
                }

            }
            xlApp.Quit();
            GC.Collect();
            return path + "保存成功";
        }
    }


    public class DataGridRow : System.Windows.Forms.DataGridView
    {
        public DataGridRow()
        {
            this.CurrentCell = null;
        }
    }

}
