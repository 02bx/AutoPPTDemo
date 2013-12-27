using System;
using System.Collections.Generic;
using System.Text;
using ATSystemModels;
using System.Data;

namespace CaseEditor
{
    public class Common
    {
        public static string caseName = "";
        public static string caseProperty = "";
        public static List<ObjectLibrary> objlist = null;
        public static int flag = 0;
        public static string excelPath = "";
        public static string sqlServer = "";
        public static string sqlName = "";
        public static string sqlPwd = "";
        public static string cfgPath = AppDomain.CurrentDomain.BaseDirectory + "config.ini";

        private string objName;

        public string ObjName
        {
            get { return objName; }
            set { objName = value; }
        }

        private string objParent;

        public string ObjParent
        {
            get { return objParent; }
            set { objParent = value; }
        }

        public void FindExcelObj(DataTable exceldt, ObjectLibrary obj, int maxCol, List<ObjectLibrary> list, string objName, string actName)
        {
            string cell = "";
            for (int i = 0; i < exceldt.Columns.Count; i++)
            {
                if (exceldt.Rows[0][i].ToString() != "")
                {
                    maxCol = i;
                    break;
                }
            }
            string[] field = new string[maxCol];
            for (int i = 1; i < exceldt.Rows.Count; i++)
            {
                string filedName = "";
                obj = new ObjectLibrary();
                for (int j = 0; j < maxCol; j++)
                {
                    cell = exceldt.Rows[i][j].ToString();
                    if (cell != "")
                    {
                        field[j] = cell;
                        obj.ObjectName = cell;
                        if (filedName.LastIndexOf("_") > 0)
                        {
                            filedName = filedName.Substring(0, filedName.LastIndexOf('_'));
                        }
                        obj.ObjectParentPath = filedName;
                        list.Add(obj);
                        break;
                    }
                    filedName += field[j] + "_";
                }
            }
            for (int i = 0; i < list.Count; i++)
            {
                obj = list[i];
                if (list[i].ObjectName == actName && list[i].ObjectParentPath.Contains(objName))
                {
                    Common.objlist.Add(list[i]);
                }
            }
        }
    }
}
