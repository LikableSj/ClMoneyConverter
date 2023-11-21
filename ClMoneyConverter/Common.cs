using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Business.Library
{
    public class Common
    {
        #region [Method] Task 중복 실행여부 확인
        static public bool IsTaskView()
        {
            var thisID = Process.GetCurrentProcess().Id;

            Process[] p = Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName);
            if (p.Length > 1)
            {
                for (int i = 0; i < p.Length; i++)
                {
                    if (p[i].Id == thisID)
                    {
                        Console.WriteLine("Task 중복 실행!!");
                        return true;
                    }
                }
            }
            return false;
        }
        #endregion




        #region ConvertDataTable
        public static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        #endregion


        #region GetItem
        public static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    var ProDescription = Attribute.IsDefined(pro, typeof(DescriptionAttribute)) ? (Attribute.GetCustomAttribute(pro, typeof(DescriptionAttribute)) as DescriptionAttribute).Description : pro.Name;

                    if (ProDescription == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    else
                        continue;
                }
            }
            return obj;
        }
        #endregion
    }
}
