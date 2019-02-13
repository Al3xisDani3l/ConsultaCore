using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

	

namespace ConsultaCore.ExcelInterop
{
   public class InteropExcelDataBase
    {
        /// <summary>
        /// Convierte una coleccion de elementos IDateable a un DataTable generico.
        /// </summary>
        /// <param name="dateables">Coleccion de elementos IDatateables</param>
        /// <returns>DataTable rellenado con las propiedades propias del objeto</returns>
        public static DataTable ConvertIDateableToDataTable(IDateable[] dateables)
        {
            DataTable table = new DataTable();

            foreach (PropertyInfo item in dateables[0].Propiedades)
            {
                var attrib = System.Attribute.GetCustomAttribute(item, typeof(ExcluibleAttribute)) as ExcluibleAttribute;

                if (attrib != null)
                {
                    if (!attrib.IsExcluible)
                    {
                        table.Columns.Add(item.Name);
                    }
                   
                }
                else
                {
                    table.Columns.Add(item.Name);
                }



            }
            for (int i = 0; i < dateables.Length; i++)
            {
                DataRow row = table.NewRow();
                foreach (PropertyInfo item in dateables[i].Propiedades)
                {
                    ExcluibleAttribute attribute = System.Attribute.GetCustomAttribute(item, typeof(ExcluibleAttribute)) as ExcluibleAttribute;
                  
                    if (attribute != null)
                    {
                        if (!attribute.IsExcluible)
                        {
                            row[item.Name] = item.GetValue(dateables[i]);
                        

                        }
                    }
                    else
                    {
                        row[item.Name] = item.GetValue(dateables[i]);
                       
                    }
                   
                }
                table.Rows.Add(row);



            }

            return table;


        }

        public static void GenerarHojaExcel(IDateable[] dateables,string path)
        {
            DataTable table = ConvertIDateableToDataTable(dateables);
        table.ExportToExcel(path);

        }


    }
}
