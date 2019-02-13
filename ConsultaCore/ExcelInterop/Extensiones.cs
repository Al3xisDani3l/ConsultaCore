using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Data;
using System.Reflection;
using System.Collections;

namespace ConsultaCore
{
    public static class Extensiones
    {
        /// <summary>
        /// Exporta un DataTable a un archivo de  Excel.
        /// </summary>
        /// <param name="DataTable">DataTable</param>
        /// <param name="ExcelFilePath">Direccion y nombre del archivo de Excel</param>
        public static void ExportToExcel(this System.Data.DataTable DataTable, string ExcelFilePath = null)
        {
            try
            {
                int ColumnsCount;

                if (DataTable == null || (ColumnsCount = DataTable.Columns.Count) == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbooks.Add();

                // single worksheet
                Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;

                object[] Header = new object[ColumnsCount];

                // column headings               
                for (int i = 0; i < ColumnsCount; i++)
                    Header[i] = DataTable.Columns[i].ColumnName;

                Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
                HeaderRange.Value = Header;
                HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                HeaderRange.Font.Bold = true;
                

                // DataCells
                int RowsCount = DataTable.Rows.Count;
                object[,] Cells = new object[RowsCount, ColumnsCount];

                for (int j = 0; j < RowsCount; j++)
                    for (int i = 0; i < ColumnsCount; i++)
                        Cells[j, i] = DataTable.Rows[j][i];

                Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;

                // check fielpath
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        Worksheet.SaveAs(ExcelFilePath);
                        Excel.Quit();
                     
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else    // no filepath is given
                {
                    Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        public static DataTable ToDataTable(this IDateable[] dateables)
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

        public static DataTable ToDataTable(this List<IDateable> dateables)
        {
            DataTable table = new DataTable();

            IDateable[] dateablesArray = dateables.ToArray();

            foreach (PropertyInfo item in dateablesArray[0].Propiedades)
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
            for (int i = 0; i < dateablesArray.Length; i++)
            {
                DataRow row = table.NewRow();
                foreach (PropertyInfo item in dateablesArray[i].Propiedades)
                {
                    ExcluibleAttribute attribute = System.Attribute.GetCustomAttribute(item, typeof(ExcluibleAttribute)) as ExcluibleAttribute;

                    if (attribute != null)
                    {
                        if (!attribute.IsExcluible)
                        {
                            row[item.Name] = item.GetValue(dateablesArray[i]);


                        }
                    }
                    else
                    {
                        row[item.Name] = item.GetValue(dateablesArray[i]);

                    }

                }
                table.Rows.Add(row);



            }

            return table;


        }

        public static void AddToDataBase(this IDateable dateable, System.Data.OleDb.OleDbConnection Conexion)
        {
            using (CoreDataBaseAccess core = new CoreDataBaseAccess(Conexion))
            {
                core.InsertarAsync(dateable);
            }
           
          
        }

        public static void AddToDataBase(this List<IDateable> List,System.Data.OleDb.OleDbConnection Conexion)
        {
            using (CoreDataBaseAccess core = new CoreDataBaseAccess(Conexion))
            {
                core.InsertarAsync(List.ToArray());
            }
          
           
        }

       



    }

   
}
