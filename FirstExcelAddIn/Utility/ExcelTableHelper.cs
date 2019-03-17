using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Excel.Extensions;
using System.Reflection;
using System.Data.OleDb;
using System.Data;

namespace FirstExcelAddIn.Utility {
    public static class ExcelTableHelper {

        public static void BindCellIntoObject<T>(Workbook workbook, Object obj) {
            Names names = workbook.Names; /*Globals.ThisAddIn.Application.ActiveWorkbook.Names;*/

            PropertyInfo[] propertyInfos = obj.GetType().GetProperties();

            foreach (Name name in Globals.ThisAddIn.Application.ActiveWorkbook.Names) {
                PropertyInfo property = propertyInfos.SingleOrDefault(o => o.Name == name.NameLocal);
                PropertyInfo propertyCell = propertyInfos.SingleOrDefault(o => o.Name == name.NameLocal + "RefersToLocal");
                Range cell = Globals.ThisAddIn.Application.Range[name.RefersToLocal];

                if (property == null || propertyCell == null || cell == null) {
                    continue;
                }

                property.SetValue(obj, cell.Value2);
                propertyCell.SetValue(obj, name.RefersToLocal);
            }
        }
        public static Microsoft.Office.Tools.Excel.ListObject BindDataTableToListObject(Worksheet worksheet, string tableName, System.Data.DataTable datasource) {
            Microsoft.Office.Tools.Excel.ListObject listObject = null;

            foreach (ListObject table in worksheet.ListObjects/*Globals.ThisAddIn.Application.ActiveSheet.ListObjects*/) {
                if (table.Name == tableName) {
                    listObject = table.GetVstoObject(Globals.Factory);

                    listObject.AutoSetDataBoundColumnHeaders = true;
                    listObject.SetDataBinding(datasource);
                    break;
                }
            }

            return listObject;
        }
        public static System.Data.DataTable GetDataTableFromListObject(Worksheet worksheet, string tableName) {
            Microsoft.Office.Tools.Excel.ListObject listObject = null;

            foreach (ListObject table in worksheet.ListObjects/*Globals.ThisAddIn.Application.ActiveSheet.ListObjects*/) {
                if (table.Name == tableName) {
                    listObject = table.GetVstoObject(Globals.Factory);
                    break;
                }
            }
            return listObject.DataSource as System.Data.DataTable;
        }

        public static System.Data.DataTable ToDataTable<T>(List<T> items) {
            System.Data.DataTable dataTable = new System.Data.DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props) {
                //Defining type of data column gives proper data table 
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names
                if (prop.CustomAttributes.Any(o => o.AttributeType.Name == "DisplayNameAttribute") &&
                    prop.CustomAttributes.Any(o => o.ConstructorArguments.Count > 0)) {

                    string columnDisplayName = prop.CustomAttributes.First(o => o.AttributeType.Name == "DisplayNameAttribute").ConstructorArguments.First().Value.ToString();

                    dataTable.Columns.Add(columnDisplayName, type);

                } else {
                    dataTable.Columns.Add(prop.Name, type);
                }
            }
            foreach (T item in items) {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++) {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }


    }
}
