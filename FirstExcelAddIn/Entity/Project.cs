using FirstExcelAddIn.Utility;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace FirstExcelAddIn {
    public class Project {

        #region Fields

        public string ProjectNo {
            get; set;
        }
        public dynamic ProjectNoRefersToLocal {
            get; set;
        }

        public string ProjectName {
            get; set;
        }
        public dynamic ProjectNameRefersToLocal {
            get; set;
        }

        public List<WBSElement> WBSElements {
            get; set;
        }

        public System.Data.DataTable DataTable_WBSElements {
            get {
                return ExcelTableHelper.ToDataTable<WBSElement>(this.WBSElements);
            }
        }

        #endregion

        #region Constructor

        public Project() {
            WBSElements = new List<WBSElement>();
        }

        #endregion

        #region Methods

        public void SetWBSElements(System.Data.DataTable dt) {
            WBSElements.Clear();

            for (int i = 0; i < dt.Rows.Count; i++) {
                WBSElement wbsElement = new WBSElement();

                wbsElement.AssetClass = dt.Rows[i][TableNameConstant.AssetClass].ToString();
                wbsElement.Location = dt.Rows[i][TableNameConstant.Location].ToString();
                wbsElement.Job = dt.Rows[i][TableNameConstant.Job].ToString();

                WBSElements.Add(wbsElement);
            }
        }

        public List<WBSElement> GetWBSElementFromExcel(string excelFileFullPath, string workSheetName, string tableName) {
            OleDbConnection oledbConn = new OleDbConnection(
                string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=""Excel 12.0;HDR=No;IMEX=1""", excelFileFullPath));

            oledbConn.Open();

            System.Data.DataTable dt = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null);

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
            DataSet dataSet = new DataSet();

            cmd.Connection = oledbConn;
            cmd.CommandText = string.Format(@"SELECT * FROM [{0}]", tableName);
            dataAdapter = new OleDbDataAdapter(cmd);
            dataAdapter.Fill(dataSet);

            List<WBSElement> wbsElements = dataSet.Tables[0].AsEnumerable().Select(s => new WBSElement {
                AssetClass = Convert.ToString(s[TableNameConstant.AssetClass] != DBNull.Value ? s[TableNameConstant.AssetClass] : ""),
                Location = Convert.ToString(s[TableNameConstant.Location] != DBNull.Value ? s[TableNameConstant.Location] : ""),
                Job = Convert.ToString(s[TableNameConstant.Job] != DBNull.Value ? s[TableNameConstant.Job] : "")
            }).ToList();

            oledbConn.Close();

            return wbsElements;
        }

        #endregion

        #region Validations

        public bool IsValid() {
            return false;
        }

        //RuleViolation...

        #endregion
    }
}
