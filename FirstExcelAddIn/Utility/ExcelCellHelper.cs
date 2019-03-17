using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FirstExcelAddIn.Utility {
    public static class ExcelCellHelper {
        public static bool HasCell(Workbook workbook, string cellName) {
            foreach (Name name in workbook.Names) {
                if (cellName == name.NameLocal) {
                    return true;
                }
            }

            return false;
        }
        public static string GetCellValue(Workbook workbook, string cellName) {
            foreach (Name name in workbook.Names) {
                if (cellName == name.NameLocal) {
                    return name.Value.ToString();
                }
            }

            return string.Empty;

        }
    }
}
