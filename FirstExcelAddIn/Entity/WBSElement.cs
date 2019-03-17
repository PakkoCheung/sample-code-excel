using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using FirstExcelAddIn.Utility;

namespace FirstExcelAddIn {
    public class WBSElement {

        [DisplayName(TableNameConstant.AssetClass)]
        public string AssetClass {
            get; set;
        }

        [DisplayName(TableNameConstant.Location)]
        public string Location {
            get; set;
        }

        [DisplayName(TableNameConstant.Job)]
        public string Job {
            get; set;
        }


    }
}
