using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace FirstExcelAddIn {
    public partial class ProjectFormSelectionPane : UserControl {
        public ProjectFormSelectionPane() {
            InitializeComponent();
        }

        private void btnAdd_Click(object sender, EventArgs e) {
            foreach (Name name in Globals.ThisAddIn.Application.Names) {
                if (name.NameLocal == "WBS_Element") {
                    Range wbs = Globals.ThisAddIn.Application.Range[name.RefersToLocal];

                    Range nextRow = wbs[1, wbs.Column];
                    nextRow.Value2 = this.tbWBS.Text;

                }
            }

            this.Visible = false;
        }
    }
}
