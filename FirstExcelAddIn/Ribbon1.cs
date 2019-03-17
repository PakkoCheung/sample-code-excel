using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System.Reflection;
using FirstExcelAddIn.Utility;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data;

namespace FirstExcelAddIn {
    public partial class Ribbon1 {

        #region Fields

        //private ProjectFormSelectionPane _projectFormSelectionPane;
        //private CustomTaskPane _customTaskPane;

        Projects projects = new Projects();
        Project editingProject = null;

        #endregion

        #region Constructor

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {
            SwitchToViewMode();
        }

        #endregion

        #region Button Methods

        private void btn_create_project_Click(object sender, RibbonControlEventArgs e) {
            //Check Is Valid Workbook
            SwitchToEditMode();

            //_projectFormSelectionPane = new ProjectFormSelectionPane();
            //_customTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_projectFormSelectionPane, "Project Input Form Selection");
            //_customTaskPane.Visible = true;
            //_customTaskPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

            editingProject = new Project();

            //Bind editing project to worksheet cell name + table
            ExcelTableHelper.BindCellIntoObject<Project>(Globals.ThisAddIn.Application.ActiveWorkbook, editingProject);
            ExcelTableHelper.BindDataTableToListObject(Globals.ThisAddIn.Application.ActiveSheet, TableNameConstant.WBSElement, editingProject.DataTable_WBSElements);
        }

        private void btn_confirm_Click(object sender, RibbonControlEventArgs e) {
            editingProject.SetWBSElements(ExcelTableHelper.GetDataTableFromListObject(Globals.ThisAddIn.Application.ActiveSheet, TableNameConstant.WBSElement));

            //TODO, check is valid

            projects.Add(editingProject);
            editingProject = null;

            SwitchToViewMode();
        }

        private void btn_cancel_Click(object sender, RibbonControlEventArgs e) {
            //TODO, add confirmation

            projects.Remove(editingProject);
            editingProject = null;

            SwitchToViewMode();
        }

        private void btn_save_Click(object sender, RibbonControlEventArgs e) {
            //TODO, Save all projects in the memory into seperated excel files
        }

        private void btn_upload_Click(object sender, RibbonControlEventArgs e) {
            //1. No need to check is valid workbook

            //Select workbook(s) from anywhere
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Filter = "Excel Files|*.xls;*.xlsx;";

            if (fileDialog.ShowDialog() == DialogResult.OK) {

                Projects validProjects = new Projects();
                var excelApplication = new Microsoft.Office.Interop.Excel.Application();

                foreach (string fileName in fileDialog.FileNames) {
                    Workbook excelWorkbook = excelApplication.Workbooks.Open(fileName);
                    Worksheet excelWorksheet = null;

                    try {
                        excelWorksheet = excelWorkbook.Sheets[AppConstant.ProjectFormWorksheetName];
                    } catch (Exception ex) {
                        //Not an valid project form workbook
                        excelWorkbook.Close();
                        continue;
                    }

                    if (!ExcelCellHelper.HasCell(excelWorkbook, AppConstant.ProjectFormName)) {
                        //Not an valid project form workbook
                        excelWorkbook.Close();
                        continue;
                    }

                    Project project = new Project();

                    ExcelTableHelper.BindCellIntoObject<Project>(excelWorkbook, project);
                    project.GetWBSElementFromExcel(fileName, AppConstant.ProjectFormWorksheetName, TableNameConstant.WBSElement);

                    //2.1 check if project is valid
                    //2.2 check if wbs is valid 

                    validProjects.Add(project);

                    excelWorkbook.Close();
                }

                //3. TODO, RFC call (All projects + wbs are validated)

                foreach (Project validProject in validProjects) {
                    string projectName = validProject.ProjectName;
                    string projectNo = validProject.ProjectNo;

                    foreach (WBSElement wbsElement in validProject.WBSElements) {

                        string assetClass = wbsElement.AssetClass;
                        string location = wbsElement.Location;
                        string job = wbsElement.Job;
                    }
                }

                //if succeed, move the C:\Users\{user}\Documents\Capital Budget Workbook\Project\Uploaded
                //if failed, don't move

                //messagebox prompt (how many project uploaded, how many failed, and open log folder)
                //output log file (how many project uploaded, project no., or error log no.+reason)

            } else {
                return;
            }
        }

        #endregion

        #region Private Methods
        private void SwitchToViewMode() {
            this.gp_Editing.Visible = true;
            this.gp_Confirmation.Visible = false;
            this.gp_sap.Visible = true;

            Globals.ThisAddIn.Application.ActiveWorkbook.Protect("FSCM");
            Globals.ThisAddIn.Application.ActiveSheet.Protect("FSCM");
        }

        private void SwitchToEditMode() {
            this.gp_Editing.Visible = false;
            this.gp_Confirmation.Visible = true;
            this.gp_sap.Visible = false;

            Globals.ThisAddIn.Application.ActiveWorkbook.Unprotect("FSCM");
            Globals.ThisAddIn.Application.ActiveSheet.Unprotect("FSCM");
        }



        #endregion
    }
}
