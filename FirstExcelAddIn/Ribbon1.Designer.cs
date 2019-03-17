namespace FirstExcelAddIn {
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.gp_sap = this.Factory.CreateRibbonGroup();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.gp_Editing = this.Factory.CreateRibbonGroup();
            this.gp_Confirmation = this.Factory.CreateRibbonGroup();
            this.btn_create_project = this.Factory.CreateRibbonButton();
            this.btn_confirm = this.Factory.CreateRibbonButton();
            this.btn_cancel = this.Factory.CreateRibbonButton();
            this.btn_upload = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btn_save = this.Factory.CreateRibbonButton();
            this.gp_sap.SuspendLayout();
            this.tab1.SuspendLayout();
            this.gp_Editing.SuspendLayout();
            this.gp_Confirmation.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // gp_sap
            // 
            this.gp_sap.Items.Add(this.btn_upload);
            this.gp_sap.Label = "SAP";
            this.gp_sap.Name = "gp_sap";
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.gp_Editing);
            this.tab1.Groups.Add(this.gp_Confirmation);
            this.tab1.Groups.Add(this.gp_sap);
            this.tab1.Label = "First VSTO";
            this.tab1.Name = "tab1";
            // 
            // gp_Editing
            // 
            this.gp_Editing.Items.Add(this.btn_save);
            this.gp_Editing.Items.Add(this.btn_create_project);
            this.gp_Editing.Label = "Editing";
            this.gp_Editing.Name = "gp_Editing";
            // 
            // gp_Confirmation
            // 
            this.gp_Confirmation.Items.Add(this.btn_confirm);
            this.gp_Confirmation.Items.Add(this.btn_cancel);
            this.gp_Confirmation.Label = "Confirmation";
            this.gp_Confirmation.Name = "gp_Confirmation";
            // 
            // btn_create_project
            // 
            this.btn_create_project.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_create_project.Label = "Create";
            this.btn_create_project.Name = "btn_create_project";
            this.btn_create_project.ShowImage = true;
            this.btn_create_project.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_create_project_Click);
            // 
            // btn_confirm
            // 
            this.btn_confirm.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_confirm.Label = "Confirm";
            this.btn_confirm.Name = "btn_confirm";
            this.btn_confirm.ShowImage = true;
            this.btn_confirm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_confirm_Click);
            // 
            // btn_cancel
            // 
            this.btn_cancel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_cancel.Label = "Cancel";
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.ShowImage = true;
            this.btn_cancel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_cancel_Click);
            // 
            // btn_upload
            // 
            this.btn_upload.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_upload.Label = "Upload";
            this.btn_upload.Name = "btn_upload";
            this.btn_upload.ShowImage = true;
            this.btn_upload.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_upload_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Connect";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Back";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // btn_save
            // 
            this.btn_save.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_save.Label = "Save";
            this.btn_save.Name = "btn_save";
            this.btn_save.ShowImage = true;
            this.btn_save.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_save_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.gp_sap.ResumeLayout(false);
            this.gp_sap.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.gp_Editing.ResumeLayout(false);
            this.gp_Editing.PerformLayout();
            this.gp_Confirmation.ResumeLayout(false);
            this.gp_Confirmation.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gp_Editing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_create_project;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gp_Confirmation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_confirm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_cancel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_upload;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gp_sap;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_save;
    }

    partial class ThisRibbonCollection {
        internal Ribbon1 Ribbon1 {
            get {
                return this.GetRibbon<Ribbon1>();
            }
        }
    }
}
