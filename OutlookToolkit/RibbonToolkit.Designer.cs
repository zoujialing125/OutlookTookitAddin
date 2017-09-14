using System;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookToolkit
{
    partial class RibbonToolkit : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonToolkit()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabToolkit = Factory.CreateRibbonTab();
            this.group1 = Factory.CreateRibbonGroup();
            this.btn_save = Factory.CreateRibbonButton();
            this.group2 = Factory.CreateRibbonGroup();
            this.btn_rply = Factory.CreateRibbonButton();
            this.btn_rplyAll = Factory.CreateRibbonButton();
            this.tabToolkit.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            this.tabToolkit.Groups.Add(this.group1);
            this.tabToolkit.Groups.Add(this.group2);
            this.tabToolkit.Label = "ToolKit";
            this.tabToolkit.Name = "tabToolkit";
            this.group1.Items.Add(this.btn_save);
            this.group1.Label = "Save";
            this.group1.Name = "group1";
            this.btn_save.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_save.Label = "Save Attachments";
            this.btn_save.Name = "btn_save";
            this.btn_save.OfficeImageId = "SaveAttachAs";
            this.btn_save.ShowImage = true;
            this.btn_save.SuperTip = "Save attachments to an appointed folder.";
            this.btn_save.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_save_Click);
            this.group2.Items.Add(this.btn_rply);
            this.group2.Items.Add(this.btn_rplyAll);
            this.group2.Label = "Reply With Attachments";
            this.group2.Name = "group2";
            this.btn_rply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_rply.Label = "Reply";
            this.btn_rply.Name = "btn_rply";
            this.btn_rply.OfficeImageId = "Reply";
            this.btn_rply.ShowImage = true;
            this.btn_rply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_rply_Click);
            this.btn_rplyAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_rplyAll.Label = "Reply All";
            this.btn_rplyAll.Name = "btn_rplyAll";
            this.btn_rplyAll.OfficeImageId = "ReplyAll";
            this.btn_rplyAll.ShowImage = true;
            this.btn_rplyAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_rplyAll_Click);

            this.Name = "RibbonToolkit";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabToolkit);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonToolkit_Load);
            this.tabToolkit.ResumeLayout(false);
            this.tabToolkit.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            base.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabToolkit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_save;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_rply;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_rplyAll;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonToolkit RibbonToolkit
        {
            get { return this.GetRibbon<RibbonToolkit>(); }
        }
    }
}
