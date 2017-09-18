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
            this.tabToolkit = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_save = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_rply = this.Factory.CreateRibbonButton();
            this.btn_rplyAll = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btn_option = this.Factory.CreateRibbonButton();
            this.tabToolkit.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabToolkit
            // 
            this.tabToolkit.Groups.Add(this.group1);
            this.tabToolkit.Groups.Add(this.group2);
            this.tabToolkit.Groups.Add(this.group3);
            this.tabToolkit.Label = "Attachment Extension";
            this.tabToolkit.Name = "tabToolkit";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_save);
            this.group1.Label = "Save Attachments";
            this.group1.Name = "group1";
            // 
            // btn_save
            // 
            this.btn_save.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_save.Label = "Save Selected";
            this.btn_save.Name = "btn_save";
            this.btn_save.OfficeImageId = "SaveAttachAs";
            this.btn_save.ShowImage = true;
            this.btn_save.SuperTip = "Save attachments to an appointed folder.";
            this.btn_save.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_save_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_rply);
            this.group2.Items.Add(this.btn_rplyAll);
            this.group2.Label = "Reply With Attachments";
            this.group2.Name = "group2";
            // 
            // btn_rply
            // 
            this.btn_rply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_rply.Label = "Reply";
            this.btn_rply.Name = "btn_rply";
            this.btn_rply.OfficeImageId = "Reply";
            this.btn_rply.ShowImage = true;
            this.btn_rply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_rply_Click);
            // 
            // btn_rplyAll
            // 
            this.btn_rplyAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_rplyAll.Label = "Reply All";
            this.btn_rplyAll.Name = "btn_rplyAll";
            this.btn_rplyAll.OfficeImageId = "ReplyAll";
            this.btn_rplyAll.ShowImage = true;
            this.btn_rplyAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_rplyAll_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn_option);
            this.group3.Label = "Options";
            this.group3.Name = "group3";
            // 
            // btn_option
            // 
            this.btn_option.Label = "Fomrat Ignore";
            this.btn_option.Name = "btn_option";
            this.btn_option.OfficeImageId = "MenuFilterMail";
            this.btn_option.ScreenTip = "Select to ingore which typs of attachments.";
            this.btn_option.ShowImage = true;
            this.btn_option.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_option_Click);
            // 
            // RibbonToolkit
            // 
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
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabToolkit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_save;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_rply;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_rplyAll;
        internal RibbonGroup group3;
        internal RibbonButton btn_option;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonToolkit RibbonToolkit
        {
            get { return this.GetRibbon<RibbonToolkit>(); }
        }
    }
}
