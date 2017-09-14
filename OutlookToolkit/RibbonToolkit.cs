using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;
using OutlookToolkit.Properties;
using System;
using System.Windows.Forms;

namespace OutlookToolkit
{
    public partial class RibbonToolkit
    {
        private void RibbonToolkit_Load(object sender, RibbonUIEventArgs e)
        {
             
        }

        private void btn_save_Click(object sender, RibbonControlEventArgs e)
        {
            string path = this.getPath();
            bool flag = path != "";
            if (flag)
            {
                Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                foreach (object current in explorer.Selection)
                {
                    bool flag2 = current is MailItem;
                    if (flag2)
                    {
                        MailItem mailItem = current as MailItem;
                        foreach (Attachment attachment in mailItem.Attachments)
                        {
                            string fileName = attachment.FileName;
                            bool flag3 = !fileName.ToLower().Contains("jpg") && !fileName.ToLower().Contains("png") && !fileName.ToLower().Contains("gif");
                            if (flag3)
                            {
                                attachment.SaveAsFile(path + "\\" + attachment.FileName);
                            }
                        }
                    }
                }
                MessageBox.Show("Save attachments finished!", "Result:", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        public string getPath()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Select the folder:";
            folderBrowserDialog.SelectedPath = Settings.Default.Folder_Path;
            string result = "";
            bool flag = folderBrowserDialog.ShowDialog() == DialogResult.OK;
            if (flag)
            {
                Settings.Default.Folder_Path = folderBrowserDialog.SelectedPath;
                Settings.Default.Save();
                result = folderBrowserDialog.SelectedPath;
            }
            return result;
        }

        private void btn_rply_Click(object sender, RibbonControlEventArgs e)
        {
            MailItem currentItem = this.GetCurrentItem();
            bool flag = currentItem != null;
            if (flag)
            {
                MailItem mailItem = currentItem.Reply();
                this.CopyAttachments(currentItem, mailItem);
                mailItem.Display(Type.Missing);
            }
        }

        private void btn_rplyAll_Click(object sender, RibbonControlEventArgs e)
        {
            MailItem currentItem = this.GetCurrentItem();
            bool flag = currentItem != null;
            if (flag)
            {
                MailItem mailItem = currentItem.ReplyAll();
                this.CopyAttachments(currentItem, mailItem);
                mailItem.Display(Type.Missing);
            }
        }

        public void CopyAttachments(MailItem objOldMail, MailItem objNewMail)
        {
            string str = SpecialDirectories.Temp + "\\";
            foreach (Attachment attachment in objOldMail.Attachments)
            {
                string fileName = attachment.FileName;
                bool flag = !fileName.ToLower().Contains("jpg") && !fileName.ToLower().Contains("png") && !fileName.ToLower().Contains("gif");
                if (flag)
                {
                    string text = str + attachment.FileName;
                    attachment.SaveAsFile(text);
                    objNewMail.Attachments.Add(text, Type.Missing, Type.Missing, attachment.FileName);
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(text);
                }
            }
        }

        public MailItem GetCurrentItem()
        {
            object varName = Globals.ThisAddIn.Application.ActiveWindow();
            string objName = Information.TypeName(varName);
            MailItem result;
            if (objName == "Inspector")
            {
                result = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
            }
            else if(objName == "Explorer")
            {
                result = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
            }
            else
            {
                result = null;
            }

            return result;
        }
    }
}
