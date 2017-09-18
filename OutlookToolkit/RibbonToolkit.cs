using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;
using OutlookToolkit.Properties;
using System;
using System.IO;
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
            if (path != "")
            {
                Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                int qtyOfAtt = 0;
                foreach (object current in explorer.Selection)
                {
                    if (current is MailItem)
                    {
                        MailItem mailItem = current as MailItem;
                        foreach (Attachment attachment in mailItem.Attachments)
                        {
                            string fileName = attachment.FileName;
                            string fullPath = Path.Combine(path, fileName);
                            string ext = Path.GetExtension(fileName);
                            bool flag3 = Settings.Default.Enable_Rule && Settings.Default.Ignore_Rule.ToLower().Contains(ext.Substring(1));
                            if (!flag3)
                            {
                                qtyOfAtt++;
                                if (File.Exists(fullPath))
                                {
                                    string newFileName = "Copy of " + fileName;
                                    var existSelection = MessageBox.Show("The file name '" + fileName + "' is already existing in the folder, please select whether to KEEP BOTH or REPLACE."
                                            , "Save Attachments", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (existSelection == DialogResult.Yes)
                                    {
                                        while (File.Exists(path + "\\" + newFileName))
                                        {
                                            newFileName = "Copy of " + newFileName;
                                        }
                                        attachment.SaveAsFile(path + "\\" + newFileName);
                                    }
                                    else
                                    {
                                        File.Delete(fullPath);
                                        attachment.SaveAsFile(fullPath);
                                    }
                                }
                                else
                                {
                                    attachment.SaveAsFile(fullPath);
                                }
                            }
                        }
                    }
                }
                if (qtyOfAtt > 0)
                {
                    MessageBox.Show("Save attachments finished, Total " + qtyOfAtt + " files!", "Save Attachments", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    MessageBox.Show("No attachment saved!", "Save Attachments", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                
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
                string ext = Path.GetExtension(fileName);
                bool flag = Settings.Default.Enable_Rule && Settings.Default.Ignore_Rule.ToLower().Contains(ext.Substring(1));
                if (!flag)
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

        private void btn_option_Click(object sender, RibbonControlEventArgs e)
        {
            FormatIgnoreSetup fmIgnore = new FormatIgnoreSetup();           
            fmIgnore.ShowDialog();
        }
    }
}
