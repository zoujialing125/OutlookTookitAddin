using OutlookToolkit.Properties;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookToolkit
{
    public partial class FormatIgnoreSetup : Form
    {
        public FormatIgnoreSetup()
        {
            InitializeComponent();
        }

        private void FormatIgnoreSetup_Load(object sender, EventArgs e)
        {
            //Show current igonring rule
            textBox1.Text = Settings.Default.Ignore_Rule;
            checkBox1.Checked = Settings.Default.Enable_Rule;
            textBox1.ReadOnly= !Settings.Default.Enable_Rule;

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip1.SetToolTip(this.textBox1, "Input the file name extensions. (E.g. jpg,png,gif)");
            toolTip1.SetToolTip(this.checkBox1, "Enable/Disable the ignore rule");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.ReadOnly = !checkBox1.Checked;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Settings.Default.Ignore_Rule = textBox1.Text;
            Settings.Default.Enable_Rule = checkBox1.Checked;
            Settings.Default.Save();
            this.Close();
        }

        private void FormatIgnoreSetup_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Settings.Default.Ignore_Rule = textBox1.Text;
        }
    }
}
