using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace xlsx2string
{
    public partial class XlsxForm : System.Windows.Forms.Form
    {
        public XlsxForm()
        {
            this.MaximizeBox = false;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private static string OnSelectedPath()
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();
            DialogResult dialog = browser.ShowDialog();
            if (dialog != DialogResult.OK) {
                return null;
            }
            if (browser.SelectedPath.Length < 4) {
                return null;
            }
            return browser.SelectedPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox3.Text = OnSelectedPath();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox2.Text = OnSelectedPath();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            List<string> list = Facade.ProcessCore(options);
            if(list != null) {
                foreach(string str in list) {
                    listBox2.Items.Add(str);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<ExportType> list = DataMemory.GetExporterTypes();
            foreach(ExportType type in list) {
                listBox1.Items.Add(type.ToString());
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void OnCheckedChangeed(object sender, EventArgs e)
        {
            CheckBox check = sender as CheckBox;
            if (check == null) {
                return;
            }
            ExportType type = check.Text.ToEnum<ExportType>();
            if (check.Checked) {
                DataMemory.SetExporterType(type);
            } else {
                DataMemory.RemExportType(type);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChangeed(sender, e);
        }
    }
}
