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

        /// <summary>
        /// 浏览输入文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.textBox3.Text = OnSelectedPath();
        }

        /// <summary>
        /// 浏览输出文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            this.textBox2.Text = OnSelectedPath();
        }

        /// <summary>
        /// 输入 路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            TextBox box = sender as TextBox;
            if (box == null) {
                return;
            }
            DataMemory.SetOptionFormSrcPath(box.Text);
        }

        /// <summary>
        /// 导出 路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            TextBox box = sender as TextBox;
            if (box == null) {
                return;
            }
            DataMemory.SetOptionFormDstPath(box.Text);
        }

        /// <summary>
        /// 检查表数据按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            OptionsForm form = DataMemory.GetOptionsFrom();
            listBox2.Items.Add(form.XlsxSrcPath);
        }

        /// <summary>
        /// 导出表格按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Facade.BeforeExporterOptionForm();
            Facade.RunXlsxForm();
            Facade.AfterExporterOptionForm();
        }

        /// <summary>
        /// 检查表信息输出框
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 导出表信息输出框
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                DataMemory.SetOptionFormType(type);
            } else {
                DataMemory.RemOptionFromType(type);
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

        /// <summary>
        /// 进度单机事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
