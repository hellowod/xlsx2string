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
        private string defaultfilePath = System.Environment.CurrentDirectory;

        public XlsxForm()
        {
            this.MaximizeBox = false;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.InitializeComponent();
            this.InitializeComponentValue();
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

        private string OnSelectedPath(string desc)
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();
            browser.Description = string.Format("请选择 {0} 目录", desc);
            if (!string.IsNullOrEmpty(defaultfilePath)) {
                browser.SelectedPath = defaultfilePath;
            }

            DialogResult dialog = browser.ShowDialog(this);
            if (dialog == DialogResult.OK) {
                defaultfilePath = browser.SelectedPath;
                return defaultfilePath;
            }
            
            return string.Empty;
        }

        /// <summary>
        /// 浏览输入文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.textBox3.Text = OnSelectedPath("表格");
        }

        private void SetInputPath(string path)
        {
            this.textBox3.Text = path;
        }

        /// <summary>
        /// 浏览输出文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            this.textBox2.Text = OnSelectedPath("导出");
        }

        private void SetOutputPath(string path)
        {
            this.textBox2.Text = path;
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
            if (!button5.Enabled) {
                return;
            }

            string error = Facade.ParseCheckerUserInput();
            if(error != null) {
                ShowMessageBox(error);
                return;
            }
            
            CheckeCallbackArgv argv = new CheckeCallbackArgv();
            argv.OnRunChanged = OnRunCheckerChanged;

            Facade.BeforeCheckerOptionForm();
            this.BeforeCheckerForm();
            Facade.RunCheckerXlsx(argv);
            this.AfterCheckerForm();
            Facade.AfterCheckerOptionForm();
        }

        private void BeforeCheckerForm()
        {
            button5.Enabled = false;
            textBox1.Text = "";
            textBox1.Refresh();
        }

        private void AfterCheckerForm()
        {
            button5.Enabled = true;
        }

        private void ShowMessageBox(string message)
        {
            MessageBox.Show(
                this, 
                message,
                "信息",
                MessageBoxButtons.OK, 
                MessageBoxIcon.Information
            );
        }

        /// <summary>
        /// 导出表格按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (!button1.Enabled) {
                return;
            }

            string error = Facade.ParseExportUserInput();
            if (error != null) {
                ShowMessageBox(error);
                return;
            }

            ExprotCallbackArgv argv = new ExprotCallbackArgv();
            argv.OnProgressChanged = OnProgressChanged;
            argv.OnRunChanged = OnRunExporterChanged;

            Facade.BeforeExporterOptionForm();
            this.BeforeExporterForm();
            Facade.RunXlsxForm(argv);
            this.AfterExporterForm();
            Facade.AfterExporterOptionForm();
        }

        private void BeforeExporterForm()
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = DataMemory.GetExportTotalCount();
            textBox4.Text = "";
            textBox4.Refresh();
            button1.Enabled = false;
        }

        private void AfterExporterForm()
        {
            button1.Enabled = true;
        }

        private void OnProgressChanged(int value)
        {
            if(value > progressBar1.Maximum) {
                progressBar1.Value = progressBar1.Maximum;
                return;
            }
            if(value < progressBar1.Minimum) {
                progressBar1.Value = progressBar1.Minimum;
                return;
            }
            progressBar1.Value = value;
        }

        private void OnRunCheckerChanged(string value)
        {
            textBox1.AppendText(string.Format("{0}   {1}\n", DateTime.Now.ToString("HH:mm:ss"), value));
        }

        private void OnRunExporterChanged(string value)
        {
            textBox4.AppendText(string.Format("{0}   {1}   ok\n", DateTime.Now.ToString("HH:mm:ss"), value));
        }

        private void OnCheckedChanged(object sender, EventArgs e)
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
            //OnCheckedChanged(sender, e);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //OnCheckedChanged(sender, e);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //OnCheckedChanged(sender, e);
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChanged(sender, e);
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            //OnCheckedChanged(sender, e);
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            //OnCheckedChanged(sender, e);
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            //OnCheckedChanged(sender, e);
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            OnCheckedChanged(sender, e);
        }

        /// <summary>
        /// 进度单机事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            //OnCheckedChanged(sender, e);
        }

        private void InitializeComponentValue()
        {
            OptionsForm option = DataMemory.ReviewOptionForm();
            if(option == null) {
                return;
            }
            SetInputPath(option.XlsxSrcPath);
            SetOutputPath(option.XlsxDstPath);

            foreach (ExportType type in option.ExporterList) {
                foreach (object obj in this.groupBox2.Controls) {
                    if (!(obj is CheckBox)) {
                        continue;
                    }
                    CheckBox b = obj as CheckBox;
                    ExportType t = b.Text.ToEnum<ExportType>();
                    if (type == t) {
                        b.Checked = true;
                    }
                }
            }
            DataMemory.SetOptionForm(option);
        }

        private void XlsxForm_Load(object sender, EventArgs e)
        {

        }

        private void XlsxForm_Close(object sender, EventArgs e)
        {
            DataMemory.CacheOptionFrom();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
    }
}
