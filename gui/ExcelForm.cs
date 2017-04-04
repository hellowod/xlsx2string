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
    public partial class ExcelForm : Form
    {
        public ExcelForm()
        {
            //AppDomain.CurrentDomain.AssemblyResolve += (sender, args) => {
            //    string resourceName = new AssemblyName(args.Name).Name + ".dll";
            //    string resource = Array.Find(this.GetType().Assembly.GetManifestResourceNames(), element => element.EndsWith(resourceName));

            //    using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource)) {
            //        Byte[] assemblyData = new Byte[stream.Length];
            //        stream.Read(assemblyData, 0, assemblyData.Length);
            //        return Assembly.Load(assemblyData);
            //    }
            //};
            InitializeComponent();
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

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();
            DialogResult dialog = browser.ShowDialog();
            if (dialog != DialogResult.OK) {
                return;
            }
            if (browser.SelectedPath.Length < 4) {
                return;
            }
            textBox3.Text = browser.SelectedPath;
            options.ExcelPath = textBox3.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();
            DialogResult dialog = browser.ShowDialog();
            if (dialog != DialogResult.OK) {
                return;
            }
            if (browser.SelectedPath.Length < 4) {
                return;
            }
            textBox2.Text = browser.SelectedPath;
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

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
