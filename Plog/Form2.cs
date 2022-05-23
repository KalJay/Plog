using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Plog
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            Global.f2 = this;
            textBox1.Text = Properties.Settings.Default.ExcelFile;
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }

        private void button1_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel Spreadsheet Files(*.xlsx)|*.xlsx|Excel Spreadsheet Macro files(*.xlsm)|*.xlsm|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    Properties.Settings.Default.ExcelFile = openFileDialog.FileName;
                    textBox1.Text = openFileDialog.FileName;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Global.f1.ResumeLayout();
            Global.f2 = null;
            Close();
        }
    }
}
