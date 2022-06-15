using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Test20;

namespace Create_Excel
{
    public partial class Form1 : Form
    {
        public string[] files;
        public Form1()
        {
            InitializeComponent();
            files = Directory.GetFiles(textBox1.Text);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    files = Directory.GetFiles(fbd.SelectedPath);
                    textBox1.Text = fbd.SelectedPath;
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        public static string[] RemoveFromArray(string[] original, int numIdx)
        {
            List<string> tmp = new List<string>(original);
            tmp.RemoveAt(numIdx);
            return tmp.ToArray();
        }
    
        private void check_andelse()
        {
            string andelse = textBox2.Text;
            string[] splt;
            for (int i = files.Length-1; i >= 0; i--)
            {
                splt = files[i].Split('.');
                if (splt[splt.Length-1] == andelse)
                {
                    remove_sokvag(files[i], i);
                }
                else
                {
                    files = RemoveFromArray(files, i);
                }
            }
        }
        private void remove_sokvag(string str, int ar_len)
        {
            string[] str_arr = str.Split('\\');
            files[ar_len] = str_arr[str_arr.Length-1];
        }
        private DataTable fill_table(DataTable tab)
        {
            tab.Columns.Add("Name");
            DataRow rowObject;
            for (int i=0; i<files.Length; i++)
            {
                rowObject = tab.NewRow();
                rowObject["Name"] = files[i];
                tab.Rows.Add(rowObject);
            }
            tab.AcceptChanges();
            return tab;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            check_andelse();
            if (files.Length != 0)
            {
                DataTable tab = new DataTable();
                tab = fill_table(tab);
                ClassExcelFile mah_class = new ClassExcelFile();
                string workbk = textBox3.Text;
                mah_class.ExportToExcelFile(tab, workbk + "\\dirfiles.xlsx");
                tab = null;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBox3.Text = fbd.SelectedPath;
                }
            }
        }
    }
}
