using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace workerwages
{
    public partial class workwages
    {
        public string path = "";
        private void workwages_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void findfile_Click(object sender, RibbonControlEventArgs e)
        {
            if (path.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("请选择文件夹路径");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("您已选择了文件夹路径" + "\r\n" + path);
                this.compute_wages(path);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            this.folderBrowserDialog1.ShowDialog();
            this.editBox1.Text = this.folderBrowserDialog1.SelectedPath;
            path = this.folderBrowserDialog1.SelectedPath;
        }

        private void compute_wages(string path1)
        {
            System.Windows.Forms.MessageBox.Show("正在计算");
            System.IO.DirectoryInfo folder = new System.IO.DirectoryInfo(path);

            foreach ( System.IO.FileInfo file in folder.GetFiles("*.*"))
            {
                System.Windows.Forms.MessageBox.Show(file.ToString());
            }



        }

        private void editBox1_TextChanged_1(object sender, RibbonControlEventArgs e)
        {
            
        }
    }
}