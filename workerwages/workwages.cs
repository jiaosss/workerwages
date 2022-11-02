using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;

namespace workerwages
{
    public partial class workwages
    {
        public string path = "";
        private void workwages_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void findfile_Click(object sender, RibbonControlEventArgs e) //计算按钮
        {
            if (path.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("请选择文件夹路径");
            }
            else
            {
                this.compute_wages(path); //开始汇总表
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            this.folderBrowserDialog1.ShowDialog();
            path = this.folderBrowserDialog1.SelectedPath;
            System.Windows.Forms.MessageBox.Show("您已选择了文件夹路径" + "\r\n" + path);
        }

        private void compute_wages(string path1) //开始汇总表
        {
            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
            //excelapp.Visible = false;
            System.Windows.Forms.MessageBox.Show("正在计算");
            //System.IO.DirectoryInfo folder = new System.IO.DirectoryInfo(path); //获取文件夹地址

            //foreach (System.IO.FileInfo file in folder.GetFiles("*.*")) //遍历文件夹excel文件
            //{
            //    System.Windows.Forms.MessageBox.Show(file.ToString());

            //}

            //string file = "C: /Users/jiaos/Desktop/工人工资/已审/合正龙腾付款单 - 锴成建筑（非建行部分）-已审.xlsx";
            excelapp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;
            //Microsoft.Office.Interop.Excel.Range rg;
            (excelapp.Range["B2:C5"]).Select();


        }

        private void editBox1_TextChanged_1(object sender, RibbonControlEventArgs e)
        {
            
        }
    }
}