using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace workerwages
{
    public partial class workwages
    {
        public string path = "";   //分表文件夹路径
        public string path1 = "";  //分表文件路径
        public string path2 = "";  //信息表路径
        public Excel.Application excelapp;
        public int number_excel = 1;
        public int number_raw = 1;
        public int number_column = 1;
       
        private void workwages_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void findfile_Click(object sender, RibbonControlEventArgs e) //计算按钮
        {
            if (path.Length == 0)
            {
                MessageBox.Show("请选择文件夹路径");
            }
            else if (path2.Length == 0)
            {
                MessageBox.Show("请选择信息表路径");
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
            MessageBox.Show("您已选择了文件夹路径" + "\r\n" + path);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog openFd = new OpenFileDialog();
            openFd.ShowDialog();
            path2 = openFd.FileName;
            MessageBox.Show("您已选择了文件" + "\r\n" + path2);

        }






        private void compute_wages(string path) //开始汇总表
        {

            System.Windows.Forms.MessageBox.Show("点击确认开始计算" + "\r\n" + "excel窗口关闭前请勿操作电脑！！！切记！！！");
            System.IO.DirectoryInfo folder = new System.IO.DirectoryInfo(path); //获取文件夹地址

            //获取当前空白excel文件
            Excel.Application new_xlapp = Globals.ThisAddIn.Application;
            Excel.Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;  //当前活动workbook
            Excel.Worksheet wsheet = (Excel.Worksheet)wbook.ActiveSheet;          //当前活动sheet


            foreach (System.IO.FileInfo file in folder.GetFiles("*.*")) //遍历文件夹excel文件
            {
                //System.IO.FileInfo file = folder.GetFiles("*.*")[4];                                                //调试用，用完删除并解除循环注释

                //当前路径
                path1 = folder.ToString() + "\\" + file.ToString();

                

                //打开指定路径excel文件
                Excel.Application xlapp = new Excel.Application();
                Excel.Workbook xlworkbook = xlapp.Workbooks.Open(path1);
                Excel.Worksheet xlworksheet = xlworkbook.Sheets[1];

                
                //寻找第一列最后一个非空单元格
                Excel.Range rng = xlapp.Range["A65535"].End[Excel.XlDirection.xlUp];
                //MessageBox.Show("A列中最后一个非空单元格是" + rng.Address[0, 0] + ",行号" + rng.Row.ToString() + ",数值" + rng.Text);


                //复制指定文件内容至新建空白文件
                    
                //Excel.Range range_open = xlworksheet.Range["A2:I" + rng.Row.ToString()];
                
                //wsheet.Range["A1:I" + rng.Row.ToString()].NumberFormat = "@";   //设置格式
                //wsheet.Range["A1:I" + rng.Row.ToString()].Value2 = range_open.Value2;                //复制内容

                //汇总
                if (number_excel == 1)
                {
                    wsheet.Range["A1:M1"].NumberFormat = "@";   //设置格式
                    //((Excel.Range)xlWorkSheet.Cells[i, j]).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    wsheet.Range["B1"].Value2 = xlworksheet.Range["B1"].Value2;     //复制表头
                    wsheet.Range["A1"].Value2 = xlworksheet.Range["C1"].Value2;
                    wsheet.Range["C1:H1"].Value2 = xlworksheet.Range["D1:I1"].Value2;
                    wsheet.Range["J1"].Value2 = "银行账户核对";
                    wsheet.Range["L1"].Value2 = "考勤表核对";

                    number_excel = number_excel + 1;
                    number_raw = number_raw + 1;

                    wsheet.Range["A" + number_raw.ToString() + ":M" + (rng.Row + number_raw - 1).ToString()].NumberFormat = "@";
                    wsheet.Range["A" + number_raw.ToString()].Value2 = file.ToString();   //复制文件名
                
                    wsheet.Range["A" + number_raw.ToString() + ":H" + number_raw.ToString()].Merge();
                
                    //复制数据
                    wsheet.Range["B" + (number_raw + 1).ToString() + ":B" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["B2:B" + rng.Row.ToString()].Value2;     //复制表头
                    wsheet.Range["A" + (number_raw + 1).ToString() + ":A" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["C2:C" + rng.Row.ToString()].Value2;
                    wsheet.Range["C" + (number_raw + 1).ToString() + ":H" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["D2:I" + rng.Row.ToString()].Value2;

                    number_raw = number_raw + rng.Row + 1;
                }
                else
                {
                    wsheet.Range["A" + number_raw.ToString() + ":M" + (rng.Row + number_raw - 1).ToString()].NumberFormat = "@";
                    wsheet.Range["A" + number_raw.ToString()].Value2 = file.ToString();   //复制文件名
                    
                    wsheet.Range["A" + number_raw.ToString() + ":H" + number_raw.ToString()].Merge();  //合并单元格
                    
                    //复制数据
                    wsheet.Range["B" + ( number_raw + 1 ).ToString() + ":B" + ( number_raw + rng.Row ).ToString()].Value2 = xlworksheet.Range["B2:B" + rng.Row.ToString()].Value2;     //复制表头
                    wsheet.Range["A" + ( number_raw + 1 ).ToString() + ":A" + ( number_raw + rng.Row ).ToString()].Value2 = xlworksheet.Range["C2:C" + rng.Row.ToString()].Value2;
                    wsheet.Range["C" + ( number_raw + 1 ).ToString() + ":H" + ( number_raw + rng.Row ).ToString()].Value2 = xlworksheet.Range["D2:I" + rng.Row.ToString()].Value2;

                    number_raw = number_raw + rng.Row + 1;


                }



                //关闭指定excel文件
                xlapp.ActiveWorkbook.Close(false);


            }

            //全局替换#N/A
            wsheet.Range["A1:L" + number_raw.ToString()].Replace("#N/A", "");

            //设置自动列宽
            wsheet.Columns.EntireColumn.AutoFit();

            //单元格居中
            wsheet.Range["A1:L" + number_raw.ToString()].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //单元格横向居中
            wsheet.Range["A1:L" + number_raw.ToString()].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;   //单元格竖向居中

            //保存汇总表
            new_xlapp.ActiveWorkbook.SaveAs(path + "\\汇总表.xlsx");

            //关闭当前工作簿
            new_xlapp.ActiveWorkbook.Close(false);
            
            //杀掉当前进程
            PublicMethod.Kill(new_xlapp);

        }


    }


    public class PublicMethod      //杀死进程
        {
            [DllImport("User32.dll", CharSet = CharSet.Auto)]
            public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
            public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
            {
                IntPtr t = new IntPtr(excel.Hwnd);//得到这个句柄，具体作用是得到这块内存入口 

                int k = 0;
                GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
                p.Kill();     //关闭进程k
            }

        }
}