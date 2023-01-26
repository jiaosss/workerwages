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
        public int number_excel = 1; //分表序号
        public int number_raw = 1;   //汇总表行号
        public int number_column = 1;
        public int number_newinformation = 1; //信息表新增序号
        public int hn = 2;//合并表格定位

        private void workwages_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void findfile_Click(object sender, RibbonControlEventArgs e) //计算按钮
        {
            if (path == string.Empty)
            {
                MessageBox.Show("请选择文件夹路径");
            }
            else if (path2 == string.Empty)
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


            //提前处理信息表
            Excel.Application xlinformation = new Excel.Application();
            Excel.Workbook inbook = xlinformation.Workbooks.Open(path2);
            Excel.Worksheet insheet1 = inbook.Sheets[1];
            Excel.Worksheet insheet2 = inbook.Sheets[2];
            Excel.Range inrng = xlinformation.Range["A65535"].End[Excel.XlDirection.xlUp];

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
                    //复制表头

                    wsheet.Range["A1:M1"].NumberFormat = "@";   //设置格式
                    wsheet.Range["B1"].Value2 = xlworksheet.Range["B1"].Value2;     //复制表头
                    wsheet.Range["A1"].Value2 = xlworksheet.Range["C1"].Value2;
                    wsheet.Range["C1:H1"].Value2 = xlworksheet.Range["D1:I1"].Value2;
                    wsheet.Range["J1"].Value2 = "银行账户核对";
                    wsheet.Range["L1"].Value2 = "考勤表核对";

                    number_excel = number_excel + 1;
                    number_raw = number_raw + 1;


                    //复制第一个表的数据

                    wsheet.Range["A" + number_raw.ToString() + ":M" + (rng.Row + number_raw - 1).ToString()].NumberFormat = "@";
                    wsheet.Range["A" + number_raw.ToString()].Value2 = file.ToString();   //复制文件名

                    wsheet.Range["A" + number_raw.ToString() + ":H" + number_raw.ToString()].Merge();

                    //复制数据
                    wsheet.Range["B" + (number_raw + 1).ToString() + ":B" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["B2:B" + rng.Row.ToString()].Value2;     //复制表头
                    wsheet.Range["A" + (number_raw + 1).ToString() + ":A" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["C2:C" + rng.Row.ToString()].Value2;
                    wsheet.Range["C" + (number_raw + 1).ToString() + ":H" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["D2:I" + rng.Row.ToString()].Value2;

                    //验证数据
                    for (int i = 1; i < rng.Row; i++)
                    {
                        string j = Convert.ToString((wsheet.Range["H" + (number_raw + i).ToString()].Value2));
                        Excel.Range ranfind = insheet1.Range["H:H"].Find(j);

                        if (ranfind != null)
                        {
                            wsheet.Range["J" + (number_raw + i).ToString()].Value2 = wsheet.Range["B" + (number_raw + i).ToString()].Value2;
                        }
                        else
                        {
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":I" + (inrng.Row + number_newinformation).ToString()].NumberFormat = "@";
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":H" + (inrng.Row + number_newinformation).ToString()].Value2 = wsheet.Range["A" + (number_raw + i).ToString() + ":H" + (number_raw + i).ToString()].Value2;
                            insheet1.Range["I" + (inrng.Row + number_newinformation).ToString()].Value2 = (DateTime.Now.ToString("yyyy.MM.dd") + "新增");
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":I" + (inrng.Row + number_newinformation).ToString()].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //单元格横向居中
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":I" + (inrng.Row + number_newinformation).ToString()].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;   //单元格竖向居中
                            wsheet.Range["J" + (number_raw + i)].Value2 = "新增";
                            number_newinformation = number_newinformation + 1;
                        }

                    }

                    number_raw = number_raw + rng.Row + 1;
                }
                else
                {
                    wsheet.Range["A" + number_raw.ToString() + ":M" + (rng.Row + number_raw - 1).ToString()].NumberFormat = "@";
                    wsheet.Range["A" + number_raw.ToString()].Value2 = file.ToString();   //复制文件名

                    wsheet.Range["A" + number_raw.ToString() + ":H" + number_raw.ToString()].Merge();  //合并单元格

                    //复制数据
                    wsheet.Range["B" + (number_raw + 1).ToString() + ":B" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["B2:B" + rng.Row.ToString()].Value2;     //复制表头
                    wsheet.Range["A" + (number_raw + 1).ToString() + ":A" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["C2:C" + rng.Row.ToString()].Value2;
                    wsheet.Range["C" + (number_raw + 1).ToString() + ":H" + (number_raw + rng.Row).ToString()].Value2 = xlworksheet.Range["D2:I" + rng.Row.ToString()].Value2;

                    //验证数据
                    for (int i = 1; i < rng.Row; i++)
                    {
                        string j = Convert.ToString((wsheet.Range["H" + (number_raw + i).ToString()].Value2));
                        Excel.Range ranfind = insheet1.Range["H:H"].Find(j);

                        if (ranfind != null)
                        {
                            wsheet.Range["J" + (number_raw + i).ToString()].Value2 = wsheet.Range["B" + (number_raw + i).ToString()].Value2;
                        }
                        else
                        {
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":I" + (inrng.Row + number_newinformation).ToString()].NumberFormat = "@";
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":H" + (inrng.Row + number_newinformation).ToString()].Value2 = wsheet.Range["A" + (number_raw + i).ToString() + ":H" + (number_raw + i).ToString()].Value2;
                            insheet1.Range["I" + (inrng.Row + number_newinformation).ToString()].Value2 = (DateTime.Now.ToString("yyyy.MM.dd") + "新增");
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":I" + (inrng.Row + number_newinformation).ToString()].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //单元格横向居中
                            insheet1.Range["A" + (inrng.Row + number_newinformation).ToString() + ":I" + (inrng.Row + number_newinformation).ToString()].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;   //单元格竖向居中
                            wsheet.Range["J" + (number_raw + i)].Value2 = "新增";
                            number_newinformation = number_newinformation + 1;
                        }

                    }

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
            new_xlapp.ActiveWorkbook.SaveAs(path + "\\汇总表" + DateTime.Now.ToString("yyyy.MM.dd") + ".xlsx"); //保存汇总表
            xlinformation.ActiveWorkbook.SaveAs(path + "\\信息表" + DateTime.Now.ToString("yyyy.MM.dd") + ".xlsx");


            //关闭当前工作簿
            new_xlapp.ActiveWorkbook.Close(false);
            xlinformation.ActiveWorkbook.Close(false);

            //杀掉当前进程
            PublicMethod.Kill(new_xlapp);
            PublicMethod.Kill(xlinformation);

        }

        //拆分表格
        private void splitexcel_Click(object sender, RibbonControlEventArgs e)
        {

            //获取当前excel文件
            Excel.Application new_xlapp = Globals.ThisAddIn.Application;
            Excel.Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;  //当前活动workbook
            Excel.Worksheet wsheet = (Excel.Worksheet)wbook.ActiveSheet;          //当前活动sheet


            string sc = JudgmentExcelColumn(wsheet);//判断最后非空列
            int hc = JudgmentExcelRaw(wsheet);   //判断最后非空行(数值)


            //删除A列重复值，做为拆分依据
            string[] Adata = new string[hc];
            for (int h = 0; h < hc; h++)
            {
                Adata[h] = Convert.ToString(wsheet.Range["A" + (h + 1).ToString()].Value2);
            }
            string[] AdataNoDvalues = Adata.Distinct().ToArray(); //删除重复值


            //开始拆分
            MessageBox.Show("请选择保存拆分表格路径");
            this.folderBrowserDialog1.ShowDialog();
            path = this.folderBrowserDialog1.SelectedPath;
            for (int i = 1; i < AdataNoDvalues.Length; i++)
            {
                Excel.Application splitexcel = new Excel.Application();
                Excel.Workbook inbook = splitexcel.Workbooks.Add();
                Excel.Worksheet insheet1 = inbook.Sheets[1];
                insheet1.Range["A1:" + sc + "1"].Value2 = wsheet.Range["A1:" + sc + "1"].Value2;
                int n = 2;
                for (int j = 1; j < hc; j++)
                {

                    if (AdataNoDvalues[i] == Convert.ToString(wsheet.Range["A" + (j + 1).ToString()].Value2))
                    {
                        insheet1.Range["A" + n.ToString() + ":" + sc + n.ToString()].NumberFormat = "@";
                        insheet1.Range["A" + n.ToString() + ":" + sc + n.ToString()].Value2 = wsheet.Range["A" + (j + 1).ToString() + ":" + sc + (j + 1).ToString()].Value2;
                        n = n + 1;
                    }
                }

                //设置自动列宽
                insheet1.Columns.EntireColumn.AutoFit();

                //单元格居中
                insheet1.Range["A1:" + sc + hc.ToString()].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //单元格横向居中
                insheet1.Range["A1:" + sc + hc.ToString()].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;   //单元格竖向居中

                //保存拆分表
                splitexcel.ActiveWorkbook.SaveAs(path + "\\" + AdataNoDvalues[i] + ".xlsx"); //保存汇总表

                //关闭当前工作簿
                splitexcel.ActiveWorkbook.Close(false);

                //杀掉当前进程
                PublicMethod.Kill(splitexcel);

            }
            MessageBox.Show("拆分完成");

        }


        public int JudgmentExcelRaw(Excel.Worksheet fristsheet)  //判断最后非空行
        {
            Excel.Worksheet FristSheet = fristsheet;
            int RowsCount = FristSheet.UsedRange.Cells.Rows.Count;
            return RowsCount;
        }




        public string JudgmentExcelColumn(Excel.Worksheet fristsheet)  //判断最后非空列
        {
            string[] ColumnName = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            int SecondColumnCount = 676;
            string[] ColumnNames = new string[SecondColumnCount];

            for (int i = 0; i < ColumnName.Length; i++)
            {
                ColumnNames[i] = ColumnName[i];
            }

            int n = 26;
            for (int i = 0; i < ColumnName.Length; i++)
            {
                for (int j = 0; j < ColumnName.Length; j++)
                {
                    if (n + j < SecondColumnCount)
                    {
                        ColumnNames[n] = ColumnName[i] + ColumnName[j];
                        n = n + 1;
                    }
                    else
                    {
                        break;
                    }
                }
                if (ColumnNames[SecondColumnCount - 1] != null)
                {
                    break;
                }

            }


            //指定要操作的Sheet
            Excel.Worksheet FristSheet = fristsheet;

            //获取该张表的总行数
            //int RowsCount = FristSheet.UsedRange.Cells.Rows.Count;
            //获取该的总列数
            int ColumnCount = FristSheet.UsedRange.Cells.Columns.Count;

            //根据所获取的最大列数,获取到该列号名称
            //string sc = ColumnNames[ColumnCount - 1].ToString() + RowsCount.ToString();


            return ColumnNames[ColumnCount - 1].ToString();
        }

        //合并表格
        private void mergeexcel_Click(object sender, RibbonControlEventArgs e)
        {
            //选择合并表格路径
            MessageBox.Show("请选择合并表格路径");
            this.folderBrowserDialog1.ShowDialog();
            path = this.folderBrowserDialog1.SelectedPath;

            //获取当前excel文件
            Excel.Application new_xlapp = Globals.ThisAddIn.Application;
            Excel.Workbook wbook = Globals.ThisAddIn.Application.ActiveWorkbook;  //当前活动workbook
            Excel.Worksheet wsheet = (Excel.Worksheet)wbook.ActiveSheet;          //当前活动sheet


            //遍历路径下文件
            System.IO.DirectoryInfo folder = new System.IO.DirectoryInfo(path);
            foreach (System.IO.FileInfo file in folder.GetFiles("*.*")) //遍历文件夹excel文件
            {
                //当前路径
                path1 = folder.ToString() + "\\" + file.ToString();

                //打开指定路径excel文件
                Excel.Application xlapp = new Excel.Application();
                Excel.Workbook xlworkbook = xlapp.Workbooks.Open(path1);
                Excel.Worksheet xlworksheet = xlworkbook.Sheets[1];

                //判断最后非空单元格
                Excel.Range rng = xlapp.Range["A65535"].End[Excel.XlDirection.xlUp];
                string sc = JudgmentExcelColumn(xlworksheet);//判断最后非空列
                int hr = rng.Row;
                //开始合并

                //复制表头
                wsheet.Range["A1:" + sc + "1"].NumberFormat = "@";
                wsheet.Range["A1:" + sc + "1"].Value2 = xlworksheet.Range["A1:" + sc + "1"].Value2;
                //复制内容
                wsheet.Range["A" + hn.ToString() + ":" + sc + (hn + hr - 2).ToString()].NumberFormat = "@";
                wsheet.Range["A" + hn.ToString() + ":" + sc + (hn + hr - 2).ToString()].Value2 = xlworksheet.Range["A2:" + sc + hr.ToString()].Value2;
                hn = (hn + hr - 1);
                                
                //关闭当前工作簿
                xlapp.ActiveWorkbook.Close(false);

                //杀掉当前进程
                PublicMethod.Kill(xlapp);
                
            }
            //设置自动列宽
            wsheet.Columns.EntireColumn.AutoFit();
            

            MessageBox.Show("合并完成");
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