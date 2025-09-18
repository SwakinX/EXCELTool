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
using System.Runtime.InteropServices;
using System.Threading;
using Spire.Xls;
using Spire.Xls.Collections;
using EXCELProcessing;
using Sunny.UI;
using System.Security.Cryptography;

namespace EXCELProcessing
{
    public partial class Form1 : UIForm
    {
        public Form1()
        {
            InitializeComponent();
            frm1 = this;
        }
        public static Form1 frm1;

        public string path;//单文件路径
        public string folder;//打开路径
        public string sfolder;//保存路径
        public List<string> names = new List<string>();//查分条件储存
        public Replace replace = new Replace();//替换
        public bool exten = false;//判断后缀名
        public bool sway = true;//拆分保存方式

        public Thread th;//进程
        public string delText = "";
        bool exist = false;//文件存在状态

        private void button2_Click(object sender, EventArgs e)//单文件拆分
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            if (b1 == true)
            {
                SetText("请选择保存目录\n");
                return;
            }
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel(*.xlsx,*.xls)|*.xlsx;*.xls|all|*.*";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                path = openFile.FileName.ToString();
            }
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            SetEable(false);
            th = new Thread(delegate () { singleSplit(path); });
            th.IsBackground = true;
            th.Start();
        }

        private void button1_Click(object sender, EventArgs e)//选择文件夹
        {
            FolderBrowserDialog ofolder = new FolderBrowserDialog();
            if (ofolder.ShowDialog() == DialogResult.OK)
            {
                folder = ofolder.SelectedPath.ToString();
                uiTextBox3.Text = folder;
            }
            if (string.IsNullOrEmpty(ofolder.SelectedPath))
            {
                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)//保存的文件夹
        {
            FolderBrowserDialog ofolder = new FolderBrowserDialog();
            if (ofolder.ShowDialog() == DialogResult.OK)
            {
                uiTextBox4.Text = ofolder.SelectedPath.ToString();
            }
            if (string.IsNullOrEmpty(ofolder.SelectedPath))
            {
                return;
            }

        }

        private void button4_Click(object sender, EventArgs e)//批量合并
        {
            names.Clear();
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                SetText("请选择目录\n");
            }
            else
            {
                SetEable(false);
                th = new Thread(delegate () { comBox(folder); });
                th.IsBackground = true;
                th.Start();
            }
        }

        private void button5_Click(object sender, EventArgs e)//批量拆分
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                SetText("请选择目录\n");
                return;
            }
            else
            {
                SetEable(false);
                th = new Thread(delegate () { splitBox(folder); });
                th.IsBackground = true;
                th.Start();

            }

        }

        private void splitDirector(string dir)//拆分遍历器
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    splitDirector(fsinfo.FullName);//递归调用
                }
                else
                {
                    Split(fsinfo.FullName);
                    SetText("Split success\n");
                }
            }
        }

        private void splitShhetsDirector(string dir)//sheets拆分遍历器
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    splitDirector(fsinfo.FullName);//递归调用
                }
                else
                {
                    SplitSheets(fsinfo.FullName);
                    SetText("\nSplit success\n");
                }
            }
        }

        private void comBox(string dir)
        {
            SetText("Start Combine\n");
            Director(dir);
            SetEable(true);
            SetText("\nCombine complete\n=========================================================\n");
        }
        private void comBox1(string dir)
        {
            SetText("Start Combine\n");
            Director1(dir);
            SetEable(true);
            SetText("\nCombine complete\n=========================================================\n");
        }
        private void comBox2(string[] dir)//选择多文件遍历
        {
            SetText("Start Combine\n");
            foreach (var item in dir)
            {
                calcultin1(item);
            }
            SetEable(true);
            exist = false;
            SetText("\nCombine complete\n=========================================================\n");
        }

        private void NormalBox(string dir)
        {
            SetText("Start Normalization\n");
            NormalDirector(dir);
            SetEable(true);
            SetText("\nNormalization complete\n=========================================================\n");
        }

        private void splitBox(string dir)
        {
            SetText("Start Split\n");
            splitDirector(dir);
            SetEable(true);
            SetText("\nSplit complete\n=========================================================\n");
        }

        private void splitSheetsbox(string dir)
        {
            SetText("Start Split\n");
            splitShhetsDirector(dir);
            SetEable(true);
            SetText("\nSplit complete\n=========================================================\n");
        }

        private void singleSplit(string p)
        {
            SetText("\nStart Split\n");
            Split(p);
            path = null;
            SetEable(true);
            SetText("\nSplit complete\n=========================================================\n");
        }

        private void NormalDirector(string dir)//合并遍历器
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    NormalDirector(fsinfo.FullName);//递归调用
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(delText) || fsinfo.FullName.Contains(delText))
                    {
                        DelHeader(fsinfo.FullName);
                    }
                }
            }
        }
        private void Director(string dir)//合并遍历器
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    Director(fsinfo.FullName);//递归调用
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(delText) || fsinfo.FullName.Contains(delText))
                    {
                        calcultin(fsinfo.FullName);
                    }
                }
            }
        }
        private void Director1(string dir)//合并遍历器
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    Director1(fsinfo.FullName);//递归调用
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(delText) || fsinfo.FullName.Contains(delText))
                    {
                        calcultin1(fsinfo.FullName);
                    }
                }
            }
        }

        private void calcultin(string dir)//合并计算
        {
            SetText("\n" + dir + "\t");//输出文件的全部路径

            Workbook workbook = new Workbook();//创表格实例
            workbook.LoadFromFile(dir);//打开表格
            Worksheet sheet = workbook.ActiveSheet;//提取活动表格
            int title = int.Parse(uiTextBox1.Text);//标题行
            if (sheet.LastRow <= title)
            {
                SetText("表为空");
                return;
            }
            string fname = Path.GetFileNameWithoutExtension(dir);//返回文件名
            string p = sfolder + "\\" + "合并" + "\\" + fname + Extension(dir);//保存路径
            bool bb = names.Contains(fname);

            if (bb == false)
            {
                names.Add(fname);//添加到列表

                Workbook workbooknew = new Workbook();//创表格实例
                workbooknew.Version = ExcelVersion.Version2013;
                workbooknew.Worksheets.Clear();
                Worksheet sheetnew = workbooknew.Worksheets.Add("sheet1");//选择单元格
                sheetnew.DefaultRowHeight = 18;
                sheetnew.DefaultColumnWidth = 25;
                //获取打开表格的全部数据
                CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
                //复制到新建表格
                range.Copy(sheetnew.Range[1, 1]);
                creatDirctory(sfolder);
                workbooknew.SaveToFile(p, ExcelVersion.Version2010);//保存表格 
                SetText("sucssus");
            }
            else
            {
                Workbook workbook2 = new Workbook();//创建表格实例
                workbook2.LoadFromFile(p);//打开已有表格
                Worksheet sheet1 = workbook2.ActiveSheet;
                sheet1.DefaultRowHeight = 18;
                sheet1.DefaultColumnWidth = 25;
                //获取打开表格的数据
                CellRange range = sheet.Range[title + 1, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
                range.Copy(sheet1.Range[sheet1.LastRow + 1, 1]);

                workbook2.Save();
                SetText("sucssus");
            }
        }

        private void Split(string dir)//拆分
        {
            //canStop = false;
            string fname = Path.GetFileNameWithoutExtension(dir);
            SetText(dir + "\n");

            Workbook workbook2 = new Workbook();//创建表格实例
            workbook2.LoadFromFile(dir);//打开已有表格
            Worksheet sheet2 = workbook2.ActiveSheet;

            //Dictionary<String, CellRange> dic = new Dictionary<String, CellRange>();

            int lColumn = sheet2.LastColumn;//获得最大列数
            int lRow = sheet2.LastRow;//获得最大行数
            int b = int.Parse(uiTextBox1.Text);//标题行
            int c = 0;//分割列号
            if (sheet2.LastRow <= b)
            {
                SetText("表格为空\n");
                return;
            }
            for (int i = 1; i <= lRow; i++)
            {
                string temp = sheet2.Range[1, i].Text.ToString();
                if (temp == uiTextBox2.Text)
                {
                    c = i; break;//获取分割列
                }
            }//检索分割列

            SetText("正在处理\n");
            //获取名称唯一值

            var crs = sheet2.Range[5, c, sheet2.LastRow, c].CellList;
            //遍历管道名称
            string[] arr = new string[crs.Count];
            for (int i = 0; i < crs.Count; i++)
            {
                arr[i] = crs[i].DisplayedText;
            }
            string[] new_arr = arr.GroupBy(p => p).Select(p => p.Key).ToArray();
            new_arr = new_arr.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
            SetText("共" + new_arr.Length.ToString() + "条管道\n");
            //创建筛选
            AutoFiltersCollection filters = sheet2.AutoFilters;
            filters.Range = sheet2.Range;
            //filters.Range = sheet2.Range[1, 1, sheet2.LastRow + 1, sheet2.LastColumn];
            if (new_arr.Length > 1)
            {
                foreach (string cr in new_arr)
                {
                    filters.Clear();
                    filters.AddFilter(c - 1, cr);
                    filters.Filter();

                    Workbook wb = new Workbook();//创表格实例
                    wb.Version = ExcelVersion.Version2013;
                    wb.Worksheets.Clear();
                    Worksheet sheet = wb.Worksheets.Add("sheet1");//选择单元格
                    sheet2.Range[1, 1, b, lColumn].Copy(sheet.Range[1, 1]);

                    int last = filters.Range.LastRow;


                    foreach (var row in filters.Range.Rows)
                    {
                        if (row.Row == 1)//跳过一行
                        {
                            continue;
                        }
                        bool bbb = sheet2.IsRowVisible(row.Row);//是否是可见行
                        if (bbb)
                        {
                            row.Copy(sheet.Range[sheet.LastRow + 1, 1]);
                        }
                        if (row.Row == last && cr == filters.Range[last, c].NumberText)
                        {
                            row.Copy(sheet.Range[sheet.LastRow + 1, 1]);
                        }//解决最后一行总是被隐藏的问题
                    } //筛选结果保存到表格
                    string k = cr.Trim();
                    string sPath = sfolder + "\\" + "拆分" + "\\" + k;
                    creatDirctory(sPath);
                    string p = sPath + "\\" + fname + Extension(dir);
                    sheet.DefaultRowHeight = 18;
                    sheet.DefaultColumnWidth = 25;
                    wb.SaveToFile(p, ExcelVersion.Version2013);

                    SetText("已完成：" + k + "\n");
                }
            }
            else
            {
                Workbook wb = new Workbook();//创表格实例
                wb.Version = ExcelVersion.Version2013;
                wb.Worksheets.Clear();
                Worksheet sheet = wb.Worksheets.Add("sheet1");//选择单元格
                sheet2.Range.Copy(sheet.Range[1, 1]);
                sheet.DefaultRowHeight = 18;
                sheet.DefaultColumnWidth = 25;
                string k = new_arr[0].Trim();
                string sPath = sfolder + "\\" + "拆分" + "\\" + k;
                creatDirctory(sPath);
                string p = sPath + "\\" + fname + Extension(dir);
                wb.SaveToFile(p, ExcelVersion.Version2013);
            }
        }

        private void SplitSheets(string dir)//Sheets拆分
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(dir);
            var Count = workbook.Worksheets.Count();
            SetText("\n" + dir);
            string fname = Path.GetFileNameWithoutExtension(dir);
            for (int i = 0; i < Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                SetText("\n" + sheet.Name);
                Workbook newbook = new Workbook();
                newbook.Worksheets.Clear();
                Worksheet sheet2 = newbook.Worksheets.AddCopy(sheet);
                if (sway == true)
                {
                    string sp = sfolder + "/" + "Sheets拆分1" + "/" + fname.Trim();
                    creatDirctory(sp);
                    string s = sp + "/" + sheet.Name;
                    if (exten == true)
                    {
                        sheet2.SaveToFile(s + ".csv", ",", Encoding.UTF8);
                    }
                    else
                    {
                        newbook.SaveToFile(s + ".xlsx", ExcelVersion.Version2016);
                    }
                }
                else
                {
                    string sp = sfolder + "/" + "Sheets拆分2" + "/" + sheet.Name.Trim();
                    creatDirctory(sp);
                    string s = sp + "/" + fname;
                    if (exten == true)
                    {
                        sheet2.SaveToFile(s + ".csv", ",", Encoding.UTF8);
                    }
                    else
                    {
                        newbook.SaveToFile(s + ".xlsx", ExcelVersion.Version2016);
                    }
                }
            }
        }

        private void calcultin1(string dir)//合并为单个文件
        {
            SetText("\n" + dir + "\n");//输出文件的全部路径

            Workbook workbook = new Workbook();//创表格实例
            workbook.LoadFromFile(dir);//打开表格
            Worksheet sheet = workbook.ActiveSheet;//提取活动表格
            int title = int.Parse(uiTextBox1.Text);//标题行
            if (sheet.LastRow <= title)
            {
                SetText("表为空");
                return;
            }
            string fname = "合并表格";//返回文件名
            string p = sfolder + "\\" + fname + ".xlsx";//保存路径

            if (exist == false)
            {
                exist = true;
                Workbook workbooknew = new Workbook();//创表格实例
                workbooknew.Version = ExcelVersion.Version2013;
                workbooknew.Worksheets.Clear();
                Worksheet sheetnew = workbooknew.Worksheets.Add("sheet1");//选择单元格
                sheetnew.DefaultRowHeight = 18;
                sheetnew.DefaultColumnWidth = 25;
                //获取打开表格的全部数据
                CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
                //复制到新建表格
                range.Copy(sheetnew.Range[1, 1]);
                creatDirctory(sfolder);
                workbooknew.SaveToFile(p, ExcelVersion.Version2013);//保存表格 
                SetText("sucssus");
            }
            else
            {
                Workbook workbook2 = new Workbook();//创建表格实例
                workbook2.LoadFromFile(p);//打开已有表格
                Worksheet sheet1 = workbook2.ActiveSheet;
                sheet1.DefaultRowHeight = 18;
                sheet1.DefaultColumnWidth = 25;
                //获取打开表格的数据
                CellRange range = sheet.Range[title + 1, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
                range.Copy(sheet1.Range[sheet1.LastRow + 1, 1]);

                workbook2.Save();
                SetText("sucssus");
            }
        }

        private void DelHeader(string dir)//保存为CSV
        {
            SetText("\n" + dir + "\t");//输出文件的全部路径

            Workbook workbook = new Workbook();//创表格实例
            workbook.LoadFromFile(dir);//打开表格
            Worksheet sheet = workbook.ActiveSheet;//提取活动表格
            int title = int.Parse(uiTextBox1.Text);//标题行
            if (sheet.LastRow <= title)
            {
                SetText("表为空");
                return;
            }
            string fname = Path.GetFileNameWithoutExtension(dir);//返回文件名
            string p = sfolder + "\\" + "标准化" + "\\" + fname + ".csv";//保存路径

            Workbook workbooknew = new Workbook();//创表格实例
            workbooknew.Version = ExcelVersion.Version2013;
            workbooknew.Worksheets.Clear();
            Worksheet sheetnew = workbooknew.Worksheets.Add("sheet1");//选择单元格
            sheetnew.DefaultRowHeight = 18;
            sheetnew.DefaultColumnWidth = 25;
            //获取打开表格的全部数据
            CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
            //复制到新建表格
            range.Copy(sheetnew.Range[1, 1]);
            //删除除第一行外的表头
            for (int i = 1; i < title; i++)
            {
                sheetnew.DeleteRow(2);
            }
            creatDirctory(sfolder);
            sheetnew.SaveToFile(p, ",", Encoding.UTF8);//保存表格 
            SetText("sucssus");
        }

        private string Extension(string dir)//判断扩展名
        {
            string extension = Path.GetExtension(dir);
            if (extension == ".xls")
            {
                return ".xlsx";
            }
            else
            {
                return ".xls";
            }
        }

        public void creatDirctory(string dir)//创建不存在的目录
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)//更新目录
        {
            sfolder = uiTextBox4.Text;
            replace.sPath = uiTextBox4.Text;
            Properties.Settings.Default.OutputText = uiTextBox4.Text;
            Properties.Settings.Default.Save();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)//自动滚动
        {
            uiRichTextBox1.SelectionStart = uiRichTextBox1.Text.Length;
            uiRichTextBox1.ScrollToCaret();
        }

        delegate void SafeSetText(string strMsg);
        public void SetText(string strMsg)//代理文本更新
        {
            SafeSetText objSet = delegate (string str)
            {
                uiRichTextBox1.AppendText(str);
            };
            this.Invoke(objSet, new object[] { strMsg });
        }

        delegate void SafeSetEable(bool strMsg);
        public void SetEable(bool strMsg)//代理pannel更新
        {
            SafeSetEable objSet = delegate (bool str)
            {
                if (str)
                {
                    panel1.Enabled = true;
                    uiButton6.Enabled = false;
                }
                else
                {
                    panel1.Enabled = false;
                    uiButton6.Enabled = true;
                }
            };
            this.Invoke(objSet, new object[] { strMsg });
        }

        public void button6_Click(object sender, EventArgs e)//终止操作
        {
            if (th == null)
            {
                return;
            }
            try
            {
                SetEable(true);
                th.Abort();
                uiRichTextBox1.AppendText("\n操作终止\n");
            }
            catch (Exception)
            {
                throw;
            }

        }

        private void button7_Click(object sender, EventArgs e)//删除
        {
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b2 == true)
            {
                SetText("请选择目录\n"); return;
            }
            DialogResult result = MessageBox.Show("确定要删除？", "请确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.OK)
            {
                th = new Thread(delegate () { delBox(folder); });
                th.IsBackground = true;
                th.Start();
            }
        }

        private void delBox(string path)//删除封装
        {
            SetEable(false);
            delFiles(path);
            SetEable(true);
            SetText("Delete complete\n=========================================================\n");
        }

        private void delFiles(string path)//删除文件
        {
            DirectoryInfo d = new DirectoryInfo(path);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    delFiles(fsinfo.FullName);//递归调用
                }
                else
                {
                    if (fsinfo.FullName.Contains(delText) && string.IsNullOrWhiteSpace(delText) == false)
                    {
                        continue;
                    }
                    else
                    {
                        File.Delete(fsinfo.FullName);
                        SetText("del: " + fsinfo.FullName + "\n");
                    }
                }
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            delText = uiTextBox5.Text;
            replace.filter = delText;
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)//打开
        {
            FolderBrowserDialog ofolder = new FolderBrowserDialog();
            if (ofolder.ShowDialog() == DialogResult.OK)
            {
                folder = ofolder.SelectedPath.ToString();
                uiTextBox3.Text = folder;
            }
            if (string.IsNullOrEmpty(ofolder.SelectedPath))
            {
                return;
            }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)//保存
        {
            FolderBrowserDialog ofolder = new FolderBrowserDialog();
            if (ofolder.ShowDialog() == DialogResult.OK)
            {
                uiTextBox4.Text = ofolder.SelectedPath.ToString();
            }
            if (string.IsNullOrEmpty(ofolder.SelectedPath))
            {
                return;
            }
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)//单文件拆分
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            if (b1 == true)
            {
                SetText("请选择保存目录\n");
                return;
            }
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel(*.xlsx,*.xls)|*.xlsx;*.xls|all|*.*";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                path = openFile.FileName.ToString();
            }
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            SetEable(false);
            th = new Thread(delegate () { singleSplit(path); });
            th.IsBackground = true;
            th.Start();
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)//批量拆分
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                SetText("请选择目录\n");
                return;
            }
            else
            {
                SetEable(false);
                th = new Thread(delegate () { splitBox(folder); });
                th.IsBackground = true;
                th.Start();

            }
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)//批量合并
        {
            names.Clear();
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                SetText("请选择目录\n");
            }
            else
            {
                SetEable(false);
                th = new Thread(delegate () { comBox(folder); });
                th.IsBackground = true;
                th.Start();
            }
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)//合并为单个文件
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                SetText("请选择目录\n");
            }
            else
            {
                SetEable(false);
                th = new Thread(delegate () { comBox1(folder); });
                th.IsBackground = true;
                th.Start();
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)//退出
        {
            DialogResult dialogResult = MessageBox.Show("是否退出", "退出", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (dialogResult == DialogResult.OK)
            {
                //Thread.ResetAbort();
                //button6_Click(sender, e);
                Environment.Exit(0);
                //Application.Exit();
            }
        }

        private void 合并为单个文件选择ToolStripMenuItem_Click(object sender, EventArgs e)//选择文件合并
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            if (b1 == true)
            {
                SetText("请选择保存目录\n");
            }
            else
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Multiselect = true;
                dlg.Filter = "Excel(*.xlsx,*.xls)|*.xlsx;*.xls|all|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    SetEable(false);
                    th = new Thread(delegate () { comBox2(dlg.FileNames); });
                    th.IsBackground = true;
                    th.Start();
                }

            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("确定要退出？", "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                Environment.Exit(0);
                //button6_Click(sender, e);
                //System.Environment.Exit(0);
                //e.Cancel = false;          //这种也可以
            }
            else
            {
                e.Cancel = true;            //取消事件的值
            }
        }

        private void 保存记录ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //保存内容到txt
            FileStream fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + "log " + DateTime.Now.ToString("yyyy-MM-dd HHmmssfff") + ".txt", FileMode.Append);
            StreamWriter sw = new StreamWriter(fs, Encoding.Default);
            sw.Write(uiRichTextBox1.Text);
            //释放资源
            sw.Close();
            fs.Close();
            SetText("保存成功！\n");
        }

        private void 保留首行表头ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            names.Clear();
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                SetText("请选择目录\n");
            }
            else
            {
                SetEable(false);
                th = new Thread(delegate () { NormalBox(folder); });
                th.IsBackground = true;
                th.Start();
            }
        }

        private void 批量替换ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            replace = new Replace();
            replace.path = uiTextBox3.Text;
            replace.sPath = sfolder;
            replace.title = int.Parse(uiTextBox1.Text);
            replace.Show();
        }

        private void uiTextBox3_TextChanged(object sender, EventArgs e)
        {
            replace.path = uiTextBox3.Text;
            folder = uiTextBox3.Text;
            Properties.Settings.Default.InputText = uiTextBox3.Text;
            Properties.Settings.Default.Save();
        }

        private void uiTextBox1_TextChanged(object sender, EventArgs e)
        {
            replace.title = int.Parse(uiTextBox1.Text);
        }

        private void 批量拆分工作表ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                SetText("请选择目录\n");
                return;
            }
            else
            {
                DialogResult result = MessageBox.Show("Yes: 以文件名称分文件夹存放\n例如：\\GS1-X5成果\\堡坎.xlsx\nNo: 以表名分文件夹存放\n例如：\\堡坎\\GS1-X5成果.xlsx", "请选择文件保存方式", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    sway = true;
                }
                else
                {
                    sway = false;
                }
                DialogResult result1 = MessageBox.Show("是否要保存为csv方便矢量化？", "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result1 == DialogResult.OK)
                {
                    exten = true;
                }
                else
                {
                    exten = false;
                }
                SetEable(false);
                th = new Thread(delegate () { splitSheetsbox(folder); });
                th.IsBackground = true;
                th.Start();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            uiTextBox3.Text = Properties.Settings.Default.InputText;
            uiTextBox4.Text = Properties.Settings.Default.OutputText;
            folder = uiTextBox3.Text;
            sfolder = uiTextBox4.Text;
            uiButton6.Enabled = false;
        }
    }
}
