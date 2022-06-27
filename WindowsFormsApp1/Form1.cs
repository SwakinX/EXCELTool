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

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string path;
        public string folder;
        public string sfolder;
        public List<string> names = new List<string>();

        Thread th;
        public string delText = "";

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
            panel1.Enabled = false;
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
                textBox3.Text = folder;
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
                textBox4.Text = ofolder.SelectedPath.ToString();
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
                panel1.Enabled = false;
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
                panel1.Enabled = false;
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

        private void comBox(string dir)
        {
            SetText("Start Combine\n");
            Director(dir);
            SetEable(true);
            SetText("\nCombine complete\n=========================================================\n");
        }

        private void splitBox(string dir)
        {
            SetText("Start Split\n");
            splitDirector(dir);
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

        private void calcultin(string dir)//合并计算
        {
            SetText("\n" + dir + "\n");//输出文件的全部路径

            Workbook workbook = new Workbook();//创表格实例
            workbook.LoadFromFile(dir);//打开表格
            Worksheet sheet = workbook.ActiveSheet;//提取活动表格
            int title = int.Parse(textBox1.Text);//标题行
            string fname = Path.GetFileNameWithoutExtension(dir);//返回文件名
            string p = sfolder + "\\" + fname + Extension(dir);//保存路径
            bool bb = names.Contains(fname);

            if (bb == false)
            {
                names.Add(fname);//添加到列表

                Workbook workbooknew = new Workbook();//创表格实例
                workbooknew.Version = ExcelVersion.Version2013;
                workbooknew.Worksheets.Clear();
                Worksheet sheetnew = workbooknew.Worksheets.Add("sheet1");//选择单元格

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
                //获取打开表格的数据
                CellRange range = sheet.Range[title + 1, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];
                range.Copy(sheet1.Range[sheet1.LastRow + 1, 1]);
                sheet1.DefaultRowHeight = 18;
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

            Dictionary<String, CellRange> dic = new Dictionary<String, CellRange>();

            int lColumn = sheet2.LastColumn;//获得最大列数
            int lRow = sheet2.LastRow;//获得最大行数
            int b = int.Parse(textBox1.Text);//标题行
            int c = 0;//分割列号
            if (sheet2.LastRow <= b)
            {
                SetText("表格为空\n");
                return;
            }
            for (int i = 1; i <= lRow; i++)
            {
                string temp = sheet2.Range[1, i].Text.ToString();
                if (temp == textBox2.Text)
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
            String[] new_arr = arr.GroupBy(p => p).Select(p => p.Key).ToArray();
            SetText("共"+new_arr.Length.ToString()+"条管道\n");
            //创建筛选
            AutoFiltersCollection filters = sheet2.AutoFilters;
            filters.Range = sheet2.Range;
            //filters.Range = sheet2.Range[1, 1, sheet2.LastRow + 1, sheet2.LastColumn];

            foreach (string cr in new_arr)
            {
                filters.Clear();
                filters.AddFilter(c-1, cr);
                filters.Filter();

                Workbook wb = new Workbook();//创表格实例
                wb.Version = ExcelVersion.Version2013;
                wb.Worksheets.Clear();
                Worksheet sheet = wb.Worksheets.Add("sheet1");//选择单元格
                sheet2.Range[1, 1, b, lColumn].Copy(sheet.Range[1, 1]);

                int last = filters.Range.LastRow;


                foreach (var row in filters.Range.Rows)
                {
                    if (row.Row == 1)
                    {
                        continue;
                    }
                    bool bbb = sheet2.IsRowVisible(row.Row);//是否是可见行
                    if (bbb)
                    {
                        row.Copy(sheet.Range[sheet.LastRow+1, 1]);
                    }
                    //if (row.Row == last && cr == filters.Range[last,c].NumberText)
                    //{
                    //    row.Copy(sheet.Range[sheet.LastRow + 1, 1]);
                    //}//解决最后一行总是被隐藏的问题
                } //筛选结果保存到表格
                string k = cr.Trim();
                string sPath = sfolder + "\\" + "拆分" + "\\" + k;
                creatDirctory(sPath);
                string p = sPath + "\\" + fname + Extension(dir);
                sheet.DefaultRowHeight = 18;
                wb.SaveToFile(p, ExcelVersion.Version2013);

                SetText("已完成：" + k + "\n");
            }
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

        private void creatDirctory(string dir)//创建不存在的目录
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)//更新目录
        {
            sfolder = textBox4.Text;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)//自动滚动
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }

        delegate void SafeSetText(string strMsg);
        private void SetText(string strMsg)//代理文本更新
        {
            SafeSetText objSet = delegate (string str)
            {
                richTextBox1.AppendText(str);
            };
           this.Invoke(objSet, new object[] { strMsg });
        }

        delegate void SafeSetEable(bool strMsg);
        private void SetEable(bool strMsg)//代理pannel更新
        {
            SafeSetEable objSet = delegate (bool str)
            {
                panel1.Enabled = str;
            };
            this.Invoke(objSet, new object[] { strMsg });
        }

        private void button6_Click(object sender, EventArgs e)//终止操作
        {
            if (th == null)
            {
                return;
            }
            try
            {
                th.Abort();
                richTextBox1.AppendText("操作终止\n");
                panel1.Enabled = true;
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
                SetText("请选择目录\n");return;
            }
            DialogResult result = MessageBox.Show("确定要删除？","请确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
            if (result==DialogResult.OK)
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
                    if (fsinfo.FullName.Contains(delText)&&string.IsNullOrWhiteSpace(delText)==false)
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
            delText = textBox5.Text;
        }
    }
}
