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
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public class PublicMethod
        {
            [DllImport("User32.dll", CharSet = CharSet.Auto)]
            public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
            public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
            {
                IntPtr t = new IntPtr(excel.Hwnd);//获得这个句柄，具体做用是获得这块内存入口 

                int k = 0;
                GetWindowThreadProcessId(t, out k);   //获得本进程惟一标志k
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //获得对进程k的引用
                p.Kill();     //关闭进程k
            }
        }
            public Form1()
        {
            InitializeComponent();
        }

        public string path;
        public string folder;
        public string sfolder;
        public List<string> names = new List<string>();
        public Microsoft.Office.Interop.Excel.Application excel;
        Thread th;
        public string delText = "";

        //public string help = "单文件拆分只需要选择保存目录/n批量拆合先选择打开目录和保存目录/n" +
        //    "拆分需要填表头行数和拆分字段/n合并只需要表头行数/n注意表头行数和要用来拆分字段是否正确！！！/n" +
        //    "当已做的存放在单独目录下时，可以在文件过滤中输入存放文件夹名，提高计算速度/n" +
        //    "\"删除\"键删除打开目录下的文件，文件过滤是要保留的文件所在的文件夹名称" +
        //    "/n====================================================/n";
        //private delegate void Mydel(string dir);
        //Mydel del;
        //private volatile bool canStop = false;

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
            //Mydel del = Split;
            //IAsyncResult iar = del.BeginInvoke(path, null, null);

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
            
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if ( b1 || b2 == true)
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
            SetText("Combine complete\n=========================================================\n");
        }

        private void splitBox(string dir)
        {
            SetText("Start Split\n");
            splitDirector(dir);
            SetEable(true);
            SetText("Split complete\n=========================================================\n");
        }

        private void singleSplit(string path)
        {
            SetText("Start Split\n");
            Split(path);
            SetEable(true);
            SetText("Split complete\n=========================================================\n");
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
                    if (string.IsNullOrWhiteSpace(delText)|| fsinfo.FullName.Contains(delText))
                    {
                        calcultin(fsinfo.FullName);
                    }
                }
            }


        }

        private void calcultin(string dir)//合并计算
        {
            //canStop = false;
            excel = new ApplicationClass
            {
                ScreenUpdating = false, //停止工作表刷新
                DisplayAlerts = false
            };
            SetText(dir+ "\n");//输出文件的全部路径
            Workbook workbook = excel.Workbooks.Open(dir);
            Worksheet sheet = (Worksheet)workbook.ActiveSheet;

            Range rng = sheet.UsedRange;
            int lColumn = rng.Columns.Count;//获得最大列数
            int lRow = rng.Rows.Count;//获得最大行数
            int b = int.Parse(textBox1.Text);//标题行

            string fname = Path.GetFileNameWithoutExtension(dir);
            bool bb = names.Contains(fname);
            if (bb == false)
            {
                names.Add(fname);//添加不存在的文件名

                Workbook workbooknew = excel.Workbooks.Add(Type.Missing);
                Worksheet sheetnew = (Worksheet)workbooknew.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Range range;
                range = (Range)sheetnew.Cells[1, 1];
                sheet.Range[sheet.Cells[1, 1], sheet.Cells[lRow, lColumn]].Copy(range);
                creatDirctory(sfolder);
                string p = sfolder + "\\" + fname + Extension(dir);
                workbooknew.SaveAs(p, XlFileFormat.xlWorkbookDefault);
                workbooknew.Close();
            }
            else
            {

                Workbook workbook2 = excel.Workbooks.Open(sfolder + "\\" + fname + Extension(dir));
                Worksheet sheet1 = (Worksheet)workbook2.ActiveSheet;
                Range range = sheet1.UsedRange;
                int Row = range.Rows.Count;//获得最大行数
                                           //int Column = range.Columns.Count;//获得最大列数
                range = (Range)sheet1.Cells[Row+1, 1];
                sheet.Range[sheet.Cells[b+1, 1], sheet.Cells[lRow, lColumn]].Copy(range);
                workbook2.Save();
                workbook2.Close();
            }
            excel.Quit();
            PublicMethod.Kill(excel);
            //canStop = true;
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);//关闭excel进程
        }

        private void Split(string dir)//拆分
        {
            //canStop = false;
            string fname = Path.GetFileNameWithoutExtension(dir);
            SetText(dir + "\n");
            Microsoft.Office.Interop.Excel.Application excel = new ApplicationClass
            {
                ScreenUpdating = false, //停止工作表刷新
                DisplayAlerts = false
            };
            Workbook workbook2 = excel.Workbooks.Open(dir);
            Worksheet sheet2 = (Worksheet)workbook2.ActiveSheet;

            Dictionary<String, Range> dic = new Dictionary<String, Range>();
            Range rng = sheet2.UsedRange;
            int lColumn = rng.Columns.Count;//获得最大列数
            int lRow = rng.Rows.Count;//获得最大行数
            int b = int.Parse(textBox1.Text);//标题行
            int c=0;
            //int c = Convert.ToInt32(textBox2.Text);//分割列
            for (int i = 1; i <= lRow; i++)
            {
                string temp = ((Range)sheet2.Cells[1, i]).Text.ToString();
                if (temp == textBox2.Text)
                {
                    c = i;break;
                }
            }
            b = b + 1;
            SetText("正在处理\n");
            for (int i = b; i < lRow + 1; i++)
            {
                string h = sheet2.Range[sheet2.Cells[i, c], sheet2.Cells[i, c]].Text.ToString();
                if (dic.ContainsKey(h))
                {
                    dic[h] = excel.Union(dic[h], sheet2.Range[sheet2.Cells[i, 1], sheet2.Cells[i, lColumn]]);
                }
                else
                {
                    dic.Add(h, sheet2.Range[sheet2.Cells[i, 1], sheet2.Cells[i, lColumn]]);
                }

            }

            for (int i = 0; i < dic.Count; i++)
            {

                Workbook workbooknew = excel.Workbooks.Add(Type.Missing);
                Worksheet sheetnew = (Worksheet)workbooknew.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Range range;
                range = (Range)sheetnew.Cells[1, 1];
                sheet2.Range[sheet2.Cells[1, 1], sheet2.Cells[b - 1, lColumn]].Copy(range);
                range = (Range)sheetnew.Cells[b, 1];
                //range = (Range)sheetnew.Cells[sheetnew.Range["A65536", "A65536"].End[XlDirection.xlUp].Row, 1];
                Dictionary<String, Range>.KeyCollection key = dic.Keys;
                string k = key.ElementAt(i);
                if (string.IsNullOrWhiteSpace(k))
                {
                    return;
                }
                dic[k].Copy(range);
                string sPath = sfolder + "\\" + "拆分" + "\\" + k;
                creatDirctory(sPath);
                //string p = sPath + "\\" + k + "中心线数据" + Extension(dir);
                string p = sPath + "\\" + fname + Extension(dir);
                workbooknew.SaveAs(p, XlFileFormat.xlWorkbookDefault);
                //workbooknew.SaveAs(sPath, XlFileFormat.xlWorkbookDefault);
                SetText("已完成："+k+"\n");
                workbooknew.Close();
            }
            workbook2.Close();
            excel.Quit();
            
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            PublicMethod.Kill(excel);//关闭excel进程
            //canStop = true;

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
                //excel.Quit();
                //PublicMethod.Kill(excel);//关闭excel进程
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
