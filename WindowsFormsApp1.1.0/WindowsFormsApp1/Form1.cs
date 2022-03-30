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
        string path;
        string folder;
        string sfolder;
        public List<string> names = new List<string>();
        public Microsoft.Office.Interop.Excel.Application excel;

        private void button2_Click(object sender, EventArgs e)//单文件拆分
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            if (b1 == true)
            {
                richTextBox1.AppendText("请选择保存目录\n");
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
            Split(path);
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
                richTextBox1.AppendText("请选择目录\n");
            }
            else
            {
                Director(folder);
                richTextBox1.AppendText("Combine success\n");
            } 
        }

        private void button5_Click(object sender, EventArgs e)//批量拆分
        {
            bool b1 = string.IsNullOrWhiteSpace(sfolder);
            bool b2 = string.IsNullOrWhiteSpace(folder);
            if (b1 || b2 == true)
            {
                richTextBox1.AppendText("请选择目录\n");
                return;
            }
            else
            {
                splitDirector(folder);
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
                    calcultin(fsinfo.FullName);
                }
            }
        }

        private void calcultin(string dir)//合并计算
        {
            excel = new ApplicationClass
            {
                ScreenUpdating = false, //停止工作表刷新
                DisplayAlerts = false
            };
            richTextBox1.AppendText(dir);//输出文件的全部路径
            richTextBox1.AppendText("\n");
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
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);//关闭excel进程
        }

        private void Split(string dir)//拆分
        {
            string fname = Path.GetFileNameWithoutExtension(dir);
            richTextBox1.AppendText(dir);
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
                string p = sPath + "\\" + fname + Extension(dir);
                workbooknew.SaveAs(p, XlFileFormat.xlWorkbookDefault);
                //workbooknew.SaveAs(sPath, XlFileFormat.xlWorkbookDefault);
                workbooknew.Close();
            }
            workbook2.Close();
            excel.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            PublicMethod.Kill(excel);//关闭excel进程
            richTextBox1.AppendText("\nSplit success\n");

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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
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
    }
}
