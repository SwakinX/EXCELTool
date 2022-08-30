using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1;
using Sunny.UI;
using Spire.Xls;
using System.IO;
using System.Threading;

namespace 表格处理工具
{
    public partial class Replace : UIForm
    {
        public Replace()
        {
            InitializeComponent();
        }
        public string path;
        public string sPath;
        public int title;
        public string filter;
        Form1 form = Form1.frm1;
        Thread th;

        /// <summary>
        /// 自动编号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }
        /// <summary>
        /// 判断表格是否为空
        /// </summary>
        /// <param name="dg"></param>
        /// <returns></returns>
        private bool isEmptydataGridView(DataGridView dg)
        {
            string a = Convert.ToString(dg.Rows[0].Cells[0].Value);
            string b = Convert.ToString(dg.Rows[0].Cells[1].Value);
            if (a.IsNullOrWhiteSpace()||b.IsNullOrWhiteSpace())
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 加载数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel(*.xlsx,*.xls,*.csv)|*.xlsx;*.xls;*.csv|all|*.*";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(openFile.FileName.ToString());
                Worksheet sheet = workbook.Worksheets[0];
                DataTable dataTable = sheet.ExportDataTable(sheet.AllocatedRange, false, false);
                uiDataGridView1.Columns.Clear();
                this.uiDataGridView1.DataSource = dataTable;
                try
                {
                    uiDataGridView1.Columns[0].HeaderText = "原始";
                    uiDataGridView1.Columns[1].HeaderText = "替换";
                }
                catch (Exception)
                {

                    throw;
                }
            }
        }

        delegate void SafeSetEable(bool strMsg);
        public void enable(bool strMsg)//代理更新
        {
            SafeSetEable objSet = delegate (bool b)
            {
                uiButton1.Enabled = b;
                uiButton2.Enabled = b;
                if (b == false)
                {
                    uiDataGridView1.ReadOnly = true;
                }
                else
                {
                    uiDataGridView1.ReadOnly = false;
                }
            };
            this.Invoke(objSet, new object[] { strMsg });
        }

        /// <summary>
        /// 开始替换
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton2_Click(object sender, EventArgs e)
        {
            bool b1 = string.IsNullOrWhiteSpace(sPath);
            bool b2 = string.IsNullOrWhiteSpace(path);
            if (b1 || b2 == true)
            {
                form.SetText("请选择目录\n");
                return;
            }
            if (isEmptydataGridView(uiDataGridView1))
            {
                form.SetText("\n请检查表格是否正确");
            }
            else
            {
                form.panel1.Enabled = false;
                enable(false);
                th = new Thread(delegate () { replaceBox(path); });
                th.IsBackground = true;
                th.Start();

            }

        }

        private void replaceBox(string dir)
        {
            form.SetText("Start Replace\n");
            replaceDirector(dir);
            enable(true);
            form.SetEable(true);
            form.SetText("\nReplace complete\n=========================================================\n");
        }
        /// <summary>
        /// 递归遍历文件
        /// </summary>
        /// <param name="dir"></param>
        private void replaceDirector(string dir)
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    replaceDirector(fsinfo.FullName);//递归调用
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(filter) || fsinfo.FullName.Contains(filter))
                    {
                        replace(fsinfo.FullName);
                    }
                }
            }
        }
        /// <summary>
        /// 替换
        /// </summary>
        /// <param name="dir"></param>
        private void replace(string dir)
        {
            Workbook workbook = new Workbook();
            string ex = Path.GetExtension(dir);
            if (ex==".csv")
                workbook.LoadFromFile(dir,",");
            else
                workbook.LoadFromFile(dir);
            Worksheet sheet = workbook.ActiveSheet;
            form.SetText(dir+"\n");
            if (sheet.LastRow <= title)
            {
                form.SetText("表为空\n");
                return;
            }

            for (int i = 0; i < uiDataGridView1.Rows.Count; i++)
            {
                string value = Convert.ToString(uiDataGridView1.Rows[i].Cells[0].Value);
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }
                CellRange[] ranges = sheet.FindAllString(value, false, false);
                foreach (var range in ranges)
                {
                    range.Text = uiDataGridView1.Rows[i].Cells[1].Value.ToString();
                    range.Style.Color = Color.Yellow;
                }
            }
            string fname = Path.GetFileNameWithoutExtension(dir);//返回文件名
            if (ex == ".xls")
                ex = ".xlsx";
            string p = sPath + "\\" + "替换" + "\\" + fname + ex;//保存路径
            sheet.DefaultRowHeight = 18;
            if (ex == ".csv")
                workbook.SaveToFile(p, ",");
            else
                workbook.SaveToFile(p,FileFormat.Version2016);
            form.SetText("sucssus\n");

        }
        /// <summary>
        /// 取消
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton3_Click(object sender, EventArgs e)
        {
            if (th == null)
            {
                return;
            }
            try
            {
                th.Abort();
                form.SetText("操作终止\n");
                enable(true);
                form.panel1.Enabled = true;
            }
            catch (Exception)
            {
                throw;
            }

        }
    }
}
