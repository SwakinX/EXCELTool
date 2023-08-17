using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace EXCELProcessing
{
    internal class Class1
    {
        public void test()
        {
            Workbook wbk = new Workbook();
            wbk.LoadFromFile(@"D:\@@@21-2-21\data.xlsx", true);

            Worksheet sht = wbk.Worksheets[0];

            var filters = sht.AutoFilters;

            filters.Range = sht.Range[1, 1, sht.LastRow, sht.LastColumn];
            filters.AddFilter(10, "Seattle");
            filters.Filter();
            //filters.Worksheet.Range;
            foreach (var item in filters.Range)
            {

                Console.WriteLine(item.Value2);

            }

            Console.WriteLine("Done!");

        }
        //private void Split(string dir)//拆分
        //{
        //    //canStop = false;
        //    string fname = Path.GetFileNameWithoutExtension(dir);
        //    SetText(dir + "\n");

        //    Workbook workbook2 = new Workbook();//创建表格实例
        //    workbook2.LoadFromFile(dir);//打开已有表格
        //    Worksheet sheet2 = workbook2.ActiveSheet;

        //    Dictionary<String, CellRange> dic = new Dictionary<String, CellRange>();

        //    int lColumn = sheet2.LastColumn;//获得最大列数
        //    int lRow = sheet2.LastRow;//获得最大行数
        //    int b = int.Parse(textBox1.Text);//标题行
        //    int c = 0;//分割列号

        //    for (int i = 1; i <= lRow; i++)
        //    {
        //        string temp = sheet2.Range[1, i].Text.ToString();
        //        if (temp == textBox2.Text)
        //        {
        //            c = i; break;//获取分割列
        //        }
        //    }
        //    b = b + 1;
        //    SetText("正在处理\n");
        //    for (int i = b; i < lRow + 1; i++)
        //    {
        //        string h = sheet2.Range[i, c].Text.ToString();//获取名称
        //        if (string.IsNullOrWhiteSpace(h))
        //        {
        //            continue;
        //        }
        //        if (dic.ContainsKey(h))
        //        {
        //            Workbook workbooknew = new Workbook();//创表格实例
        //            workbooknew.Version = ExcelVersion.Version2013;
        //            //workbooknew.Worksheets.Clear();
        //            Worksheet sheetnew = workbooknew.Worksheets[0];//选择单元格
        //            CellRange r = sheetnew.Range["A1"];

        //            dic[h].Copy(r);
        //            sheet2.Range[i, 1, i, lColumn].Copy(sheetnew.Range[sheetnew.LastRow + 1, 1]);
        //            //var d = dic[h].AddCombinedRange(sheet2.Range[i, 1, i, lColumn]);
        //            //var d = sheet2.Range[i, 1, i, lColumn].AddCombinedRange(dic[h]);

        //            //workbooknew.SaveToFile("save.xlsx",ExcelVersion.Version2013);
        //            dic[h] = sheetnew.Range[1, 1, sheetnew.LastRow, sheetnew.LastColumn];
        //            workbooknew.Dispose();

        //        }
        //        else
        //        {
        //            dic.Add(h, sheet2.Range[i, 1, i, lColumn]);
        //        }

        //    }

        //    for (int i = 0; i < dic.Count; i++)
        //    {
        //        Workbook workbooknew = new Workbook();//创表格实例
        //        workbooknew.Version = ExcelVersion.Version2013;
        //        workbooknew.Worksheets.Clear();
        //        Worksheet sheetnew = workbooknew.Worksheets.Add("sheet1");//选择单元格
        //        CellRange range;

        //        sheet2.Range[1, 1, b - 1, lColumn].Copy(sheetnew.Range[1, 1]);
        //        range = sheetnew.Range[b, 1];
        //        Dictionary<String, CellRange>.KeyCollection key = dic.Keys;
        //        string k = key.ElementAt(i);
        //        dic[k].Copy(range);
        //        string sPath = sfolder + "\\" + "拆分" + "\\" + k;
        //        creatDirctory(sPath);
        //        string p = sPath + "\\" + fname + Extension(dir);
        //        workbooknew.SaveToFile(p, ExcelVersion.Version2013);

        //        SetText("已完成：" + k + "\n");

        //    }
        //}
    }
}
