using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EXCELProcessing;
using Sunny.UI;

namespace EXCELProcessing
{
    static class Program
    {

        
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

                TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
                //处理UI线程异常 
                Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
                //处理非UI线程异常 
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {
                string str = GetExceptionMsg(ex, string.Empty);
                MessageBox.Show(str, "系统错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private static void TaskScheduler_UnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            throw new NotImplementedException();
        }


        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Form1 form = Form1.frm1;

            string str = GetExceptionMsg(e.ExceptionObject as Exception, e.ToString());
            form.SetText(str);
            form.SetEable(true);
            Exception error = e.ExceptionObject as Exception;
            MessageBox.Show(error.Message, "系统异常提示信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            while (true)
            {//循环处理，否则应用程序将会退出
                Thread.Sleep(2 * 1000);
            };
        }

        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            Form1 form = Form1.frm1;

            string str = GetExceptionMsg(e.Exception, e.ToString());
 
            form.SetText(str);
            form.SetEable(true);
            MessageBox.Show(e.Exception.Message, "系统异常提示信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            throw new NotImplementedException();
        }
        // <summary> 
        /// 生成自定义异常消息 
        /// </summary> 
        /// <param name="ex">异常对象</param> 
        /// <param name="backStr">备用异常消息：当ex为null时有效</param> 
        /// <returns>异常字符串文本</returns> 
        private static string GetExceptionMsg(Exception ex, string backStr)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("****************************异常文本****************************");
            sb.AppendLine("【出现时间】：" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"));
            if (ex != null)
            {
                sb.AppendLine("【异常类型】：" + ex.GetType().Name);
                sb.AppendLine("【异常信息】：" + ex.Message);
                sb.AppendLine("【堆栈调用】：" + ex.StackTrace);
                sb.AppendLine("【异常方法】：" + ex.TargetSite);
            }
            else
            {
                sb.AppendLine("【未处理异常】：" + backStr);
            }
            sb.AppendLine("***************************************************************");

            return sb.ToString();
        }
    }
}
