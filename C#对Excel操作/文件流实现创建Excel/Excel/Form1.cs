using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string text1, text2, text3, text4, text5;

            text1 = textBox1.Text;
            text2 = textBox2.Text;
            text3 = textBox3.Text;
            text4 = textBox4.Text;
            text5 = textBox5.Text;

            if (string.IsNullOrEmpty(text1) && string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text3) && string.IsNullOrEmpty(text4) && string.IsNullOrEmpty(text5))
            {
                MessageBox.Show("不能为空！");
                return;
            }


            DemoXls(text1, text2, text3, text4, text5);
        }


        /// <summary>
        /// Excel模板
        ///    DemoXls
        /// </summary>
        /// <returns></returns>
        public static void DemoXls(string text1, string text2, string text3, string text4, string text5)
        {
            string pp = @"C:\Users\Administrator\Desktop";

            string c1, c2, c3, c4, c5, c6, c7;
            c1 = "1";
            c2 = "11.10";
            c3 = "0.2";
            c4 = "客满中心";
            c5 = "文件";
            c6 = "是";
            c7 = "鼠天哪";


            if (!Directory.Exists(pp))
            {
                Directory.CreateDirectory(pp);
            }

            pp = pp + "\\" + "Example.xls";

            try
            {
                FindAndKillProcessByName("EXCEL");  //结束excel进程
                FindAndKillProcessByName("et");     //结束wps进程

                Thread.Sleep(1000);                 //写入excel延时
                FileStream fs = new FileStream(pp, FileMode.Append);
                StreamWriter fsw = new StreamWriter(fs, Encoding.GetEncoding("GB2312"));

                fsw.WriteLine(c1 + "\t" + c2 + "\t" + text1 + "\t" + c3 + "\t" + text2 + "\t" + text3 + "\t" + text4+ "\t" + c4 + "\t" + c5 + "\t" + c6 + "\t" + text5 + "\t" + c7);
                // fsw.WriteLine();    //空一行

                fsw.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #region 结束进程
        /// <summary>
        /// 结束进程
        ///    FindAndKillProcessByName
        /// </summary>
        /// <param name="name"></param>
        public static void FindAndKillProcessByName(string name)
        {
            foreach (Process winProc in Process.GetProcessesByName(name))
            {
                if (winProc.ProcessName.Equals(name))
                {
                    winProc.Kill();
                }
            }
        }
        #endregion


    }
}
