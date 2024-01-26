using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace AkizukiHistory
{
    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var newForm = new MainForm();

            if (args.Length > 0)
                newForm.folder = new System.IO.DirectoryInfo(args[0]);
            if(args.Length > 2)
            {
                newForm.userID = System.Web.HttpUtility.UrlDecode(args[1]);
                newForm.password = System.Web.HttpUtility.UrlDecode(args[2]);
            }

            Application.Run(newForm);
        }
    }
}