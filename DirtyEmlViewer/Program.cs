using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DirtyEmlViewer
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(String[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Form form = null;
            if (args.Length > 0 && !String.IsNullOrWhiteSpace(args[0]))
            {
                if (File.Exists(args[0]))
                {
                    var file = new FileInfo(args[0]);
                    form = new Form1(file.Directory, file);
                }
                else if (Directory.Exists(args[0]))
                {
                    var dir = new DirectoryInfo(args[0]);
                    form = new Form1(dir);
                }
            }
            else
            {
                form = new Form1();
            }

            Application.Run(form);

        }
    }
}
