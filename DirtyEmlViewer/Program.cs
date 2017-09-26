using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DirtyEmlViewer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(String[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (File.Exists(args[0]))
            {
                var file = new FileInfo(args[0]);
                Application.Run(new Form1(file.Directory, file));
            }
            else if (Directory.Exists(args[0]))
            {
                var dir = new DirectoryInfo(args[0]);
                Application.Run(new Form1(dir));
            }
            else
            {
                MessageBox.Show($"Cannot find specified file or directory.\r\n{args[0]}", "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
