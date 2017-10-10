using MimeKit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DirtyEmlViewer
{

    public partial class Form1 : Form
    {
        private class Mail
        {
            public FileInfo File { get; }
            public MimeMessage Message { get; }

            public Mail(FileInfo file)
            {
                File = file;
                Message = MimeMessage.Load(file.FullName);
            }
        }

        private DirectoryInfo _emailsDir;


        public Form1(DirectoryInfo emailsDir = null, FileInfo file = null)
        {
            _emailsDir = emailsDir;

            InitializeComponent();

            if (file != null)
            {
                LoadMessage(file);
            }
            else if (emailsDir != null)
            {
                LoadMessages();
            }
            else
            {
                OpenFolderDialog();
            }
        }

        private void LoadMessages()
        {
            Clear();

            if (_emailsDir.Exists)
            {
                var files = new List<FileInfo>();
                files.AddRange(_emailsDir.GetFiles("*.eml", SearchOption.AllDirectories));
                files.AddRange(_emailsDir.GetFiles("*.mht", SearchOption.AllDirectories));

                files.OrderByDescending(x => x.CreationTime).ToList().ForEach(LoadMessage);

                SetStatus($@"Loaded {listView1.Items.Count} email(s) from {_emailsDir.FullName}");
            }
        }

        private void SetStatus(String status)
        {
            toolStripStatusLabel1.Text = status;
        }

        private void Clear()
        {
            listView1.Items.Clear();
            webBrowser1.DocumentText = "";
            toolStripStatusLabel1.Text = "";
        }

        private void LoadMessage(FileInfo emlFile)
        {
            listView1.Items.Add(ParseEmail(emlFile));

            SetStatus($@"Loaded file {emlFile.FullName}");
        }

        private ListViewItem ParseEmail(FileInfo emlFile)
        {
            var str = new List<String>();

            var mail = new Mail(emlFile);

            str.Add(mail.Message.Subject);
            str.Add(mail.Message.To.FirstOrDefault()?.ToString());
            str.Add(mail.Message.Date.ToString("u"));

            var item = new ListViewItem(str.ToArray())
            {
                Tag = mail
            };

            return item;
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            OpenMessage(e.Item);
        }

        private void OpenMessage(ListViewItem item)
        {
            viewAttchmentsToolStripMenuItem.Enabled = false;

            var mail = (Mail)item.Tag;

            webBrowser1.DocumentText = mail.Message.HtmlBody;

            if (mail.Message.Attachments.Any())
            {
                viewAttchmentsToolStripMenuItem.Enabled = true;
            }
        }

        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            var pattern = new Regex(@"^((http[s]?|ftp):\/)?\/?([^:\/\s]+)((\/\w+)*\/)([\w\-\.]+[^#?\s]+)(.*)?(#[\w\-]+)?$");
            var match = pattern.Match(e.Url.ToString());

            if (match.Success)
            {
                String link = match.Groups[0].Value;
                Process.Start(link);
                e.Cancel = true;
            }
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                var item = listView1.SelectedItems[0];
                var mail = (Mail)item.Tag;

                try
                {
                    mail.File.Delete();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            LoadMessages();
        }

        private void reloadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadMessages();
        }

        private void viewHeadersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                var item = listView1.SelectedItems[0];
                var mail = (Mail)item.Tag;

                var strBuilder = new StringBuilder();

                foreach (var header in mail.Message.Headers)
                {
                    strBuilder.AppendLine($"{header.Field:30}: {header.Value}");
                }

                MessageBox.Show(strBuilder.ToString(), "Headers", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Quick and Dirty Eml file Viewer\r\nBy http://github.com/b1thunt3r", "Dirty Eml Viewer", MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void deleteAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_emailsDir.Exists)
            {
                var files = new List<FileInfo>();
                files.AddRange(_emailsDir.GetFiles("*.eml", SearchOption.AllDirectories));
                files.AddRange(_emailsDir.GetFiles("*.mht", SearchOption.AllDirectories));

                files.ForEach(file =>
                {
                    file.Delete();
                });
            }

            LoadMessages();
        }

        private void loadDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFolderDialog();
        }

        private void OpenFolderDialog()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                _emailsDir = new DirectoryInfo(folderBrowserDialog1.SelectedPath);
                LoadMessages();
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            Clear();
            var paths = (String[])e.Data.GetData(DataFormats.FileDrop);

            foreach (var path in paths)
            {
                if (Directory.Exists(path))
                {
                    _emailsDir = new DirectoryInfo(path);
                    LoadMessages();
                } else if (File.Exists(path))
                {
                    LoadMessage(new FileInfo(path));
                }
            }
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            Form1_DragEnter(sender, e);
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            Form1_DragDrop(sender, e);
        }

        private void webBrowser1_Navigating_1(object sender, WebBrowserNavigatingEventArgs e)
        {
            var pattern = new Regex(@"^((http[s]?|ftp):\/)?\/?([^:\/\s]+)((\/\w+)*\/)([\w\-\.]+[^#?\s]+)(.*)?(#[\w\-]+)?$");
            var match = pattern.Match(e.Url.ToString());

            if (match.Success)
            {
                String link = match.Groups[0].Value;
                Process.Start(link);
                e.Cancel = true;
            }
        }
    }
}
