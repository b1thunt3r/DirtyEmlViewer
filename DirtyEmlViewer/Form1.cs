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

        private readonly DirectoryInfo _emailsDir;
        private readonly FileInfo _selectedFile;


        public Form1(DirectoryInfo emailsDir, FileInfo file = null)
        {
            _emailsDir = emailsDir;
            _selectedFile = file;

            InitializeComponent();

            listView1.Columns.AddRange(new[]
            {
                new ColumnHeader("Subject"),
                new ColumnHeader("To"),
                new ColumnHeader("Recivied")
            });

            LoadMessages(true);
        }

        private void LoadMessages(Boolean selectFile = false)
        {
            listView1.Items.Clear();
            webBrowser1.DocumentText = "";
            toolStripStatusLabel1.Text = "";

            ListViewItem item = null;
            //dirPath = @"C:\temp\scope-email";
            if (_emailsDir.Exists)
            {
                var files = new List<FileInfo>();
                files.AddRange(_emailsDir.GetFiles("*.eml", SearchOption.AllDirectories));
                files.AddRange(_emailsDir.GetFiles("*.mht", SearchOption.AllDirectories));

                files = files.OrderByDescending(x => x.CreationTime).ToList();

                files.ForEach(file =>
                {
                    var tItem = listView1.Items.Add(ParseEmail(file));

                    if (_selectedFile != null && file.FullName == _selectedFile.FullName)
                    {
                        item = tItem;
                    }

                });

                if (listView1.Items.Count > 0 && item == null)
                {
                    item = listView1.Items[0];
                }

                toolStripStatusLabel1.Text = $@"Loaded {listView1.Items.Count} email(s) from {_emailsDir.FullName}";
            }

            if (item != null)
            {
                item.Selected = true;
                item.Focused = true;
            }
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
    }
}
