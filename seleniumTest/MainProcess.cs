using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.Windows.Forms;

namespace ScraperBase
{
    class MainProcess
    {
        public Form MyForm { get; set; }
        public RichTextBox MyRichTextBox { get; set; }
        public String FileName { get; set; }

        public void Run()
        {
            Log("Starting email job on thread ID " + Thread.CurrentThread.ManagedThreadId);

            FileInfo fileInfo = new FileInfo(this.FileName);
            try
            {
                this.Log(string.Concat("Loading file ", fileInfo.Name));
                ExcelPackage package = new ExcelPackage(fileInfo);
                try
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.First<ExcelWorksheet>();
                    this.Log("Loading info...");
                    for (int i = 2; i < sheet.Dimension.End.Row + 1; i++)
                    {
                        var emailFrom = sheet.Cells[i, 1].Text;
                        var password = sheet.Cells[i, 2].Text;
                        var emailTo = sheet.Cells[i, 3].Text;
                        var subject = sheet.Cells[i, 4].Text;
                        var imagePath = sheet.Cells[i, 5].Text;

                        var msg = new MailMessage(emailFrom, emailTo, subject, "");
                        msg.Attachments.Add(new Attachment(imagePath));

                        var client = new SmtpClient("smtp.gmail.com", 587)
                        {
                            Credentials = new NetworkCredential(emailFrom, password),
                            EnableSsl = true
                        };

                        client.Send(msg);

                        this.Log(string.Concat("Loaded data on row ", i - 1));
                       
                        Thread.Sleep(1000);
                        this.Log(string.Concat("Sent email ", i - 1));
                        Thread.Sleep(10000);
                    }
                    this.Log("Finished task");
                }
                finally
                {
                    if (package != null)
                    {
                        ((IDisposable)package).Dispose();
                    }
                }
            }
            catch (Exception exception)
            {
                this.LogException(exception);
            }
        }

        private void LogException(Exception ex)
        {
            MyForm.Invoke(new Action(
                delegate ()
                {
                    MyRichTextBox.Text += ex.Message + Environment.NewLine;
                    MyRichTextBox.Text += ex.StackTrace + Environment.NewLine;
                    MyRichTextBox.SelectionStart = MyRichTextBox.Text.Length;
                    MyRichTextBox.ScrollToCaret();
                }));

            MyForm.Invoke(new Action(
                delegate ()
                {
                    MyForm.Refresh();
                }));
        }

        private void Log(String text)
        {
            MyForm.Invoke(new Action(
                delegate ()
                {
                    MyRichTextBox.Text += text + Environment.NewLine;
                    MyRichTextBox.SelectionStart = MyRichTextBox.Text.Length;
                    MyRichTextBox.ScrollToCaret();
                }));

            MyForm.Invoke(new Action(
                delegate ()
                {
                    MyForm.Refresh();
                }));
        }
    }
}
