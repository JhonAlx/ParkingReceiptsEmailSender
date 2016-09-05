using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace seleniumTest
{
    class MainProcess
    {
        public Form MyForm { get; set; }
        public RichTextBox MyRichTextBox { get; set; }
        public String FolderName { get; set; }
        
        private const String ACCESS_URL = "";
        
        private IWebDriver driver;

        public void Run()
        {
            Log("Starting scraping job on thread ID " + Thread.CurrentThread.ManagedThreadId);

            var options = new ChromeOptions();
            driver = new ChromeDriver(options);

            Log("Trying to access extension on thread ID: " + Thread.CurrentThread.ManagedThreadId);

            FillData();
        }

        private void FillData()
        {
            driver.Quit();
            Log("Process finished on thread ID " + Thread.CurrentThread.ManagedThreadId);
        }

        private void LogException(Exception ex)
        {
            MyForm.Invoke(new Action(
                delegate()
                {
                    MyRichTextBox.Text += ex.Message + Environment.NewLine;
                    MyRichTextBox.Text += ex.StackTrace + Environment.NewLine;
                    MyRichTextBox.SelectionStart = MyRichTextBox.Text.Length;
                    MyRichTextBox.ScrollToCaret();
                }));

            MyForm.Invoke(new Action(
                delegate()
                {
                    MyForm.Refresh();
                }));
        }

        private void Log(String text)
        {
            MyForm.Invoke(new Action(
                delegate()
                {
                    MyRichTextBox.Text += text + Environment.NewLine;
                    MyRichTextBox.SelectionStart = MyRichTextBox.Text.Length;
                    MyRichTextBox.ScrollToCaret();
                }));

            MyForm.Invoke(new Action(
                delegate()
                {
                    MyForm.Refresh();
                }));
        }
    }
}
