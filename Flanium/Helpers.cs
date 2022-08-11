﻿using System.Data;
using System.IO.Compression;
using System.Net;
    using System.Net.Mail;
    using System.Reflection;
using System.Windows.Forms;
    using FlaUI.Core.AutomationElements;
    using FlaUI.UIA3;
    using Microsoft.Office.Interop.Excel;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using SAPFEWSELib;
    using SapROTWr;
    using Application = FlaUI.Core.Application;
    using DataTable = System.Data.DataTable;
    using Range = Microsoft.Office.Interop.Excel.Range;

    namespace Flanium
    {
        internal class Helpers
        {
            public static Dictionary<object, object> SapSession(string userName, string sapPw, string sapConnection, string sapClientId, string sapshcutPath = @"C:\Program Files (x86)\SAP\FrontEnd\SapGui\sapshcut.exe")
            {
                Dictionary<object, object> output = new Dictionary<object, object>();

                var app = Application.Launch(sapshcutPath, "-user=" + userName + " -pw=" + sapPw + " -system=SID -sysname=" + '"' + sapConnection + '"' + " -language=EN -client=" + sapClientId);

                RetrySAPConnection:
                GuiApplication sapp = null;
                GuiConnection conn = null;
                GuiSession session = null;
                try
                {
                    CSapROTWrapper sapRotWrapper = new CSapROTWrapper();
                    object sapGuilRot = sapRotWrapper.GetROTEntry("SAPGUI");
                    object engine = sapGuilRot.GetType().InvokeMember("GetScriptingEngine", BindingFlags.InvokeMethod, null, sapGuilRot, null);
                    sapp = (GuiApplication)engine;
                    conn = (GuiConnection)sapp.Connections.ElementAt(0);
                    session = (GuiSession)conn.Children.ElementAt(0);
                }
                catch
                {
                    // ignored
                }

                if (session == null)
                    goto RetrySAPConnection;

                session.ActiveWindow.SendVKey(0);
                session.ActiveWindow.SendVKey(0);
                AutomationElement window = WinEvents.GetWindowByLinq(x => x.Name.Contains("SAP Easy Access"));
                
                output.Add("WINDOW", window);
                output.Add("SAPGUI", sapp);
                output.Add("SAPCONNECTION", conn);
                output.Add("SAPSESSION", session);
                return output;

            }

            //Check if folder contains files
            public static bool FolderContainsFiles(string folderPath)
            {
                try
                {
                    if (Directory.GetFiles(folderPath).Length > 0)
                    {
                        return true;
                    }

                    return false;
                }
                catch
                {
                    return false;
                }

            }

            //Delete duplicate files from folder
            public static void DeleteDuplicateFiles(string folderPath, string duplicateIdentifier = "(")
            {
                string[] files = Directory.GetFiles(folderPath);

                foreach (var file in files)
                {
                    if (file.Contains(duplicateIdentifier))
                    {
                        File.Delete(file);
                    }
                }

            }

            //Create Folder
            public static void CreateFolder(string path)
            {
                if (!Directory.Exists(path))
                {
                    if (path != null) Directory.CreateDirectory(path);
                }
            }

            //Delete folder
            public static void DeleteFolder(string folderPath)
            {

                    if (Directory.Exists(folderPath))
                    {
                        Directory.Delete(folderPath, true);
                    }

            }

            //Archive folder

            public static void ArchiveFolder(string folderPath, string otherPath)
            {
                try
                {
                    if (Directory.Exists(folderPath))
                    {
                        ZipFile.CreateFromDirectory(folderPath, otherPath + ".zip");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }


            //Delete file
            public static void DeleteFile(string filePath)
            {

                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

            }
            //Move file 
            public static void MoveFile(string filePath, string directoryPath)
            {
                if (Directory.Exists(directoryPath))
                {
                    if (File.Exists(directoryPath + "\\" + Path.GetFileName(filePath)))
                    {
                        File.Delete(directoryPath + "\\" + Path.GetFileName(filePath));
                    }
                    File.Move(filePath, directoryPath + "\\" + Path.GetFileName(filePath));
                }
                else
                {
                    if (directoryPath == null) return;
                    Directory.CreateDirectory(directoryPath);
                    File.Move(filePath, directoryPath + "\\" + Path.GetFileName(filePath));
                }
            }

            //Move Files
            public static void MoveFiles(List<string> files, string folderPath, string directoryPath)
            {
                if (Directory.Exists(folderPath))
                {
                    foreach (var file in files)
                    {
                        if (File.Exists(directoryPath + "\\" + Path.GetFileName(file)))
                        {
                            File.Delete(directoryPath + "\\" + Path.GetFileName(file));
                        }
                        MoveFile(folderPath + "\\" + file, directoryPath);
                        Console.WriteLine("\nMoved file: " + file + " to: " + directoryPath);
                    }
                }
                else
                {
                    Directory.CreateDirectory(directoryPath);
                    foreach (var file in files)
                    {
                        MoveFile(folderPath + "\\" + file, directoryPath);
                        Console.WriteLine("\nMoved file: " + file + " to: " + directoryPath);
                    }
                }
            }
            public static DataTable ExcelToDataTable(string path, int sheet = 1, int rowToDelete = -1, bool deleteFile = true)
            {
                Console.WriteLine(" ");
                Console.WriteLine("----------------------------");
                Console.WriteLine("Converting Excel Sheet 1 to DataTable");
                Console.WriteLine("----------------------------");
                Console.WriteLine(" ");

                DataTable dataTable = new DataTable();

                //Open Excel Workbook interop

                var xlApp = new Microsoft.Office.Interop.Excel.Application();
                var xlWorkBook = xlApp.Workbooks.Open(path);
                var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(sheet);

                if (rowToDelete!=-1)
                    ((Range)xlWorkSheet.Rows[rowToDelete]).Delete();
                
                var myCells = xlWorkSheet.UsedRange;
                myCells.Select();
                myCells.Copy();

                //Get Clipboard value
                string clipboardText = Clipboard.GetText(TextDataFormat.UnicodeText);
                string[] excelRows = clipboardText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                //Create datatablecolumns from index 0
                for (int i = 0; i < excelRows[0].Split('\t').Length; i++)
                {
                    dataTable.Columns.Add(excelRows[0].Split('\t')[i]);
                }

                //Create datatablerows from index 1+
                for (int i = 1; i < excelRows.Length; i++)
                {
                    dataTable.Rows.Add();
                    for (int j = 0; j < excelRows[i].Split('\t').Length; j++)
                    {
                        try
                        {
                            dataTable.Rows[i - 1][j] = excelRows[i].Split('\t')[j];
                        }
                        catch (Exception)
                        {
                            // ignored
                        }
                    }
                }


                //Close Excel Workbook

                xlApp.DisplayAlerts = false;

                xlWorkBook.Close();
                xlApp.Quit();

                if(deleteFile)
                    File.Delete(path);

                return dataTable;

            }
            
            public static DataTable ProcessDataTable(ChromeDriver chromeDriver, DataTable inputTable)
            {
                Console.WriteLine(" ");
                Console.WriteLine("----------------------------");
                Console.WriteLine("Filtering the DataTable");
                Console.WriteLine("----------------------------");
                Console.WriteLine(" ");

                var table = inputTable.AsEnumerable()
                    .Where(r => r["Remaining TAT"].ToString() == "0") 
                    .OrderBy(r => r["Remaining TAT"])
                    .CopyToDataTable();

                //Sorting the Table
                //DataView dataView = inputTable.DefaultView;
                //dataView.Sort = "Remaining TAT ASC";
                //inputTable = dataView.ToTable();

                Console.WriteLine("Finished filtering the DataTable.");

                return table;
            }
            public static void Highlight(IWebElement context, ChromeDriver driver)
            {
                var rc = context;
                var script = @"arguments[0].style.cssText = ""border-width: 2px; border-style: solid; border-color: red""; ";
                driver.ExecuteScript(script, rc);
                Thread.Sleep(3000);
                var clear = @"arguments[0].style.cssText = ""border-width: 0px; border-style: solid; border-color: red""; ";
                driver.ExecuteScript(clear, rc);

            }
            public static List<string> DownloadListener(ChromeDriver chromeDriver)
            {
                List<string> downloads = new List<string>();
                chromeDriver.SwitchTo().NewWindow(WindowType.Tab);
                chromeDriver.Navigate().GoToUrl("chrome://downloads/");
                chromeDriver.SwitchTo().Window(chromeDriver.WindowHandles.Last());
                var downloadsHandle = chromeDriver.CurrentWindowHandle;

            //Check for danger button "Keep" and press it.
            DownloadChecker: 
                Thread.Sleep(1000);
                chromeDriver.Navigate().Refresh();
                var downloadsManager = chromeDriver.FindElement(By.TagName("downloads-manager"));
                var downloadsManagerShadowRoot = downloadsManager.GetShadowRoot();
                var mainContainer = downloadsManagerShadowRoot.FindElement(By.Id("mainContainer"));
                var ironList = mainContainer.FindElement(By.TagName("iron-list"));
                var downloadsItemList = ironList.FindElements(By.TagName("downloads-item")).ToList();
                ISearchContext dlShadowRoot = null;

                //Check if item is blocked
                try { dlShadowRoot = downloadsItemList[0].GetShadowRoot(); }
                catch { goto Finish; }

                try
                {
                    var statusDanger = dlShadowRoot.FindElement(By.Id("dangerous")).FindElements(By.TagName("span"))[1].FindElements(By.TagName("cr-button"))[0];
                    if (statusDanger != null && statusDanger.Displayed)
                    {
                        statusDanger.Click();
                        var keepAnywayButton = new UIA3Automation().GetDesktop().FindAllChildren().Single(x => x.Name.Contains("Downloads - Google Chrome") || x.Name.Contains("Confirm download")).FindFirstByXPath("//Button[@Name='Keep anyway']");
                        keepAnywayButton.AsButton().Invoke();

                        goto DownloadChecker;
                    }

                }
                catch
                {
                    // ignored
                }

                //Check if item is downloading
                try
                {
                    var statusDownloading = dlShadowRoot.FindElement(By.Id("description")).Text;

                    if (statusDownloading != "" && statusDownloading.Contains("/s -"))
                        goto DownloadChecker;
                }
                catch
                {
                    // ignored
                }

                //Check if item finished downloading and remove it from the list
                try
                {
                    var shadowRoot = downloadsItemList[0].GetShadowRoot();
                    var closeButton = shadowRoot.FindElement(By.ClassName("icon-clear"));
                    downloads.Add(shadowRoot.FindElement(By.Id("file-link")).Text);
                    chromeDriver.ExecuteScript("arguments[0].click();", closeButton);
                    if (closeButton != null)
                        goto DownloadChecker;
                }
                catch
                {
                    // ignored
                }

                Finish:
                chromeDriver.SwitchTo().Window(downloadsHandle).Close();
                return downloads;

            }

            public static void CloseTab(ChromeDriver chromeDriver, string url, string anchorXpath = "", int timeoutTries = 10, int msWaitTime = 1000)
            {
                Console.WriteLine(" ");
                Console.WriteLine("----------------------------");
                Console.WriteLine("Closing Tab with URL: " + url);
                Console.WriteLine("----------------------------");
                Console.WriteLine(" ");

                bool queueBreak = false;
                Thread.Sleep(msWaitTime);
                while (timeoutTries != 0)
                {
                    foreach (var windows in chromeDriver.WindowHandles)
                    {
                        var windowUrl = chromeDriver.SwitchTo().Window(windows).Url;

                        if (windowUrl.Contains(url))
                        {
                            if (anchorXpath != "")
                            {
                                if (WebEvents.FindWebElementByXPath(chromeDriver, anchorXpath) != null)
                                {
                                    chromeDriver.SwitchTo().Window(windows).Close();
                                    queueBreak = true;
                                    break;
                                }
                            }

                            if (anchorXpath == "")
                            {
                                chromeDriver.SwitchTo().Window(windows).Close();
                                queueBreak = true;
                                break;
                            }

                        }
                    }

                    if (queueBreak)
                    {
                        break;
                    }

                    Thread.Sleep(msWaitTime);
                    timeoutTries--;
                }
            }
            public static void SendEmail(string from, string[] recipients, string subject, string body, string server, string userName, string password, int port)
            {
                // Create a message and set up the recipients.
                MailMessage message = new MailMessage();


                message.From = new MailAddress(from);
                foreach (var r in recipients)
                {
                    message.To.Add(r);
                }
                message.Subject = subject;
                message.Body = body;


                //Send the message.
                SmtpClient client = new SmtpClient(server);
                client.Credentials = new NetworkCredential { UserName = userName, Password = password };
                client.Port = port;

                try
                {
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught in CreateMessageWithAttachment(): {0}",
                        ex);
                }
            }

        }
    }