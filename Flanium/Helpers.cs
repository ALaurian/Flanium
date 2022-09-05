using System.Data;
using System.Diagnostics;
using System.IO.Compression;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SAPFEWSELib;
using SapROTWr;
using Application = FlaUI.Core.Application;
using static Flanium.WebEvents.Search;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;
using TimeSpan = ABI.System.TimeSpan;


namespace Flanium;

public class Helpers
{
    public class SystemOperations
    {
        public class Files
        {
            public static void DeleteDuplicateFiles(string folderPath, string duplicateIdentifier = "(")
            {
                var files = Directory.GetFiles(folderPath);

                foreach (var file in files)
                    if (file.Contains(duplicateIdentifier)) File.Delete(file);
            }
            public static void DeleteFile(string filePath)
            {
                if (File.Exists(filePath)) File.Delete(filePath);
            }

            public static string CreateFolder(string folderPath)
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch
                {
                    
                }
                return folderPath;
            }
            public static void MoveFile(string filePath, string directoryPath)
            {
                    File.Move(filePath, directoryPath,true); 
            }

            public static string RenameFile(string filePath, string newName)
            {
                if (File.Exists(filePath))
                {
                    File.Move(filePath, filePath.Replace(Path.GetFileName(filePath), "") + newName);
                    File.Delete(filePath);
                    return filePath.Replace(Path.GetFileName(filePath), "") + newName;
                }

                return null;
            }
        
            public static void MoveFiles(List<string> files, string folderPath, string directoryPath)
            {
                if (Directory.Exists(folderPath))
                {
                    foreach (var file in files)
                    {
                        if (File.Exists(directoryPath + "\\" + Path.GetFileName(file))) File.Delete(directoryPath + "\\" + Path.GetFileName(file));

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

        }

        public class Folders
        {
            public static bool FolderContainsFiles(string folderPath)
            {
                try
                {
                    if (Directory.GetFiles(folderPath).Length > 0) return true;

                    return false;
                }
                catch
                {
                    return false;
                }
            }

            public static void CreateFolder(string path)
            {
                if (!Directory.Exists(path))
                    if (path != null) Directory.CreateDirectory(path);
            }
        
            public static void DeleteFolder(string folderPath)
            {
                if (Directory.Exists(folderPath)) Directory.Delete(folderPath, true);
            }
        
            public static void ArchiveFolder(string folderPath, string otherPath)
            {
                try
                {
                    if (Directory.Exists(folderPath)) ZipFile.CreateFromDirectory(folderPath, otherPath + ".zip");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            } 
        }

        
    }

    public class SAP
    {
        public static Dictionary<object, object> OpenSession(string userName, string sapPw, string sapConnection,
            string sapClientId, string sapshcutPath = @"C:\Program Files (x86)\SAP\FrontEnd\SapGui\sapshcut.exe")
        {
            var output = new Dictionary<object, object>();

            var app = Application.Launch(sapshcutPath,
                "-user=" + userName + " -pw=" + sapPw + " -system=SID -sysname=" + '"' + sapConnection + '"' +
                " -language=EN -client=" + sapClientId);

            RetrySAPConnection:
            GuiApplication sapp = null;
            GuiConnection conn = null;
            GuiSession session = null;
            try
            {
                var sapRotWrapper = new CSapROTWrapper();
                var sapGuilRot = sapRotWrapper.GetROTEntry("SAPGUI");
                var engine = sapGuilRot.GetType().InvokeMember("GetScriptingEngine", BindingFlags.InvokeMethod, null,
                    sapGuilRot, null);
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
            AutomationElement window = WinEvents.Search.GetWindow("*[contains(@Name,'" + "SAP Easy Access" + "']");
        
            output.Add("WINDOW", window);
            output.Add("SAPGUI", sapp);
            output.Add("SAPCONNECTION", conn);
            output.Add("SAPSESSION", session);
            return output;
        }
   
    }
    
    public static List<string> HandleDownloads(ChromeDriver chromeDriver, double refreshRateInSeconds = 0.250)
    {
        var downloads = new List<string>();
        chromeDriver.SwitchTo().Window(chromeDriver.WindowHandles.First());
        chromeDriver.SwitchTo().NewWindow(WindowType.Tab);
        chromeDriver.Navigate().GoToUrl("chrome://downloads/");
        chromeDriver.SwitchTo().Window(chromeDriver.WindowHandles.Last());
        var downloadsHandle = chromeDriver.CurrentWindowHandle;
        Thread.Sleep(1000);
        
        //Check for danger button "Keep" and press it.
        DownloadChecker:
        Thread.Sleep(System.TimeSpan.FromSeconds(refreshRateInSeconds));
        chromeDriver.Navigate().Refresh();
        var downloadsManager = chromeDriver.FindElement(By.TagName("downloads-manager"));
        var downloadsManagerShadowRoot = downloadsManager.GetShadowRoot();
        var mainContainer = downloadsManagerShadowRoot.FindElement(By.Id("mainContainer"));
        var ironList = mainContainer.FindElement(By.TagName("iron-list"));
        var downloadsItemList = ironList.FindElements(By.TagName("downloads-item")).ToList();
        ISearchContext dlShadowRoot = null;

        //Check if item is blocked
        try
        {
            dlShadowRoot = downloadsItemList[0].GetShadowRoot();
        }
        catch
        {
            goto Finish;
        }

        try
        {
            var statusDanger = dlShadowRoot.FindElement(By.Id("dangerous")).FindElements(By.TagName("span"))[1]
                .FindElements(By.TagName("cr-button"))[0];
            if (statusDanger != null && statusDanger.Displayed)
            {
                    
                statusDanger.Click();
                Thread.Sleep(100);
                var keepAnywayButton = new UIA3Automation().GetDesktop().FindAllChildren()
                    .Single(x =>
                        x.Name.Contains("Downloads - Google Chrome") || x.Name.Contains("Confirm download"))
                    .FindFirstByXPath("//Button[@Name='Keep anyway']");

                WinEvents.Action.Click(keepAnywayButton, true);
                //keepAnywayButton.AsButton().Invoke();
                // Thread.Sleep(500);
                    
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

    public static void SendEmail(string from, string[] recipients, string subject, string body, string server,
        string userName, string password, int port)
    {
        // Create a message and set up the recipients.
        var message = new MailMessage();


        message.From = new MailAddress(from);
        foreach (var r in recipients) message.To.Add(r);

        message.Subject = subject;
        message.Body = body;
        message.IsBodyHtml = true;

        
        //Send the message.
        var client = new SmtpClient(server);
        client.Credentials = new NetworkCredential { UserName = userName, Password = password };
        client.Port = port;

        try
        {
            client.Send(message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Could not send the message to {0} due to the following error: {1}", string.Join(";",recipients), ex.Message);
        }
    }
}