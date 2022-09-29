using System.IO.Compression;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SAPFEWSELib;
using SapROTWr;
using Application = FlaUI.Core.Application;


namespace Flanium;

public class Helpers
{
    private static string GetTimeStamp()
    {
        return "[" + DateTime.Now.Day + "." + DateTime.Now.Month.ToString("00") + "." +
               DateTime.Now.Year + " " + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" +
               DateTime.Now.Second + "]";
    }

    public class SystemOperations
    {
        public class Files
        {
            public static void CopyFile(string source, string destination)
            {
                File.Copy(source, destination,true);
            }
            
            public static void CopyFiles(string source, string destination, List<string> files, bool overwrite)
            {
                if (Directory.Exists(destination))
                {
                    foreach (var file in files)
                    {
                        if (File.Exists(destination + "\\" + Path.GetFileName(file)))
                            File.Delete(destination + "\\" + Path.GetFileName(file));

                        CopyFile(source + "\\" + file, destination + "\\" + file);
                        Console.WriteLine(GetTimeStamp() + "Moved file: " + file + " to: " + destination + "\n");
                    }
                }
                else
                {
                    Directory.CreateDirectory(destination);
                    foreach (var file in files)
                    {
                        MoveFile(source + "\\" + file, destination + "\\" + file);
                        Console.WriteLine(GetTimeStamp() + "Moved file: " + file + " to: " + destination + "\n");
                    }
                }
            }
            
            public static void DeleteDuplicateFiles(string folderPath, string duplicateIdentifier = "(")
            {
                var files = Directory.GetFiles(folderPath);

                foreach (var file in files)
                    if (file.Contains(duplicateIdentifier))
                    {
                        File.Delete(file);
                        Console.WriteLine(GetTimeStamp() + "Deleted duplicate file: " + file + "\n");
                    }
            }

            public static void DeleteFile(string filePath)
            {
                if (File.Exists(filePath)) File.Delete(filePath);

                Console.WriteLine(GetTimeStamp() + "Deleted file found at: " + filePath + "\n");
            }

            public static string CreateFolder(string folderPath)
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                    Console.WriteLine(GetTimeStamp() + "Created folder at: " + folderPath + "\n");
                }
                catch
                {
                }

                return folderPath;
            }

            public static void MoveFile(string filePath, string directoryPath)
            {
                File.Move(filePath, directoryPath, true);
                Console.WriteLine(GetTimeStamp() + "Moved file: " + filePath + " to " + directoryPath + "\n");
            }

            public static string RenameFile(string filePath, string newName)
            {
                if (File.Exists(filePath))
                {
                    File.Move(filePath, filePath.Replace(Path.GetFileName(filePath), "") + newName);
                    File.Delete(filePath);
                    Console.WriteLine(GetTimeStamp() + "Renamed file: " + filePath + " to " + newName + "\n");
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
                        if (File.Exists(directoryPath + "\\" + Path.GetFileName(file)))
                            File.Delete(directoryPath + "\\" + Path.GetFileName(file));

                        MoveFile(folderPath + "\\" + file, directoryPath + "\\" + file);
                        Console.WriteLine(GetTimeStamp() + "Moved file: " + file + " to: " + directoryPath + "\n");
                    }
                }
                else
                {
                    Directory.CreateDirectory(directoryPath);
                    foreach (var file in files)
                    {
                        MoveFile(folderPath + "\\" + file, directoryPath + "\\" + file);
                        Console.WriteLine(GetTimeStamp() + "Moved file: " + file + " to: " + directoryPath + "\n");
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
                    if (Directory.GetFiles(folderPath).Length > 0)
                    {
                        Console.WriteLine(GetTimeStamp() + "Folder contains files: " + folderPath + "\n");
                        return true;
                    }

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
                    if (path != null)
                        Directory.CreateDirectory(path);

                Console.WriteLine(GetTimeStamp() + "Created folder at: " + path + "\n");
            }

            public static void DeleteFolder(string folderPath)
            {
                if (Directory.Exists(folderPath)) Directory.Delete(folderPath, true);

                Console.WriteLine(GetTimeStamp() + "Deleted folder at: " + folderPath + "\n");
            }

            public static void ArchiveFolder(string folderPath, string otherPath)
            {
                try
                {
                    if (Directory.Exists(folderPath)) ZipFile.CreateFromDirectory(folderPath, otherPath + ".zip");
                    Console.WriteLine(
                        GetTimeStamp() + "Archived folder: " + folderPath + " to: " + otherPath + "\n");
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
                var engine = sapGuilRot.GetType().InvokeMember("GetScriptingEngine", BindingFlags.InvokeMethod,
                    null,
                    sapGuilRot, null);
                sapp = (GuiApplication) engine;
                conn = (GuiConnection) sapp.Connections.ElementAt(0);
                session = (GuiSession) conn.Children.ElementAt(0);
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

            Console.WriteLine(GetTimeStamp() + "Opened SAP session for user: " + userName + "\n");
            return output;
        }
    }

    public static List<string> GetDownloadedFiles(WebDriver _driver, double refreshRateInSeconds = 0.250)
    {
        var downloads = new List<string>();
        _driver.SwitchTo().Window(_driver.WindowHandles.First());
        _driver.SwitchTo().NewWindow(WindowType.Tab);
        _driver.Navigate().GoToUrl("chrome://downloads/");
        _driver.SwitchTo().Window(_driver.WindowHandles.Last());
        var downloadsHandle = _driver.CurrentWindowHandle;
        Thread.Sleep(1000);

        //Check for danger button "Keep" and press it.
        DownloadChecker:
        Thread.Sleep(TimeSpan.FromSeconds(refreshRateInSeconds));
        _driver.Navigate().Refresh();
        var downloadsManager = _driver.FindElement(By.TagName("downloads-manager"));
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
            _driver.ExecuteScript("arguments[0].click();", closeButton);
            if (closeButton != null)
                goto DownloadChecker;
        }
        catch
        {
            // ignored
        }

        Finish:
        _driver.SwitchTo().Window(downloadsHandle).Close();
        Console.WriteLine(GetTimeStamp() + "Downloaded files: " + string.Join(", ", downloads) + "\n");
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
        client.Credentials = new NetworkCredential {UserName = userName, Password = password};
        client.Port = port;

        try
        {
            client.Send(message);
            Console.WriteLine(GetTimeStamp() + "Email sent to: " + string.Join(", ", recipients) + "\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Could not send the message to {0} due to the following error: {1}",
                string.Join(";", recipients), ex.Message);
        }
    }
}