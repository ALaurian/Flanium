using System.Diagnostics;
using FlaUI.UIA3;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;

namespace Flanium;

public class Initializers
{
    public static FirefoxDriver InitializeFirefox(PageLoadStrategy loadStrategy = PageLoadStrategy.Normal)
    {
        var options = new FirefoxOptions();
        var service = FirefoxDriverService.CreateDefaultService();

        service.SuppressInitialDiagnosticInformation = true;
        service.HideCommandPromptWindow = false;
        options.PageLoadStrategy = loadStrategy;
        options.AcceptInsecureCertificates = true;

        return new FirefoxDriver(service, options);
    }

    public static ChromeDriver InitializeChrome(string appPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", string port = "8080", PageLoadStrategy loadStrategy = PageLoadStrategy.Normal)
    {
        var process = new Process();
        var processStartInfo = new ProcessStartInfo
        {
            FileName = appPath,
            Arguments = "--remote-debugging-port=" + port,
            UseShellExecute = false,
        };
        process.StartInfo = processStartInfo;
            
        var options = new ChromeOptions
        {
            DebuggerAddress = "localhost:8080"
        };
            
        options.AddArgument("--no-sandbox");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--disable-crash-reporter");
        options.AddArgument("--disable-extensions");
        options.AddArgument("--disable-in-process-stack-traces");
        options.AddArgument("--disable-logging");
        options.AddArgument("--disable-dev-shm-usage");
        options.AddArgument("--log-level=3");
        options.AddArgument("--output=/dev/null");
        options.AddArgument("--safebrowsing-disable-download-protection");
        options.AddArgument("--safebrowsing-disable-extension-blacklist");
        options.AddArgument("--force-renderer-accessibility");
            
            
        var service = ChromeDriverService.CreateDefaultService();
        service.EnableVerboseLogging = false;
        service.SuppressInitialDiagnosticInformation = true;
        service.HideCommandPromptWindow = false;

        options.PageLoadStrategy = loadStrategy;
        process.Start();
        var Driver = new ChromeDriver(service, options);
            
        return Driver;
    }
}