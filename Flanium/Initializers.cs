﻿using OpenQA.Selenium;
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
    public static ChromeDriver InitializeChrome(PageLoadStrategy loadStrategy = PageLoadStrategy.Normal)
    {
        var options = new ChromeOptions();
        
        var service = ChromeDriverService.CreateDefaultService();
        service.EnableVerboseLogging = false;
        service.SuppressInitialDiagnosticInformation = true;
        service.HideCommandPromptWindow = false;
        
        options.PageLoadStrategy = loadStrategy;
        
        options.AddArgument("--no-sandbox");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--disable-crash-reporter");
        options.AddArgument("--disable-extensions");
        options.AddArgument("--disable-in-process-stack-traces");
        options.AddArgument("--disable-logging");
        options.AddArgument("--disable-dev-shm-usage");
        options.AddArgument("--log-level=3");
        options.AddArgument("--output=/dev/null");


        options.AcceptInsecureCertificates = true;

        options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
        options.AddUserProfilePreference("profile.default_content_settings.popups", 0);
        options.AddUserProfilePreference("download.default_directory", @"C:\Users\" + Environment.UserName + @"\Downloads");
        options.AddUserProfilePreference("download.prompt_for_download", false);
        options.AddUserProfilePreference("safebrowsing.enabled", false);

        options.AddArgument("--safebrowsing-disable-download-protection");
        options.AddArgument("--safebrowsing-disable-extension-blacklist");
            
        return new ChromeDriver(service, options);

    }
}