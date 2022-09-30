using System.Data;
using System.Diagnostics;
using System.Management;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using log4net;
using log4net.Appender;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.Extensions;
using Polly;
using WinRT;

[assembly: log4net.Config.XmlConfigurator(ConfigFile = "log4net.config")]

namespace Flanium;

/// <summary>
/// Web Controller class.
/// </summary>
public class WebBrowser
{
    public enum BrowserType
    {
        None,
        Chrome,
        Firefox
    }

    /// <summary>
    /// Disposes of the WebDriver and closes the Chrome window.
    /// </summary>
    public void Dispose()
    {
        _browser.GetDriver().Dispose();
        _browser.GetProcess().CloseMainWindow();
    }

    private WebBrowser _browser;
    private static log4net.ILog _logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

    /// <summary>
    /// WebBrowser constructor.
    /// </summary>
    /// <param name="browserType">Sets the browser type, if None then the Browser will be disabled.</param>
    /// <param name="pageLoadStrategy">Sets the strategy, Selenium documentation covers PageLoadStrategy behavior.</param>
    /// <param name="debugMode">Enables or disables debugging in the console.</param>
    public WebBrowser(BrowserType browserType = BrowserType.None, PageLoadStrategy pageLoadStrategy = PageLoadStrategy.Default, bool debugMode = false)
    {
        if (debugMode == true)
        {
            _logger.Info("Debug mode is on.");
        }
        else
        {
            _logger.Info("Debug mode is off.");
            var root = ((log4net.Repository.Hierarchy.Hierarchy) LogManager.GetRepository()).Root;
            root.RemoveAppender("Console");
        }

        if (browserType != BrowserType.None)
            _browser = new WebBrowser(browserType, pageLoadStrategy);
    }

    private static WebDriver _driver { get; set; }
    private static Process _browserProcess { get; set; }

    private SelectionStrategy _selectionStrategy = SelectionStrategy.Selenium;

    /// <summary>
    /// Gets the Selenium WebDriver object.
    /// </summary>
    /// <returns>WebDriver</returns>
    public WebDriver GetDriver()
    {
        return _driver;
    }

    /// <summary>
    /// Gets the Chrome window process.
    /// </summary>
    /// <returns></returns>
    public Process GetProcess()
    {
        return _browserProcess;
    }

    private BrowserType _selectedBrowserType { get; }

    /// <summary>
    /// WebBrowser constructor.
    /// </summary>
    /// <param name="browserType">Can be None, Chrome or Firefox.</param>
    /// <param name="pageLoadStrategy">See Selenium docs for PageLoadStrategy behavior.</param>
    public WebBrowser(BrowserType browserType, PageLoadStrategy pageLoadStrategy = PageLoadStrategy.Default)
    {
        switch (browserType)
        {
            case BrowserType.Chrome:
                var objectList = Initializers.InitializeChrome(loadStrategy: pageLoadStrategy);
                _driver = objectList[1] as WebDriver;
                _browserProcess = objectList[0] as Process;
                _selectedBrowserType = browserType;
                break;
            case BrowserType.Firefox:
                _driver = Initializers.InitializeFirefox(pageLoadStrategy);
                _selectedBrowserType = browserType;
                break;
        }

        _logger.Info("WebSurfer has started succesfully.");
    }

    /// <summary>
    /// Waits for 1 second, can be configured.
    /// </summary>
    /// <param name="seconds"></param>
    public void Delay(int seconds = 1)
    {
        Thread.Sleep(TimeSpan.FromSeconds(seconds));
    }

    /// <summary>
    /// Navigates to a URL.
    /// </summary>
    /// <param name="url">URL of the website, valid URLs contain 'https://'</param>
    public void NavigateTo(string url)
    {
        _driver.Navigate().GoToUrl(url);
    }

    /// <summary>
    /// Opens a new Window or Tab, depending on the WindowType parameter, then switches to it.
    /// </summary>
    /// <param name="windowType"></param>
    public void Open(WindowType windowType)
    {
        _driver.SwitchTo().NewWindow(windowType);
        _driver.SwitchTo().Window(_driver.WindowHandles.Last());
    }

    /// <summary>
    /// Closes the Window or Tab with the given URL.
    /// </summary>
    /// <param name="url"></param>
    /// <param name="retries"></param>
    /// <param name="retryInterval"></param>
    public void Close(string url, int retries = 15, int retryInterval = 1)
    {
        _driver.SwitchTo().Window(_driver.WindowHandles.First());
        var window = Policy.HandleResult<string>(result => result == null)
            .WaitAndRetry(retries, interval => TimeSpan.FromMilliseconds(retryInterval))
            .Execute(() =>
            {
                var windows = _driver.WindowHandles;
                foreach (var w in windows)
                {
                    _driver.SwitchTo().DefaultContent();
                    try
                    {
                        if (_driver.SwitchTo().Window(w).Url.Contains(url))
                        {
                            _logger.Info("     Found tab with URL: " + url + "");
                            return w;
                        }
                    }
                    catch
                    {
                        // ignored
                    }
                }

                return null;
            });

        if (window is null)
        {
            _logger.Info("Could not find tab with URL: " + url + "");
        }
        else
        {
            _logger.Info("Closing tab with URL: " + url + "");
            _driver.SwitchTo().Window(window).Close();
        }
    }

    /// <summary>
    /// Closes the Window or Tab with the given URL and where the element is present.
    /// </summary>
    /// <param name="url"></param>
    /// <param name="element"></param>
    /// <param name="retries"></param>
    /// <param name="retryInterval"></param>
    public void Close(string url, IWebElement element,
        int retries = 15, int retryInterval = 1)
    {
        _driver.SwitchTo().Window(_driver.WindowHandles.First());
        var window = Policy.HandleResult<string>(result => result == null)
            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
            .Execute(() =>
            {
                var windows = _driver.WindowHandles;
                foreach (var w in windows)
                {
                    _driver.SwitchTo().DefaultContent();

                    try
                    {
                        if (_driver.SwitchTo().Window(w).Url.Contains(url))
                        {
                            try
                            {
                                if (element != null)
                                {
                                    _logger.Info("     Found tab containing URL: " + url + " and Element: " + element.Text + "");
                                    return w;
                                }
                            }
                            catch
                            {
                                return null;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                    }
                }

                return null;
            });

        if (window is null)
        {
            _logger.Info("Could not find tab with URL: " + url + "");
        }
        else
        {
            _logger.Info("Closing tab with URL: " + url + " and Element: " + element.Text + "");
            _driver.SwitchTo().Window(window).Close();
        }
    }

    /// <summary>
    /// Highlights an element for 3 seconds, configurable.
    /// </summary>
    /// <param name="element"></param>
    /// <param name="duration"></param>
    public void Highlight(IWebElement element, int duration = 3000)
    {
        var origStyle = _driver.ExecuteScript("arguments[0].style;", element);
        var script =
            @"arguments[0].style.cssText = ""border-width: 2px; border-style: solid; border-color: red"";";
        _driver.ExecuteScript(script, element);
        Thread.Sleep(duration);
        _driver.ExecuteScript("arguments[0].style=arguments[1]", new[] {element, origStyle});
    }

    /// <summary>
    /// Waits, manages dangerous files and removes files from the download history while also returning the file path.
    /// </summary>
    /// <param name="refreshRateInSeconds"></param>
    /// <returns>A list of files that were downloaded.</returns>
    public List<string> GetDownloadedFiles(double refreshRateInSeconds = 0.250)
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

        _logger.Info("Downloaded files: " + string.Join(", ", downloads) + "");
        return downloads;
    }

    public enum SelectionStrategy
    {
        None,
        Javascript,
        Selenium
    }

    /// <summary>
    /// Sets the framework to be used universally in this WebBrowser session, however, each Action can have their own SelectionStrategy.
    /// </summary>
    /// <param name="selectionStrategy"></param>
    public void SetSelectionStrategy(SelectionStrategy selectionStrategy)
    {
        _selectionStrategy = selectionStrategy;
    }

    /// <summary>
    /// Clicks an element.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns>The clicked element.</returns>
    public IWebElement Click(string XPath, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Actions.Click(XPath),
                SelectionStrategy.Selenium => FS.Actions.Click(XPath),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Actions.Click(XPath),
            SelectionStrategy.Selenium => FS.Actions.Click(XPath),
            _ => null
        };
    }
    /// <returns></returns>
    /// <summary>
    /// Sets the value or text of an element, whichever is applicable.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="value">The value which the element will obtain.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns>The edited element.</returns>
    public IWebElement SetValue(string XPath, string value, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Actions.SetValue(XPath, value),
                SelectionStrategy.Selenium => FS.Actions.SetValue(XPath, value),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Actions.SetValue(XPath, value),
            SelectionStrategy.Selenium => FS.Actions.SetValue(XPath, value),
            _ => null
        };
    }

    /// <summary>
    /// Gets the text or value of an element, whichever is applicable.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns>The text of the element.</returns>
    public string GetText(string XPath, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Actions.GetText(XPath),
                SelectionStrategy.Selenium => FS.Actions.GetText(XPath),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Actions.GetText(XPath),
            SelectionStrategy.Selenium => FS.Actions.GetText(XPath),
            _ => null
        };
    }

    /// <summary>
    /// Waits for an element to vanish.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns>Returns the element if it has not vanished or returns null if it has.</returns>
    public IWebElement WaitElementVanish(string XPath, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Actions.WaitElementVanish(XPath),
                SelectionStrategy.Selenium => FS.Actions.WaitElementVanish(XPath),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Actions.WaitElementVanish(XPath),
            SelectionStrategy.Selenium => FS.Actions.WaitElementVanish(XPath),
            _ => null
        };
    }

    /// <summary>
    /// Waits for an element to vanish forever, use carefully.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns>Returns the element if it has not vanished or returns null if it has.</returns>
    public IWebElement WaitForeverElementVanish(string XPath, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Actions.WaitForeverElementVanish(XPath),
                SelectionStrategy.Selenium => FS.Actions.WaitForeverElementVanish(XPath),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Actions.WaitForeverElementVanish(XPath),
            SelectionStrategy.Selenium => FS.Actions.WaitForeverElementVanish(XPath),
            _ => null
        };
    }

    /// <summary>
    /// Simulates hover over an element.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns>Returns the element that was hovered.</returns>
    public IWebElement Hover(string XPath, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Actions.Hover(XPath),
                SelectionStrategy.Selenium => FS.Actions.Hover(XPath),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Actions.Hover(XPath),
            SelectionStrategy.Selenium => FS.Actions.Hover(XPath),
            _ => null
        };
    }

    /// <summary>
    /// Finds all the immediate children of the element.
    /// </summary>
    /// <param name="element">The element from which to extract the children.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns>A list of IWebElements.</returns>
    public List<IWebElement> FindAllChildren(IWebElement element, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Searchers.FindAllChildren(element),
                SelectionStrategy.Selenium => FS.Searchers.FindAllChildren(element),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Searchers.FindAllChildren(element),
            SelectionStrategy.Selenium => FS.Searchers.FindAllChildren(element),
            _ => null
        };
    }

    /// <summary>
    /// Finds all the descendants of the element, this includes the immediate children and any nesting children inside of them, searches recursively.
    /// </summary>
    /// <param name="element">The element from which to extract the children.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <returns></returns>
    public List<IWebElement> FindAllDescendants(IWebElement element, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Searchers.FindAllDescendants(element),
                SelectionStrategy.Selenium => FS.Searchers.FindAllDescendants(element),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Searchers.FindAllDescendants(element),
            SelectionStrategy.Selenium => FS.Searchers.FindAllDescendants(element),
            _ => null
        };
    }

    /// <summary>
    /// Finds any WebElement by XPath, no matter the iframe it is in.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <param name="FrameType">The type of the frame to look into, e.g. 'iframe' or 'frame' or 'object' for legacy apps.</param>
    /// <returns>The searched element.</returns>
    public IWebElement FindWebElementByXPath(string XPath, SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None,
        string FrameType = "iframe")
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Searchers.FindWebElementByXPath(XPath, FrameType),
                SelectionStrategy.Selenium => FS.Searchers.FindWebElementByXPath(XPath, FrameType),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Searchers.FindWebElementByXPath(XPath, FrameType),
            SelectionStrategy.Selenium => FS.Searchers.FindWebElementByXPath(XPath, FrameType),
            _ => null
        };
    }

    /// <summary>
    /// Waits for an element to appear.
    /// </summary>
    /// <param name="XPath">Element XPath.</param>
    /// <param name="overwriteSelectionStrategy">Allows you to select a different selection strategy, such as Javascript or Selenium.</param>
    /// <param name="FrameType">The type of the frame to look into, e.g. 'iframe' or 'frame' or 'object' for legacy apps.</param>
    /// <returns>The element that appeared.</returns>
    public IWebElement WaitElementAppear(string XPath, string FrameType = "iframe", SelectionStrategy overwriteSelectionStrategy = SelectionStrategy.None)
    {
        if (overwriteSelectionStrategy != SelectionStrategy.None)
        {
            return overwriteSelectionStrategy switch
            {
                SelectionStrategy.Javascript => JS.Searchers.WaitElementAppear(XPath, FrameType: FrameType),
                SelectionStrategy.Selenium => FS.Searchers.WaitElementAppear(XPath, FrameType: FrameType),
                _ => null
            };
        }

        return _selectionStrategy switch
        {
            SelectionStrategy.Javascript => JS.Searchers.WaitElementAppear(XPath, FrameType: FrameType),
            SelectionStrategy.Selenium => FS.Searchers.WaitElementAppear(XPath, FrameType: FrameType),
            _ => null
        };
    }
    public static class JS
    {
        public static class Actions
        {
            public static IWebElement Hover(string xPath, int retries = 15,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            _logger.Info("Hovered over Element: " + xPath + " (" + element.Text + ")");

                            _driver.ExecuteJavaScript(
                                "arguments[0].dispatchEvent(new MouseEvent('mouseover', {bubbles: true}))", element);
                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;
                _logger.Info("Failed to hover over Element: " + xPath + "");

                throw new Exception(" Failed to hover over element with XPath: " + xPath + "");
            }

            public static IWebElement WaitForeverElementVanish(string xPath,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result != null)
                    .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element == null)
                        {
                            _logger.Info("Element vanished: " + xPath + "");

                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;

                _logger.Info("Element did not vanish: " + xPath + " ...how? ");
                return null;
            }

            public static IWebElement WaitElementVanish(string xPath, int retries = 60,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result != null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element == null)
                        {
                            _logger.Info("Element vanished: " + xPath + "");
                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;


                _logger.Info("Element did not vanish: " + xPath + "");
                return null;
            }

            public static string GetText(string xPath, int retries = 15, int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            if (_driver.ExecuteJavaScript<string>("return arguments[0].innerText", element) ==
                                null &&
                                _driver.ExecuteJavaScript<string>("return arguments[0].value", element) == null)
                                return null;

                            return element;
                        }

                        return null;
                    });

                if (element != null)
                {
                    if (_driver.ExecuteJavaScript<string>("return arguments[0].value", element) != null)
                    {
                        _logger.Info("Got value from Element: " + xPath + " (" +
                                     _driver.ExecuteJavaScript<string>("return arguments[0].value", element) + ")");

                        return _driver.ExecuteJavaScript<string>("return arguments[0].value", element);
                    }

                    if (_driver.ExecuteJavaScript<string>("return arguments[0].innerText", element) != null)
                    {
                        _logger.Info("Got text from Element: " + xPath + " (" +
                                     _driver.ExecuteJavaScript<string>("return arguments[0].innerText", element) +
                                     ")");

                        return _driver.ExecuteJavaScript<string>("return arguments[0].innerText", element);
                    }
                }


                _logger.Info("Failed to get text from Element: " + xPath + "");
                return "";
            }

            public static IWebElement SetValue(string xPath, string text, int retries = 15,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            _logger.Info("Set value of Element: " + xPath + " (" + element.Text + ")");
                            try
                            {
                                _driver.ExecuteJavaScript("arguments[0].value = ''", element);
                                _driver.ExecuteJavaScript("arguments[0].value = arguments[1];", element, text);
                            }
                            catch
                            {
                                // ignored
                            }

                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                _logger.Info("Failed to set value of Element: " + xPath + "");
                throw new Exception(" Failed to set value of element with XPath: " + xPath + "");
            }

            public static IWebElement Click(string xPath, int retries = 15,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            _logger.Info("Clicked Element: " + xPath + " (" + element.Text + ")");

                            try
                            {
                                _driver.ExecuteScript("arguments[0].click();", element);
                                return element;
                            }
                            catch (Exception e)
                            {
                                if (e is ElementClickInterceptedException or StaleElementReferenceException)
                                    return null;
                            }
                        }

                        return null;
                    });


                if (element != null) return element;


                _logger.Info("Failed to click Element: " + xPath + "");
                throw new Exception(" Failed to use Javascript to click element with XPath: " + xPath + "");
            }
        }

        public static class Searchers
        {
            public static List<IWebElement> FindAllChildren(IWebElement element)
            {
                var children = _driver.ExecuteJavaScript<List<IWebElement>>(
                    "return arguments[0].childNodes", element).ToList();
                return children;
            }

            public static List<IWebElement> FindAllDescendants(IWebElement element)
            {
                var descendants = _driver.ExecuteJavaScript<object>("return arguments[0].querySelectorAll('*');", element)
                    .As<IEnumerable<IWebElement>>().ToList();
                return descendants;
            }

            public static IWebElement FindWebElementByXPath(string XPath,
                string FrameType = "iframe")
            {
                IWebElement element;

                _logger.Info("         Searching for Element: " + XPath + "");

                var wHandles = _driver.WindowHandles.Reverse();

                foreach (var w in wHandles)
                {
                    //============================================================
                    _driver.SwitchTo().Window(w);
                    _driver.SwitchTo().DefaultContent();

                    try
                    {
                        element = _driver.ExecuteJavaScript<IWebElement>(
                            "return document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;",
                            XPath);
                        if (element != null)
                        {
                            _logger.Info("     Found Element: " + XPath + " (" + element.Text + ")");

                            return element;
                        }
                    }
                    catch
                    {
                    }

                    //============================================================
                    var iframes = _driver.FindElements(By.TagName(FrameType)).ToList();
                    var index = 0;

                    Reset:
                    for (; index < iframes.Count; index++)
                    {
                        retry:
                        try
                        {
                            var item = iframes[index];

                            try
                            {
                                _driver.SwitchTo().Frame(item);
                                try
                                {
                                    element = _driver.ExecuteJavaScript<IWebElement>(
                                        "return document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;",
                                        XPath);
                                    if (element != null)
                                    {
                                        _logger.Info("     Found Element: " + XPath + " (" + element.Text + ")");
                                        return element;
                                    }
                                }
                                catch
                                {
                                }
                            }
                            catch
                            {
                                _driver.SwitchTo().ParentFrame();
                                goto retry;
                            }


                            var children = _driver.FindElements(By.TagName(FrameType)).ToList();

                            if (children.Count > 0)
                                if (iframes.Any(x => children.Any(y => y.Equals(x))) == false)
                                {
                                    iframes.InsertRange(index + 1, children);
                                    goto Reset;
                                }
                        }
                        catch
                        {
                        }
                    }
                }

                return null;
            }

            public static IWebElement WaitElementAppear(string xPath, int retries = 15,
                int retryInterval = 1, string FrameType = "iframe")
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = FindWebElementByXPath(xPath, FrameType);
                        if (element != null)
                        {
                            _logger.Info("Element Element: " + xPath + "");

                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;


                _logger.Info("Element did not appear: " + xPath + "");
                return null;
            }
        }
    }

    /// <summary>
    /// The Selenium class library.
    /// </summary>
    public static class FS
    {
        public static class Actions
        {
            public static DataTable DataScraping(IWebElement tableElement)
            {
                var table = WebEvents.Search.FindAllDescendants(tableElement as WebElement)
                    .Single(e => e.TagName == "table");

                var headers = table.FindElements(By.TagName("thead"));
                WebElement header = null;
                foreach (var h in headers)
                {
                    if (h.GetAttribute("innerHtml") == null)
                    {
                        continue;
                    }

                    if (h.GetAttribute("innerHtml") != null)
                    {
                        header = h as WebElement;
                        break;
                    }
                }

                var headerColumns = WebEvents.Search.FindAllDescendants(header).Where(e => e.TagName == "th");
                var headerColumnNames = headerColumns.Select(e => e.Text).ToList();

                var dataTable = new DataTable();

                foreach (var headerColumnName in headerColumnNames)
                {
                    dataTable.Columns.Add(headerColumnName);
                    dataTable.Columns[headerColumnName].DataType = typeof(string);
                }

                var tableBody = table.FindElement(By.TagName("tbody"));
                var rows = tableBody.FindElements(By.TagName("tr"));

                foreach (var row in rows)
                {
                    var rowColumns = WebEvents.Search.FindAllDescendants(row as WebElement)
                        .Where(e => e.TagName == "td");
                    var rowColumnValues = rowColumns.Select(e => e.Text).ToList();
                    dataTable.Rows.Add(rowColumnValues.ToArray());
                }

                return dataTable;
            }


            public static IWebElement WaitElementVanish(string xPath, int retries = 60,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result != null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element == null)
                        {
                            _logger.Info("Element vanished: " + xPath + "");
                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;


                _logger.Info("Element did not vanish: " + xPath + "");
                return null;
            }

            public static IWebElement WaitForeverElementVanish(string xPath,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result != null)
                    .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element == null)
                        {
                            _logger.Info("Element vanished: " + xPath + "");
                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;


                _logger.Info("Element did not vanish: " + xPath + "");
                return null;
            }

            public static string GetText(string xPath, int retries = 15, int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            if (element.Text == "")
                                return null;
                            if (element.GetAttribute("value") == "")
                                return null;

                            return element;
                        }

                        return null;
                    });

                if (element != null)
                {
                    if (element.Text != "")
                    {
                        _logger.Info("Got Text of Element: " + xPath + " (" + element.Text + ")");
                        return element.Text;
                    }

                    if (element.GetAttribute("value") != "")
                    {
                        _logger.Info("Got Value of Element: " + xPath + " (" + element.GetAttribute("value") + ")");
                        return element.GetAttribute("value");
                    }
                }


                _logger.Info("Could not get text of Element: " + xPath + "");
                return "";
            }

            public static IWebElement Hover(string xPath, int retries = 15,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            _logger.Info("Hovered over Element: " + xPath + " (" + element.Text + ")");
                            var hover = new OpenQA.Selenium.Interactions.Actions(_driver);
                            hover.MoveToElement(element).Perform();
                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                _logger.Info("Could not hover over Element: " + xPath + "");
                throw new Exception(" Failed to hover over element with XPath: " + xPath + "");
            }

            public static bool WaitForAlert(List<string> acceptedAlertText, int retries = 15,
                int retryInterval = 1)
            {
                var alert = Policy.HandleResult<IAlert>(alert => alert == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval)).Execute(() =>
                    {
                        IAlert alert = null;
                        try
                        {
                            alert = _driver.SwitchTo().Alert();
                        }
                        catch (Exception e)
                        {
                        }

                        return alert;
                    });

                if (alert == null)
                {
                    return false;
                }

                var alertText = alert.Text.Trim();
                if (acceptedAlertText.Any(x => alertText.Contains(x)))
                {
                    alert.Accept();
                    return true;
                }

                alert.Dismiss();
                return false;
            }

            public static IWebElement Click(string xPath, int retries = 15,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            _logger.Info("Clicked Element: " + xPath + " (" + element.Text + ")");
                            try
                            {
                                element.Click();
                            }
                            catch (Exception e)
                            {
                                if (e is ElementClickInterceptedException or StaleElementReferenceException)
                                    return null;
                            }

                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                _logger.Info("Could not click Element: " + xPath + "");
                throw new Exception(" Failed to click element with XPath: " + xPath + "");
            }

            public static IWebElement SetValue(string xPath, string text, int retries = 15,
                int retryInterval = 1)
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        var element = Searchers.FindWebElementByXPath(xPath);
                        if (element != null)
                        {
                            _logger.Info("Set Value of Element: " + xPath + " (" + text + ")");
                            try
                            {
                                element.Clear();
                                element.SendKeys(text);
                            }
                            catch (ElementNotInteractableException e)
                            {
                                element.Click();
                                Thread.Sleep(100);
                                element.Clear();
                                element.SendKeys(text);
                            }


                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                _logger.Info("Could not set value of Element: " + xPath + "");
                throw new Exception(" Failed to set value of element with XPath: " + xPath + "");
            }
        }

        public static class Searchers
        {
            public static List<IWebElement> FindAllChildren(IWebElement element)
            {
                return element.FindElements(By.XPath(".//*")).ToList();
            }

            public static List<IWebElement> FindAllDescendants(IWebElement element)
            {
                return element.FindElements(By.CssSelector("*")).ToList();
            }

            public static IWebElement FindWebElementByXPath(string XPath,
                string FrameType = "iframe")
            {
                IWebElement element = null;

                _logger.Info("         Searching for Element: " + XPath + "");

                var wHandles = _driver.WindowHandles.Reverse().ToArray();

                foreach (var w in wHandles)
                {
                    //============================================================
                    _driver.SwitchTo().Window(w);
                    _driver.SwitchTo().DefaultContent();

                    try
                    {
                        element = _driver.FindElement(By.XPath(XPath));
                        if (element != null)
                        {
                            _logger.Info("     Found Element: " + XPath + " (" + element.Text + ")");
                            return element;
                        }
                    }
                    catch
                    {
                    }

                    //============================================================
                    var iframes = _driver.FindElements(By.TagName(FrameType)).ToList<IWebElement>();
                    var index = 0;

                    Reset:
                    for (; index < iframes.Count; index++)
                    {
                        retry:
                        try
                        {
                            var item = iframes[index];

                            try
                            {
                                _driver.SwitchTo().Frame(item);
                                try
                                {
                                    element = _driver.FindElement(By.XPath(XPath));
                                    if (element != null)
                                    {
                                        _logger.Info("     Found Element: " + XPath + " (" + element.Text + ")");
                                        return element;
                                    }
                                }
                                catch
                                {
                                }
                            }
                            catch
                            {
                                _driver.SwitchTo().ParentFrame();
                                goto retry;
                            }


                            var children = _driver.FindElements(By.TagName(FrameType)).ToList();

                            if (children.Count > 0)
                            {
                                if (iframes.Any(x => children.Any(y => y.Equals(x))) == false)
                                {
                                    iframes.InsertRange(index + 1, children);
                                    goto Reset;
                                }
                            }
                        }
                        catch
                        {
                        }
                    }
                }

                return element;
            }

            public static IWebElement FindWebElementByXPathWindowless(string XPath,
                string FrameType = "iframe")
            {
                IWebElement element = null;

                _logger.Info("         Searching for Element: " + XPath + "");

                _driver.SwitchTo().DefaultContent();

                try
                {
                    element = _driver.FindElement(By.XPath(XPath));
                    if (element != null)
                    {
                        _logger.Info("     Found Element: " + XPath + " (" + element.Text + ")");
                        return element;
                    }
                }
                catch
                {
                }

                //============================================================
                var iframes = _driver.FindElements(By.TagName(FrameType)).ToList<IWebElement>();
                var index = 0;

                Reset:
                for (; index < iframes.Count; index++)
                {
                    retry:
                    try
                    {
                        var item = iframes[index];

                        try
                        {
                            _driver.SwitchTo().Frame(item);
                            try
                            {
                                element = _driver.FindElement(By.XPath(XPath));
                                if (element != null)
                                {
                                    _logger.Info("     Found Element: " + XPath + " (" + element.Text + ")");
                                    return element;
                                }
                            }
                            catch
                            {
                            }
                        }
                        catch
                        {
                            _driver.SwitchTo().ParentFrame();
                            goto retry;
                        }


                        var children = _driver.FindElements(By.TagName(FrameType)).ToList();

                        if (children.Count > 0)
                        {
                            if (iframes.Any(x => children.Any(y => y.Equals(x))) == false)
                            {
                                iframes.InsertRange(index + 1, children);
                                goto Reset;
                            }
                        }
                    }
                    catch
                    {
                    }
                }

                return element;
            }

            public static IWebElement WaitElementAppear(string xPath, int retries = 15,
                int retryInterval = 1, string FrameType = "iframe")
            {
                var element = Policy.HandleResult<IWebElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() =>
                    {
                        IWebElement element = FindWebElementByXPath(xPath, FrameType);
                        if (element != null)
                        {
                            _logger.Info("Element appeared: " + xPath + "");
                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                _logger.Info("Element did not appear: " + xPath + "");
                return null;
            }
        }
    }
}