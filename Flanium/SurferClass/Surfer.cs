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

public class Surfer
{
    private WebBrowser _browser;
    private WinBrowser _winBrowser;
    private static log4net.ILog _logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

    public Surfer(BrowserType browserType = BrowserType.None, PageLoadStrategy pageLoadStrategy = PageLoadStrategy.Default, bool debugMode = false)
    {
        if (debugMode == true)
        {
            _logger.Info("Debug mode is on.");
        }
        else
        {
            _logger.Info("Debug mode is off.");
            var root = ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository()).Root;
            root.RemoveAppender("Console");
        }

        if (browserType != BrowserType.None)
            _browser = new WebBrowser(browserType, pageLoadStrategy);
        _winBrowser = new WinBrowser();
    }

    public WebBrowser Browser => _browser;
    public WinBrowser winBrowser => _winBrowser;

    public void Dispose()
    {
        _browser.GetDriver().Dispose();
    }

    public enum BrowserType
    {
        None,
        Chrome,
        Firefox
    }

    public class WebBrowser
    {
        private static WebDriver _driver { get; set; }

        public WebDriver GetDriver()
        {
            return _driver;
        }

        private BrowserType _selectedBrowserType { get; }

        public JS JavaScript = new();
        public FS FlaniumScript = new();

        public WebBrowser(BrowserType browserType, PageLoadStrategy pageLoadStrategy = PageLoadStrategy.Normal)
        {
            switch (browserType)
            {
                case BrowserType.Chrome:
                    _driver = Initializers.InitializeChrome(loadStrategy: pageLoadStrategy);
                    _selectedBrowserType = browserType;
                    break;
                case BrowserType.Firefox:
                    _driver = Initializers.InitializeFirefox(pageLoadStrategy);
                    _selectedBrowserType = browserType;
                    break;
            }
            _logger.Info("WebSurfer has started succesfully.");
        }

        public void Delay(int seconds)
        {
            Thread.Sleep(TimeSpan.FromSeconds(seconds));
        }

        public void NavigateTo(string url)
        {
            _driver.Navigate().GoToUrl(url);
        }

        public void Open(WindowType windowType)
        {
            _driver.SwitchTo().NewWindow(windowType);
            _driver.SwitchTo().Window(_driver.WindowHandles.Last());
        }

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

            if(window is null)
            {
                _logger.Info("Could not find tab with URL: " + url  + "");
            }
            else
            {
                _logger.Info("Closing tab with URL: " + url + " and Element: " + element.Text + "");
                _driver.SwitchTo().Window(window).Close();
            }
        }

        public void Highlight(IWebElement element, int duration = 3000)
        {
            var origStyle = _driver.ExecuteScript("arguments[0].style;", element);
            var script =
                @"arguments[0].style.cssText = ""border-width: 2px; border-style: solid; border-color: red"";";
            _driver.ExecuteScript(script, element);
            Thread.Sleep(duration);
            _driver.ExecuteScript("arguments[0].style=arguments[1]", new[] {element, origStyle});
        }

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

        public class JS
        {
            public Actions Action = new();
            public Searchers Searcher = new();

            public class Actions
            {
                private Searchers Searcher = new();

                public IWebElement Hover(string xPath, int retries = 15,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public IWebElement WaitForeverElementVanish(string xPath,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result != null)
                        .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public IWebElement WaitElementVanish(string xPath, int retries = 60,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result != null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public string GetText(string xPath, int retries = 15, int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public IWebElement SetValue(string xPath, string text, int retries = 15,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public IWebElement Click(string xPath, int retries = 15,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

            public class Searchers
            {
                public List<IWebElement> FindAllChildren(IWebElement element)
                {
                    var children = _driver.ExecuteJavaScript<List<IWebElement>>(
                        "return arguments[0].childNodes", element).ToList();
                    return children;
                }

                public List<IWebElement> FindAllDescendants(IWebElement element)
                {
                    var descendants = _driver.ExecuteJavaScript<object>("return arguments[0].querySelectorAll('*');", element)
                        .As<IEnumerable<IWebElement>>().ToList();
                    return descendants;
                }

                public IWebElement FindWebElementByXPathWindowless(string XPath,
                    string FrameType = "iframe")
                {
                    IWebElement element = null;

                    _logger.Info("         Searching for Element: " + XPath + "");

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

                    return element;
                }

                public IWebElement FindWebElementByXPath(string XPath,
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

                public IWebElement WaitElementAppear(string xPath, int retries = 15,
                    int retryInterval = 1, string FrameType = "iframe")
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = new Searchers().FindWebElementByXPath(xPath, FrameType);
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

        public class FS
        {
            public Searchers Searcher = new();
            public Actions Action = new();

            public class Actions
            {
                private Searchers Searcher = new();

                public DataTable DataScraping(IWebElement tableElement)
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


                public IWebElement WaitElementVanish(string xPath, int retries = 60,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result != null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public IWebElement WaitForeverElementVanish(string xPath,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result != null)
                        .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public string GetText(string xPath, int retries = 15, int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public IWebElement Hover(string xPath, int retries = 15,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public bool WaitForAlert(List<string> acceptedAlertText, int retries = 15,
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

                public IWebElement Click(string xPath, int retries = 15,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

                public IWebElement SetValue(string xPath, string text, int retries = 15,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            var element = Searcher.FindWebElementByXPath(xPath);
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

            public class Searchers
            {
                public List<IWebElement> FindAllChildren(IWebElement element)
                {
                    return element.FindElements(By.XPath("*")).ToList();
                }

                public List<IWebElement> FindAllDescendants(IWebElement element)
                {
                    return element.FindElements(By.CssSelector("*")).ToList();
                }

                public IWebElement FindWebElementByXPath(string XPath,
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

                public IWebElement FindWebElementByXPathWindowless(string XPath,
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

                public IWebElement WaitElementAppear(string xPath, int retries = 15,
                    int retryInterval = 1)
                {
                    var element = Policy.HandleResult<IWebElement>(result => result == null)
                        .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                        .Execute(() =>
                        {
                            IWebElement element = new Searchers().FindWebElementByXPath(xPath);
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

    public class WinBrowser
    {
        public Searchers Search = new();
        public Actions Action = new();

        public WinBrowser()
        {
            _logger.Info("WinSurfer has started succesfully.");
        }

        public class Searchers
        {
            public int GetDriverProcessId(ChromeDriverService driverService)
            {
                //Get all the childs generated by the driver like conhost, chrome.exe...
                var mos = new ManagementObjectSearcher(
                    $"Select * From Win32_Process Where ParentProcessID={driverService.ProcessId}");
                foreach (var mo in mos.Get())
                {
                    var processId = Convert.ToInt32(mo["ProcessID"]);
                    return processId;
                }

                return 0;
            }

            public Window GetWindow(int processId)
            {
                var automation = new UIA3Automation();
                var allProcesses = Process.GetProcesses();
                //loop through processes
                foreach (var item in allProcesses)
                    if (item.Id == processId)
                        return automation.FromHandle(item.MainWindowHandle).AsWindow();

                return null;
            }

            public Window GetWindow(string xPath, int retries = 15, double retryInterval = 1)
            {
                var automation = new UIA3Automation();
                var desktop = automation.GetDesktop();
                var window = Polly.Policy.HandleResult<Window>(result => result == null)
                    .WaitAndRetry(retries, retryAttempt => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() => desktop.FindFirstByXPath(xPath).AsWindow());

                return window;
            }

            public List<Window> GetWindows(string xPath, int retries = 15, double retryInterval = 1)
            {
                var automation = new UIA3Automation();
                var desktop = automation.GetDesktop();
                var window = Polly.Policy.HandleResult<List<Window>>(result => result.Count == 0)
                    .WaitAndRetry(retries, retryAttempt => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() => desktop.FindAllByXPath(xPath).Cast<Window>().ToList());

                return window;
            }

            public AutomationElement FindElement(Window window, string xPath, int retries = 15,
                double retryInterval = 1)
            {
                var element = Policy.HandleResult<AutomationElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() => window.FindFirstByXPath(xPath));

                return element;
            }

            public List<AutomationElement> FindElements(Window window, string xPath, int retries = 15,
                double retryInterval = 1)
            {
                var elements = Policy.HandleResult<List<AutomationElement>>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() => window.FindAllByXPath(xPath).ToList());

                return elements;
            }

            public AutomationElement FindElementInElement(AutomationElement element, string xPath,
                int retries = 15, double retryInterval = 1)
            {
                var welement = Policy.HandleResult<AutomationElement>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() => element.FindFirstByXPath(xPath));

                return welement;
            }

            public List<AutomationElement> FindElementsInElement(AutomationElement element, string xPath,
                int retries = 15, double retryInterval = 1)
            {
                var welements = Policy.HandleResult<List<AutomationElement>>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                    .Execute(() => element.FindAllByXPath(xPath).ToList());

                return welements;
            }
        }

        public class Actions
        {
            public string DesktopScreenshot(string saveToPath)
            {
                var automation = new UIA3Automation();
                var desktop = automation.GetDesktop();
                desktop.CaptureToFile(saveToPath);
                return saveToPath;
            }

            public string ElementScreenshot(AutomationElement element, string saveToPath)
            {
                element.CaptureToFile(saveToPath);
                return saveToPath;
            }

            public string GetText(AutomationElement element, int retries = 15, double retryInterval = 1)
            {
                var retryGetText = Policy.HandleResult<string>(result => result == null)
                    .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                var retryGetTextResult = retryGetText.Execute(() =>
                {
                    try
                    {
                        return element.AsTextBox().Text;
                    }
                    catch (Exception e)
                    {
                        return null;
                    }
                });
                return retryGetTextResult;
            }

            public bool SendText(AutomationElement element, string text, bool eventTrigger = false, int retries = 15,
                double retryInterval = 1)
            {
                if (element != null)
                {
                    if (eventTrigger == false)
                    {
                        var retryClick = Policy.HandleResult<bool>(result => result == true)
                            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                        var retryResult = retryClick.Execute(() =>
                        {
                            try
                            {
                                element.AsTextBox().Text = text;

                                while (element.AsTextBox().Text != text) Thread.Sleep(250);
                            }
                            catch (Exception e)
                            {
                                // ignored
                            }

                            return false;
                        });

                        return retryResult;
                    }
                    else
                    {
                        var retryClick = Policy.HandleResult<bool>(result => result == true)
                            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                        var retryResult = retryClick.Execute(() =>
                        {
                            try
                            {
                                element.AsTextBox().Enter(text);

                                while (element.AsTextBox().Text != text) Thread.Sleep(250);
                            }
                            catch (Exception e)
                            {
                                // ignored
                            }

                            return false;
                        });

                        return retryResult;
                    }
                }

                return false;
            }

            public bool Click(AutomationElement element, bool invoke = false, int retries = 15,
                double retryInterval = 1)
            {
                if (element != null)
                {
                    if (invoke == false)
                    {
                        var retryClick = Policy.HandleResult<bool>(result => result == true)
                            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                        var retryResult = retryClick.Execute(() =>
                        {
                            try
                            {
                                element.Click();
                                Thread.Sleep(100);
                            }
                            catch (Exception e)
                            {
                                // ignored
                            }

                            return false;
                        });

                        return retryResult;
                    }
                    else
                    {
                        var retryClick = Policy.HandleResult<bool>(result => result == true)
                            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                        var retryResult = retryClick.Execute(() =>
                        {
                            try
                            {
                                element.AsButton().Invoke();
                                Thread.Sleep(100);
                            }
                            catch (Exception e)
                            {
                                // ignored
                            }

                            return false;
                        });

                        return retryResult;
                    }
                }

                return false;
            }

            public bool DoubleClick(AutomationElement element, bool invoke = false, int retries = 15,
                double retryInterval = 1)
            {
                if (element != null)
                {
                    if (invoke == false)
                    {
                        var retryClick = Policy.HandleResult<bool>(result => result == true)
                            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                        var retryResult = retryClick.Execute(() =>
                        {
                            try
                            {
                                element.DoubleClick();
                                Thread.Sleep(100);
                            }
                            catch (Exception e)
                            {
                                // ignored
                            }

                            return false;
                        });

                        return retryResult;
                    }
                    else
                    {
                        var retryClick = Policy.HandleResult<bool>(result => result == true)
                            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                        var retryResult = retryClick.Execute(() =>
                        {
                            try
                            {
                                element.AsButton().Invoke();
                                element.AsButton().Invoke();
                                Thread.Sleep(100);
                            }
                            catch (Exception e)
                            {
                                // ignored
                            }

                            return false;
                        });

                        return retryResult;
                    }
                }

                return false;
            }

            public bool RightClick(AutomationElement element, bool invoke = false, int retries = 15,
                double retryInterval = 1)
            {
                if (element != null)
                {
                    if (invoke == false)
                    {
                        var retryClick = Policy.HandleResult<bool>(result => result == true)
                            .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval));
                        var retryResult = retryClick.Execute(() =>
                        {
                            try
                            {
                                element.RightClick();
                                Thread.Sleep(100);
                            }

                            catch (Exception e)
                            {
                                // ignored
                            }

                            return false;
                        });

                        return retryResult;
                    }
                }

                return false;
            }

            public void SendKeys(VirtualKeyShort keyShort)
            {
                FlaUI.Core.Input.Keyboard.Press(keyShort);
                FlaUI.Core.Input.Keyboard.Release(keyShort);
            }

            public void SendKeyCombination(VirtualKeyShort[] keyShorts)
            {
                var keyCombination = FlaUI.Core.Input.Keyboard.Pressing(keyShorts);
                keyCombination.Dispose();
            }
        }
    }
}