using System.Data;
using Flanium;
using FlaUI.UIA3;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using Polly;

namespace WebSurfer;

public enum BrowserType
{
    Chrome,
    Firefox
}

public class WebBrowser
{
    private static string GetTimeStamp()
    {
        return "[" + DateTime.Now.Day + "." + DateTime.Now.Month.ToString("00") + "." +
               DateTime.Now.Year + " " + DateTime.Now.Hour + ":" + DateTime.Now.Minute + ":" +
               DateTime.Now.Second + "]";
    }

    public WebBrowser ConnectOrchestrator()
    {
        return this;
    }
    private static WebDriver _driver { get; set; }
    private BrowserType _selectedBrowserType { get; }

    public JS JavaScript = new();
    public FS FlaniumScript = new();

    public WebBrowser(BrowserType browserType, PageLoadStrategy pageLoadStrategy = PageLoadStrategy.Normal)
    {
        switch (browserType)
        {
            case BrowserType.Chrome:
                _driver = Initializers.InitializeChrome(pageLoadStrategy);
                Console.Clear();
                _selectedBrowserType = browserType;
                break;
            case BrowserType.Firefox:
                _driver = Initializers.InitializeFirefox(pageLoadStrategy);
                Console.Clear();
                _selectedBrowserType = browserType;
                break;
        }
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
                            Console.WriteLine("     Found tab with URL: " + url + "\n");
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

        Console.WriteLine("Closing tab with URL: " + url + "\n");
        _driver.SwitchTo().Window(window).Close();
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
                                    Console.WriteLine("     Found tab containing URL: " + url + " and Element: " +
                                                      element.Text + "\n");
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

        Console.WriteLine("Closing tab containing URL: " + url + " and Element: " +
                          element.Text + "\n");
        _driver.SwitchTo().Window(window).Close();
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
        Console.WriteLine(GetTimeStamp() + "Downloaded files: " + string.Join(", ", downloads) + "\n");
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
                            Console.WriteLine(GetTimeStamp() + "Hovered over Element: " + xPath + " (" + element.Text +
                                              ")\n");
                            _driver.ExecuteJavaScript(
                                "arguments[0].dispatchEvent(new MouseEvent('mouseover', {bubbles: true}))", element);
                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;
                Console.WriteLine(GetTimeStamp() + "Failed to hover over Element: " + xPath + "\n");
                throw new Exception("\n Failed to hover over element with XPath: " + xPath + "\n");
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
                            Console.WriteLine(GetTimeStamp() + "Element vanished: " + xPath + "\n");
                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;

                Console.WriteLine(GetTimeStamp() + "Element did not vanish: " + xPath + " ...how? \n");
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
                            Console.WriteLine(GetTimeStamp() + "Element vanished: " + xPath + "\n");
                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;

                Console.WriteLine(GetTimeStamp() + "Element did not vanish: " + xPath + "\n");
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
                        Console.WriteLine(GetTimeStamp() + "Got value of Element: " + xPath + " (" +
                                          _driver.ExecuteJavaScript<string>("return arguments[0].value", element) +
                                          ")\n");
                        return _driver.ExecuteJavaScript<string>("return arguments[0].value", element);
                    }

                    if (_driver.ExecuteJavaScript<string>("return arguments[0].innerText", element) != null)
                    {
                        Console.WriteLine(GetTimeStamp() + "Got value of Element: " + xPath + " (" +
                                          _driver.ExecuteJavaScript<string>("return arguments[0].innerText",
                                              element) +
                                          ")\n");
                        return _driver.ExecuteJavaScript<string>("return arguments[0].innerText", element);
                    }
                }

                Console.WriteLine(GetTimeStamp() + "No value was found for Element: " + xPath + "\n");
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
                            Console.WriteLine(GetTimeStamp() + "Set Value of Element: " + xPath + " (" + element.Text +
                                              ")\n");
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
                Console.WriteLine(GetTimeStamp() + "Failed to set value of Element: " + xPath + "\n");
                throw new Exception("\n Failed to set value of element with XPath: " + xPath + "\n");
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
                            Console.WriteLine(
                                GetTimeStamp() + "Clicked Element: " + xPath + " (" + element.Text + ")\n");
                            _driver.ExecuteScript("arguments[0].click();", element);
                            return element;
                        }

                        return null;
                    });


                if (element != null) return element;

                Console.WriteLine(GetTimeStamp() + "Failed to click element with XPath: " + xPath + "\n");
                throw new Exception("\n Failed to use Javascript to click element with XPath: " + xPath + "\n");
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
                var descendants = _driver.ExecuteJavaScript<List<IWebElement>>(
                    "return arguments[0].querySelectorAll('*')", element).ToList();
                return descendants;
            }

            public IWebElement FindWebElementByXPathWindowless(string XPath,
                string FrameType = "iframe")
            {
                IWebElement element = null;
                Console.WriteLine(GetTimeStamp() + "         Searching for Element: " + XPath + "\n");

                _driver.SwitchTo().DefaultContent();

                try
                {
                    element = _driver.ExecuteJavaScript<IWebElement>(
                        "return document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;",
                        XPath);
                    if (element != null)
                    {
                        Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" + element.Text +
                                          ")\n");
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
                                    Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" +
                                                      element.Text + ")\n");
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
                IWebElement element = null;
                Console.WriteLine(GetTimeStamp() + "         Searching for Element: " + XPath + "\n");

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
                            Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" + element.Text +
                                              ")\n");
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
                                        Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" +
                                                          element.Text + ")\n");
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


                return element;
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
                            Console.WriteLine(GetTimeStamp() + "Element appeared: " + xPath + "\n");
                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                Console.WriteLine(GetTimeStamp() + "Element did not appear: " + xPath + "\n");
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
                            //Get timestamp


                            Console.WriteLine("Element vanished: " + xPath + "\n");
                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;

                Console.WriteLine("Element did not vanish: " + xPath + "\n");
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
                            Console.WriteLine("Element vanished: " + xPath + "\n");
                            return null;
                        }

                        return element;
                    });

                if (element != null) return element;

                Console.WriteLine("Element did not vanish: " + xPath + " ...how? \n");
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
                        Console.WriteLine("Got Text of Element: " + xPath + " (" + element.Text + ")\n");
                        return element.Text;
                    }

                    if (element.GetAttribute("value") != "")
                    {
                        Console.WriteLine("Got Value of Element: " + xPath + " (" + element.GetAttribute("value") +
                                          ")\n");
                        return element.GetAttribute("value");
                    }
                }

                Console.WriteLine("No text or value was found for Element: " + xPath + "\n");
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
                            Console.WriteLine("Hovered over Element: " + xPath + " (" + element.Text + ")\n");
                            var hover = new OpenQA.Selenium.Interactions.Actions(_driver);
                            hover.MoveToElement(element).Perform();
                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;
                Console.WriteLine("Failed to hover over Element: " + xPath + "\n");
                throw new Exception("\n Failed to hover over element with XPath: " + xPath + "\n");
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
                            Console.WriteLine("Clicked Element: " + xPath + " (" + element.Text + ")\n");
                            try
                            {
                                element.Click();
                            }
                            catch (Exception e)
                            {
                                if (e is ElementClickInterceptedException or StaleElementReferenceException)
                                    return null;
                                throw;
                            }

                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                Console.WriteLine("Failed to click Element: " + xPath + "\n");
                throw new Exception("\n Failed to click element with XPath: " + xPath + "\n");
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
                            Console.WriteLine("Set Value of Element: " + xPath + " (" + element.Text + ")\n");
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
                Console.WriteLine("Failed to set value of Element: " + xPath + "\n");
                throw new Exception("\n Failed to set value of element with XPath: " + xPath + "\n");
            }
        }

        public class Searchers
        {
            public List<IWebElement> FindAllChildren(IWebElement element)
            {
                return element.FindElements(By.XPath(".//*")).ToList();
            }

            public List<IWebElement> FindAllDescendants(IWebElement element)
            {
                return element.FindElements(By.CssSelector("*")).ToList();
            }

            public IWebElement FindWebElementByXPath(string XPath,
                string FrameType = "iframe")
            {
                IWebElement element = null;
                Console.WriteLine(GetTimeStamp() + "         Searching for Element: " + XPath + "\n");

                var wHandles = _driver.WindowHandles.Reverse();

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
                            Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" + element.Text +
                                              ")\n");
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
                                        Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" +
                                                          element.Text + ")\n");
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
                Console.WriteLine(GetTimeStamp() + "         Searching for Element: " + XPath + "\n");

                _driver.SwitchTo().DefaultContent();

                try
                {
                    element = _driver.FindElement(By.XPath(XPath));
                    if (element != null)
                    {
                        Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" + element.Text +
                                          ")\n");
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
                                    Console.WriteLine(GetTimeStamp() + "     Found Element: " + XPath + " (" +
                                                      element.Text + ")\n");
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
                            Console.WriteLine(GetTimeStamp() + "Element appeared: " + xPath + "\n");
                            return element;
                        }

                        return null;
                    });

                if (element != null) return element;

                Console.WriteLine(GetTimeStamp() + "Element did not appear: " + xPath + "\n");
                return null;
            }
        }
    }
}