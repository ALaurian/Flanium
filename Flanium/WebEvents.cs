using System.ComponentModel.Design;
using System.Data;
using System.Windows.Forms;
using FlaUI.Core;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools.V102.Page;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using Polly;
using static Flanium.WebEvents.Search;

namespace Flanium;

public class WebEvents
{
    public class ActionJs
    {
        public static IWebElement Hover(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = SearchJs.FindWebElementByXPath(chromeDriver, xPath);
                    if (element != null)
                    {
                        Console.WriteLine("Hovered over Element: " + xPath + " (" + element.Text + ")\n");
                        chromeDriver.ExecuteJavaScript(
                            "arguments[0].dispatchEvent(new MouseEvent('mouseover', {bubbles: true}))", element);
                        return element;
                    }

                    return null;
                });

            if (element != null) return element;
            Console.WriteLine("Failed to hover over Element: " + xPath + "\n");
            throw new Exception("\n Failed to hover over element with XPath: " + xPath + "\n");
        }

        public static IWebElement WaitForeverElementVanish(ChromeDriver chromeDriver, string xPath,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = SearchJs.FindWebElementByXPath(chromeDriver, xPath);
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

        public static IWebElement WaitElementVanish(ChromeDriver chromeDriver, string xPath, int retries = 60,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = SearchJs.FindWebElementByXPath(chromeDriver, xPath);
                    if (element == null)
                    {
                        Console.WriteLine("Element vanished: " + xPath + "\n");
                        return null;
                    }

                    return element;
                });

            if (element != null) return element;

            Console.WriteLine("Element did not vanish: " + xPath + "\n");
            return null;
        }

        public static string GetText(ChromeDriver chromeDriver, string xPath, int retries = 15, int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = SearchJs.FindWebElementByXPath(chromeDriver, xPath);
                    if (element != null)
                    {
                        if (chromeDriver.ExecuteJavaScript<string>("return arguments[0].innerText", element) == null &&
                            chromeDriver.ExecuteJavaScript<string>("return arguments[0].value", element) == null)
                            return null;

                        return element;
                    }

                    return null;
                });

            if (element != null)
            {
                if (chromeDriver.ExecuteJavaScript<string>("return arguments[0].value", element) != null)
                {
                    Console.WriteLine("Got value of Element: " + xPath + " (" +
                                      chromeDriver.ExecuteJavaScript<string>("return arguments[0].value", element) +
                                      ")\n");
                    return chromeDriver.ExecuteJavaScript<string>("return arguments[0].value", element);
                }

                if (chromeDriver.ExecuteJavaScript<string>("return arguments[0].innerText", element) != null)
                {
                    Console.WriteLine("Got value of Element: " + xPath + " (" +
                                      chromeDriver.ExecuteJavaScript<string>("return arguments[0].innerText", element) +
                                      ")\n");
                    return chromeDriver.ExecuteJavaScript<string>("return arguments[0].innerText", element);
                }
            }

            Console.WriteLine("No value was found for Element: " + xPath + "\n");
            return "";
        }

        public static IWebElement SetValue(ChromeDriver chromeDriver, string xPath, string text, int retries = 15,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = WebEvents.SearchJs.FindWebElementByXPath(chromeDriver, xPath);
                    if (element != null)
                    {
                        Console.WriteLine("Set Value of Element: " + xPath + " (" + element.Text + ")\n");
                        try
                        {
                            chromeDriver.ExecuteJavaScript("arguments[0].value = ''", element);
                            chromeDriver.ExecuteJavaScript("arguments[0].value = arguments[1];", element, text);
                        }
                        catch
                        {
                        }

                        return element;
                    }

                    return null;
                });

            if (element != null) return element;
            Console.WriteLine("Failed to set value of Element: " + xPath + "\n");
            throw new Exception("\n Failed to set value of element with XPath: " + xPath + "\n");
        }


        public static IWebElement Click(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = WebEvents.SearchJs.FindWebElementByXPath(chromeDriver, xPath);
                    if (element != null)
                    {
                        Console.WriteLine("Used Javascript to click Element: " + xPath + " (" + element.Text + ")\n");
                        chromeDriver.ExecuteScript("arguments[0].click();", element);
                        return element;
                    }

                    return null;
                });


            if (element != null) return element;

            Console.WriteLine("Failed to use Javascript to click element with XPath: " + xPath + "\n");
            throw new Exception("\n Failed to use Javascript to click element with XPath: " + xPath + "\n");
        }
    }

    public class SearchJs
    {
        public static List<IWebElement> FindAllChildren(ChromeDriver chromeDriver, IWebElement element)
        {
            var children = chromeDriver.ExecuteJavaScript<List<IWebElement>>(
                "return arguments[0].childNodes", element).ToList();
            return children;
        }

        public static List<IWebElement> FindAllDescendants(ChromeDriver chromeDriver, IWebElement element)
        {
            var descendants = chromeDriver.ExecuteJavaScript<List<IWebElement>>(
                "return arguments[0].querySelectorAll('*')", element).ToList();
            return descendants;
        }

        public static IWebElement FindWebElementByXPathWindowless(ChromeDriver chromeDriver, string XPath,
            string FrameType = "iframe")
        {
            IWebElement element = null;
            Console.WriteLine("         Searching for Element: " + XPath + "\n");

            chromeDriver.SwitchTo().DefaultContent();

            try
            {
                element = chromeDriver.ExecuteJavaScript<IWebElement>(
                    "return document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;",
                    XPath);
                if (element != null)
                {
                    Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                    return element;
                }
            }
            catch
            {
            }

            //============================================================
            var iframes = chromeDriver.FindElements(By.TagName(FrameType)).ToList();
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
                        chromeDriver.SwitchTo().Frame(item);
                        try
                        {
                            element = chromeDriver.ExecuteJavaScript<IWebElement>(
                                "return document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;",
                                XPath);
                            if (element != null)
                            {
                                Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                                return element;
                            }
                        }
                        catch
                        {
                        }
                    }
                    catch
                    {
                        chromeDriver.SwitchTo().ParentFrame();
                        goto retry;
                    }


                    var children = chromeDriver.FindElements(By.TagName(FrameType)).ToList();

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

        public static IWebElement FindWebElementByXPath(ChromeDriver chromeDriver, string XPath,
            string FrameType = "iframe")
        {
            IWebElement element = null;
            Console.WriteLine("         Searching for Element: " + XPath + "\n");

            var wHandles = chromeDriver.WindowHandles.Reverse();

            foreach (var w in wHandles)
            {
                //============================================================
                chromeDriver.SwitchTo().Window(w);
                chromeDriver.SwitchTo().DefaultContent();

                try
                {
                    element = chromeDriver.ExecuteJavaScript<IWebElement>(
                        "return document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;",
                        XPath);
                    if (element != null)
                    {
                        Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                        return element;
                    }
                }
                catch
                {
                }

                //============================================================
                var iframes = chromeDriver.FindElements(By.TagName(FrameType)).ToList();
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
                            chromeDriver.SwitchTo().Frame(item);
                            try
                            {
                                element = chromeDriver.ExecuteJavaScript<IWebElement>(
                                    "return document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;",
                                    XPath);
                                if (element != null)
                                {
                                    Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                                    return element;
                                }
                            }
                            catch
                            {
                            }
                        }
                        catch
                        {
                            chromeDriver.SwitchTo().ParentFrame();
                            goto retry;
                        }


                        var children = chromeDriver.FindElements(By.TagName(FrameType)).ToList();

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

        public static IWebElement WaitElementAppear(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1, string FrameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    IWebElement element = SearchJs.FindWebElementByXPath(chromeDriver, xPath, FrameType);
                    if (element != null)
                    {
                        Console.WriteLine("Element appeared: " + xPath + "\n");
                        return element;
                    }

                    return null;
                });

            if (element != null) return element;

            Console.WriteLine("Element did not appear: " + xPath + "\n");
            return null;
        }
    }

    public class Search
    {
        public static List<IWebElement> FindAllChildren(IWebElement element)
        {
            return element.FindElements(By.XPath(".//*")).ToList();
        }

        public static List<IWebElement> FindAllDescendants(IWebElement element)
        {
            return element.FindElements(By.CssSelector("*")).ToList();
        }

        public static IWebElement FindWebElementByXPath(ChromeDriver chromeDriver, string XPath,
            string FrameType = "iframe")
        {
            IWebElement element = null;
            Console.WriteLine("         Searching for Element: " + XPath + "\n");

            var wHandles = chromeDriver.WindowHandles.Reverse();

            foreach (var w in wHandles)
            {
                //============================================================
                chromeDriver.SwitchTo().Window(w);
                chromeDriver.SwitchTo().DefaultContent();

                try
                {
                    element = chromeDriver.FindElement(By.XPath(XPath));
                    if (element != null)
                    {
                        Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                        return element;
                    }
                }
                catch
                {
                }

                //============================================================
                var iframes = chromeDriver.FindElements(By.TagName(FrameType)).ToList<IWebElement>();
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
                            chromeDriver.SwitchTo().Frame(item);
                            try
                            {
                                element = chromeDriver.FindElement(By.XPath(XPath));
                                if (element != null)
                                {
                                    Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                                    return element;
                                }
                            }
                            catch
                            {
                            }
                        }
                        catch
                        {
                            chromeDriver.SwitchTo().ParentFrame();
                            goto retry;
                        }


                        var children = chromeDriver.FindElements(By.TagName(FrameType)).ToList();

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

        public static IWebElement FindWebElementByXPathWindowless(ChromeDriver chromeDriver, string XPath,
            string FrameType = "iframe")
        {
            IWebElement element = null;
            Console.WriteLine("         Searching for Element: " + XPath + "\n");

            chromeDriver.SwitchTo().DefaultContent();

            try
            {
                element = chromeDriver.FindElement(By.XPath(XPath));
                if (element != null)
                {
                    Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                    return element;
                }
            }
            catch
            {
            }

            //============================================================
            var iframes = chromeDriver.FindElements(By.TagName(FrameType)).ToList<IWebElement>();
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
                        chromeDriver.SwitchTo().Frame(item);
                        try
                        {
                            element = chromeDriver.FindElement(By.XPath(XPath));
                            if (element != null)
                            {
                                Console.WriteLine("     Found Element: " + XPath + " (" + element.Text + ")\n");
                                return element;
                            }
                        }
                        catch
                        {
                        }
                    }
                    catch
                    {
                        chromeDriver.SwitchTo().ParentFrame();
                        goto retry;
                    }


                    var children = chromeDriver.FindElements(By.TagName(FrameType)).ToList();

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


        public static IWebElement WaitElementAppear(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    IWebElement element = FindWebElementByXPath(chromeDriver, xPath);
                    if (element != null)
                    {
                        Console.WriteLine("Element appeared: " + xPath + "\n");
                        return element;
                    }

                    return null;
                });

            if (element != null) return element;

            Console.WriteLine("Element did not appear: " + xPath + "\n");
            return null;
        }
    }

    public class Action
    {
        public static DataTable DataScraping(IWebElement tableElement)
        {
            var table = WebEvents.Search.FindAllDescendants(tableElement as WebElement)
                .Single(e => e.TagName == "table");

            var headers = table.FindElements(By.TagName("thead"));
            WebElement header = null;
            foreach (var h in headers)
            {
                if(h.GetAttribute("innerHtml") == null)
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
                var rowColumns = WebEvents.Search.FindAllDescendants(row as WebElement).Where(e => e.TagName == "td");
                var rowColumnValues = rowColumns.Select(e => e.Text).ToList();
                dataTable.Rows.Add(rowColumnValues.ToArray());
            }
        
            return dataTable;
        }

        public static void CloseTab(ChromeDriver chromeDriver, string url, int retries = 15, int retryInterval = 1)
        {
            chromeDriver.SwitchTo().Window(chromeDriver.WindowHandles.First());
            var window = Policy.HandleResult<string>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromMilliseconds(retryInterval))
                .Execute(() =>
                {
                    var windows = chromeDriver.WindowHandles;
                    foreach (var w in windows)
                    {
                        chromeDriver.SwitchTo().DefaultContent();
                        try
                        {
                            if (chromeDriver.SwitchTo().Window(w).Url.Contains(url))
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
            chromeDriver.SwitchTo().Window(window).Close();
        }

        public static void CloseTabAnchorable(ChromeDriver chromeDriver, string url, string elementXPath,
            int retries = 15, int retryInterval = 1)
        {
            chromeDriver.SwitchTo().Window(chromeDriver.WindowHandles.First());
            var window = Policy.HandleResult<string>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var windows = chromeDriver.WindowHandles;
                    foreach (var w in windows)
                    {
                        chromeDriver.SwitchTo().DefaultContent();

                        try
                        {
                            if (chromeDriver.SwitchTo().Window(w).Url.Contains(url))
                            {
                                try
                                {
                                    if (FindWebElementByXPathWindowless(chromeDriver, elementXPath) != null)
                                    {
                                        Console.WriteLine("     Found tab containing URL: " + url +
                                                          " and with Element with XPath: " + elementXPath + "\n");
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

            Console.WriteLine("Closing tab containing URL: " + url + " and with Element with XPath: " + elementXPath +
                              "\n");
            chromeDriver.SwitchTo().Window(window).Close();
        }

        public static IWebElement WaitElementVanish(ChromeDriver chromeDriver, string xPath, int retries = 60,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath);
                    if (element == null)
                    {
                        Console.WriteLine("Element vanished: " + xPath + "\n");
                        return null;
                    }

                    return element;
                });

            if (element != null) return element;

            Console.WriteLine("Element did not vanish: " + xPath + "\n");
            return null;
        }

        public static IWebElement WaitForeverElementVanish(ChromeDriver chromeDriver, string xPath,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath);
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

        public static string GetText(ChromeDriver chromeDriver, string xPath, int retries = 15, int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath);
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
                    Console.WriteLine("Got Value of Element: " + xPath + " (" + element.GetAttribute("value") + ")\n");
                    return element.GetAttribute("value");
                }
            }

            Console.WriteLine("No text or value was found for Element: " + xPath + "\n");
            return "";
        }

        public static IWebElement Hover(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath);
                    if (element != null)
                    {
                        Console.WriteLine("Hovered over Element: " + xPath + " (" + element.Text + ")\n");
                        var hover = new Actions(chromeDriver);
                        hover.MoveToElement(element).Perform();
                        return element;
                    }

                    return null;
                });

            if (element != null) return element;
            Console.WriteLine("Failed to hover over Element: " + xPath + "\n");
            throw new Exception("\n Failed to hover over element with XPath: " + xPath + "\n");
        }


        public static bool WaitForAlert(ChromeDriver chromeDriver, List<string> acceptedAlertText, int retries = 15,
            int retryInterval = 1)
        {
            var alert = Policy.HandleResult<IAlert>(alert => alert == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval)).Execute(() =>
                {
                    IAlert alert = null;
                    try
                    {
                        alert = chromeDriver.SwitchTo().Alert();
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


        public static IWebElement Click(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath);
                    if (element != null)
                    {
                        Console.WriteLine("Clicked Element: " + xPath + " (" + element.Text + ")\n");
                        try
                        {
                            element.Click();
                        }
                        catch (ElementClickInterceptedException)
                        {
                            return null;
                        }

                        return element;
                    }

                    return null;
                });

            if (element != null) return element;

            Console.WriteLine("Failed to click Element: " + xPath + "\n");
            throw new Exception("\n Failed to click element with XPath: " + xPath + "\n");
        }

        public static IWebElement SetValue(ChromeDriver chromeDriver, string xPath, string text, int retries = 15,
            int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath);
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

        public static void Highlight(ChromeDriver driver, IWebElement element, int duration = 3000)
        {
            var origStyle = driver.ExecuteScript("arguments[0].style;", element);
            var script = @"arguments[0].style.cssText = ""border-width: 2px; border-style: solid; border-color: red"";";
            driver.ExecuteScript(script, element);
            Thread.Sleep(duration);
            driver.ExecuteScript("arguments[0].style=arguments[1]", new[] {element, origStyle});
        }
    }
}