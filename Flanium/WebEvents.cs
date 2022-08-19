using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using Polly;
using static Flanium.WebEvents.Search;

namespace Flanium;

public class WebEvents
{
    public class Search
    {
        public static List<IWebElement> FindAllChildren(WebElement element)
        {
            return element.FindElements(By.XPath("/*")).ToList();
        }

        public static List<IWebElement> FindAllDescendants(WebElement element)
        {
            return element.FindElements(By.CssSelector("*")).ToList();
        }
        
        public static IWebElement FindWebElementByXPath(ChromeDriver chromeDriver, string xPath,
            string frameType = "iframe")
        {
            Console.WriteLine("         Searching for Element: " + xPath + "\n");
            Console.WriteLine();

            IWebElement element;
            try
            {
                chromeDriver.SwitchTo().DefaultContent();
                element = chromeDriver.FindElement(By.XPath(xPath));
                if (element != null)
                {
                    Console.WriteLine("     Found Element: " + xPath + " (" + element.Text + ")\n");
                    Console.WriteLine();
                    return element;
                }
            }
            catch (Exception)
            {
                // ignored
            }

            var windowHandles = chromeDriver.WindowHandles;

            foreach (var window in windowHandles)
            {
                chromeDriver.SwitchTo().Window(window);

                var frames = chromeDriver.FindElements(By.TagName(frameType));

                foreach (var frame in frames)
                {
                    chromeDriver.SwitchTo().DefaultContent();
                    chromeDriver.SwitchTo().Frame(frame);
                    try
                    {
                        element = chromeDriver.FindElement(By.XPath(xPath));
                        if (element != null)
                        {
                            Console.WriteLine("     Found Element: " + xPath + " (" + element.Text + ")\n");
                            Console.WriteLine();
                            return element;
                        }
                    }
                    catch (Exception)
                    {
                        // ignored
                    }


                    var frameChildren = chromeDriver.FindElements(By.TagName(frameType));
                    foreach (var child in frameChildren)
                    {
                        chromeDriver.SwitchTo().DefaultContent();
                        chromeDriver.SwitchTo().Frame(frame);
                        chromeDriver.SwitchTo().Frame(child);

                        try
                        {
                            element = chromeDriver.FindElement(By.XPath(xPath));
                            if (element != null)
                            {
                                Console.WriteLine("     Found Element: " + xPath + " (" + element.Text + ")\n");
                                Console.WriteLine();
                                return element;
                            }
                        }
                        catch (Exception)
                        {
                            // ignored
                        }
                    }
                }
            }

            Console.WriteLine("Failed to find Element: " + xPath + "\n");
            Console.WriteLine();
            return null;
        }

        public static IWebElement WaitElementAppear(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    IWebElement element = FindWebElementByXPath(chromeDriver, xPath, frameType);
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
        public static IWebElement WaitElementVanish(ChromeDriver chromeDriver, string xPath, int retries = 60,
            int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
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
            string frameType = "iframe", int retryInterval = 1)
        {
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
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

        public static string GetText(ChromeDriver chromeDriver, string xPath, int retries = 15, int retryInterval = 1,
            string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
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
            int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
                    if (element != null)
                    {
                        Console.WriteLine("Hovered over Element: " + xPath + " (" + element.Text + ")\n");
                        var hover = new Actions(chromeDriver);
                        hover.MoveToElement(element).Perform();
                        chromeDriver.SwitchTo().DefaultContent();
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
            var windows = chromeDriver.WindowHandles;

            var alert = Policy.HandleResult<IAlert>(alert => alert == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval)).Execute(() =>
                {
                    IAlert alert = null;
                    foreach (var w in windows)
                    {
                        chromeDriver.SwitchTo().Window(w);
                        alert = chromeDriver.SwitchTo().Alert();
                    }

                    return alert;
                });

            if (alert != null)
            {
                if (acceptedAlertText.Contains(alert.Text))
                {
                    alert.Accept();
                    return true;
                }

                alert.Dismiss();
                return false;
            }

            return false;
        }
        public static IWebElement JsClick(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
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
        public static IWebElement Click(ChromeDriver chromeDriver, string xPath, int retries = 15,
            int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
                    if (element != null)
                    {
                        Console.WriteLine("Clicked Element: " + xPath + " (" + element.Text + ")\n");
                        element.Click();
                        return element;
                    }

                    return null;
                });

            if (element != null) return element;

            Console.WriteLine("Failed to click Element: " + xPath + "\n");
            throw new Exception("\n Failed to click element with XPath: " + xPath + "\n");

            // Helpers.CreateMessageWithAttachment("Failed to click element with XPath: " + xPath + "\n");
        }

        public static IWebElement SetValue(ChromeDriver chromeDriver, string xPath, string text, int retries = 15,
            int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
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
            driver.ExecuteScript("arguments[0].style=arguments[1]", new[]{element,origStyle});
        }
    }



}