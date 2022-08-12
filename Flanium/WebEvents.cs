using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using Polly;

namespace Flanium
{
    public class WebEvents
    {
        
        /// <summary>
        /// Helper method for searching for an element in every window and iframe of that window.
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance. </param>
        /// <param name="xPath"> Represents the XPath of the element. </param>
        /// <param name="frameType"> Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns the element that was searched.</returns>
        public static IWebElement FindWebElementByXPath(ChromeDriver chromeDriver, string xPath, string frameType = "iframe")
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

        /// <summary>
        /// This method is used to search for an alert, and accept it given a list of strings that would contain the alert text.
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance. </param>
        /// <param name="acceptedAlertText"> Represents the list of strings that would match to the alert text. </param>
        /// <param name="retries"> Represents the number of times this method will retry.</param>
        /// <param name="retryInterval">Represents the amount of time in seconds to wait before each retry.</param>
        /// <returns>True if accepted, False if dismissed.</returns>
        public static bool WaitForAlert(ChromeDriver chromeDriver, List<string> acceptedAlertText, int retries = 15, int retryInterval = 1)
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
        
        /// <summary>
        /// This method is used to click an element using Javascript by searching for it via XPath.
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance.</param>
        /// <param name="xPath"> Represents the element XPath.</param>
        /// <param name="retries"> Represents the number of times this method will retry. </param>
        /// <param name="retryInterval"> Represents the amount of time in seconds to wait before each retry.</param>
        /// <param name="frameType">Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns the element that was clicked.</returns>
        /// <exception cref="Exception"> Throws a custom Exception. </exception>
        public static IWebElement JsClick(ChromeDriver chromeDriver, string xPath, int retries = 15, int retryInterval = 1, string frameType = "iframe")
        {

            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
                    if(element != null)
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

        /// <summary>
        /// This method is used to click an element by searching for it via XPath.
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance.</param>
        /// <param name="xPath"> Represents the element XPath.</param>
        /// <param name="retries"> Represents the number of times this method will retry. </param>
        /// <param name="retryInterval"> Represents the amount of time in seconds to wait before each retry.</param>
        /// <param name="frameType">Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns the element that was clicked.</returns>
        /// <exception cref="Exception"> Throws a custom Exception. </exception>
        public static IWebElement Click(ChromeDriver chromeDriver, string xPath, int retries = 15, int retryInterval = 1, string frameType = "iframe")
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

        /// <summary>
        /// This method is used to set the value of an element by searching for it via XPath.
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance.</param>
        /// <param name="xPath"> Represents the element XPath.</param>
        /// <param name="text"> Represents the text value to set the element value to.</param>
        /// <param name="retries"> Represents the number of times this method will retry. </param>
        /// <param name="retryInterval"> Represents the amount of time in seconds to wait before each retry.</param>
        /// <param name="frameType">Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns the element that whose value or text was changed.</returns>
        /// <exception cref="Exception"> Throws a custom Exception. </exception>
        public static IWebElement SetValue(ChromeDriver chromeDriver, string xPath, string text, int retries = 15, int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
                    if (element != null)
                    {
                        Console.WriteLine("Set Value of Element: " + xPath + " (" + element.Text + ")\n");
                        element.Clear();
                        element.SendKeys(text);
                        
                        return element;
                    }
                    return null;
                });
            
            if (element != null) return element;
            Console.WriteLine("Failed to set value of Element: " + xPath + "\n");
            throw new Exception("\n Failed to set value of element with XPath: " + xPath + "\n");
        }
        
        /// <summary>
        /// This method is used to hover over an element by searching for it via XPath.
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance.</param>
        /// <param name="xPath"> Represents the element XPath.</param>
        /// <param name="retries"> Represents the number of times this method will retry. </param>
        /// <param name="retryInterval"> Represents the amount of time in seconds to wait before each retry.</param>
        /// <param name="frameType">Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns the element that was hovered. </returns>
        /// <exception cref="Exception"> Throws a custom Exception. </exception>
        public static IWebElement Hover(ChromeDriver chromeDriver, string xPath, int retries = 15, int retryInterval = 1, string frameType = "iframe")
        {
            var element = Policy.HandleResult<IWebElement>(result => result == null)
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element = FindWebElementByXPath(chromeDriver, xPath, frameType);
                    if (element != null)
                    {
                        Console.WriteLine("Hovered over Element: " + xPath + " (" + element.Text + ")\n");
                        Actions hover = new Actions(chromeDriver);
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

        /// <summary>
        /// This method is used to get the text of an element by searching for it via XPath.
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance.</param>
        /// <param name="xPath"> Represents the element XPath.</param>
        /// <param name="retries"> Represents the number of times this method will retry. </param>
        /// <param name="retryInterval"> Represents the amount of time in seconds to wait before each retry.</param>
        /// <param name="frameType">Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns null if no text or value is found.</returns>
        public static string GetText(ChromeDriver chromeDriver, string xPath, int retries = 15, int retryInterval = 1, string frameType = "iframe")
        {
            
            var element = Policy.HandleResult<IWebElement>(result => result.Text == "" && result.GetAttribute("value") == "")
                .WaitAndRetry(retries, interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() => FindWebElementByXPath(chromeDriver, xPath, frameType));

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

            Console.WriteLine("No text or value was found for Element: " + xPath + "\n");
            return "";
        }
        
        /// <summary>
        /// This method is used to wait for an element by searching for it via XPath (it waits forever).
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance.</param>
        /// <param name="xPath"> Represents the element XPath.</param>
        /// <param name="retryInterval"> Represents the amount of time in seconds to wait before each retry.</param>
        /// <param name="frameType">Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns null if element did not vanish.</returns>
        public static IWebElement WaitForeverElementVanish(ChromeDriver chromeDriver, string xPath, string frameType = "iframe", int retryInterval = 1)
        {
            
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                                       .WaitAndRetryForever(interval => TimeSpan.FromSeconds(retryInterval))
                                       .Execute(() =>
                                       {
                                           var element =  FindWebElementByXPath(chromeDriver, xPath, frameType);
                                           if(element == null)
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
        /// <summary>
        /// This method is used to wait for an element by searching for it via XPath (it waits forever).
        /// </summary>
        /// <param name="chromeDriver"> Represents the chromeDriver instance.</param>
        /// <param name="xPath"> Represents the element XPath.</param>
        /// <param name="retries"> Represents the number of times this method will retry. </param>
        /// <param name="retryInterval"> Represents the amount of time in seconds to wait before each retry.</param>
        /// <param name="frameType">Represents the frame type, it can be 'iframe', 'frame', 'object' etc.</param>
        /// <returns>Returns null if element did not vanish.</returns>
        public static IWebElement WaitElementVanish(ChromeDriver chromeDriver, string xPath, int retries=60,int retryInterval = 1,string frameType = "iframe")
        {
            
            var element = Policy.HandleResult<IWebElement>(result => result != null)
                .WaitAndRetry(retries,interval => TimeSpan.FromSeconds(retryInterval))
                .Execute(() =>
                {
                    var element =  FindWebElementByXPath(chromeDriver, xPath, frameType);
                    if(element == null)
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

    }
}
