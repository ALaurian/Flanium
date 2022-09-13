﻿using System.Diagnostics;
using System.Management;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using OpenQA.Selenium.Chrome;
using Polly;

namespace WinSurfer;

public class WinBrowser
{
    public Searchers Search = new();
        public Actions Action = new();

        public WinBrowser()
        {
            Console.WriteLine("WinSurfer has started successfully.");
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