![2255624 (1) (1)](https://user-images.githubusercontent.com/110975879/193143158-30d426a1-428c-4715-b505-f8fb48b13285.png)
# Flanium 


Flanium is an all-purpose RPA Library created for Windows.

Flanium is made by joining two frameworks:
 1. **Selenium** - Selenium is a suite of tools for automating web browsers.
 2. **FlaUi** - FlaUI is a .NET library which helps with automated UI testing of Windows applications (Win32, WinForms, WPF, Store Apps, ...).
 
 Flanium only supports Windows and Chrome browser **for now**. [12/08/2022]
 
 Flanium now supports Windows, Chrome and Firefox **partially**. [30/09/2022]
 
# [Documentation can be found here.](https://github.com/ALaurian/Flanium/wiki)
 
# Your very first web automation (a short tutorial on WebEvents)

**First start off by opening Visual Studio (not Visual Studio Code) or maybe Jetbrains Rider, and create a Console Application in .NET 6 (Windows), then follow the instructions below.**

The initializers in the Initializers library are dumb-proof, but you can make your own if you want, but these are generally the ones I would start with when making my web based RPA.

```csharp
var chromeWeb = Initializers.InitializeChrome(Initializers.InitializeService);
```

This will create your Chrome browser with a plethora of useful arguments and other superpowers, it disables all of the logging that the base chromedriver has, so it leaves the console application open for writing without interruptions.

Now we have it open, then we would need to navigate to a webpage, let's choose https://www.codeproject.com/.

```csharp
chromeWeb.Navigate().GoToUrl("https://www.codeproject.com/");
```

Then we want to click an element, I want to click the "Quick Answers" button on the homepage, I would do it like this:

```csharp
WebEvents.Action.Click(chromeWeb,"//*[@id='ctl00_TopNavBar_Answers']");
```

That's it, now you've clicked the "Quick Answers" button.

Now perhaps I am curious about what text does the first answer have in the table that just appeared.

```csharp
var answerText = WebEvents.Action.GetText(chromeWeb, "//*[@id='ctl00_ctl00_MC_AMC_Entries_ctl01_QuestionRow_H']");
```

I now have retrieved the text of the first question, for me it was "How do I pass variables from one PHP page to another".

Since I have all the information I need, I will now close the chrome session like this:

```csharp
chromeWeb.Dispose();
```

And now we're done, our very first automation done easily.

## Automatic Logging

All of the event libraries have built in logging for them, so you will see the results of the code in the Console Application.

Example: ![image](https://user-images.githubusercontent.com/110975879/184256363-f316f713-0712-4391-8a41-70b3bb64e1a9.png)

The logging no longer looks like in the example image, but the general idea is the same.
