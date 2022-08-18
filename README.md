# Flanium

Flanium is an all-purpose RPA Library created for Windows.

Flanium is made by joining two frameworks:
 1. **Selenium** - Selenium is a suite of tools for automating web browsers.
 2. **FlaUi** - FlaUI is a .NET library which helps with automated UI testing of Windows applications (Win32, WinForms, WPF, Store Apps, ...).
 
 Flanium only supports Windows and Chrome browser **for now**. [12/08/2022]
 
# [Documentation can be found here.](https://github.com/ALaurian/Flanium/blob/main/Documentation/LibraryDB.md)
 
# Knowledge requirements

You will need a basic knowledge of Linq to use the WinEvents library and a medium to advanced knowledge of XPath syntax to be able to use the WebEvents library.
It goes without saying that you should also know a bit of C# to be able to use it properly.

### XPath
The w3cschools XPath tutorial is pretty good for XPaths and covers pretty much everything, you can find it here: https://www.w3schools.com/xml/xpath_intro.asp.


### Linq [optional]
#### [There are methods implemented that can be used with XPath, Linq methods are 2 times quicker than XPath methods but may be unstable.]
Linq is also easy to understand, at least up to the level that you will need to be able to use it, for example, a Linq query would look like this:

```csharp
var Window = WinEvents.Linq.GetWindowByLinq(x => x.Name == "Google");
```

What is between the parantheses is the Linq query, x represents the AutomationElement you will be searching for, and x.Name represents the Name property of the x object which is a string, then you compare the x.Name with "Google" by using "==", if it returns true, then voila, you have your window, you can add multiple conditions by linking them with "&&" (and) and "||" (or), see example below:

```csharp
var Window = WinEvents.Linq.GetWindowByLinq(x => x.Name == "yes" && x.AutomationId == "1001");
```

**or**

```csharp
var Window = WinEvents.Linq.GetWindowByLinq(x => x.Name == "yes" || x.AutomationId == "1001");
```

The first version with "&&" will search for a Window that has both the Name property equal to "yes" and the AutomationId property equal to "1001" while the other will chose one that either has the Name equal to "yes" or the AutomationId equal to "1001".

### Remarks

I tried to make it as easy as possible to use for beginner coders or people with little to no coding experience, you can probably make your own RPA with just barely any coding experience by following a simple tutorial or the examples I wrote below.

# Your very first web automation

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
