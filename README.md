# Flanium

Flanium is an all-purpose RPA Library created for Windows.

Flanium is made by joining two frameworks:
 1. **Selenium** - Selenium is a suite of tools for automating web browsers.
 2. **FlaUi** - FlaUI is a .NET library which helps with automated UI testing of Windows applications (Win32, WinForms, WPF, Store Apps, ...).
 
# Knowledge requirements

You will need a basic knowledge of Linq to use the WinEvents library and a medium to advanced knowledge of XPath syntax to be able to use the WebEvents library.
It goes without saying that you should also know a bit of C# to be able to use it properly.

### XPath
The w3cschools XPath tutorial is pretty good for XPaths and covers pretty much everything, you can find it here: https://www.w3schools.com/xml/xpath_intro.asp.

### Linq
Linq is also easy to understand, at least up to the level that you will need to be able to use it, for example, a Linq query would look like this:

```csharp
var Window = WinEvents.GetWindowByLinq(x => x.Name == "Google");
```

What is between the parantheses is the Linq query, x represents the AutomationElement you will be searching for, and x.Name represents the Name property of the x object which is a string, then you compare the x.Name with "Google" by using "==", if it returns true, then voila, you have your window, you can add multiple conditions by linking them with "&&" (and) and "||" (or), see example below:

```csharp
var Window = WinEvents.GetWindowByLinq(x => x.Name == "yes" && x.AutomationId == "1001");
```

**or**

```csharp
var Window = WinEvents.GetWindowByLinq(x => x.Name == "yes" || x.AutomationId == "1001");
```

The first version with "&&" will search for a Window that has both the Name property equal to "yes" and the AutomationId property equal to "1001" while the other will chose one that either has the Name equal to "yes" or the AutomationId equal to "1001".

### Polly

While I wouldn't necessarily post this as a knowledge requirement since this is for more advanced users.. but knowing how to use Polly will make your life a lot easier when using this framework, find it here -> https://github.com/App-vNext/Polly.

### Remarks

I tried to make it as easy as possible to use for beginner coders or people with little to no coding experience, you can probably make your own RPA with just barely any coding experience by following a simple tutorial or the examples I wrote below.

# The list of toys

Attempt | WinEvents | WebEvents | Initializers | Helpers |
--- | --- | --- | --- |--- 
1  |  GetDriverProcessId  |  FindWebElementByXPath  |  InitializeService  |  OpenSapSession
2  |  GetWindowByProcessId  |  WaitForAlert  |  InitializeChrome  |  FolderContainsFiles
3  |  GetWindowByLinq  |  JsClick  |    |  DeleteDuplicateFiles
4  |  GetWindowsByLinq  |  Click  |    |  CreateFolder
5  |  FindElementByLinq  |  SetValue  |    |  DeleteFolder
6  |  FindElementsByLinq  |  Hover  |    |  ArchiveFolder
7  |    |  GetText  |    |  DeleteFile
8  |    |  WaitForeverElementVanish  |    |  MoveFile
9  |    |  WaitElementVanish  |    |  MoveFiles
10  |    |    |    |  ExcelToDataTable
11  |    |    |    |  Highlight
12  |    |    |    |  HandleDownloads
13  |    |    |    |  CloseTab
14  |    |    |    |  SendEmail

If you have any experience with FlaUI you probably already noticed that there is no XPath searching method in the WinEvents library, that's because from my experiments, searching by Linq query is.. a lot quicker (400ms up to a few seconds depending on the application).

# The library in practice
