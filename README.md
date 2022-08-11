# Flanium

Flanium is an all-purpose RPA Library created for Windows.

Flanium is made by joining two frameworks:
 1. **Selenium** - Selenium is a suite of tools for automating web browsers.
 2. **FlaUi** - FlaUI is a .NET library which helps with automated UI testing of Windows applications (Win32, WinForms, WPF, Store Apps, ...).
 
 # Version Information
 
 **Version 1.0 capabilities and support:**
 
 * Chrome.
 * Windows.
    
# Documentation
This section will go through the various libraries and what they do.

## WinEvents Library

### GetDriverProcessId(ChromeDriverService driverService) : returns int

* This method is used to retrieve the current process ID of the <c>Chrome</c> browser attached to the ChromeDriverService.

Parameters:
>**driverService** - Represents the ChromeDriverService instance.

### GetWindow(int processId) : returns AutomationElement

* This method is used to retrieve the window of the main window handle the process ID has.

Parameters:
>**processId** - Represents the process ID.


## WebEvents Library

## Initializers Library

## Helpers Library
