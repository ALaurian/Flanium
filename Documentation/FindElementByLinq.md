# FindElementByLinq

* This method is used to retrieve the Element in a specific Window based on a Linq query. It is used to find the Element based on a specific property.

## Parameters

* **AutomationElement window** : Represents the Window of the application in which to search this element.
* **Func<AutomationElement, bool> linq** : Represents the Linq query.

## Returns

```csharp
AutomationElement
```

## Examples

```csharp
var LinqElement = WinEvents.FindElementByLinq(AppWindow, x => x.Name == "Address bar");
```

# [Back to main documentation.](https://github.com/ALaurian/Flanium/blob/main/Documentation/LibraryDB.md)
