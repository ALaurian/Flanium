# GetWindowByLinq

* This method is used to retrieve the Window by using a Linq query. It is used to find the Window based on a specific property. 

## Parameters

* **Func<AutomationElement, bool> linq** : Represents the Linq query.

## Returns

```csharp
Window
```

## Examples

```csharp
var LinqWindow = WinEvents.Linq.GetWindowByLinq(x => x.Name.StartsWith("Explorer"));
```

# [Back to main documentation.](https://github.com/ALaurian/Flanium/blob/main/Documentation/LibraryDB.md)
