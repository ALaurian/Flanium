# GetWindowsByLinq

* This method is used to retrieve the Windows by using a Linq query. It is used to find the Windows based on a specific property.

## Parameters

* **Func<AutomationElement, bool> linq** : Represents the Linq query.

## Returns

```csharp
List of Window
```

## Examples

```csharp
var LinqWindow = WinEvents.GetWindowByLinq(x => x.Name.StartsWith("Explorer"));
```

# [Back to main documentation.](https://github.com/ALaurian/Flanium/blob/main/Documentation/LibraryDB.md)
