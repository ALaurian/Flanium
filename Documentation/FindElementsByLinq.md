# FindElementsByLinq

* This method is used to retrieve the Elements in a specific Window based on a Linq query. It is used to find the Elements based on a specific property.

## Parameters

* **Window window** : Represents the Window of the application in which to search this element.
* **Func<AutomationElement, bool> linq** : Represents the Linq query.
* **int retries** : Represents the amount of time that the method will retry.
* **double retryInterval** : Represents the interval at which the retry will occur.

## Returns

```csharp
List of AutomationElement
```

## Examples

```csharp
var LinqElements = WinEvents.Linq.FindElementsByLinq(AppWindow, x => x.Name == "Address bar");
```

# [Back to main documentation.](https://github.com/ALaurian/Flanium/blob/main/Documentation/LibraryDB.md)
