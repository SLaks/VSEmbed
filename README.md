VSEmbed
=======

VSEmbed lets you embed the Visual Studio editor &amp; theme architecture in your own programs.

![screenshot](https://pbs.twimg.com/media/B1dX6NxCMAAv6iZ.png:large)

#Usage
This is not quite ready for consumption yet; I will soon publish NuGet packages.

To initialize Visual Studio, you need the following code:

```C#
VsLoader.LoadLatest();    // Or .Load(new Version(...))
VsServiceProvider.Initialize();
VsMefContainerBuilder.CreateDefaultContainer(); // Only needed for editor embedding
```

If you're already using MEF, you can use the other methods in `VsMefContainerBuilder` to add the editor services to your MEF container.

#Caveats
 - The end-user must have a version (2012+) of Visual Studio (including Express editions) installed for this to run.
 - The Roslyn editor services will only work if Dev14 (any version) is installed.


