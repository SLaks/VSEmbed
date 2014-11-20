VSEmbed
=======

VSEmbed lets you embed the Visual Studio editor &amp; theme architecture in your own programs.

![better screenshpt](https://pbs.twimg.com/media/B2CVhsmCAAAUPmg.png:large)
![screenshot](https://pbs.twimg.com/media/B1dX6NxCMAAv6iZ.png:large)

#Usage
This is not quite ready for consumption yet; I will soon publish NuGet packages.

To initialize Visual Studio, you need the following code:

```C#
VsLoader.LoadLatest();    // Or .Load(new Version(...))
VsServiceProvider.Initialize();
VsMefContainerBuilder.CreateDefault().Build(); // Only needed for editor embedding
```

The last line can only be JITted after initializing VsLoader (because `Build()` returns an `IComponentModel`, which is defined in a VS assembly), so you should put it in a separate method and call that method after setting up VsLoader.

You _must_ create the MEF container using `VsMefContainerBuilder`; it will use Visual Studio 2015's [new version of MEF](http://blog.slaks.net/2014-11-16/mef2-roslyn-visual-studio-compatibility/) (where available) to support Roslyn's MEF2 exports.  
If you're already using MEF, you can call `WithFilteredCatalogs(assemblies)` or `WithCatalog(types)` to add your own assemblies to the MEF container.   Note that `VsMefContainerBuilder` is immutable; these methods return new instances with the new catalogs added.

#Using Roslyn
After loading Dev14, you can set the ContentType of an ITextBuffer to `C#` or `VisualBasic` to activate the Roslyn editors.  However, you will also need to link the ITextBuffer to a Roslyn [Workspace](http://source.roslyn.codeplex.com/#Microsoft.CodeAnalysis.Workspaces/Workspace/Workspace.cs) to activate the language services.  In addition, the workspace must have an `IWorkCoordinatorRegistrationService` registered to run diagnostics in the background.

To do all this, use my `EditorWorkspace` class, and call `CreateDocument()` to create a new document linked to a text buffer.  You can also inherit this class to provide additional behavior.

If you create a Roslyn-powered buffer and do not link it to a workspace, I have a buffer listener which will create a simple workspace with a few references for you.

#Caveats
 - The end-user must have a version (2012+) of Visual Studio (including Express editions) installed for this to run.
 - The Roslyn editor services will only work if VS2015 Preview (or later builds) is installed.
  - To make it support older Dev14 CTPs, use Reflection to call `MefHostService` if `MefV1HostServices` does not exist, and re-add the older `XmlDocumentationProvider` code that was replaced in this commit.
 - XML doc comments are not shown
  - This is caused by https://roslyn.codeplex.com/workitem/406
 - If Visual Studio 2012 assemblies are in the GAC, other versions will not load properly.
 - Code snippets are not implemented.
 - Peek does not work.
  - To make Peek work, implement & export `IPeekResultPresenter` & `IPeekResultPresentation`, and create a WpfTextViewHost in `Create()`.  Note that peek only operates on file paths.
- Rename with preview does not work.
 - To make this work, implement `IVsPreviewChangesService` and add it to the ServiceProvider.
