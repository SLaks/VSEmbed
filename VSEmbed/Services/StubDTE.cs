using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
namespace VSEmbed.Services {
	// Used by KeyBindingHelper.GetKeyBinding, which is used by VSLightBulbPresenterStyle.
	class StubDTE : DTE {
		public Commands Commands { get; } = new StubCommands();

		public Document ActiveDocument { get { throw new NotImplementedException(); } }
		public dynamic ActiveSolutionProjects { get { throw new NotImplementedException(); } }
		public Window ActiveWindow { get { throw new NotImplementedException(); } }
		public AddIns AddIns { get { throw new NotImplementedException(); } }
		public DTE Application { get { throw new NotImplementedException(); } }
		public dynamic CommandBars { get { throw new NotImplementedException(); } }
		public string CommandLineArguments { get { throw new NotImplementedException(); } }
		public ContextAttributes ContextAttributes { get { throw new NotImplementedException(); } }
		public Debugger Debugger { get { throw new NotImplementedException(); } }
		public vsDisplay DisplayMode {
			get { throw new NotImplementedException(); }
			set { throw new NotImplementedException(); }
		}
		public Documents Documents { get { throw new NotImplementedException(); } }
		public DTE DTE { get { throw new NotImplementedException(); } }
		public string Edition { get { throw new NotImplementedException(); } }
		public Events Events { get { throw new NotImplementedException(); } }
		public string FileName { get { throw new NotImplementedException(); } }
		public Find Find { get { throw new NotImplementedException(); } }
		public string FullName { get { throw new NotImplementedException(); } }
		public Globals Globals { get { throw new NotImplementedException(); } }
		public ItemOperations ItemOperations { get { throw new NotImplementedException(); } }
		public int LocaleID { get { throw new NotImplementedException(); } }
		public Macros Macros { get { throw new NotImplementedException(); } }
		public DTE MacrosIDE { get { throw new NotImplementedException(); } }
		public Window MainWindow { get { throw new NotImplementedException(); } }
		public vsIDEMode Mode { get { throw new NotImplementedException(); } }
		public string Name { get { throw new NotImplementedException(); } }
		public ObjectExtenders ObjectExtenders { get { throw new NotImplementedException(); } }
		public string RegistryRoot { get { throw new NotImplementedException(); } }
		public SelectedItems SelectedItems { get { throw new NotImplementedException(); } }
		public Solution Solution { get { throw new NotImplementedException(); } }
		public SourceControl SourceControl { get { throw new NotImplementedException(); } }
		public StatusBar StatusBar { get { throw new NotImplementedException(); } }
		public bool SuppressUI {
			get { throw new NotImplementedException(); }
			set { throw new NotImplementedException(); }
		}
		public UndoContext UndoContext { get { throw new NotImplementedException(); } }
		public bool UserControl {
			get { throw new NotImplementedException(); }
			set { throw new NotImplementedException(); }
		}
		public string Version { get { throw new NotImplementedException(); } }
		public WindowConfigurations WindowConfigurations { get { throw new NotImplementedException(); } }
		public Windows Windows { get { throw new NotImplementedException(); } }
		public void ExecuteCommand(string CommandName, string CommandArgs = "") { throw new NotImplementedException(); }
		public dynamic GetObject(string Name) { throw new NotImplementedException(); }
		public bool get_IsOpenFile(string ViewKind, string FileName) { throw new NotImplementedException(); }
		public Properties get_Properties(string Category, string Page) { throw new NotImplementedException(); }
		public wizardResult LaunchWizard(string VSZFile, ref object[] ContextParams) { throw new NotImplementedException(); }
		public Window OpenFile(string ViewKind, string FileName) { throw new NotImplementedException(); }
		public void Quit() { throw new NotImplementedException(); }
		public string SatelliteDllPath(string Path, string Name) { throw new NotImplementedException(); }
	}

	internal class StubCommands : Commands {
		public Command Item(object index, int ID = -1) => null;

		public int Count { get { throw new NotImplementedException(); } }
		public DTE DTE { get { throw new NotImplementedException(); } }
		public DTE Parent { get { throw new NotImplementedException(); } }
		public void Add(string Guid, int ID, ref object Control) { throw new NotImplementedException(); }
		public dynamic AddCommandBar(string Name, vsCommandBarType Type, [IDispatchConstant]object CommandBarParent, int Position = 1) { throw new NotImplementedException(); }
		public Command AddNamedCommand(AddIn AddInInstance, string Name, string ButtonText, string Tooltip, bool MSOButton, int Bitmap, ref object[] ContextUIGUIDs, int vsCommandDisabledFlagsValue = 16) { throw new NotImplementedException(); }
		public void CommandInfo(object CommandBarControl, out string Guid, out int ID) { throw new NotImplementedException(); }
		public IEnumerator GetEnumerator() { throw new NotImplementedException(); }
		public void Raise(string Guid, int ID, ref object CustomIn, ref object CustomOut) { throw new NotImplementedException(); }
		public void RemoveCommandBar(object CommandBar) { throw new NotImplementedException(); }
	}
}