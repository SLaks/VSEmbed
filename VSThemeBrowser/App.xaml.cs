using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace VSThemeBrowser {
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application {
		public App() {
			VisualStudio.VsLoader.Initialize(new Version(12, 0, 0, 0));
			VisualStudio.FakeServiceProvider.Initialize();
		}
		protected override void OnStartup(StartupEventArgs e) {
			base.OnStartup(e);
		}
	}
}
