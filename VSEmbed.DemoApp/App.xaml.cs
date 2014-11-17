using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using VSEmbed;

namespace VSEmbed.DemoApp {
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application {
		public App() {
			VsLoader.Load(new Version(14, 0, 0, 0));
			VsServiceProvider.Initialize();
			VsMefContainerBuilder.CreateDefault().Build();
		}
		protected override void OnStartup(StartupEventArgs e) {
			base.OnStartup(e);
		}
	}
}
