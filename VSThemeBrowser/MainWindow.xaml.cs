using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.VisualStudio.PlatformUI.Shell;
using Microsoft.VisualStudio.PlatformUI.Shell.Controls;

namespace VSThemeBrowser {
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : CustomChromeWindow {
		public MainWindow() {
			InitializeComponent();
			ViewManager.Instance.Initialize(new ContentControl());
		}
	}
}
