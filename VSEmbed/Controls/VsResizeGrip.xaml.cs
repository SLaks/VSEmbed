using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace VSEmbed.Controls {
	/// <summary>
	/// Interaction logic for VsResizeGrip.xaml
	/// </summary>
	public partial class VsResizeGrip : ResizeGrip {
		private const int SizeBottomLeft = 7;
		private const int SizeBottomRight = 8;
		///<summary>Forwards mouse events to the underlying window.</summary>
		protected override void OnMouseDown(MouseButtonEventArgs e) {
			base.OnMouseDown(e);
			if (e.ChangedButton == MouseButton.Left) {
				int num = (FlowDirection == FlowDirection.LeftToRight) ? 8 : 7;
				HwndSource hwndSource = (HwndSource)PresentationSource.FromVisual(this);
				if (hwndSource != null) {
					NativeMethods.SendMessage(hwndSource.Handle, 274, (IntPtr)(61440 + num), IntPtr.Zero);
				}
			}
		}
	}
}
