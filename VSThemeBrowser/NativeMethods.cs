using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace VSThemeBrowser {
	static class NativeMethods {
		// Microsoft.VisualStudio.PlatformUI.NativeMethods
		[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		public static extern IntPtr SendMessage(IntPtr hWnd, int nMsg, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		[return: MarshalAs(UnmanagedType.Bool)]
		internal static extern bool SystemParametersInfo(int uiAction, int uiParam, ref NONCLIENTMETRICS pvParam, int fWinIni);
	}
	internal struct NONCLIENTMETRICS {
		public int cbSize;
		public int iBorderWidth;
		public int iScrollWidth;
		public int iScrollHeight;
		public int iCaptionWidth;
		public int iCaptionHeight;
		public LOGFONT lfCaptionFont;
		public int iSmCaptionWidth;
		public int iSmCaptionHeight;
		public LOGFONT lfSmCaptionFont;
		public int iMenuWidth;
		public int iMenuHeight;
		public LOGFONT lfMenuFont;
		public LOGFONT lfStatusFont;
		public LOGFONT lfMessageFont;
		public int iPaddedBorderWidth;
	}
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	internal struct LOGFONT {
		public int lfHeight;
		public int lfWidth;
		public int lfEscapement;
		public int lfOrientation;
		public int lfWeight;
		public byte lfItalic;
		public byte lfUnderline;
		public byte lfStrikeOut;
		public byte lfCharSet;
		public byte lfOutPrecision;
		public byte lfClipPrecision;
		public byte lfQuality;
		public byte lfPitchAndFamily;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
		public string lfFaceName;
	}
}
