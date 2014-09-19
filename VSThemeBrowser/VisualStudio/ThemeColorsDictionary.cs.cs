using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using Microsoft.Internal.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;

namespace VSThemeBrowser.VisualStudio {
	public class ThemeColorsDictionary : ResourceDictionary {

		// We must access everything from these classes using dynamic due to NoPIA conflicts.
		// The compiler gives some errors since we do not have the right PIA, and the runtime
		// gives more errors because NoPIA doesn't unify for managed implementations.
		dynamic currentTheme;
		readonly dynamic service;
		public ThemeColorsDictionary() {
			if (ServiceProvider.GlobalProvider.GetService(new Guid("FD57C398-FDE3-42c2-A358-660F269CBE43")) != null)
				return;	// Do nothing when hosted in VS
						//AssemblyResolverHack.AddHandler();
			FakeServiceProvider.Initialize();
			service = new Microsoft.VisualStudio.Platform.WindowManagement.ColorThemeService();
			//service = Activator.CreateInstance(Type.GetType("Microsoft.VisualStudio.Platform.WindowManagement.ColorThemeService, Microsoft.VisualStudio.Platform.WindowManagement"));
			ThemeIndex = 0;
		}
		int themeIndex;
		public int ThemeIndex {
			get { return themeIndex; }
			set { themeIndex = value; LoadTheme(value); }
		}

		static Color ToColorFromRgba(uint colorValue) {
			return Color.FromArgb((byte)(colorValue >> 24), (byte)colorValue, (byte)(colorValue >> 8), (byte)(colorValue >> 16));
		}
		static SolidColorBrush GetBrush(Color color) {
			var brush = new SolidColorBrush(color);
			brush.Freeze();
			return brush;
		}
		// Loosely based on Microsoft.VisualStudio.Platform.WindowManagement.ResourceSynchronizer.AddSolidColorKeys()
		public void LoadTheme(int index) {
			if (service == null)
				return;
			Clear();

			currentTheme = service.Themes[index % service.Themes.Count];
			foreach (ColorName colorName in service.ColorNames) {
				IVsColorEntry entry = currentTheme[colorName];
				if (entry == null) continue;

				if (entry.BackgroundType != 0) {
					ThemeResourceKey bgBrush = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.BackgroundBrush);
					ThemeResourceKey bgColor = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.BackgroundColor);

					var color = ToColorFromRgba(entry.Background);
					Add(bgBrush, GetBrush(color));
					Add(bgColor, color);

					int colorId = VsColorFromName(colorName);
					if (colorId != 0) {
						Add(VsColors.GetColorKey(colorId), color);
						Add(VsBrushes.GetBrushKey(colorId), GetBrush(color));
					}
				}
				if (entry.ForegroundType != 0) {
					ThemeResourceKey fgBrush = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.ForegroundBrush);
					ThemeResourceKey fgColor = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.ForegroundColor);
					var color = ToColorFromRgba(entry.Foreground);
					Add(fgBrush, GetBrush(color));
					Add(fgColor, color);
				}
			}
			AddFonts();
		}
		void AddFonts() {
			var dialogFont = System.Drawing.SystemFonts.CaptionFont;

			Add(VsFonts.EnvironmentFontFamilyKey, new FontFamily(dialogFont.Name));
			Add(VsFonts.EnvironmentFontSizeKey, dialogFont.Size);
			AddCaptionFonts();
		}
		private void AddCaptionFonts() {
			NONCLIENTMETRICS nONCLIENTMETRICS = default(NONCLIENTMETRICS);
			nONCLIENTMETRICS.cbSize = Marshal.SizeOf(typeof(NONCLIENTMETRICS));
			if (!NativeMethods.SystemParametersInfo(41, nONCLIENTMETRICS.cbSize, ref nONCLIENTMETRICS, 0)) {
				Add(VsFonts.CaptionFontFamilyKey, this[VsFonts.EnvironmentFontFamilyKey]);
				Add(VsFonts.CaptionFontSizeKey, this[VsFonts.EnvironmentFontSizeKey]);
				Add(VsFonts.CaptionFontWeightKey, FontWeights.Normal);
				return;
			}
			FontFamily value = new FontFamily(nONCLIENTMETRICS.lfCaptionFont.lfFaceName);
			double num = this.FontSizeFromLOGFONTHeight(nONCLIENTMETRICS.lfCaptionFont.lfHeight);
			FontWeight fontWeight = FontWeight.FromOpenTypeWeight(nONCLIENTMETRICS.lfCaptionFont.lfWeight);
			Add(VsFonts.CaptionFontFamilyKey, value);
			Add(VsFonts.CaptionFontSizeKey, num);
			Add(VsFonts.CaptionFontWeightKey, fontWeight);
		}
		private double FontSizeFromLOGFONTHeight(int lfHeight) {
			return Math.Abs(lfHeight) * DpiHelper.DeviceToLogicalUnitsScalingFactorY;
		}


		// Microsoft.VisualStudio.Platform.WindowManagement.ColorNameTranslator
		static readonly Guid environmentColors = new Guid("{624ed9c3-bdfd-41fa-96c3-7c824ea32e3d}");
		static int VsColorFromName(ColorName colorName) {
			if (colorName.Category != environmentColors)
				return 0;
			int result;
			VsColors.TryGetColorIDFromBaseKey("VsColor." + colorName.Name, out result);
			return result;
		}
	}
}
namespace Microsoft.Internal.VisualStudio.Shell.Interop {
	[CompilerGenerated, Guid("413D8344-C0DB-4949-9DBC-69C12BADB6AC"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), TypeIdentifier]
	[ComImport]
	public interface IVsColorTheme {
		IVsColorEntry this[[In] ColorName Name] {
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}
		Guid ThemeId { get; }
		string Name {
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}
		bool IsUserVisible { get; }
		void Apply();
	}
	[CompilerGenerated, Guid("BBE70639-7AD9-4365-AE36-9877AF2F973B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), TypeIdentifier]
	[ComImport]
	public interface IVsColorEntry {
		ColorName ColorName { get; }
		byte BackgroundType { get; }
		byte ForegroundType { get; }
		uint Background { get; }
		uint Foreground { get; }
		uint BackgroundSource { get; }
		uint ForegroundSource { get; }
	}

	[CompilerGenerated, TypeIdentifier("EF2A7BE1-84AF-4E47-A2CF-056DF55F3B7A", "Microsoft.Internal.VisualStudio.Shell.Interop.ColorName")]
	[StructLayout(LayoutKind.Sequential, Pack = 8)]
	public struct ColorName {
		public Guid Category;
		[MarshalAs(UnmanagedType.BStr)]
		public string Name;
	}
}