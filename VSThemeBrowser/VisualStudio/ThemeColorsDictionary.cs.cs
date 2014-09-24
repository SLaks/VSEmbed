using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using Microsoft.Internal.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using VSThemeBrowser.Controls;

namespace VSThemeBrowser.VisualStudio {
	public class ThemeColorsDictionary : ConditionalResourceDictionary {

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

			// Add an empty dictionary for the loader to replace.
			MergedDictionaries.Add(new ResourceDictionary());
			ThemeIndex = 0;

			typeof(EnvironmentRenderCapabilities)
				.GetProperty("VisualEffectsAllowed")
				.SetValue(EnvironmentRenderCapabilities.Current, 1 | 2);

			Themes = Enumerable.Range(0, (int)service.Themes.Count)
				.Select(i => new ColorTheme(i, service.Themes[i]))
				.ToList()
				.AsReadOnly();
		}
		bool designerOnly;
		///<summary>Gets or sets whether the theme resources should only be added in the designer.</summary>
		///<remarks>
		/// Use this property to create a theme dictionary for the designer 
		/// while still inheriting an application-level theme at runtime.
		///</remarks>
		public bool DesignerOnly {
			get { return designerOnly; }
			set {
				designerOnly = value;
				if (value) {
					if (IsDesignMode)
						LoadTheme(new Random().Next(Themes.Count));
					else
						MergedDictionaries[0].Clear();
				}
			}
		}

		public ReadOnlyCollection<ColorTheme> Themes { get; private set; }
		public ColorTheme CurrentTheme {
			get { return Themes[ThemeIndex]; }
			set { ThemeIndex = value.Index; }
		}
		int themeIndex;
		public int ThemeIndex {
			get { return themeIndex; }
			set { themeIndex = value; LoadTheme(value); }
		}


		private bool IsDesignMode {
			get { return (bool)DesignerProperties.IsInDesignModeProperty.GetMetadata(typeof(DependencyObject)).DefaultValue; }
		}

		public void LoadTheme(int index) {
			if (service == null || (DesignerOnly && IsDesignMode))
				return;
			var newDictionary = new ResourceDictionary();

			currentTheme = service.Themes[index % service.Themes.Count];

			AddSolidColorKeys(newDictionary);
			AddGradients(newDictionary);
			AddTextureBrushes(newDictionary);
			AddFonts(newDictionary);

			// Replace the old dictionary as a single operation to avoid extra lookups
			MergedDictionaries[0] = newDictionary;
		}

		#region AddSolidColorKeys
		// Copied from ResourceSynchronizer and modified to use currentTheme and to add actual values instead of deferred keys.
		void AddSolidColorKeys(ResourceDictionary newDictionary) {
			foreach (ColorName colorName in service.ColorNames) {
				IVsColorEntry entry = currentTheme[colorName];
				if (entry == null)
					continue;

				if (entry.BackgroundType != 0) {
					ThemeResourceKey bgBrush = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.BackgroundBrush);
					ThemeResourceKey bgColor = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.BackgroundColor);

					var color = ToColorFromRgba(entry.Background);
					newDictionary.Add(bgBrush, GetBrush(color));
					newDictionary.Add(bgColor, color);

					int colorId = VsColorFromName(colorName);
					if (colorId != 0) {
						newDictionary.Add(VsColors.GetColorKey(colorId), color);
						newDictionary.Add(VsBrushes.GetBrushKey(colorId), GetBrush(color));
					}
				}
				if (entry.ForegroundType != 0) {
					ThemeResourceKey fgBrush = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.ForegroundBrush);
					ThemeResourceKey fgColor = new ThemeResourceKey(entry.ColorName.Category, entry.ColorName.Name, ThemeResourceKeyType.ForegroundColor);
					var color = ToColorFromRgba(entry.Foreground);
					newDictionary.Add(fgBrush, GetBrush(color));
					newDictionary.Add(fgColor, color);
				}
			}
		}
		static Color ToColorFromRgba(uint colorValue) {
			return Color.FromArgb((byte)(colorValue >> 24), (byte)colorValue, (byte)(colorValue >> 8), (byte)(colorValue >> 16));
		}
		static SolidColorBrush GetBrush(Color color) {
			var brush = new SolidColorBrush(color);
			brush.Freeze();
			return brush;
		}
		// Microsoft.VisualStudio.Platform.WindowManagement.ColorNameTranslator
		static readonly Guid environmentColors = new Guid("{624ed9c3-bdfd-41fa-96c3-7c824ea32e3d}");
		static int VsColorFromName(ColorName colorName) {
			if (colorName.Category != environmentColors)
				return 0;
			int result;
			VsColors.TryGetColorIDFromBaseKey(colorName.Name, out result);
			return result;
		}
		#endregion


		// All of these types are internal.
		static readonly Assembly WindowManagement = typeof(Microsoft.VisualStudio.Platform.WindowManagement.ColorThemeService).Assembly;
		static readonly Type ResourceSynchronizer = WindowManagement.GetType("Microsoft.VisualStudio.Platform.WindowManagement.ResourceSynchronizer");

		// This method doesn't use the instance at all, and adds values directly, so I can use it as-is.
		static readonly Action<ResourceDictionary> AddTextureBrushes =
			(Action<ResourceDictionary>)Delegate.CreateDelegate(
				typeof(Action<ResourceDictionary>),
				null,
				ResourceSynchronizer.GetMethod("AddTextureBrushes", BindingFlags.NonPublic | BindingFlags.Instance));

		#region AddGradients
		static readonly IEnumerable<object> gradients = (IEnumerable<object>)ResourceSynchronizer.GetField("gradients", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null);

		// Ugly hack to call a method that takes an internal type as a parameter
		static Func<object, TResult> MakeRuntimeTypeFunc<TResult>(Type privateType, MethodInfo method) {
			return (Func<object, TResult>)MakeFunc1Method
				.MakeGenericMethod(privateType, typeof(TResult))
				.Invoke(null, new[] { method });
		}
		static readonly MethodInfo MakeFunc1Method = typeof(ThemeColorsDictionary).GetMethod("MakeFunc1", BindingFlags.Static | BindingFlags.NonPublic);
		static Func<object, TResult> MakeFunc1<TPrivate, TResult>(MethodInfo method) {
			var typedDelegate = (Func<TPrivate, TResult>)Delegate.CreateDelegate(typeof(Func<TPrivate, TResult>), method);
			return o => typedDelegate((TPrivate)o);
		}

		static Func<object, T2, TResult> MakeRuntimeTypeFunc<T2, TResult>(Type privateType, MethodInfo method) {
			return (Func<object, T2, TResult>)MakeFunc2Method
				.MakeGenericMethod(privateType, typeof(T2), typeof(TResult))
				.Invoke(null, new[] { method });
		}
		static readonly MethodInfo MakeFunc2Method = typeof(ThemeColorsDictionary).GetMethod("MakeFunc2", BindingFlags.Static | BindingFlags.NonPublic);
		static Func<object, T2, TResult> MakeFunc2<TPrivate, T2, TResult>(MethodInfo method) {
			var typedDelegate = (Func<TPrivate, T2, TResult>)Delegate.CreateDelegate(typeof(Func<TPrivate, T2, TResult>), method);
			return (o1, o2) => typedDelegate((TPrivate)o1, o2);
		}

		static readonly Type Gradient = WindowManagement.GetType("Microsoft.VisualStudio.Platform.WindowManagement.Gradient");
		static readonly Func<object, object> GradientKey
			= MakeRuntimeTypeFunc<object>(Gradient, Gradient.GetProperty("Key").GetMethod);
		static readonly Func<object, ResourceDictionary, Brush> GradientCreateBrush
			= MakeRuntimeTypeFunc<ResourceDictionary, Brush>(Gradient, Gradient.GetMethod("CreateBrush"));

		void AddGradients(ResourceDictionary newDictionary) {
			foreach (var gradient in gradients) {
				newDictionary.Add(GradientKey(gradient), GradientCreateBrush(gradient, newDictionary));
			}
		}
		#endregion

		#region AddFonts
		// Copied from ResourceSynchronizer and modified to not use VS native font utilities
		void AddFonts(ResourceDictionary newDictionary) {
			var dialogFont = System.Drawing.SystemFonts.CaptionFont;

			newDictionary.Add("VsFont.EnvironmentFontFamily", new FontFamily(dialogFont.Name));
			newDictionary.Add("VsFont.EnvironmentFontSize", (double)dialogFont.Size);
			AddCaptionFonts(newDictionary);
		}

		private void AddCaptionFonts(ResourceDictionary newDictionary) {
			NONCLIENTMETRICS nONCLIENTMETRICS = default(NONCLIENTMETRICS);
			nONCLIENTMETRICS.cbSize = Marshal.SizeOf(typeof(NONCLIENTMETRICS));
			if (!NativeMethods.SystemParametersInfo(41, nONCLIENTMETRICS.cbSize, ref nONCLIENTMETRICS, 0)) {
				newDictionary.Add("VsFont.CaptionFontFamily", this["VsFont.EnvironmentFontFamilyKey"]);
				newDictionary.Add("VsFont.CaptionFontSize", this["VsFont.EnvironmentFontSizeKey"]);
				newDictionary.Add("VsFont.CaptionFontWeight", FontWeights.Normal);
				return;
			}
			FontFamily captionFont = new FontFamily(nONCLIENTMETRICS.lfCaptionFont.lfFaceName);
			double captionSize = FontSizeFromLOGFONTHeight(nONCLIENTMETRICS.lfCaptionFont.lfHeight);
			FontWeight fontWeight = FontWeight.FromOpenTypeWeight(nONCLIENTMETRICS.lfCaptionFont.lfWeight);
			newDictionary.Add("VsFont.CaptionFontFamily", captionFont);
			newDictionary.Add("VsFont.CaptionFontSize", captionSize);
			newDictionary.Add("VsFont.CaptionFontWeight", fontWeight);
		}
		private double FontSizeFromLOGFONTHeight(int lfHeight) {
			return Math.Abs(lfHeight) * DpiHelper.DeviceToLogicalUnitsScalingFactorY;
		}
		#endregion
	}
	public class ColorTheme {
		public ColorTheme(int index, object theme) {
			Index = index;

			// The VS ColorTheme.Name property tries to read the translated
			// name from the resource ID using IVsResourceManager, which is
			// native code.  I extract the name from the ID instead.
			theme.GetType().GetMethod("RealizeLocalizedResourceID", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(theme, null);
			var nameId = (string)theme.GetType().GetField("_localizedNameResourceId", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(theme);
			if (nameId[0] == '@') {
				nameId = nameId.Replace("@ThemeName_", "");
				var end = nameId.IndexOf(";");
				if (end > 0)
					nameId = nameId.Remove(end);
			}
			Name = nameId;
		}
		public int Index { get; private set; }
		public string Name { get; private set; }
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