using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MedControl.UI
{
	// Helpers to apply Windows 11 Fluent effects (Mica/Acrylic) and rounded corners without external libs
	public static class FluentEffects
	{
		// Public entry point: try Mica, fall back to Acrylic/Blur
		public static void ApplyWin11Mica(Form form)
		{
			if (form == null || form.IsDisposed) return;
			try
			{
				form.BackColor = Color.FromArgb(245, 246, 248); // subtle light background
				EnableMica(form.Handle);
				SetRoundedCorners(form.Handle, 2); // DWMWCP_ROUND
			}
			catch
			{
				// Fallback to Acrylic/Accent blur if Mica not supported
				try { EnableBlur(form.Handle); } catch { /* ignore */ }
			}
		}

		// Reset backdrop to default (disable Mica/Acrylic effects)
		public static void ResetWin11Backdrop(Form form)
		{
			if (form == null || form.IsDisposed) return;
			try
			{
				DisableBackdrop(form.Handle);
				SetRoundedCorners(form.Handle, 0); // default corners
			}
			catch { }
		}

		public static void SetRoundedCorners(IntPtr hWnd, int cornerPreference)
		{
			// 33 = DWMWA_WINDOW_CORNER_PREFERENCE, values: 0=Default,1=DoNotRound,2=Round,3=RoundSmall
			const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
			if (hWnd == IntPtr.Zero) return;
			DwmSetWindowAttribute(hWnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref cornerPreference, sizeof(int));
		}

		private static void EnableMica(IntPtr hWnd)
		{
			if (hWnd == IntPtr.Zero) return;
			// 38 = DWMWA_SYSTEMBACKDROP_TYPE. 2 = Mica, 3 = Acrylic (on Win11)
			const int DWMWA_SYSTEMBACKDROP_TYPE = 38;
			int mica = 2;
			DwmSetWindowAttribute(hWnd, DWMWA_SYSTEMBACKDROP_TYPE, ref mica, sizeof(int));
		}

		private static void DisableBackdrop(IntPtr hWnd)
		{
			if (hWnd == IntPtr.Zero) return;
			const int DWMWA_SYSTEMBACKDROP_TYPE = 38;
			int none = 1; // DWMSBT_NONE
			DwmSetWindowAttribute(hWnd, DWMWA_SYSTEMBACKDROP_TYPE, ref none, sizeof(int));
		}

		private static void EnableBlur(IntPtr hWnd)
		{
			if (hWnd == IntPtr.Zero) return;
			// Use SetWindowCompositionAttribute with ACCENT_ENABLE_BLURBEHIND
			var accent = new ACCENT_POLICY
			{
				AccentState = ACCENT_STATE.ACCENT_ENABLE_BLURBEHIND,
				AccentFlags = 0,
				GradientColor = 0,
				AnimationId = 0
			};
			int size = Marshal.SizeOf(accent);
			IntPtr accentPtr = Marshal.AllocHGlobal(size);
			try
			{
				Marshal.StructureToPtr(accent, accentPtr, false);
				var data = new WINDOWCOMPOSITIONATTRIBDATA
				{
					Attribute = WINDOWCOMPOSITIONATTRIB.WCA_ACCENT_POLICY,
					Data = accentPtr,
					SizeOfData = size
				};
				SetWindowCompositionAttribute(hWnd, ref data);
			}
			finally
			{
				Marshal.FreeHGlobal(accentPtr);
			}
		}

		[DllImport("dwmapi.dll")]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

		[DllImport("user32.dll")]
		private static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WINDOWCOMPOSITIONATTRIBDATA data);

		private enum WINDOWCOMPOSITIONATTRIB
		{
			WCA_ACCENT_POLICY = 19
		}

		private enum ACCENT_STATE
		{
			ACCENT_DISABLED = 0,
			ACCENT_ENABLE_GRADIENT = 1,
			ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
			ACCENT_ENABLE_BLURBEHIND = 3,
			ACCENT_ENABLE_ACRYLICBLURBEHIND = 4,
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct ACCENT_POLICY
		{
			public ACCENT_STATE AccentState;
			public int AccentFlags;
			public int GradientColor;
			public int AnimationId;
		}

		[StructLayout(LayoutKind.Sequential)]
		private struct WINDOWCOMPOSITIONATTRIBDATA
		{
			public WINDOWCOMPOSITIONATTRIB Attribute;
			public IntPtr Data;
			public int SizeOfData;
		}
	}
}
