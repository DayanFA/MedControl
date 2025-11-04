using System;
using System.Drawing;
using System.Windows.Forms;

namespace MedControl.UI
{
	// Simple flat renderer that respects theme colors for MenuStrip/ToolStrip
	public class ModernMenuRenderer : ToolStripProfessionalRenderer
	{
		private readonly Color _bg;
		private readonly Color _fg;
		private readonly Color _sel;
		private readonly Color _selBorder;

		public ModernMenuRenderer(Color background, Color foreground, Color selection, Color selectionBorder)
			: base(new FlatColorTable(background, selection))
		{
			_bg = background; _fg = foreground; _sel = selection; _selBorder = selectionBorder;
		}

		protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
		{
			// no border
		}

		protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
		{
			Rectangle rc = new Rectangle(Point.Empty, e.Item.Size);
			if (e.Item.Selected || e.Item.Pressed)
			{
				using var b = new SolidBrush(_sel);
				e.Graphics.FillRectangle(b, rc);
				using var p = new Pen(_selBorder);
				rc.Width -= 1; rc.Height -= 1;
				e.Graphics.DrawRectangle(p, rc);
			}
			else
			{
				using var b = new SolidBrush(_bg);
				e.Graphics.FillRectangle(b, rc);
			}
			e.Item.ForeColor = _fg;
		}

		private class FlatColorTable : ProfessionalColorTable
		{
			private readonly Color _bg;
			private readonly Color _sel;
			public FlatColorTable(Color background, Color selection)
			{ _bg = background; _sel = selection; }
			public override Color MenuItemSelected => _sel;
			public override Color MenuItemSelectedGradientBegin => _sel;
			public override Color MenuItemSelectedGradientEnd => _sel;
			public override Color MenuItemPressedGradientBegin => _sel;
			public override Color MenuItemPressedGradientEnd => _sel;
			public override Color ToolStripDropDownBackground => _bg;
			public override Color ToolStripGradientBegin => _bg;
			public override Color ToolStripGradientMiddle => _bg;
			public override Color ToolStripGradientEnd => _bg;
			// Do not override UseSystemColors; keep framework default
		}
	}
}

