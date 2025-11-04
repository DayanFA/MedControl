using System;
using System.Windows.Forms;
using System.Drawing;

namespace MedControl.UI
{
	// Painel quadrado: mantém largura = altura para cartões de chave
	public class SquareCardPanel : Panel
	{
		// Tamanho sugerido para inicialização
		public int SizeHint { get; set; } = 180;

		protected override void OnCreateControl()
		{
			base.OnCreateControl();
			try
			{
				Width = SizeHint;
				Height = SizeHint;
			}
			catch { }
		}

		protected override void OnSizeChanged(EventArgs e)
		{
			base.OnSizeChanged(e);
			try
			{
				if (Height != Width)
				{
					// Mantém quadrado tomando a largura como referência
					Height = Width;
				}
			}
			catch { }
		}
	}
}
