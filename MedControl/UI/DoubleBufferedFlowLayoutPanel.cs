using System.Windows.Forms;

namespace MedControl.UI
{
    // Minimiza flicker em FlowLayoutPanel durante grandes atualizações
    public class DoubleBufferedFlowLayoutPanel : FlowLayoutPanel
    {
        public DoubleBufferedFlowLayoutPanel()
        {
            this.DoubleBuffered = true;
            this.ResizeRedraw = true;
        }
    }
}
