namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for PUCost.xaml
    /// </summary>
    public partial class PUCost : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal PUCostVM vm = new PUCostVM();

        public PUCost()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
