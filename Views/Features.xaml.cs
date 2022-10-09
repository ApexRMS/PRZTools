namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for Features.xaml
    /// </summary>
    public partial class Features : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal FeaturesVM vm = new FeaturesVM();
        public Features()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
