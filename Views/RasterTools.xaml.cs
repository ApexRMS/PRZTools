namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for RasterTools.xaml
    /// </summary>
    public partial class RasterTools : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal RasterToolsVM vm = new RasterToolsVM();

        public RasterTools()
        {
            InitializeComponent();
            this.DataContext = vm;
        }

    }
}
