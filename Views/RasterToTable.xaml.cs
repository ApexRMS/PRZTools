namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for RasterToTable.xaml
    /// </summary>
    public partial class RasterToTable : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal RasterToTableVM vm = new RasterToTableVM();

        public RasterToTable()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
