namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for ExportToWTW.xaml
    /// </summary>
    public partial class ExportWTW : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal ExportWTWVM vm = new ExportWTWVM();

        public ExportWTW()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
