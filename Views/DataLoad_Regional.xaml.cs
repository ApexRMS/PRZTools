namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for DataLoad_Regional.xaml
    /// </summary>
    public partial class DataLoad_Regional : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal DataLoad_RegionalVM vm = new DataLoad_RegionalVM();

        public DataLoad_Regional()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
