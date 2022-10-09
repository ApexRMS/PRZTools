namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for DataLoad_National.xaml
    /// </summary>
    public partial class DataLoad_National : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal DataLoad_NationalVM vm = new DataLoad_NationalVM();

        public DataLoad_National()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
