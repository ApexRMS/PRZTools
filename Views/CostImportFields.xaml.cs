namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for CostImportFields.xaml
    /// </summary>
    public partial class CostImportFields : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal CostImportFieldsVM vm = new CostImportFieldsVM();

        public CostImportFields()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
