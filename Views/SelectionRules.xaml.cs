namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for PUStatus.xaml
    /// </summary>
    public partial class SelectionRules : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public SelectionRulesVM vm = new SelectionRulesVM();

        public SelectionRules()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
