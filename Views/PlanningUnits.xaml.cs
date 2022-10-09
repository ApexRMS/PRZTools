namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for PlanningUnits.xaml
    /// </summary>
    public partial class PlanningUnits : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal PlanningUnitsVM vm = new PlanningUnitsVM();

        public PlanningUnits()
        {
            InitializeComponent();
            this.DataContext = vm;
        }

    }
}
