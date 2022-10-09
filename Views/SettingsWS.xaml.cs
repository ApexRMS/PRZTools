namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for WorkspaceSettings.xaml
    /// </summary>
    public partial class SettingsWS : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal SettingsWSVM vm = new SettingsWSVM();

        public SettingsWS()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
