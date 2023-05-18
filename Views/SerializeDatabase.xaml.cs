namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for DataLoad_National.xaml
    /// </summary>
    public partial class SerializeDatabase : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal SerializeDatabaseVM vm = new SerializeDatabaseVM();

        public SerializeDatabase()
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
