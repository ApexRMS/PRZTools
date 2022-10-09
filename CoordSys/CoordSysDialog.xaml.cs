using System.Windows;
using ArcGIS.Core.Geometry;

namespace NCC.PRZTools
{
    /// <summary>
    /// Interaction logic for CoordSysDialog.xaml
    /// </summary>
    public partial class CoordSysDialog : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        internal CoordSysDialogVM vm = new CoordSysDialogVM();

        public CoordSysDialog()
        {
            InitializeComponent();

            this.DataContext = vm;

            this.CoordinateSystemsControl.SelectedSpatialReferenceChanged += (s, args) => {
                vm.SelectedSpatialReference = args.SpatialReference;
            };
        }

        public SpatialReference SpatialReference => vm.SelectedSpatialReference;

        private void Close_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
