using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Controls;
using ArcGIS.Desktop.Framework.Contracts;

namespace NCC.PRZTools
{
    public class CoordSysDialogVM : PropertyChangedBase // TODO: Possibly remove the interface
    {
        private bool _showVCS = false;
        private SpatialReference _sr;
        private CoordinateSystemsControlProperties _props = null;

        public CoordSysDialogVM()
        {
            UpdateCoordinateControlProperties();
        }

        public string SelectedCoordinateSystemName => _sr != null ? _sr.Name : "";

        public SpatialReference SelectedSpatialReference
        {
            get => _sr;
            set
            {
                SetProperty(ref _sr, value, () => SelectedSpatialReference);
                UpdateCoordinateControlProperties();
            }
        }

        public bool ShowVCS
        {
            get => _showVCS;
            set
            {
                bool changed = SetProperty(ref _showVCS, value, () => ShowVCS);

                if (changed)
                {
                    UpdateCoordinateControlProperties();
                }
            }
        }

        public CoordinateSystemsControlProperties ControlProperties
        {
            get => _props;
            set => SetProperty(ref _props, value, () => ControlProperties);
        }

        private void UpdateCoordinateControlProperties()
        {
            var map = MapView.Active?.Map;
            var props = new CoordinateSystemsControlProperties()
            {
                Map = map,
                SpatialReference = this._sr,
                ShowVerticalCoordinateSystems = this.ShowVCS
            };
            this.ControlProperties = props;
        }
    }
}