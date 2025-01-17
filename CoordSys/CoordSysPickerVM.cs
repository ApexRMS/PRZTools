﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Threading;
using ArcGIS.Desktop.Framework.Contracts;

namespace NCC.PRZTools
{
    internal class CoordSystemInfo
    {
        public string Region { get; set; }
        public string Name { get; set; }
        public string InternalName { get; set; }
        public int WKID { get; set; }
        public string WKT { get; set; }
    }

    internal class CoordSystemRegion
    {
        public string Name { get; set; }
        public List<CoordSystemInfo> CoordSystemInfos { get; set; }
    }

    internal class CoordSystemCollection
    {
        public string Name { get; set; }
        public List<CoordSystemRegion> Regions { get; set; }
    }

    internal class CoordSysPickerVM : PropertyChangedBase
    {
        private bool _executing = false;
        private CoordSystemInfo _selectedCoordSystemInfo = null;
        private object _selectedObject = null;

        private DispatcherTimer _timer = null;
        private double _progressValue = 1.0;
        private double _maxProgressValue = 100;
        private static readonly object _theLock = new object();

        private static ObservableCollection<CoordSystemCollection> _coordSystems = null;

        public CoordSysPickerVM()
        {
            IsExecutingQuery = _coordSystems == null;
            BindingOperations.CollectionRegistering += BindingOperations_CollectionRegistering;

            //We only need to load the coordinate systems the first time
            if (_coordSystems == null)
            {
                //init timer
                _timer = new DispatcherTimer()
                {
                    Interval = TimeSpan.FromMilliseconds(25d),
                    IsEnabled = false
                };
                _timer.Tick += (o, e) => {
                    //update the progress bar
                    _progressValue += 1.0;
                    if (_progressValue > _maxProgressValue)
                        _progressValue = 1.0;
                    System.Windows.Application.Current.Dispatcher.Invoke(() => {
                        NotifyPropertyChanged("ProgressValue");
                    });
                };

                LoadSpatialReferencesAsync();
            }
        }

        void BindingOperations_CollectionRegistering(object sender, CollectionRegisteringEventArgs e)
        {
            //register all the collections
            BindingOperations.EnableCollectionSynchronization(_coordSystems, _theLock);

            //unregister - we only need this event once
            BindingOperations.CollectionRegistering -= BindingOperations_CollectionRegistering;
        }

        public double ProgressValue => _progressValue;

        public double MaxProgressValue => _maxProgressValue;

        public string SelectedCoordinateSystemName => _selectedCoordSystemInfo != null ? _selectedCoordSystemInfo.Name : "";

        public object SelectedObject
        {
            get => _selectedObject;
            set
            {
                _selectedObject = value;
                if (_selectedObject == null)
                    return;
                if ((_selectedObject as CoordSystemInfo) != null)
                {
                    _selectedCoordSystemInfo = (CoordSystemInfo)_selectedObject;
                    NotifyPropertyChanged("SelectedCoordinateSystemName");
                }
            }
        }

        public CoordSystemInfo SelectedCoordSystemInfo => _selectedCoordSystemInfo;

        public bool IsExecutingQuery
        {
            get => _executing;
            set => SetProperty(ref _executing, value, () => IsExecutingQuery);
        }

        public ObservableCollection<CoordSystemCollection> CoordinateSystems => _coordSystems;

        private async void LoadSpatialReferencesAsync()
        {

            //init - placeholder
            IsExecutingQuery = true;
            _timer.Start();

            _coordSystems = new ObservableCollection<CoordSystemCollection>();
            _coordSystems.Add(new CoordSystemCollection()
            {
                Name = "Loading Geographic Coordinate Systems...",
                Regions = null
            });

            _coordSystems.Add(new CoordSystemCollection()
            {
                Name = "Loading Projected Coordinate Systems...",
                Regions = null
            });

            try
            {
                await System.Threading.Tasks.Task.Run(() => {
                    Dictionary<string, Dictionary<string, CoordSystemRegion>> regions =
                    new Dictionary<string, Dictionary<string, CoordSystemRegion>>();
                    regions.Add("Geographic Coordinate Systems", new Dictionary<string, CoordSystemRegion>());//Geographic
                    regions.Add("Projected Coordinate Systems", new Dictionary<string, CoordSystemRegion>());//Projected

                    string content = Properties.Settings.Default.COORDSYSPICKER_SROUT;//Stored coordinate system information

                    int startIndex = 0;
                    string coordSystemType = "]]Geographic Coordinate Systems";
                    string remainder = "";
                    int idx = 9999;
                    while (idx > 0)
                    {
                        idx = content.IndexOf(coordSystemType);
                        if (idx < 0 && coordSystemType == "]]Geographic Coordinate Systems")
                        {
                            coordSystemType = "]]Projected Coordinate Systems";
                            idx = content.IndexOf(coordSystemType);
                        }
                        //now if it is less than zero, we are on the last coordinate system record
                        //Note: "+2" to skip the ']]' in the search string
                        string sr_record = content.Substring(0, idx > 0 ? idx + 2 : content.Length);

                        //each sr comes in blocks of 4
                        string[] srs = sr_record.Split(new char[] { '|' });
                        //Get the name components first - includes the coordinate system type
                        //and the region
                        string[] names = srs[0].Split(new char[] { '/' });
                        var srCategory = names[0];
                        string sep = "";
                        string srRegion = "";
                        for (int n = 1; n < names.Length - 1; n++)
                        {
                            srRegion += sep + names[n];
                            sep = ", ";
                        }
                        var name = string.Format("{0} ({1})", names[names.Length - 1], srs[1]);
                        var wkid = Int32.Parse(srs[2]);
                        var wkt = srs[3];

                        //save it
                        CoordSystemRegion region = null;
                        if (!regions[srCategory].ContainsKey(srRegion))
                        {
                            region = new CoordSystemRegion();
                            region.CoordSystemInfos = new List<CoordSystemInfo>();
                            region.Name = srRegion;
                            regions[srCategory].Add(srRegion, region);
                        }
                        else
                        {
                            region = regions[srCategory][srRegion];
                        }
                        CoordSystemInfo cs = new CoordSystemInfo()
                        {
                            Name = name,
                            InternalName = srs[1],
                            Region = srRegion,
                            WKID = wkid,
                            WKT = wkt
                        };
                        region.CoordSystemInfos.Add(cs);

                        if (idx > 0)
                        {
                            remainder = content.Substring(idx + 2);
                            content = remainder;
                        }
                    }

                    //save them
                    _coordSystems.Clear();
                    _coordSystems = null;
                    _coordSystems = new ObservableCollection<CoordSystemCollection>();
                    _coordSystems.Add(new CoordSystemCollection()
                    {
                        Name = "Geographic Coordinate Systems",
                        Regions = regions["Geographic Coordinate Systems"].Values.ToList()
                    });
                    _coordSystems.Add(new CoordSystemCollection()
                    {
                        Name = "Projected Coordinate Systems",
                        Regions = regions["Projected Coordinate Systems"].Values.ToList()
                    });
                });
            }
            finally
            {
                _timer.Stop();
                IsExecutingQuery = false;
                NotifyPropertyChanged("CoordinateSystems");
            }
        }
    }
}