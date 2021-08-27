﻿using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.Raster;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Controls;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ProMsgBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using PRZC = NCC.PRZTools.PRZConstants;
using PRZH = NCC.PRZTools.PRZHelper;
using PRZM = NCC.PRZTools.PRZMethods;

namespace NCC.PRZTools
{
    public class CFGeneratorVM : PropertyChangedBase
    {
        public CFGeneratorVM()
        {
        }


        #region Properties

        private bool _cfTableExists;
        public bool CFTableExists
        {
            get => _cfTableExists;
            set => SetProperty(ref _cfTableExists, value, () => CFTableExists);
        }

        private bool _puvcfTableExists;
        public bool PUVCFTableExists
        {
            get => _puvcfTableExists;
            set => SetProperty(ref _puvcfTableExists, value, () => PUVCFTableExists);
        }

        private string _gridCaption;
        public string GridCaption
        {
            get => _gridCaption; set => SetProperty(ref _gridCaption, value, () => GridCaption);
        }

        private string _defaultThreshold = Properties.Settings.Default.DEFAULT_CF_THRESHOLD;
        public string DefaultThreshold
        {
            get => _defaultThreshold;

            set
            {
                SetProperty(ref _defaultThreshold, value, () => DefaultThreshold);
                Properties.Settings.Default.DEFAULT_CF_THRESHOLD = value;
                Properties.Settings.Default.Save();
            }
        }

        private string _defaultTarget = Properties.Settings.Default.DEFAULT_CF_TARGET;
        public string DefaultTarget
        {
            get => _defaultTarget;

            set
            {
                SetProperty(ref _defaultTarget, value, () => DefaultTarget);
                Properties.Settings.Default.DEFAULT_CF_TARGET = value;
                Properties.Settings.Default.Save();
            }
        }

        private ObservableCollection<ConservationFeature> _conservationFeatures = new ObservableCollection<ConservationFeature>();
        public ObservableCollection<ConservationFeature> ConservationFeatures
        {
            get { return _conservationFeatures; }
            set
            {
                _conservationFeatures = value;
                SetProperty(ref _conservationFeatures, value, () => ConservationFeatures);
            }
        }

        private ConservationFeature _selectedConservationFeature;
        public ConservationFeature SelectedConservationFeature
        {
            get => _selectedConservationFeature; set => SetProperty(ref _selectedConservationFeature, value, () => SelectedConservationFeature);
        }



        private ProgressManager _pm = ProgressManager.CreateProgressManager(50);    // initialized to min=0, current=0, message=""
        public ProgressManager PM
        {
            get => _pm; set => SetProperty(ref _pm, value, () => PM);
        }

        #endregion

        #region Commands

        private ICommand _cmdConstraintDoubleClick;
        public ICommand CmdGridDoubleClick => _cmdConstraintDoubleClick ?? (_cmdConstraintDoubleClick = new RelayCommand(() => GridDoubleClick(), () => true));

        private ICommand _cmdBuildCFTable;
        public ICommand CmdBuildCFTable => _cmdBuildCFTable ?? (_cmdBuildCFTable = new RelayCommand(() => BuildCFTable(), () => true));

        private ICommand _cmdClearLog;
        public ICommand CmdClearLog => _cmdClearLog ?? (_cmdClearLog = new RelayCommand(() =>
        {
            PRZH.UpdateProgress(PM, "", false, 0, 1, 0);
        }, () => true));


        #endregion

        #region Methods

        public async Task OnProWinLoaded()
        {
            try
            {
                // Initialize the Progress Bar & Log
                PRZH.UpdateProgress(PM, "", false, 0, 1, 0);

                // checkboxes
                CFTableExists = await PRZH.CFTableExists();
                PUVCFTableExists = await PRZH.PUVCFTableExists();

                // Populate the Grid
                bool Populated = await PopulateCFGrid();

            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
        }

        private async Task<bool> PopulateCFGrid()
        {
            try
            {
                // Clear the contents of the Conflicts observable collection
                ConservationFeatures.Clear();

                if (!await PRZH.CFTableExists())
                {
                    // format stuff appropriately if no table exists
                    GridCaption = "Conservation Features Table not yet built...";
                    //return true;
                }

                // CF Table Exists - retrieve data from it









                //// PUStatus Table exists, retrieve the data
                //Dictionary<int, string> DICT_IN = new Dictionary<int, string>();    // Dictionary where key = Area Column Indexes, value = IN Constraint Layer to which it applies
                //Dictionary<int, string> DICT_EX = new Dictionary<int, string>();    // Dictionary where key = Area Column Indexes, value = EX Constraint Layer to which it applies

                //List<Field> fields = null;  // List of Planning Unit Status table fields

                //// Populate the Dictionaries
                //await QueuedTask.Run(async () =>
                //{
                //    using (Table table = await PRZH.GetStatusInfoTable())
                //    using (RowCursor rowCursor = table.Search(null, false))
                //    {
                //        TableDefinition tDef = table.GetDefinition();
                //        fields = tDef.GetFields().ToList();

                //        // Get the first row (I only need one row)
                //        if (rowCursor.MoveNext())
                //        {
                //            using (Row row = rowCursor.Current)
                //            {
                //                // Now, get field info and row value info that's present in all rows exactly the same (that's why I only need one row)
                //                for (int i = 0; i < fields.Count; i++)
                //                {
                //                    Field field = fields[i];

                //                    if (field.Name.EndsWith("_Name"))
                //                    {
                //                        string constraint_name = row[i].ToString();         // this is the name of the constraint layer to which columns i to i+2 apply

                //                        if (field.Name.StartsWith("IN"))
                //                        {
                //                            DICT_IN.Add(i + 2, constraint_name);    // i + 2 is the Area field, two columns to the right of the Name field
                //                        }
                //                        else if (field.Name.StartsWith("EX"))
                //                        {
                //                            DICT_EX.Add(i + 2, constraint_name);    // i + 2 is the Area field, two columns to the right of the Name field
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }
                //});

                //// Build a List of Unique Combinations of IN and EX Layers
                //string c_ConflictNumber = "CONFLICT";
                //string c_LayerName_Include = "'INCLUDE' Constraint";
                //string c_LayerName_Exclude = "'EXCLUDE' Constraint";
                //string c_PUCount = "PLANNING UNIT COUNT";
                //string c_AreaFieldIndex_Include = "IndexIN";
                //string c_AreaFieldIndex_Exclude = "IndexEX";
                //string c_ConflictExists = "Exists";

                //DataTable DT = new DataTable();
                //DT.Columns.Add(c_ConflictNumber, Type.GetType("System.Int32"));
                //DT.Columns.Add(c_LayerName_Include, Type.GetType("System.String"));
                //DT.Columns.Add(c_LayerName_Exclude, Type.GetType("System.String"));
                //DT.Columns.Add(c_PUCount, Type.GetType("System.Int32"));
                //DT.Columns.Add(c_AreaFieldIndex_Include, Type.GetType("System.Int32"));
                //DT.Columns.Add(c_AreaFieldIndex_Exclude, Type.GetType("System.Int32"));
                //DT.Columns.Add(c_ConflictExists, Type.GetType("System.Boolean"));

                //foreach (int IN_AreaFieldIndex in DICT_IN.Keys)
                //{
                //    string IN_LayerName = DICT_IN[IN_AreaFieldIndex];

                //    foreach (int EX_AreaFieldIndex in DICT_EX.Keys)
                //    {
                //        string EX_LayerName = DICT_EX[EX_AreaFieldIndex];

                //        DataRow DR = DT.NewRow();

                //        DR[c_LayerName_Include] = IN_LayerName;
                //        DR[c_LayerName_Exclude] = EX_LayerName;
                //        DR[c_PUCount] = 0;
                //        DR[c_AreaFieldIndex_Include] = IN_AreaFieldIndex;
                //        DR[c_AreaFieldIndex_Exclude] = EX_AreaFieldIndex;
                //        DR[c_ConflictExists] = false;
                //        DT.Rows.Add(DR);
                //    }
                //}

                //// For each row in DataTable, query PU Status for pairs having area>0 in both IN and EX
                //int conflict_number = 1;
                //int IN_AreaField_Index;
                //int EX_AreaField_Index;
                //string IN_AreaField_Name = "";
                //string EX_AreaField_Name = "";

                //foreach (DataRow DR in DT.Rows)
                //{
                //    IN_AreaField_Index = (int)DR[c_AreaFieldIndex_Include];
                //    EX_AreaField_Index = (int)DR[c_AreaFieldIndex_Exclude];

                //    IN_AreaField_Name = fields[IN_AreaField_Index].Name;
                //    EX_AreaField_Name = fields[EX_AreaField_Index].Name;

                //    string where_clause = IN_AreaField_Name + @" > 0 And " + EX_AreaField_Name + @" > 0";

                //    QueryFilter QF = new QueryFilter();
                //    QF.SubFields = IN_AreaField_Name + "," + EX_AreaField_Name;
                //    QF.WhereClause = where_clause;

                //    int row_count = 0;

                //    await QueuedTask.Run(async () =>
                //    {
                //        using (Table table = await PRZH.GetStatusInfoTable())
                //        {
                //            row_count = table.GetCount(QF);
                //        }
                //    });

                //    if (row_count > 0)
                //    {
                //        DR[c_ConflictNumber] = conflict_number++;
                //        DR[c_PUCount] = row_count;
                //        DR[c_ConflictExists] = true;
                //    }
                //}


                //// Filter out only those DataRows where conflict exists
                //DataView DV = DT.DefaultView;
                //DV.RowFilter = c_ConflictExists + " = true";

                //// Finally, populate the Observable Collection

                //List<StatusConflict> l = new List<StatusConflict>();
                //foreach (DataRowView DRV in DV)
                //{
                //    StatusConflict sc = new StatusConflict();

                //    sc.include_layer_name = DRV[c_LayerName_Include].ToString();
                //    sc.include_area_field_index = (int)DRV[c_AreaFieldIndex_Include];
                //    sc.exclude_layer_name = DRV[c_LayerName_Exclude].ToString();
                //    sc.exclude_area_field_index = (int)DRV[c_AreaFieldIndex_Exclude];
                //    sc.conflict_num = (int)DRV[c_ConflictNumber];
                //    sc.pu_count = (int)DRV[c_PUCount];

                //    l.Add(sc);
                //}

                //// Sort them
                //l.Sort((x, y) => x.conflict_num.CompareTo(y.conflict_num));

                //// Set the property
                //_conflicts = new ObservableCollection<StatusConflict>(l);
                //NotifyPropertyChanged(() => Conflicts);

                //int count = DV.Count;

                //ConflictGridCaption = "Planning Unit Status Conflict Listing (" + ((count == 1) ? "1 conflict)" : count.ToString() + " conflicts)");

                List<ConservationFeature> l = new List<ConservationFeature>();
                l.Add(new ConservationFeature
                {
                    cf_id = 1,
                    cf_name = "Blue Meadows",
                    area_m2 = 123,
                    area_ac = 2,
                    area_ha = 1,
                    area_km2 = 0.3
                });

                l.Add(new ConservationFeature
                {
                    cf_id = 2,
                    cf_name = "Green Acres",
                    area_m2 = 124,
                    area_ac = 3,
                    area_ha = 2,
                    area_km2 = 1.3
                });

                // Sort them
                l.Sort((x, y) => x.cf_id.CompareTo(y.cf_id));

                // Set the property
                _conservationFeatures = new ObservableCollection<ConservationFeature>(l);
                NotifyPropertyChanged(() => ConservationFeatures);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                PRZH.UpdateProgress(PM, PRZH.WriteLog(ex.Message, LogMessageType.ERROR), true);
                return false;
            }
        }

        private async Task<bool> BuildCFTable()
        {
            int val = 0;

            try
            {
                #region INITIALIZATION AND USER INPUT VALIDATION

                // Initialize a few thingies
                Map map = MapView.Active.Map;

                // Initialize ProgressBar and Progress Log
                int max = 50;
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Initializing the CF Table Generator..."), false, max, ++val);

                // Validation: Ensure the Project Geodatabase Exists
                string gdbpath = PRZH.GetProjectGDBPath();
                if (!await PRZH.ProjectGDBExists())
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Project Geodatabase not found: " + gdbpath, LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show("Project Geodatabase not found at this path:" +
                                   Environment.NewLine +
                                   gdbpath +
                                   Environment.NewLine + Environment.NewLine +
                                   "Please specify a valid Project Workspace.", "Validation");

                    return false;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Project Geodatabase is OK: " + gdbpath), true, ++val);
                }

                // Validation: Ensure the Planning Unit FC exists
                string pufcpath = PRZH.GetPlanningUnitFCPath();
                if (!await PRZH.PlanningUnitFCExists())
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Planning Unit Feature Class not found in the Project Geodatabase.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    return false;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Planning Unit Feature Class is OK: " + pufcpath), true, ++val);
                }

                // Validation: Ensure that the Planning Unit Feature Layer exists
                if (!PRZH.FeatureLayerExists_PU(map))
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Planning Unit Feature Layer not found in the map.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show("Planning Unit Feature Layer not present in the map.  Please reload PRZ layers");
                    return false;
                }

                // Validation: Ensure the Default Threshold is valid
                string threshold_text = string.IsNullOrEmpty(DefaultThreshold) ? "0" : ((DefaultThreshold.Trim() == "") ? "0" : DefaultThreshold.Trim());

                if (!double.TryParse(threshold_text, out double threshold_double))
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Missing or invalid Default Threshold value", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show("Please specify a valid Default Threshold value.  Value must be a number between 0 and 100 (inclusive)", "Validation");
                    return false;
                }
                else if (threshold_double < 0 | threshold_double > 100)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Missing or invalid Default Threshold value", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show("Please specify a valid Default Threshold value.  Value must be a number between 0 and 100 (inclusive)", "Validation");
                    return false;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Default Threshold = " + threshold_text), true, ++val);
                }

                // Validation: Ensure the Default Target is valid
                string target_text = string.IsNullOrEmpty(DefaultTarget) ? "0" : ((DefaultTarget.Trim() == "") ? "0" : DefaultTarget.Trim());

                if (!double.TryParse(target_text, out double target_double))
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Missing or invalid Default Target value", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show("Please specify a valid Default Target value.  Value must be a number between 0 and 100 (inclusive)", "Validation");
                    return false;
                }
                else if (target_double < 0 | target_double > 100)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Missing or invalid Default Target value", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show("Please specify a valid Default Target value.  Value must be a number between 0 and 100 (inclusive)", "Validation");
                    return false;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Default Target = " + target_text), true, ++val);
                }


                //// Validation: Ensure three required Layers are present
                //if (!PRZH.PRZLayerExists(map, PRZLayerNames.STATUS_INCLUDE) || !PRZH.PRZLayerExists(map, PRZLayerNames.STATUS_EXCLUDE) || !PRZH.PRZLayerExists(map, PRZLayerNames.PU))
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Layers are missing.  Please reload PRZ layers.", LogMessageType.VALIDATION_ERROR), true, ++val);
                //    ProMsgBox.Show("PRZ Layers are missing.  Please reload the PRZ Layers and try again.", "Validation");
                //    return false;
                //}

                //// Validation: Ensure that at least one Feature Layer is present in either of the two group layers
                //var LIST_IncludeFL = PRZH.GetFeatureLayers_STATUS_INCLUDE(map);
                //var LIST_ExcludeFL = PRZH.GetFeatureLayers_STATUS_EXCLUDE(map);

                //if (LIST_IncludeFL == null || LIST_ExcludeFL == null)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> Unable to retrieve contents of Status Include or Status Exclude Group Layers.  Please reload PRZ layers.", LogMessageType.VALIDATION_ERROR), true, ++val);
                //    ProMsgBox.Show("Unable to retrieve contents of Status Include or Status Exclude Group Layers.  Please reload the PRZ Layers and try again.", "Validation");
                //    return false;
                //}

                //if (LIST_IncludeFL.Count == 0 && LIST_ExcludeFL.Count == 0)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Validation >> No Feature Layers found within Status Include or Status Exclude group layers.", LogMessageType.VALIDATION_ERROR), true, ++val);
                //    ProMsgBox.Show("There must be at least one Feature Layer within either the Status INCLUDE or the Status EXCLUDE group layers.", "Validation");
                //    return false;
                //}

                //// Validation: Prompt User for permission to proceed
                //if (ProMsgBox.Show("If you proceed, the Planning Unit Status table will be overwritten if it exists in the Project Geodatabase." +
                //   Environment.NewLine + Environment.NewLine +
                //   "Additionally, the contents of the 'status' field in the Planning Unit Feature Class will be updated." +
                //   Environment.NewLine + Environment.NewLine +
                //   "Do you wish to proceed?" +
                //   Environment.NewLine + Environment.NewLine +
                //   "Choose wisely...",
                //   "TABLE OVERWRITE WARNING",
                //   System.Windows.MessageBoxButton.OKCancel, System.Windows.MessageBoxImage.Exclamation,
                //   System.Windows.MessageBoxResult.Cancel) == System.Windows.MessageBoxResult.Cancel)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("User bailed out of Status Calculation."), true, ++val);
                //    return false;
                //}

                #endregion

                //#region PREPARE THE LAYER DATATABLES

                //// Create Include and Exclude Data Tables
                //// Each DataTable will contain one row per layer in that category (e.g. Include DT might contain 3 rows, one row per layer within the INclude group layer)
                //DataTable DT_IncludeLayers = new DataTable("INCLUDE");    // doesn't really need a name
                //DT_IncludeLayers.Columns.Add(PRZC.c_FLD_DATATABLE_STATUS_LAYER, typeof(Layer));
                //DT_IncludeLayers.Columns.Add(PRZC.c_FLD_DATATABLE_STATUS_INDEX, Type.GetType("System.Int32"));
                //DT_IncludeLayers.Columns.Add(PRZC.c_FLD_DATATABLE_STATUS_NAME, Type.GetType("System.String"));
                //DT_IncludeLayers.Columns.Add(PRZC.c_FLD_DATATABLE_STATUS_THRESHOLD, Type.GetType("System.Double"));
                //DT_IncludeLayers.Columns.Add(PRZC.c_FLD_DATATABLE_STATUS_STATUS, Type.GetType("System.Int32"));

                //DataTable DT_ExcludeLayers = DT_IncludeLayers.Clone();

                //if (!await PopulateLayerTable(PRZLayerNames.STATUS_INCLUDE, DT_IncludeLayers))
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("INCLUDE Layers >> Unable to populate Data Table.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Unable to populate the INCLUDE Data Table.", "INCLUDE Layers");
                //    return false;
                //}

                //if (!await PopulateLayerTable(PRZLayerNames.STATUS_EXCLUDE, DT_ExcludeLayers))
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("EXCLUDE Layers >> Unable to populate Data Table.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Unable to populate the EXCLUDE Data Table.", "EXCLUDE Layers");
                //    return false;
                //}

                //// Ensure that at least one row exists in one of the 2 DataTables.  Otherwise, quit

                //if (DT_IncludeLayers.Rows.Count == 0 && DT_ExcludeLayers.Rows.Count == 0)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Status Calculator >> No valid Polygon Feature Layers found from INCLUDE or EXCLUDE group layers.", LogMessageType.VALIDATION_ERROR), true, ++val);
                //    ProMsgBox.Show("There must be at least one valid Polygon Feature Layer in either the INCLUDE or EXCLUDE group layers.", "Status Calculator");
                //    return false;
                //}

                //#endregion

                //#region POPULATE 2 DICTIONARIES:  PUID -> AREA_M, and PUID -> STATUS

                //Dictionary<int, double> DICT_PUID_and_assoc_area_m2 = new Dictionary<int, double>();
                //Dictionary<int, int> DICT_PUID_and_assoc_status = new Dictionary<int, int>();

                //await QueuedTask.Run(async () =>
                //{
                //    using (Geodatabase gdb = await PRZH.GetProjectGDB())
                //    using (FeatureClass puFC = gdb.OpenDataset<FeatureClass>(PRZC.c_FC_PLANNING_UNITS))
                //    using (RowCursor rowCursor = puFC.Search(null, false))
                //    {
                //        // Get the Definition
                //        FeatureClassDefinition fcDef = puFC.GetDefinition();

                //        while (rowCursor.MoveNext())
                //        {
                //            using (Row row = rowCursor.Current)
                //            {
                //                int puid = (int)row[PRZC.c_FLD_PUFC_ID];
                //                double a = (double)row[PRZC.c_FLD_PUFC_AREA_M];
                //                int status = (int)row[PRZC.c_FLD_PUFC_STATUS];

                //                // store this id -> area KVP in the 1st dictionary
                //                DICT_PUID_and_assoc_area_m2.Add(puid, a);

                //                // store this id -> status KVP in the 2nd dictionary
                //                DICT_PUID_and_assoc_status.Add(puid, status);
                //            }
                //        }
                //    }
                //});

                //#endregion

                //// Start a stopwatch
                //Stopwatch stopwatch = new Stopwatch();
                //stopwatch.Start();

                //// Some GP variables
                //IReadOnlyList<string> toolParams;
                //IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                //string toolOutput;

                //#region BUILD THE STATUS INFO TABLE

                //string sipath = PRZH.GetStatusInfoTablePath();

                //// Delete the existing Status Info table, if it exists

                //if (await PRZH.StatusInfoTableExists())
                //{
                //    // Delete the existing Status Info table
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting Status Info Table..."), true, ++val);
                //    toolParams = Geoprocessing.MakeValueArray(sipath, "");
                //    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                //    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, GPExecuteToolFlags.RefreshProjectItems);
                //    if (toolOutput == null)
                //    {
                //        PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting the Status Info table.  GP Tool failed or was cancelled by user", LogMessageType.ERROR), true, ++val);
                //        return false;
                //    }
                //    else
                //    {
                //        PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleted the existing Status Info Table..."), true, ++val);
                //    }
                //}

                //// Copy PU FC rows into new Status Info table
                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Copying Planning Unit FC Attributes into new Status Info table..."), true, ++val);
                //toolParams = Geoprocessing.MakeValueArray(pufcpath, sipath, "");
                //toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                //toolOutput = await PRZH.RunGPTool("CopyRows_management", toolParams, toolEnvs, GPExecuteToolFlags.RefreshProjectItems);
                //if (toolOutput == null)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error copying PU FC rows to Status Info table.  GP Tool failed or was cancelled by user", LogMessageType.ERROR), true, ++val);
                //    return false;
                //}
                //else
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Status Info Table successfully created and populated..."), true, ++val);
                //}

                //// Delete all fields but OID and PUID from Status Info table
                //List<string> LIST_DeleteFields = new List<string>();

                //using (Table tab = await PRZH.GetStatusInfoTable())
                //{
                //    if (tab == null)
                //    {
                //        ProMsgBox.Show("Error getting Status Info Table :(");
                //        return false;
                //    }

                //    await QueuedTask.Run(() =>
                //    {
                //        TableDefinition tDef = tab.GetDefinition();
                //        List<Field> fields = tDef.GetFields().Where(f => f.FieldType != FieldType.OID && f.Name != PRZC.c_FLD_PUFC_ID).ToList();

                //        foreach (Field field in fields)
                //        {
                //            LIST_DeleteFields.Add(field.Name);
                //        }
                //    });
                //}

                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Removing unnecessary fields from the Status Info table..."), true, ++val);
                //toolParams = Geoprocessing.MakeValueArray(sipath, LIST_DeleteFields);
                //toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                //toolOutput = await PRZH.RunGPTool("DeleteField_management", toolParams, toolEnvs, GPExecuteToolFlags.RefreshProjectItems);
                //if (toolOutput == null)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting fields from Status Info table.  GP Tool failed or was cancelled by user", LogMessageType.ERROR), true, ++val);
                //    return false;
                //}
                //else
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Status Info Table fields successfully deleted"), true, ++val);
                //}

                //// Now index the PUID field in the Status Info table
                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Indexing Planning Unit ID field in the Status Info table..."), true, ++val);
                //toolParams = Geoprocessing.MakeValueArray(sipath, PRZC.c_FLD_PUFC_ID, "ix" + PRZC.c_FLD_PUFC_ID, "", "");
                //toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                //toolOutput = await PRZH.RunGPTool("AddIndex_management", toolParams, toolEnvs, GPExecuteToolFlags.RefreshProjectItems);
                //if (toolOutput == null)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error indexing fields.  GP Tool failed or was cancelled by user", LogMessageType.ERROR), true, ++val);
                //    return false;
                //}

                //// Add 2 additional fields to Status Info
                //string fldQuickStatus = PRZC.c_FLD_STATUSINFO_QUICKSTATUS + " LONG 'Quick Status' # # #;";
                //string fldConflict = PRZC.c_FLD_STATUSINFO_CONFLICT + " LONG 'Conflict' # # #;";

                //string flds = fldQuickStatus + fldConflict;

                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Adding fields to Status Info Table..."), true, ++val);
                //toolParams = Geoprocessing.MakeValueArray(sipath, flds);
                //toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                //toolOutput = await PRZH.RunGPTool("AddFields_management", toolParams, toolEnvs, GPExecuteToolFlags.RefreshProjectItems);
                //if (toolOutput == null)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error adding fields to Status Info Table.  GP Tool failed or was cancelled by user", LogMessageType.ERROR), true, ++val);
                //    return false;
                //}

                //// Add INCLUDE & EXCLUDE layer-based columns to Status Info table
                //if (!await AddLayerFields(PRZLayerNames.STATUS_INCLUDE, DT_IncludeLayers))
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error adding INCLUDE layer fields to Status Info Table.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Error adding the INCLUDE layer fields to the Status Info table.", "");
                //    return false;
                //}

                //if (!await AddLayerFields(PRZLayerNames.STATUS_EXCLUDE, DT_ExcludeLayers))
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error adding EXCLUDE layer fields to Status Info Table.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Error adding the EXCLUDE layer fields to the Status Info table.", "");
                //    return false;
                //}

                //#endregion

                //#region INTERSECT THE VARIOUS LAYERS

                //if (!await IntersectConstraintLayers(PRZLayerNames.STATUS_INCLUDE, DT_IncludeLayers, DICT_PUID_and_assoc_area_m2))
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error intersecting the INCLUDE layers.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Error intersecting the INCLUDE layers.", "");
                //    return false;
                //}
                //else
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Successfully intersected all INCLUDE layers (if any were present)."), true, ++val);
                //}

                //if (!await IntersectConstraintLayers(PRZLayerNames.STATUS_EXCLUDE, DT_ExcludeLayers, DICT_PUID_and_assoc_area_m2))
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error intersecting the EXCLUDE layers.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Error intersecting the EXCLUDE layers.", "");
                //    return false;
                //}
                //else
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Successfully disoolved all EXCLUDE layers (if any were present)."), true, ++val);
                //}

                //#endregion

                //#region UPDATE QUICKSTATUS AND CONFLICT FIELDS

                //Dictionary<int, int> DICT_PUID_and_QuickStatus = new Dictionary<int, int>();
                //Dictionary<int, int> DICT_PUID_and_Conflict = new Dictionary<int, int>();

                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Calculating Status Conflicts and QuickStatus"), true, ++val);

                //try
                //{
                //    await QueuedTask.Run(async () =>
                //    {
                //        using (Table table = await PRZH.GetStatusInfoTable())
                //        {
                //            TableDefinition tDef = table.GetDefinition();

                //            // Get list of INCLUDE layer Area fields
                //            var INAreaFields = tDef.GetFields().Where(f => f.Name.StartsWith("IN") && f.Name.EndsWith("_Area")).ToList();

                //            // Get list of EXCLUDE layer Area fields
                //            var EXAreaFields = tDef.GetFields().Where(f => f.Name.StartsWith("EX") && f.Name.EndsWith("_Area")).ToList();

                //            using (RowCursor rowCursor = table.Search(null, false))
                //            {
                //                while (rowCursor.MoveNext())
                //                {
                //                    using (Row row = rowCursor.Current)
                //                    {
                //                        int puid = (int)row[PRZC.c_FLD_PUFC_ID];

                //                        bool hasIN = false;
                //                        bool hasEX = false;

                //                        // Determine if there are any INCLUDE area fields having values > 0 for this PU ID
                //                        foreach (Field fld in INAreaFields)
                //                        {
                //                            double test = (double)row[fld.Name];
                //                            if (test > 0)
                //                            {
                //                                hasIN = true;
                //                            }
                //                        }

                //                        // Determine if there are any EXCLUDE area fields having values > 0 for this PU ID
                //                        foreach (Field fld in EXAreaFields)
                //                        {
                //                            double test = (double)row[fld.Name];
                //                            if (test > 0)
                //                            {
                //                                hasEX = true;
                //                            }
                //                        }

                //                        // Update the QuickStatus and Conflict information

                //                        // Both INCLUDE and EXCLUDE occur in PU:
                //                        if (hasIN && hasEX)
                //                        {
                //                            // Flag as a conflicted planning unit
                //                            row[PRZC.c_FLD_STATUSINFO_CONFLICT] = 1;
                //                            DICT_PUID_and_Conflict.Add(puid, 1);

                //                            // Set the 'winning' QuickStatus based on user-specified setting
                //                            if (SelectedOverrideOption == c_OVERRIDE_INCLUDE)
                //                            {
                //                                row[PRZC.c_FLD_STATUSINFO_QUICKSTATUS] = 2;
                //                                DICT_PUID_and_QuickStatus.Add(puid, 2);
                //                            }
                //                            else if (SelectedOverrideOption == c_OVERRIDE_EXCLUDE)
                //                            {
                //                                row[PRZC.c_FLD_STATUSINFO_QUICKSTATUS] = 3;
                //                                DICT_PUID_and_QuickStatus.Add(puid, 3);
                //                            }

                //                        }

                //                        // INCLUDE only:
                //                        else if (hasIN)
                //                        {
                //                            // Flag as a no-conflict planning unit
                //                            row[PRZC.c_FLD_STATUSINFO_CONFLICT] = 0;
                //                            DICT_PUID_and_Conflict.Add(puid, 0);

                //                            // Set the Status
                //                            row[PRZC.c_FLD_STATUSINFO_QUICKSTATUS] = 2;
                //                            DICT_PUID_and_QuickStatus.Add(puid, 2);
                //                        }

                //                        // EXCLUDE only:
                //                        else if (hasEX)
                //                        {
                //                            // Flag as a no-conflict planning unit
                //                            row[PRZC.c_FLD_STATUSINFO_CONFLICT] = 0;
                //                            DICT_PUID_and_Conflict.Add(puid, 0);

                //                            // Set the Status
                //                            row[PRZC.c_FLD_STATUSINFO_QUICKSTATUS] = 3;
                //                            DICT_PUID_and_QuickStatus.Add(puid, 3);
                //                        }

                //                        // Neither:
                //                        else
                //                        {
                //                            // Flag as a no-conflict planning unit
                //                            row[PRZC.c_FLD_STATUSINFO_CONFLICT] = 0;
                //                            DICT_PUID_and_Conflict.Add(puid, 0);

                //                            // Set the Status
                //                            row[PRZC.c_FLD_STATUSINFO_QUICKSTATUS] = 0;
                //                            DICT_PUID_and_QuickStatus.Add(puid, 0);
                //                        }

                //                        // update the row
                //                        row.Store();
                //                    }
                //                }
                //            }
                //        }

                //    });
                //}
                //catch (Exception ex)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error updating the Status Info Quickstatus and Conflict fields.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Error updating Status Info Table quickstatus and conflict fields" + Environment.NewLine + Environment.NewLine + ex.Message, "");
                //    return false;
                //}

                //#endregion

                //#region UPDATE PLANNING UNIT FC QUICKSTATUS AND CONFLICT COLUMNS

                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Updating Planning Unit FC Status Column"), true, ++val);

                //try
                //{
                //    await QueuedTask.Run(async () =>
                //    {
                //        using (Table table = await PRZH.GetPlanningUnitFC())    // Get the Planning Unit FC attribute table
                //        using (RowCursor rowCursor = table.Search(null, false))
                //        {
                //            while (rowCursor.MoveNext())
                //            {
                //                using (Row row = rowCursor.Current)
                //                {
                //                    int puid = (int)row[PRZC.c_FLD_PUFC_ID];

                //                    if (DICT_PUID_and_QuickStatus.ContainsKey(puid))
                //                    {
                //                        row[PRZC.c_FLD_PUFC_STATUS] = DICT_PUID_and_QuickStatus[puid];
                //                    }
                //                    else
                //                    {
                //                        row[PRZC.c_FLD_PUFC_STATUS] = -1;
                //                    }

                //                    if (DICT_PUID_and_Conflict.ContainsKey(puid))
                //                    {
                //                        row[PRZC.c_FLD_PUFC_CONFLICT] = DICT_PUID_and_Conflict[puid];
                //                    }
                //                    else
                //                    {
                //                        row[PRZC.c_FLD_PUFC_CONFLICT] = -1;
                //                    }

                //                    row.Store();
                //                }
                //            }
                //        }

                //    });
                //}
                //catch (Exception ex)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error updating the Status Info Quickstatus and Conflict fields.", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show("Error updating Status Info Table quickstatus and conflict fields" + Environment.NewLine + Environment.NewLine + ex.Message, "");
                //    return false;
                //}

                //#endregion

                #region WRAP THINGS UP

                // Populate the Grid
                bool Populated = await PopulateCFGrid();

                //// Compact the Geodatabase
                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Compacting the geodatabase..."), true, ++val);
                //toolParams = Geoprocessing.MakeValueArray(gdbpath);
                //toolOutput = await PRZH.RunGPTool("Compact_management", toolParams, null, GPExecuteToolFlags.None);
                //if (toolOutput == null)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error compacting the geodatabase. GP Tool failed or was cancelled by user", LogMessageType.ERROR), true, ++val);
                //    return false;
                //}

                //// Refresh the Map & TOC
                //if (!await PRZM.ValidatePRZGroupLayers())
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error validating PRZ layers...", LogMessageType.ERROR), true, ++val);
                //    return false;
                //}

                //// Wrap things up
                //stopwatch.Stop();
                //string message = PRZH.GetElapsedTimeMessage(stopwatch.Elapsed);
                //PRZH.UpdateProgress(PM, PRZH.WriteLog("Construction completed successfully!"), true, 1, 1);
                //PRZH.UpdateProgress(PM, PRZH.WriteLog(message), true, 1, 1);

                //ProMsgBox.Show("Construction Completed Sucessfully!" + Environment.NewLine + Environment.NewLine + message);

                return true;

                #endregion
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                PRZH.UpdateProgress(PM, PRZH.WriteLog(ex.Message, LogMessageType.ERROR), true, ++val);
                return false;
            }
        }



        private async Task<bool> GridDoubleClick()
        {
            try
            {
                ProMsgBox.Show("click");

                //if (SelectedConflict != null)
                //{
                //    string LayerName_IN = SelectedConflict.include_layer_name;
                //    string LayerName_EX = SelectedConflict.exclude_layer_name;

                //    int AreaFieldIndex_IN = SelectedConflict.include_area_field_index;
                //    int AreaFieldIndex_EX = SelectedConflict.exclude_area_field_index;

                //    ProMsgBox.Show("Include Index: " + AreaFieldIndex_IN.ToString() + "   Layer: " + LayerName_IN);
                //    ProMsgBox.Show("Exclude Index: " + AreaFieldIndex_EX.ToString() + "   Layer: " + LayerName_EX);

                //    // Query the Status Info table for all records (i.e. PUs) where field IN ix > 0 and field EX ix > 0
                //    // Save the PUIDs in a list

                //    List<int> PlanningUnitIDs = new List<int>();

                //    await QueuedTask.Run(async () =>
                //    {
                //        using (Table table = await PRZH.GetStatusInfoTable())
                //        using (TableDefinition tDef = table.GetDefinition())
                //        {
                //            // Get the field names
                //            var fields = tDef.GetFields();
                //            string area_field_IN = fields[AreaFieldIndex_IN].Name;
                //            string area_field_EX = fields[AreaFieldIndex_EX].Name;

                //            ProMsgBox.Show("IN Field Name: " + area_field_IN + "    EX Field Name: " + area_field_EX);

                //            QueryFilter QF = new QueryFilter();
                //            QF.SubFields = PRZC.c_FLD_PUFC_ID + "," + area_field_IN + "," + area_field_EX;
                //            QF.WhereClause = area_field_IN + @" > 0 And " + area_field_EX + @" > 0";

                //            using (RowCursor rowCursor = table.Search(QF, false))
                //            {
                //                while (rowCursor.MoveNext())
                //                {
                //                    using (Row row = rowCursor.Current)
                //                    {
                //                        int puid = (int)row[PRZC.c_FLD_PUFC_ID];
                //                        PlanningUnitIDs.Add(puid);
                //                    }
                //                }
                //            }
                //        }
                //    });

                //    // validate the number of PUIDs returned
                //    if (PlanningUnitIDs.Count == 0)
                //    {
                //        ProMsgBox.Show("No Planning Unit IDs retrieved.  That's very strange, there should be at least 1.  Ask JC about this.");
                //        return true;
                //    }

                //    // I now have a list of PUIDs.  I need to select the associated features in the PUFC
                //    Map map = MapView.Active.Map;

                //    await QueuedTask.Run(() =>
                //    {
                //        // Get the Planning Unit Feature Layer
                //        FeatureLayer featureLayer = PRZH.GetFeatureLayer_PU(map);

                //        // Clear Selection
                //        featureLayer.ClearSelection();

                //        // Build QueryFilter
                //        QueryFilter QF = new QueryFilter();
                //        string puid_list = string.Join(",", PlanningUnitIDs);
                //        QF.WhereClause = PRZC.c_FLD_PUFC_ID + " In (" + puid_list + ")";

                //        // Do the actual selection
                //        using (Selection selection = featureLayer.Select(QF, SelectionCombinationMethod.New))   // selection happens here
                //        {
                //            // do nothing?
                //        }
                //    });
                //}

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }



        #endregion

        #region Event Handlers


        #endregion

    }
}