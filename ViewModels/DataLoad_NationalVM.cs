﻿using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Raster;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ProMsgBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using PRZC = NCC.PRZTools.PRZConstants;
using PRZH = NCC.PRZTools.PRZHelper;

namespace NCC.PRZTools
{
    public class DataLoad_NationalVM : PropertyChangedBase
    {
        public DataLoad_NationalVM()
        {
        }

        #region FIELDS

        private CancellationTokenSource _cts = null;
        private ProgressManager _pm = ProgressManager.CreateProgressManager(50);
        private bool _operation_Cmd_IsEnabled;
        private bool _operationIsUnderway = false;
        private Cursor _proWindowCursor;

        private bool _pu_exists = false;
        private bool _pu_isnat = false;
        private bool _natdb_exists = false;

        private Map _map;

        #region COMMANDS

        private ICommand _cmdLoadNationalData;
        private ICommand _cmdCancel;
        private ICommand _cmdClearLog;

        #endregion

        #region COMPONENT STATUS INDICATORS

        // Planning Unit Dataset
        private string _compStat_Img_PlanningUnits_Path;
        private string _compStat_Txt_PlanningUnits_Label;

        // National DB
        private string _compStat_Img_NatDB_Path;
        private string _compStat_Txt_NatDB_Label;

        #endregion

        #region OPERATION STATUS INDICATORS

        private Visibility _opStat_Img_Visibility = Visibility.Collapsed;
        private string _opStat_Txt_Label;

        #endregion

        #endregion

        #region PROPERTIES

        public ProgressManager PM
        {
            get => _pm; set => SetProperty(ref _pm, value, () => PM);
        }

        public bool Operation_Cmd_IsEnabled
        {
            get => _operation_Cmd_IsEnabled;
            set => SetProperty(ref _operation_Cmd_IsEnabled, value, () => Operation_Cmd_IsEnabled);
        }

        public bool OperationIsUnderway
        {
            get => _operationIsUnderway;
        }

        public Cursor ProWindowCursor
        {
            get => _proWindowCursor;
            set => SetProperty(ref _proWindowCursor, value, () => ProWindowCursor);
        }

        #region COMPONENT STATUS INDICATORS

        // Planning Units Dataset
        public string CompStat_Img_PlanningUnits_Path
        {
            get => _compStat_Img_PlanningUnits_Path;
            set => SetProperty(ref _compStat_Img_PlanningUnits_Path, value, () => CompStat_Img_PlanningUnits_Path);
        }

        public string CompStat_Txt_PlanningUnits_Label
        {
            get => _compStat_Txt_PlanningUnits_Label;
            set => SetProperty(ref _compStat_Txt_PlanningUnits_Label, value, () => CompStat_Txt_PlanningUnits_Label);
        }

        // National Database
        public string CompStat_Img_NatDB_Path
        {
            get => _compStat_Img_NatDB_Path;
            set => SetProperty(ref _compStat_Img_NatDB_Path, value, () => CompStat_Img_NatDB_Path);
        }

        public string CompStat_Txt_NatDB_Label
        {
            get => _compStat_Txt_NatDB_Label;
            set => SetProperty(ref _compStat_Txt_NatDB_Label, value, () => CompStat_Txt_NatDB_Label);
        }

        #endregion

        #region OPERATION STATUS INDICATORS

        public Visibility OpStat_Img_Visibility
        {
            get => _opStat_Img_Visibility;
            set => SetProperty(ref _opStat_Img_Visibility, value, () => OpStat_Img_Visibility);
        }

        public string OpStat_Txt_Label
        {
            get => _opStat_Txt_Label;
            set => SetProperty(ref _opStat_Txt_Label, value, () => OpStat_Txt_Label);
        }

        #endregion

        #region COMMANDS

        public ICommand CmdLoadNationalData => _cmdLoadNationalData ?? (_cmdLoadNationalData = new RelayCommand(async () =>
        {
            // Change UI to Underway
            StartOpUI();

            // Start the operation
            using (_cts = new CancellationTokenSource())
            {
                await LoadNationalData(_cts.Token);
            }

            // Set source to null (it's already disposed)
            _cts = null;

            // Validate controls
            await ValidateControls();

            // Reset UI to Idle
            ResetOpUI();

        }, () => true, true, false));

        public ICommand CmdCancel => _cmdCancel ?? (_cmdCancel = new RelayCommand(() =>
        {
            if (_cts != null)
            {
                // Optionally notify the user or prompt the user here

                // Cancel the operation
                _cts.Cancel();
            }
        }, () => _cts != null, true, false));

        public ICommand CmdClearLog => _cmdClearLog ?? (_cmdClearLog = new RelayCommand(() =>
        {
            PRZH.UpdateProgress(PM, "", false, 0, 1, 0);
        }, () => true, true, false));

        #endregion

        #endregion

        #region METHODS

        public async Task OnProWinLoaded()
        {
            try
            {
                // get reference to the active map
                _map = MapView.Active.Map;

                // Initialize the Progress Bar & Log
                PRZH.UpdateProgress(PM, "", false, 0, 1, 0);

                // Configure a few controls
                await ValidateControls();

                // Reset the UI
                ResetOpUI();
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
        }

        private async Task LoadNationalData(CancellationToken token)
        {
            bool edits_are_disabled = !Project.Current.IsEditingEnabled;
            int val = 0;
            int max = 50;

            try
            {
                #region INITIALIZATION

                #region EDITING CHECK

                // Check for currently unsaved edits in the project
                if (Project.Current.HasEdits)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("ArcGIS Pro Project has unsaved edits.  Please save all edits before proceeding.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show("This ArcGIS Pro Project has some unsaved edits.  Please save all edits before proceeding.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("ArcGIS Pro Project has no unsaved edits.  Proceeding..."), true, ++val);
                }

                // If editing is disabled, enable it temporarily (and disable again in the finally block)
                if (edits_are_disabled)
                {
                    if (!await Project.Current.SetIsEditingEnabledAsync(true))
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Editing on this Project has been disabled.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Editing on this Project has been disabled.");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("ArcGIS Pro editing enabled."), true, ++val);
                    }
                }

                #endregion

                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                GPExecuteToolFlags toolFlags_GPRefresh = GPExecuteToolFlags.GPThread | GPExecuteToolFlags.RefreshProjectItems;
                string toolOutput;

                // Initialize ProgressBar and Progress Log
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Initializing the National Data Loader..."), false, max, ++val);

                // Ensure the Project Geodatabase Exists
                string gdbpath = PRZH.GetPath_ProjectGDB();
                var tryex_gdb = await PRZH.GDBExists_Project();

                if (!tryex_gdb.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Project Geodatabase not found: '{gdbpath}'.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Project Geodatabase not found at {gdbpath}.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Project Geodatabase found at '{gdbpath}'."), true, ++val);
                }

                // Planning Unit existence
                var tryex_pudata = await PRZH.PUDataExists();
                if (!tryex_pudata.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Planning Units dataset not found.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Planning Units dataset not found.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Planning Units dataset exists."), true, ++val);
                }

                // Ensure that the Planning Unit Data is national-enabled
                if (!tryex_pudata.national_enabled)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Planning Units data is not configured for national data.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Planning Units data is not configured for national data.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Planning Units data is configured for national data."), true, ++val);
                }

                // Ensure the National db exists
                string natpath = PRZH.GetPath_NatGDB();
                var tryexists_nat = await PRZH.GDBExists_Nat();
                if (!tryexists_nat.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Valid National Geodatabase not found: '{natpath}'.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Valid National Geodatabase not found at {natpath}.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"National Geodatabase is OK: '{natpath}'."), true, ++val);
                }

                // Capture the Planning Unit Spatial Reference
                SpatialReference PlanningUnitSR = await QueuedTask.Run(() =>
                {
                    var tryget_ras = PRZH.GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);
                    using (RasterDataset rasterDataset = tryget_ras.rasterDataset)
                    using (RasterDatasetDefinition rasterDef = rasterDataset.GetDefinition())
                    {
                        return rasterDef.GetSpatialReference();
                    }
                });

                if (ProMsgBox.Show($"If you proceed, any existing National Theme and Element tables in the Project Geodatabase will be overwritten." + 
                    Environment.NewLine + Environment.NewLine +
                    $"Do you wish to proceed?",
                    "Overwrite Geodatabase?",
                    System.Windows.MessageBoxButton.OKCancel, System.Windows.MessageBoxImage.Warning,
                    System.Windows.MessageBoxResult.Cancel) == System.Windows.MessageBoxResult.Cancel)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("User cancelled operation."), true, ++val);
                    return;
                }

                // Start a stopwatch
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                #endregion

                #region DELETE EXISTING GEODATABASE OBJECTS

                // Delete the National Theme Table if present
                if ((await PRZH.TableExists_Project(PRZC.c_TABLE_NATPRJ_THEMES)).exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting the {PRZC.c_TABLE_NATPRJ_THEMES} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_NATPRJ_THEMES);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting the {PRZC.c_TABLE_NATPRJ_THEMES} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error deleting the {PRZC.c_TABLE_NATPRJ_THEMES} table.");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Table deleted successfully."), true, ++val);
                    }
                }

                PRZH.CheckForCancellation(token);

                // Delete the National Element table if present
                if ((await PRZH.TableExists_Project(PRZC.c_TABLE_NATPRJ_ELEMENTS)).exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting the {PRZC.c_TABLE_NATPRJ_ELEMENTS} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_NATPRJ_ELEMENTS);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting the {PRZC.c_TABLE_NATPRJ_ELEMENTS} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error deleting the {PRZC.c_TABLE_NATPRJ_ELEMENTS} table.");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Table deleted successfully."), true, ++val);
                    }
                }

                PRZH.CheckForCancellation(token);

                // Delete any national element tables
                var tryget_tables = await PRZH.GetNationalElementTables();

                if (!tryget_tables.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving list of national element tables.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving list of national element tables.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"{tryget_tables.tables.Count} national element tables found."), true, ++val);
                }

                if (tryget_tables.tables.Count > 0)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting {tryget_tables.tables.Count} national element tables..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(string.Join(";", tryget_tables.tables));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting the national element tables ({string.Join(";", tryget_tables.tables)}.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error deleting the national element tables.");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"National element tables deleted successfully."), true, ++val);
                    }
                }

                PRZH.CheckForCancellation(token);

                // Delete and rebuild National FDS
                // delete...
                var tryex_natfds = await PRZH.FDSExists_Project(PRZC.c_FDS_NATIONAL_ELEMENTS);
                if (tryex_natfds.exists)
                {
                    // delete it
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting {PRZC.c_FDS_NATIONAL_ELEMENTS} FDS..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_FDS_NATIONAL_ELEMENTS);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting {PRZC.c_FDS_NATIONAL_ELEMENTS} FDS.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error deleting {PRZC.c_FDS_NATIONAL_ELEMENTS} FDS.");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"FDS deleted successfully."), true, ++val);
                    }
                }

                PRZH.CheckForCancellation(token);

                // (re)build!
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Creating {PRZC.c_FDS_NATIONAL_ELEMENTS} feature dataset..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_FDS_NATIONAL_ELEMENTS, PlanningUnitSR);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(
                    workspace: gdbpath,
                    overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("CreateFeatureDataset_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error creating feature dataset.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error creating feature dataset.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Feature dataset created."), true, ++val);
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region NATIONAL TABLES AND NATIONAL FEATURE CLASSES

                // Process the national tables
                var trynatdb = await ProcessNationalDbTables(token);

                if (!trynatdb.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error processing National DB tables.\n{trynatdb.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error processing National DB tables.\n{trynatdb.message}.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"National data loaded successfully."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // TODO: Visualize national database data without Feature classes
                //// Generate the National Element spatial datasets
                //var tryspat = await GenerateSpatialDatasets(token);
                //if (!tryspat.success)
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error generating national spatial datasets.\n{tryspat.message}", LogMessageType.ERROR), true, ++val);
                //    ProMsgBox.Show($"Error generating national spatial datasets.\n{tryspat.message}.");
                //    return;
                //}
                //else
                //{
                //    PRZH.UpdateProgress(PM, PRZH.WriteLog($"National spatial datasets generated successfully."), true, ++val);
                //}

                #endregion

                PRZH.CheckForCancellation(token);

                #region WRAP UP

                // Compact the Geodatabase
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Compacting the Geodatabase..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(gdbpath);
                toolOutput = await PRZH.RunGPTool("Compact_management", toolParams, null, toolFlags_GPRefresh);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error compacting the geodatabase. GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show("Error compacting the geodatabase.");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Geodatabase compacted."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Refresh the Map & TOC
                if (!(await PRZH.RedrawPRZLayers(_map)).success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error redrawing the PRZ layers.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error redrawing the PRZ layers.");
                    return;
                }

                // Final message
                stopwatch.Stop();
                string message = PRZH.GetElapsedTimeMessage(stopwatch.Elapsed);
                PRZH.UpdateProgress(PM, PRZH.WriteLog("National data load completed successfully!"), true, 1, 1);
                PRZH.UpdateProgress(PM, PRZH.WriteLog(message), true, 1, 1);

                ProMsgBox.Show("National data load completed successfully!" + Environment.NewLine + Environment.NewLine + message);

                #endregion
            }
            catch (OperationCanceledException)
            {
                // Cancelled by user
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Operation cancelled by user.", LogMessageType.CANCELLATION), true, ++val);
                ProMsgBox.Show($"Operation cancelled by user.");
            }
            catch (Exception ex)
            {
                PRZH.UpdateProgress(PM, PRZH.WriteLog(ex.Message, LogMessageType.CANCELLATION), true, ++val);
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                // reset disabled editing status
                if (edits_are_disabled)
                {
                    await Project.Current.SetIsEditingEnabledAsync(false);
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("ArcGIS Pro editing disabled."), true, max, ++val);
                }
            }
        }

        private async Task<(bool success, string message)> ProcessNationalDbTables(CancellationToken token)
        {
            int val = PM.Current;
            int max = PM.Max;

            try
            {
                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                string toolOutput;

                #region RETRIEVE AND PREPARE INFO FROM NATIONAL DATABASE

                // COPY THE ELEMENT TABLE
                string gdbpath = PRZH.GetPath_ProjectGDB();
                string natdbpath = PRZH.GetPath_NatGDB();

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Copying the {PRZC.c_TABLE_NATSRC_ELEMENTS} table..."), true, ++val);
                var q_elem = await PRZH.GetNatDBQualifiedName(PRZC.c_TABLE_NATSRC_ELEMENTS);
                string inputelempath = Path.Combine(natdbpath, q_elem.qualified_name);
                toolParams = Geoprocessing.MakeValueArray(inputelempath, PRZC.c_TABLE_NATPRJ_ELEMENTS, "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("Copy_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying the {PRZC.c_TABLE_NATSRC_ELEMENTS} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error copying the {PRZC.c_TABLE_NATSRC_ELEMENTS} table.");
                    return (false, "table copy error.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Table copied successfully."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // ALTER ALIAS NAME OF ELEMENT TABLE
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Altering table alias..."), true, ++val);
                await QueuedTask.Run(() =>
                {
                    var tryget_projectgdb = PRZH.GetGDB_Project();

                    if (!tryget_projectgdb.success)
                    {
                        throw new Exception("Error opening project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_projectgdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_NATPRJ_ELEMENTS))
                    using (TableDefinition tblDef = table.GetDefinition())
                    {
                        // Get the Table Description
                        TableDescription tblDescr = new TableDescription(tblDef);
                        tblDescr.AliasName = "National Elements";

                        // get the schemabuilder
                        SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);
                        schemaBuilder.Modify(tblDescr);
                        var success = schemaBuilder.Build();
                    }
                });

                // INSERT EXTRA FIELDS INTO ELEMENT TABLE
                string fldElemPresence = PRZC.c_FLD_TAB_NATELEMENT_PRESENCE + $" SHORT 'Presence' # {(int)ElementPresence.Absent} '" + PRZC.c_DOMAIN_PRESENCE + "';";
                string flds = fldElemPresence;

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Adding fields to the copied {PRZC.c_TABLE_NATPRJ_ELEMENTS} table..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_NATPRJ_ELEMENTS, flds);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("AddFields_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error adding fields to the copied {PRZC.c_TABLE_NATPRJ_ELEMENTS} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error adding fields to the copied {PRZC.c_TABLE_NATPRJ_ELEMENTS} table.");
                    return (false, "field addition error.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Fields added successfully."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // COPY THE THEMES TABLE
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Copying the {PRZC.c_TABLE_NATSRC_THEMES} table..."), true, ++val);
                var q_theme = await PRZH.GetNatDBQualifiedName(PRZC.c_TABLE_NATSRC_THEMES);
                string inputthemepath = Path.Combine(natdbpath, q_theme.qualified_name);
                toolParams = Geoprocessing.MakeValueArray(inputthemepath, PRZC.c_TABLE_NATPRJ_THEMES, "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("Copy_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying the {PRZC.c_TABLE_NATSRC_THEMES} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error copying the {PRZC.c_TABLE_NATSRC_THEMES} table.");
                    return (false, "table copy error.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Table copied successfully."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // ALTER ALIAS NAME OF THEME TABLE
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Altering table alias..."), true, ++val);
                await QueuedTask.Run(() =>
                {
                    var tryget_projectgdb = PRZH.GetGDB_Project();

                    if (!tryget_projectgdb.success)
                    {
                        throw new Exception("Error opening project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_projectgdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_NATPRJ_THEMES))
                    using (TableDefinition tblDef = table.GetDefinition())
                    {
                        // Get the Table Description
                        TableDescription tblDescr = new TableDescription(tblDef);
                        tblDescr.AliasName = "National Themes";

                        // get the schemabuilder
                        SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);
                        schemaBuilder.Modify(tblDescr);
                        var success = schemaBuilder.Build();
                    }
                });

                // INSERT EXTRA FIELDS INTO THEME TABLE
                string fldThemePresence = PRZC.c_FLD_TAB_NATTHEME_PRESENCE + $" SHORT 'Presence' # {(int)ElementPresence.Absent} '" + PRZC.c_DOMAIN_PRESENCE + "';";
                flds = fldThemePresence;

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Adding fields to the copied {PRZC.c_TABLE_NATPRJ_THEMES} table..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_NATPRJ_THEMES, flds);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("AddFields_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error adding fields to the copied {PRZC.c_TABLE_NATPRJ_THEMES} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error adding fields to the copied {PRZC.c_TABLE_NATPRJ_THEMES} table.");
                    return (false, "field addition error.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Fields added successfully."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Get the National Themes from the copied national themes table
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieving national themes..."), true, ++val);
                var theme_outcome = await PRZH.GetNationalThemes();
                if (!theme_outcome.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national themes.\n{theme_outcome.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving national themes.\n{theme_outcome.message}");
                    return (false, "error retrieving themes.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {theme_outcome.themes.Count} national themes."), true, ++val);
                }
                List<NatTheme> themes = theme_outcome.themes;

                PRZH.CheckForCancellation(token);

                // Get the National Elements from the copied national elements table
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieving national elements..."), true, ++val);
                var elem_outcome = await PRZH.GetNationalElements();
                if (!elem_outcome.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national elements.\n{elem_outcome.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving national elements.\n{elem_outcome.message}");
                    return (false, "error retrieving elements.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {elem_outcome.elements.Count} national elements."), true, ++val);
                }
                List<NatElement> elements = elem_outcome.elements;

                var try_setup_table_format = await PRZH.SetElementTableNamingFormat((elements ?? new List<NatElement>()).Count > 0 ? elements[0].ElementID : -1);

                if(!try_setup_table_format.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Unable to detect the format of element presence table names." + Environment.NewLine + $"{try_setup_table_format.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show("Error while preparing to process element tables." + Environment.NewLine + $"{try_setup_table_format.message}");
                    return (false, "Error while preparing to process element tables.");
                }

                PRZH.CheckForCancellation(token);

                #endregion

                PRZH.CheckForCancellation(token);

                #region RETRIEVE INTERSECTING ELEMENTS

                // Construct dictionary of planning units / national grid ids
                var outcome = await PRZH.GetCellNumbersAndPUIDs();
                if (!outcome.success)
                {
                    throw new Exception("Error constructing PUID dictionary, try rebuilding planning units.");
                }
                Dictionary<long, int> study_area_cells = outcome.dict;

                // Load tile metadata for study area and national database
                var tryread_studyarea_tiles = PRZH.ReadBinary(PRZH.GetPath_ProjectTilesMetadataPath());
                var tryread_natdata_tiles = PRZH.ReadBinary(PRZH.GetPath_NationalDatabaseElementsTileMetadataPath());

                if (!tryread_studyarea_tiles.success)
                {
                    throw new Exception(tryread_studyarea_tiles.message);
                }
                HashSet<int> study_area_tiles = (HashSet<int>)tryread_studyarea_tiles.obj;

                if (!tryread_natdata_tiles.success)
                {
                    throw new Exception(tryread_natdata_tiles.message);
                }
                Dictionary<int, HashSet<int>> natdata_tiles = (Dictionary<int, HashSet<int>>)tryread_natdata_tiles.obj;

                // Setup up lists to track which elements and themes are present
                HashSet<int> elements_present = new HashSet<int>();
                HashSet<int> themes_present = new HashSet<int>();

                // Refresh project-scope elements folder
                string output_elements_folder = PRZH.GetPath_ProjectNationalElementsSubfolder();
                if(Directory.Exists(output_elements_folder))
                {
                    Directory.Delete(output_elements_folder, true);
                }
                Directory.CreateDirectory(output_elements_folder);

                // Find intersecting cells
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Finding intersecting cells in National Database..."), true, ++val);
                string input_elements_folder = PRZH.GetPath_NationalDatabaseElementsSubfolder();
                int progress = 0;
                var options = new ParallelOptions() { MaxDegreeOfParallelism = 3 };


                Parallel.ForEach(elements, options, element =>
                {
                    progress++;
                    if (progress % 100 == 0)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Processing national element {progress}/{elements.Count()}."), true, ++val);
                    }

                    // Determine which if any tiles overlap
                    natdata_tiles[element.ElementID].IntersectWith(study_area_tiles);

                    // Skip elements with no overlapping tiles
                    if (natdata_tiles[element.ElementID].Count == 0)
                        return;

                    // Read in relevant element data tiles from national database
                    Dictionary<long, double> element_cells = new Dictionary<long, double>();
                    foreach (int tile_id in natdata_tiles[element.ElementID])
                    {
                        string tile_filepath = Path.Combine(input_elements_folder, $"{element.ElementTable}-{tile_id}.bin");
                        var tryread_tile = PRZH.ReadBinary(tile_filepath);

                        if (!tryread_tile.success)
                        {
                            throw new Exception(tryread_tile.message);
                        }
                        Dictionary<long, double> tile_cells = (Dictionary<long, double>)tryread_tile.obj;

                        if (tile_cells.Count > 0)
                        {
                            element_cells = element_cells.Concat(tile_cells).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                        }
                    }

                    // Find cells that intersect study area
                    Dictionary<long, double> intersecting_cells = element_cells.Where(x => study_area_cells.ContainsKey(x.Key)).ToDictionary(x => x.Key, x => x.Value);

                    if (intersecting_cells.Count() > 0)
                    {
                        var trywrite = PRZH.WriteBinary(intersecting_cells, $"{Path.Combine(output_elements_folder, element.ElementTable)}.bin");

                        if (!trywrite.success)
                        {
                            throw new Exception(trywrite.message);
                        }
                    }
                });

                // Identify elements that are present by id
                foreach(string f in Directory.GetFiles(output_elements_folder))
                {
                    int element_id = Convert.ToInt32(Path.GetFileNameWithoutExtension(f).Substring(1));
                    elements_present.Add(element_id);
                }

                // Update the table
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Updating the {PRZC.c_TABLE_NATPRJ_ELEMENTS} table {PRZC.c_FLD_TAB_NATELEMENT_PRESENCE} field..."), true, ++val);
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = PRZH.GetGDB_Project();

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_NATPRJ_ELEMENTS))
                    using (RowCursor rowCursor = table.Search(null, false))
                    {
                        geodatabase.ApplyEdits(() =>
                        {
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    if (elements_present.Contains((int)row[PRZC.c_FLD_TAB_NATELEMENT_ELEMENT_ID]))
                                    {
                                        row[PRZC.c_FLD_TAB_NATELEMENT_PRESENCE] = (int)ElementPresence.Present;
                                        int theme_id = (int)row[PRZC.c_FLD_TAB_NATELEMENT_THEME_ID];
                                        if(!themes_present.Contains(theme_id))
                                        {
                                            themes_present.Add(theme_id);
                                        }
                                    }
                                    else
                                    {
                                        row[PRZC.c_FLD_TAB_NATELEMENT_PRESENCE] = (int)ElementPresence.Absent;
                                    }

                                    row.Store();
                                }
                            }
                        });
                    }
                });

                #endregion

                PRZH.CheckForCancellation(token);

                #region UPDATE THE LOCAL THEME TABLE PRESENCE FIELD

                // Update the table
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Updating the {PRZC.c_TABLE_NATPRJ_THEMES} table {PRZC.c_FLD_TAB_NATTHEME_PRESENCE} field..."), true, ++val);
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = PRZH.GetGDB_Project();

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_NATPRJ_THEMES))
                    using (RowCursor rowCursor = table.Search(null, false))
                    {
                        geodatabase.ApplyEdits(() =>
                        {
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    int theme_id = (int)row[PRZC.c_FLD_TAB_NATTHEME_THEME_ID];

                                    if (themes_present.Contains(theme_id))
                                    {
                                        row[PRZC.c_FLD_TAB_NATTHEME_PRESENCE] = (int)ElementPresence.Present;
                                    }
                                    else
                                    {
                                        row[PRZC.c_FLD_TAB_NATTHEME_PRESENCE] = (int)ElementPresence.Absent;
                                    }

                                    row.Store();
                                }
                            }
                        });
                    }
                });

                #endregion

                PRZH.CheckForCancellation(token);

                #region UPDATE REGIONAL THEME DOMAIN

/*                Dictionary<int, string> national_values = new Dictionary<int, string>();
                foreach (NatTheme theme in themes)
                {
                    if (theme.ThemeID <= 1000)
                    {
                        national_values.Add(theme.ThemeID, theme.ThemeName);
                    }
                }

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Updating regional themes domain..."), true, ++val);
                var tryupdate = await PRZH.UpdateRegionalThemesDomain(national_values);
                if (!tryupdate.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error updating regional themes domain", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error updating regional themes domain.");
                    return (false, "error updating domain.");
                }*/

                #endregion

                // we're done here
                return (true, "success");
            }
            catch (OperationCanceledException cancelex)
            {
                throw cancelex;
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private async Task<(bool success, string message)> GenerateSpatialDatasets(CancellationToken token)
        {
            int val = PM.Current;
            int max = PM.Max;

            try
            {
                #region GET NATIONAL ELEMENT INFOS

                // Get list of national element tables (e.g. n00010)
                var tryget_LIST_elemtables_nat = await PRZH.GetNationalElementTables();
                if (!tryget_LIST_elemtables_nat.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving list of national element tables.\n{tryget_LIST_elemtables_nat.message}", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving list of national element tables.\n{tryget_LIST_elemtables_nat.message}");
                    return (false, "error retrieving nat table list.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Found {tryget_LIST_elemtables_nat.tables.Count} national element tables"), true, ++val);
                }

                List<string> LIST_NatElemTables = tryget_LIST_elemtables_nat.tables;

                // If no national element tables are found, return.
                if (LIST_NatElemTables.Count == 0)
                {
                    // there are no national tables, so there is no spatial data to process
                    return (true, "no national tables to process (this is OK).");
                }

                PRZH.CheckForCancellation(token);

                // ASSEMBLE LISTS OF NATIONAL ELEMENTS
                // Get All Nat Elements where presence = yes
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieving all present national elements..."), true, ++val);
                var tryget_all = await PRZH.GetNationalElements(null, null, ElementPresence.Present);
                if (!tryget_all.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national elements.\n{tryget_all.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving national elements.\n{tryget_all.message}");
                    return (false, "error retrieving elements.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"{tryget_all.elements.Count} national element(s) retrieved."), true, ++val);
                }

                List<NatElement> LIST_NatElements = tryget_all.elements;

                // Ensure at least one element in list
                if (LIST_NatElements.Count == 0)
                {
                    return (true, "no national elements to process (this is OK).");
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region PREPARE THE BASE FEATURE CLASS

                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                string toolOutput;

                // Get project gdb path
                string gdbpath = PRZH.GetPath_ProjectGDB();

                // Get the fds fc path
                string base_fc = "pu_fc";
                string base_fc_path = PRZH.GetPath_Project(base_fc, PRZC.c_FDS_NATIONAL_ELEMENTS).path;

                // Get the Planning Unit SR
                SpatialReference PlanningUnitSR = await QueuedTask.Run(() =>
                {
                    var tryget_ras = PRZH.GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);
                    using (RasterDataset rasterDataset = tryget_ras.rasterDataset)
                    using (RasterDatasetDefinition rasterDef = rasterDataset.GetDefinition())
                    {
                        return rasterDef.GetSpatialReference();
                    }
                });

                // Copy the Planning Units FC into nat fds
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Copying the {PRZC.c_FC_PLANNING_UNITS} fc..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_PLANNING_UNITS, base_fc_path);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(
                    workspace: gdbpath,
                    overwriteoutput: true,
                    outputCoordinateSystem: PlanningUnitSR);
                toolOutput = await PRZH.RunGPTool("CopyFeatures_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying {PRZC.c_FC_PLANNING_UNITS}.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error copying {PRZC.c_FC_PLANNING_UNITS}");
                    return (false, "fc copy error.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("feature class copied successfully."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Delete all but id field
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting unnecessary fields..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(base_fc, PRZC.c_FLD_FC_PU_ID, "KEEP_FIELDS");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                toolOutput = await PRZH.RunGPTool("DeleteField_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting fields.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show("Error deleting fields.");
                    return (false, "field deletion error.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("fields deleted."), true, ++val);
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region CYCLE THROUGH ALL NATIONAL ELEMENTS

                // LOOP THROUGH ELEMENTS
                for (int i = 0; i < LIST_NatElements.Count; i++)
                {
                    // Get the element
                    NatElement element = LIST_NatElements[i];

                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Processing Element #{element.ElementID}: {element.ElementName}"), true, ++val);

                    var tryget_elemtablename = PRZH.GetNationalElementTableName(element.ElementID);
                    if (!tryget_elemtablename.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving table name for nat element {element.ElementID}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Error retrieving nat element table name.");
                        return (false, "error retrieving nat element table name.");
                    }
                    string table_name = tryget_elemtablename.table_name;

                    PRZH.CheckForCancellation(token);

                    // Ensure element table exists
                    var tryex_elemtable = await PRZH.TableExists_Project(table_name);
                    if (!tryex_elemtable.exists)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"element table {table_name} not found.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"element table {table_name} not found.");
                        return (false, $"element table {table_name} not found.");
                    }

                    // Copy base fc
                    string elem_fc_name = $"fc_{table_name}";
                    string elem_fc_path = PRZH.GetPath_Project(elem_fc_name, PRZC.c_FDS_NATIONAL_ELEMENTS).path;

                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Creating the {elem_fc_name} feature class..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(base_fc_path, elem_fc_path);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(
                        workspace: gdbpath,
                        overwriteoutput: true,
                        outputCoordinateSystem: PlanningUnitSR);
                    toolOutput = await PRZH.RunGPTool("CopyFeatures_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error creating {elem_fc_name} feature class.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error creating {elem_fc_name} feature class");
                        return (false, "fc creation error.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("feature class created successfully."), true, ++val);
                    }

                    PRZH.CheckForCancellation(token);

                    // JOIN FIELDS
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Joining fields..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(elem_fc_name, PRZC.c_FLD_FC_PU_ID, table_name, PRZC.c_FLD_TAB_NAT_ELEMVAL_PU_ID);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(
                        workspace: gdbpath,
                        overwriteoutput: true);
                    toolOutput = await PRZH.RunGPTool("JoinField_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error joining table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error joining table.");
                        return (false, "table join error.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("table joined successfully."), true, ++val);
                    }

                    PRZH.CheckForCancellation(token);

                    // DELETE ROWS WHERE ID_1 IS NULL
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting unjoined rows..."), true, ++val);
                    await QueuedTask.Run(() =>
                    {
                        var tryget_gdb = PRZH.GetGDB_Project();

                        using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                        using (Table table = geodatabase.OpenDataset<Table>(elem_fc_name))
                        {
                            geodatabase.ApplyEdits(() =>
                            {
                                table.DeleteRows(new QueryFilter { WhereClause = $"{PRZC.c_FLD_TAB_NAT_ELEMVAL_PU_ID}_1 IS NULL" });
                            });
                        }
                    });
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting unjoined rows..."), true, ++val);

                    PRZH.CheckForCancellation(token);

                    // DELETE UNNECESSARY FIELD
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting extra id field..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(elem_fc_name, $"{PRZC.c_FLD_TAB_NAT_ELEMVAL_PU_ID}_1", "DELETE_FIELDS");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("DeleteField_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting field.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Error deleting field.");
                        return (false, "field deletion error.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("field deleted."), true, ++val);
                    }

                    PRZH.CheckForCancellation(token);

                    // index the puid field
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Indexing {PRZC.c_FLD_FC_PU_ID} field in {elem_fc_name} feature class..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(elem_fc_name, PRZC.c_FLD_FC_PU_ID, "ix" + PRZC.c_FLD_FC_PU_ID, "", "");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(
                        workspace: gdbpath,
                        overwriteoutput: true);
                    toolOutput = await PRZH.RunGPTool("AddIndex_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Error indexing field.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Error indexing field.");
                        return (false, "error indexing field.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Field indexed successfully."), true, ++val);
                    }

                    PRZH.CheckForCancellation(token);

                    // index the cell number field
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Indexing {PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER} field in {elem_fc_name} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(elem_fc_name, PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER, "ix" + PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER, "", "");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(
                        workspace: gdbpath,
                        overwriteoutput: true);
                    toolOutput = await PRZH.RunGPTool("AddIndex_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Error indexing field.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Error indexing field.");
                        return (false, "error indexing field.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Field indexed successfully."), true, ++val);
                    }

                    PRZH.CheckForCancellation(token);

                    // ALTER ALIAS NAME OF FEATURE CLASS
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Altering feature class alias..."), true, ++val);
                    await QueuedTask.Run(() =>
                    {
                        var tryget_projectgdb = PRZH.GetGDB_Project();

                        using (Geodatabase geodatabase = tryget_projectgdb.geodatabase)
                        using (Table table = geodatabase.OpenDataset<Table>(elem_fc_name))
                        using (TableDefinition tblDef = table.GetDefinition())
                        {
                            // Get the Table Description
                            TableDescription tblDescr = new TableDescription(tblDef);
                            tblDescr.AliasName = $"{table_name}: {element.ElementName}";

                            // get the schemabuilder
                            SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);
                            schemaBuilder.Modify(tblDescr);
                            var success = schemaBuilder.Build();
                        }
                    });
                }

                #endregion

                PRZH.CheckForCancellation(token);

                // DELETE THE TEMP FC
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting the {base_fc} feature class..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(base_fc);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting the {base_fc} feature class.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error deleting the {base_fc} feature class.");
                    return (false, $"Error deleting the {base_fc} feature class.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"feature class deleted successfully."), true, ++val);
                }

                // we're done here
                return (true, "success");
            }
            catch (OperationCanceledException cancelex)
            {
                throw cancelex;
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private async Task ValidateControls()
        {
            try
            {
                // Planning Unit existence & national-enabled-ness
                var tryex_pudata = await PRZH.PUDataExists();
                _pu_exists = tryex_pudata.exists;
                _pu_isnat = tryex_pudata.national_enabled;

                if (_pu_exists & _pu_isnat)
                {
                    CompStat_Txt_PlanningUnits_Label = "Planning Units exist and are configured for National data.";
                    CompStat_Img_PlanningUnits_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_Yes16.png";
                }
                else if (_pu_exists)
                {
                    CompStat_Txt_PlanningUnits_Label = "Planning Units exist but are not configured for National data.";
                    CompStat_Img_PlanningUnits_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_No16.png";
                }
                else
                {
                    CompStat_Txt_PlanningUnits_Label = "Planning Units do not exist. Build them.";
                    CompStat_Img_PlanningUnits_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_No16.png";
                }

                // National Database existence
                _natdb_exists = (await PRZH.GDBExists_Nat()).exists;

                if (_natdb_exists)
                {
                    CompStat_Txt_NatDB_Label = "National Database exists.";
                    CompStat_Img_NatDB_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_Yes16.png";
                }
                else
                {
                    CompStat_Txt_NatDB_Label = "National Database does not exist or is invalid.";
                    CompStat_Img_NatDB_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_No16.png";
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
        }

        private void StartOpUI()
        {
            _operationIsUnderway = true;
            Operation_Cmd_IsEnabled = false;
            OpStat_Img_Visibility = Visibility.Visible;
            OpStat_Txt_Label = "Processing...";
            ProWindowCursor = Cursors.Wait;
        }

        private void ResetOpUI()
        {
            ProWindowCursor = Cursors.Arrow;
            Operation_Cmd_IsEnabled = _pu_exists & _pu_isnat & _natdb_exists;
            OpStat_Img_Visibility = Visibility.Hidden;
            OpStat_Txt_Label = "Idle";
            _operationIsUnderway = false;
        }


        #endregion


    }
}
