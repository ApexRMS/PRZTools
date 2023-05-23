using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
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
using System.Collections;
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
    public class DataLoad_RegionalVM : PropertyChangedBase
    {
        public DataLoad_RegionalVM()
        {
        }

        #region FIELDS

        private CancellationTokenSource _cts = null;
        private ProgressManager _pm = ProgressManager.CreateProgressManager(50);
        private bool _operation_Cmd_IsEnabled;
        private bool _operationIsUnderway = false;
        private Cursor _proWindowCursor;

        private bool _pu_exists = false;
        private bool _regdb_exists = false;

        private Map _map;

        #region COMMANDS

        private ICommand _cmdLoadRegionalData;
        private ICommand _cmdCancel;
        private ICommand _cmdClearLog;

        #endregion

        #region COMPONENT STATUS INDICATORS

        // Planning Unit Dataset
        private string _compStat_Img_PlanningUnits_Path;
        private string _compStat_Txt_PlanningUnits_Label;

        // Regional Data Folder
        private string _compStat_Img_RegDB_Path;
        private string _compStat_Txt_RegDB_Label;

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

        // Regional Database
        public string CompStat_Img_RegDB_Path
        {
            get => _compStat_Img_RegDB_Path;
            set => SetProperty(ref _compStat_Img_RegDB_Path, value, () => CompStat_Img_RegDB_Path);
        }

        public string CompStat_Txt_RegDB_Label
        {
            get => _compStat_Txt_RegDB_Label;
            set => SetProperty(ref _compStat_Txt_RegDB_Label, value, () => CompStat_Txt_RegDB_Label);
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

        public ICommand CmdLoadRegionalData => _cmdLoadRegionalData ?? (_cmdLoadRegionalData = new RelayCommand(async () =>
        {
            // Change UI to Underway
            StartOpUI();

            // Start the operation
            using (_cts = new CancellationTokenSource())
            {
                await LoadRegionalData(_cts.Token);
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

        private async Task LoadRegionalData(CancellationToken token)
        {
            bool edits_are_disabled = !Project.Current.IsEditingEnabled;
            int val = 0;
            int max = 17;

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

                // If editing is disabled, enable it temporarily (and disable again in the finally block)
                if (edits_are_disabled)
                {
                    if (!await Project.Current.SetIsEditingEnabledAsync(true))
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Unable to enabled editing for this ArcGIS Pro Project.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Unable to enabled editing for this ArcGIS Pro Project.");
                        return;
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
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Initializing the Regional Data Extractor..."), false, max, ++val);

                // Ensure the Project Geodatabase Exists
                string gdbpath = PRZH.GetPath_ProjectGDB();
                var tryex_gdb = await PRZH.GDBExists_Project();

                if (!tryex_gdb.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Project Geodatabase not found: '{gdbpath}'.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Project Geodatabase not found at {gdbpath}.");
                    return;
                }

                // Planning Unit existence
                var tryex_pudata = await PRZH.PUDataExists();
                if (!tryex_pudata.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Planning Units dataset not found.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Planning Units dataset not found.");
                    return;
                }

                // Ensure the Regional db exists
                string regpath = PRZH.GetPath_RegGDB();
                var tryexists_reg = await PRZH.GDBExists_Reg();
                if (!tryexists_reg.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Valid Regional Geodatabase not found: '{regpath}'.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Valid Regional Geodatabase not found at {regpath}.");
                    return;
                }

                if (ProMsgBox.Show($"If you proceed, any existing Regional Theme and Element tables in the Project Geodatabase will be overwritten." +
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

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting existing project tables..."), true, ++val);

                // Delete the Regional Theme Table if present
                if ((await PRZH.TableExists_Project(PRZC.c_TABLE_REGPRJ_THEMES)).exists)
                {
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_REGPRJ_THEMES);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting the {PRZC.c_TABLE_REGPRJ_THEMES} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error deleting the {PRZC.c_TABLE_REGPRJ_THEMES} table.");
                        return;
                    }
                }

                PRZH.CheckForCancellation(token);

                // Delete the Regional Element table if present
                if ((await PRZH.TableExists_Project(PRZC.c_TABLE_REGPRJ_ELEMENTS)).exists)
                {
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_REGPRJ_ELEMENTS);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting the {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error deleting the {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.");
                        return;
                    }
                }

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done deleting existing project tables..."), true, ++val);

                #endregion

                PRZH.CheckForCancellation(token);

                #region Process Regional tables

                // Process the regional tables

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Processing regional database..."), true, ++val);

                var tryregdb = await ProcessRegionalDbTables(token);

                if (!tryregdb.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error processing Regional DB tables.\n{tryregdb.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error processing Regional DB tables.\n{tryregdb.message}.");
                    return;
                }
                else
                {
                    val = PM.Current;
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done processing regional database..."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                #endregion

                // TODO: Visualize regional database data without Feature classes

                PRZH.CheckForCancellation(token);

                #region WRAP UP

                PRZH.UpdateProgress(PM, PRZH.WriteLog("Wrapping up..."), true, ++val);

                // Compact the Geodatabase
                toolParams = Geoprocessing.MakeValueArray(gdbpath);
                toolOutput = await PRZH.RunGPTool("Compact_management", toolParams, null, toolFlags_GPRefresh);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error compacting the geodatabase. GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show("Error compacting the geodatabase.");
                    return;
                }

                PRZH.CheckForCancellation(token);

                // TODO: Should this be enabled by default? Could be slow. Decide after re-implementing visualizations
                /*                // Refresh the Map & TOC
                                if (!(await PRZH.RedrawPRZLayers(_map)).success)
                                {
                                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error redrawing the PRZ layers.", LogMessageType.ERROR), true, ++val);
                                    ProMsgBox.Show($"Error redrawing the PRZ layers.");
                                    return;
                                }*/

                // Final message
                stopwatch.Stop();
                string message = PRZH.GetElapsedTimeMessage(stopwatch.Elapsed);
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Regional data extraction completed successfully!"), true, 1, 1);
                PRZH.UpdateProgress(PM, PRZH.WriteLog(message), true, 1, 1);

                ProMsgBox.Show("Regional data extraction completed successfully!" + Environment.NewLine + Environment.NewLine + message);

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

        private async Task<(bool success, string message)> ProcessRegionalDbTables(CancellationToken token)
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

                #region Initialize project tables

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Initializing project tables..."), true, ++val);

                // COPY THE ELEMENT TABLE
                string gdbpath = PRZH.GetPath_ProjectGDB();
                string regdbpath = PRZH.GetPath_RegGDB();

                var q_elem = await PRZH.GetRegDBQualifiedName(PRZC.c_TABLE_REGSRC_ELEMENTS);
                string inputelempath = Path.Combine(regdbpath, q_elem.qualified_name);
                toolParams = Geoprocessing.MakeValueArray(inputelempath, PRZC.c_TABLE_REGPRJ_ELEMENTS, "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("Copy_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying the {PRZC.c_TABLE_REGSRC_ELEMENTS} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error copying the {PRZC.c_TABLE_REGSRC_ELEMENTS} table.");
                    return (false, "table copy error.");
                }

                PRZH.CheckForCancellation(token);

                // ALTER ALIAS NAME OF ELEMENT TABLE
                await QueuedTask.Run(() =>
                {
                    var tryget_projectgdb = PRZH.GetGDB_Project();

                    if (!tryget_projectgdb.success)
                    {
                        throw new Exception("Error opening project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_projectgdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_REGPRJ_ELEMENTS))
                    using (TableDefinition tblDef = table.GetDefinition())
                    {
                        // Get the Table Description
                        TableDescription tblDescr = new TableDescription(tblDef);
                        tblDescr.AliasName = "Regional Elements";

                        // get the schemabuilder
                        SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);
                        schemaBuilder.Modify(tblDescr);
                        var success = schemaBuilder.Build();
                    }
                });

                // INSERT EXTRA FIELDS INTO ELEMENT TABLE
                string fldElemPresence = PRZC.c_FLD_TAB_REGELEMENT_PRESENCE + $" SHORT 'Presence' # {(int)ElementPresence.Absent} '" + PRZC.c_DOMAIN_PRESENCE + "';";
                string flds = fldElemPresence;

                toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_REGPRJ_ELEMENTS, flds);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("AddFields_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error adding fields to the copied {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error adding fields to the copied {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.");
                    return (false, "field addition error.");
                }

                PRZH.CheckForCancellation(token);

                // COPY THE THEMES TABLE
                var q_theme = await PRZH.GetRegDBQualifiedName(PRZC.c_TABLE_REGSRC_THEMES);
                string inputthemepath = Path.Combine(regdbpath, q_theme.qualified_name);
                toolParams = Geoprocessing.MakeValueArray(inputthemepath, PRZC.c_TABLE_REGPRJ_THEMES, "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("Copy_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying the {PRZC.c_TABLE_REGSRC_THEMES} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error copying the {PRZC.c_TABLE_REGSRC_THEMES} table.");
                    return (false, "table copy error.");
                }

                PRZH.CheckForCancellation(token);

                // ALTER ALIAS NAME OF THEME TABLE
                await QueuedTask.Run(() =>
                {
                    var tryget_projectgdb = PRZH.GetGDB_Project();

                    if (!tryget_projectgdb.success)
                    {
                        throw new Exception("Error opening project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_projectgdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_REGPRJ_THEMES))
                    using (TableDefinition tblDef = table.GetDefinition())
                    {
                        // Get the Table Description
                        TableDescription tblDescr = new TableDescription(tblDef);
                        tblDescr.AliasName = "Regional Themes";

                        // get the schemabuilder
                        SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);
                        schemaBuilder.Modify(tblDescr);
                        var success = schemaBuilder.Build();
                    }
                });

                // INSERT EXTRA FIELDS INTO THEME TABLE
                string fldThemePresence = PRZC.c_FLD_TAB_REGTHEME_PRESENCE + $" SHORT 'Presence' # {(int)ElementPresence.Absent} '" + PRZC.c_DOMAIN_PRESENCE + "';";
                flds = fldThemePresence;

                toolParams = Geoprocessing.MakeValueArray(PRZC.c_TABLE_REGPRJ_THEMES, flds);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("AddFields_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error adding fields to the copied {PRZC.c_TABLE_REGPRJ_THEMES} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error adding fields to the copied {PRZC.c_TABLE_REGPRJ_THEMES} table.");
                    return (false, "field addition error.");
                }

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done initializing project tables..."), true, ++val);

                #endregion

                PRZH.CheckForCancellation(token);

                #region RETRIEVE AND PREPARE INFO FROM REGIONAL DATABASE

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Parsing regional themes and elements..."), true, ++val);

                // Get the Regional Themes from the copied regional themes table
                var theme_outcome = await PRZH.GetRegionalThemes();
                if (!theme_outcome.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional themes.\n{theme_outcome.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving regional themes.\n{theme_outcome.message}");
                    return (false, "error retrieving themes.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {theme_outcome.themes.Count} regional themes."), true, ++val);
                }
                List<RegTheme> themes = theme_outcome.themes;

                PRZH.CheckForCancellation(token);

                // Get the Regional Elements from the copied regional elements table
                var elem_outcome = await PRZH.GetRegionalElements();
                if (!elem_outcome.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional elements.\n{elem_outcome.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving regional elements.\n{elem_outcome.message}");
                    return (false, "error retrieving elements.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {elem_outcome.elements.Count} regional elements."), true, ++val);
                }
                List<RegElement> elements = elem_outcome.elements;

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done parsing regional themes and elements..."), true, ++val);

                #endregion

                PRZH.CheckForCancellation(token);

                #region RETRIEVE INTERSECTING ELEMENTS

                int log_every = 500;
                int steps = elements.Count() / log_every;
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Finding intersecting cells in Regional Database..."), true, max + steps, ++val);

                #region Setup
                // Construct dictionary of planning units / regional grid ids
                var outcome = await PRZH.GetCellNumbersAndPUIDs();
                if (!outcome.success)
                {
                    throw new Exception("Error constructing PUID dictionary, try rebuilding planning units.");
                }
                Dictionary<long, int> study_area_cells = outcome.dict;

                // Load tile metadata for study area and regional database
                var tryread_studyarea_tiles = await PRZH.ReadBinary(PRZH.GetPath_ProjectTilesMetadataPath());
                var tryread_regdata_tiles = await PRZH.ReadBinary(PRZH.GetPath_RegionalDatabaseElementsTileMetadataPath());

                if (!tryread_studyarea_tiles.success)
                {
                    throw new Exception(tryread_studyarea_tiles.message);
                }
                HashSet<int> study_area_tiles = (HashSet<int>)tryread_studyarea_tiles.obj;

                if (!tryread_regdata_tiles.success)
                {
                    throw new Exception(tryread_regdata_tiles.message);
                }
                Dictionary<int, HashSet<int>> regdata_tiles = (Dictionary<int, HashSet<int>>)tryread_regdata_tiles.obj;

                // Setup up lists to track which elements, themes, and tiles are present
                Dictionary<string, HashSet<int>> element_tiles_present = new Dictionary<string, HashSet<int>>();
                HashSet<int> elements_present = new HashSet<int>();
                HashSet<int> themes_present = new HashSet<int>();

                // Refresh project-scope elements folder
                string output_elements_folder = PRZH.GetPath_ProjectRegionalElementsSubfolder();
                if (Directory.Exists(output_elements_folder))
                {
                    Directory.Delete(output_elements_folder, true);
                }
                Directory.CreateDirectory(output_elements_folder);

                #endregion

                PRZH.CheckForCancellation(token);

                // Find intersecting cells
                string input_elements_folder = PRZH.GetPath_RegionalDatabaseElementsSubfolder();
                int progress = 0;

                await Parallel.ForEachAsync(elements, async (element, token) =>
                {
                    progress++;
                    if (progress % log_every == 0)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done processing {progress} / {elements.Count()} regional elements."), true, ++val);
                    }

                    // Determine which if any tiles overlap
                    regdata_tiles[element.ElementID].IntersectWith(study_area_tiles);

                    // Skip elements with no overlapping tiles
                    if (regdata_tiles[element.ElementID].Count == 0)
                        return;

                    // Read in relevant element data tiles from regional database
                    foreach (int tile_id in regdata_tiles[element.ElementID])
                    {
                        string tile_filepath = Path.Combine(input_elements_folder, $"{element.ElementTable}-{tile_id}.bin");
                        var tryread_tile = await PRZH.ReadBinary(tile_filepath);

                        if (!tryread_tile.success)
                        {
                            throw new Exception(tryread_tile.message);
                        }

                        // Identify overlapping cells
                        // TODO: Could this be faster by splitting sutdy_area_cells by tile to limit search space?
                        Dictionary<long, double> intersecting_cells = ((Dictionary<long, double>)tryread_tile.obj).Where(x => study_area_cells.ContainsKey(x.Key)).ToDictionary(x => x.Key, x => x.Value);

                        if (intersecting_cells.Count() == 0) continue;

                        // Write to project
                        var trywrite = await PRZH.WriteBinary(intersecting_cells, Path.Combine(output_elements_folder, $"{element.ElementTable}-{tile_id}.bin"));

                        if (!trywrite.success)
                        {
                            throw new Exception(trywrite.message);
                        }
                    }

                    PRZH.CheckForCancellation(token);
                });


                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done finding intersecting cells..."), true, ++val);

                #endregion

                PRZH.CheckForCancellation(token);

                #region Update the element presence field

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Updating project tables..."), true, ++val);

                // Identify elements that are present by id
                foreach (string f in Directory.GetFiles(output_elements_folder))
                {
                    string[] element_info = Path.GetFileNameWithoutExtension(f).Split('-', 2);
                    int element_id = Convert.ToInt32(element_info[0].Substring(1));
                    int tile_id = Convert.ToInt32(element_info[1]);

                    if (!elements_present.Contains(element_id))
                    {
                        elements_present.Add(element_id);
                        element_tiles_present.Add(element_info[0], new HashSet<int>());
                    }
                    element_tiles_present[element_info[0]].Add(tile_id);
                }

                // Save list of intersecting elements and tiles to dictionary
                var trywrite_tiles = await PRZH.WriteBinary(element_tiles_present, PRZH.GetPath_ProjectRegionalElementTilesMetadataPath());
                if (!trywrite_tiles.success)
                {
                    throw new Exception(trywrite_tiles.message);
                }

                // Update the table
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = PRZH.GetGDB_Project();

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_REGPRJ_ELEMENTS))
                    using (RowCursor rowCursor = table.Search(null, false))
                    {
                        geodatabase.ApplyEdits(() =>
                        {
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    if (elements_present.Contains((int)row[PRZC.c_FLD_TAB_REGELEMENT_ELEMENT_ID]))
                                    {
                                        row[PRZC.c_FLD_TAB_REGELEMENT_PRESENCE] = (int)ElementPresence.Present;
                                        int theme_id = (int)row[PRZC.c_FLD_TAB_REGELEMENT_THEME_ID];
                                        if (!themes_present.Contains(theme_id))
                                        {
                                            themes_present.Add(theme_id);
                                        }
                                    }
                                    else
                                    {
                                        row[PRZC.c_FLD_TAB_REGELEMENT_PRESENCE] = (int)ElementPresence.Absent;
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
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = PRZH.GetGDB_Project();

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_REGPRJ_THEMES))
                    using (RowCursor rowCursor = table.Search(null, false))
                    {
                        geodatabase.ApplyEdits(() =>
                        {
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    int theme_id = (int)row[PRZC.c_FLD_TAB_REGTHEME_THEME_ID];

                                    if (themes_present.Contains(theme_id))
                                    {
                                        row[PRZC.c_FLD_TAB_REGTHEME_PRESENCE] = (int)ElementPresence.Present;
                                    }
                                    else
                                    {
                                        row[PRZC.c_FLD_TAB_REGTHEME_PRESENCE] = (int)ElementPresence.Absent;
                                    }

                                    row.Store();
                                }
                            }
                        });
                    }
                });

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done updating project tables..."), true, ++val);

                #endregion

                PRZH.CheckForCancellation(token);

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


        // TODO: Remove?
        /*private async Task<(bool success, string message)> ProcessRegionalFolder(
            RegionalDataSubfolder subFolderType,
            Dictionary<int, long> DICT_PUID_and_cellnumbers,
            CancellationToken token)
        {
            int val = PM.Current;
            int max = PM.Max;

            try
            {
                #region INITIALIZE

                // Initialize a few objects and names
                string gdbpath = PRZH.GetPath_ProjectGDB();
                string regpath = PRZH.GetPath_RegionalDataFolder();
                string regpath_subdir = PRZH.GetPath_RegionalDataSubfolder(subFolderType);
                string pu_ras_path = PRZH.GetPath_Project(PRZC.c_RAS_PLANNING_UNITS).path;

                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                string toolOutput;

                // Set some values based on subdir type
                string gl_name = "";
                ElementType elementType = ElementType.Unknown;
                ElementTheme elementTheme = ElementTheme.Unknown;

                switch (subFolderType)
                {
                    case RegionalDataSubfolder.GOALS:
                        gl_name = PRZC.c_GROUPLAYER_REG_GOALS;
                        elementType = ElementType.Goal;
                        elementTheme = ElementTheme.RegionalGoal;
                        break;

                    case RegionalDataSubfolder.WEIGHTS:
                        gl_name = PRZC.c_GROUPLAYER_REG_WEIGHTS;
                        elementType = ElementType.Weight;
                        elementTheme = ElementTheme.RegionalWeight;
                        break;

                    case RegionalDataSubfolder.INCLUDES:
                        gl_name = PRZC.c_GROUPLAYER_REG_INCLUDES;
                        elementType = ElementType.Include;
                        elementTheme = ElementTheme.RegionalInclude;
                        break;

                    case RegionalDataSubfolder.EXCLUDES:
                        gl_name = PRZC.c_GROUPLAYER_REG_EXCLUDES;
                        elementType = ElementType.Exclude;
                        elementTheme = ElementTheme.RegionalExclude;
                        break;
                }

                // Retrieve the Planning Unit Spatial Reference (from the FC)
                SpatialReference PlanningUnitSR = await QueuedTask.Run(() =>
                {
                    var tryget_fc = PRZH.GetFC_Project(PRZC.c_FC_PLANNING_UNITS);
                    using (FeatureClass featureClass = tryget_fc.featureclass)
                    using (FeatureClassDefinition fcDef = featureClass.GetDefinition())
                    {
                        return fcDef.GetSpatialReference();
                    }
                });

                // Retrieve Raster Cell Size and Extent
                double pu_cell_size = 0;
                Envelope extent_pu = null;

                await QueuedTask.Run(() =>
                {
                    var tryget = PRZH.GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);

                    using (RasterDataset RD = tryget.rasterDataset)
                    using (Raster raster = RD.CreateFullRaster())
                    {
                        pu_cell_size = raster.GetMeanCellSize().Item1;
                        extent_pu = raster.GetExtent();
                    }
                });

                // Calculate some pu raster-related values
                double poly_to_ras_cellsize = pu_cell_size / 10.0;    // TODO: Is this the best way to determine cell size?
                double poly_to_ras_cellarea = poly_to_ras_cellsize * poly_to_ras_cellsize;  // square meters

                #endregion

                PRZH.CheckForCancellation(token);

                #region PREPARE GROUP LAYER

                // Determine if the group layer exists already - if so, delete it
                GroupLayer GL = _map.GetLayersAsFlattenedList().OfType<GroupLayer>().FirstOrDefault(gl => string.Equals(gl.Name, gl_name, StringComparison.OrdinalIgnoreCase));
                if (GL != null)
                {
                    await QueuedTask.Run(() => { _map.RemoveLayer(GL); });
                }

                // Create the new group layer
                GL = await QueuedTask.Run(() => { return LayerFactory.Instance.CreateGroupLayer(_map, _map.Layers.Count, gl_name); });

                // Create and load the raster layer
                Uri uri = new Uri(pu_ras_path);
                LayerCreationParams creationParams = new LayerCreationParams(uri)
                {
                    Name = "pu",
                    IsVisible = false,
                    MapMemberPosition = MapMemberPosition.AddToBottom
                };

                Layer pu_layer = null;
                await QueuedTask.Run(() =>
                {
                    pu_layer = LayerFactory.Instance.CreateLayer<RasterLayer>(creationParams, GL);
                });
                RasterLayer pu_rl = (RasterLayer)pu_layer;

                #endregion

                PRZH.CheckForCancellation(token);

                #region ASSEMBLE LAYERS

                // find all lyrx files
                var layer_files = Directory.EnumerateFiles(regpath_subdir, "*.lyrx", SearchOption.TopDirectoryOnly);

                // Lists of individual FL and RL CIM Documents.  I can turn these into layers later.  OPTION A
                List<CIMLayerDocument> cimDocuments_FL = new List<CIMLayerDocument>();
                List<CIMLayerDocument> cimDocuments_RL = new List<CIMLayerDocument>();

                // Lists of lyrx paths and associated lists of FL or RL CIMLayerDocuments.  OPTION B
                List<(string lyrx_path, List<CIMLayerDocument> cimLyrDocs)> FL_CIMDocs = new List<(string lyrx_path, List<CIMLayerDocument> cimLyrDocs)>();
                List<(string lyrx_path, List<CIMLayerDocument> cimLyrDocs)> RL_CIMDocs = new List<(string lyrx_path, List<CIMLayerDocument> cimLyrDocs)>();

                await QueuedTask.Run(() =>
                {
                    foreach (var layer_file in layer_files)
                    {
                        // Get the lyrx file's underlying CIMLayerDocument and all FL and RL cimdefinitions in it
                        LayerDocument layerDocument = new LayerDocument(layer_file);
                        CIMLayerDocument cimLayerDocument = layerDocument.GetCIMLayerDocument();

                        // Get a list of any feature or raster layer definitions in the CIMLayerDocument
                        List<CIMFeatureLayer> cimFLDefs = cimLayerDocument.LayerDefinitions.OfType<CIMFeatureLayer>().ToList();
                        List<CIMRasterLayer> cimRLDefs = cimLayerDocument.LayerDefinitions.OfType<CIMRasterLayer>().ToList();

                        // Move on if this lyrx file contains no feature or raster layers
                        if (cimFLDefs.Count == 0 & cimRLDefs.Count == 0)
                        {
                            continue;
                        }

                        // Process the FL CIMDefinitions
                        if (cimFLDefs.Count > 0)
                        {
                            List<CIMLayerDocument> tempCIMs = new List<CIMLayerDocument>();

                            foreach (var cimFLDef in cimFLDefs)
                            {
                                // Create a new CIMLayerDocument
                                CIMLayerDocument tempCIM = new CIMLayerDocument();

                                // Set properties
                                tempCIM.Layers = new string[] { cimFLDef.URI };
                                tempCIM.LayerDefinitions = new CIMDefinition[] { cimFLDef };

                                // Store the CIM Doc
                                cimDocuments_FL.Add(tempCIM);
                                tempCIMs.Add(tempCIM);
                            }

                            FL_CIMDocs.Add((layer_file, tempCIMs));
                        }

                        // Process the RL CIMDefinitions
                        if (cimRLDefs.Count > 0)
                        {
                            List<CIMLayerDocument> tempCIMs = new List<CIMLayerDocument>();

                            foreach (var cimRLDef in cimRLDefs)
                            {
                                // Create a new CIMLayerDocument
                                CIMLayerDocument tempCIM = new CIMLayerDocument();

                                // Set properties
                                tempCIM.Layers = new string[] { cimRLDef.URI };
                                tempCIM.LayerDefinitions = new CIMDefinition[] { cimRLDef };

                                // Store the CIM Doc
                                cimDocuments_RL.Add(tempCIM);
                                tempCIMs.Add(tempCIM);
                            }

                            RL_CIMDocs.Add((layer_file, tempCIMs));
                        }
                    }
                });

                PRZH.CheckForCancellation(token);

                // Build lists of layers
                List<(string lyrx_path, List<FeatureLayer> FLs)> lyrx_FLs = new List<(string lyrx_path, List<FeatureLayer> FLs)>();
                List<(string lyrx_path, List<RasterLayer> RLs)> lyrx_RLs = new List<(string lyrx_path, List<RasterLayer> RLs)>();

                // Load the feature layers here
                foreach (var o in FL_CIMDocs)
                {
                    List<FeatureLayer> fls = new List<FeatureLayer>();

                    foreach (var a in o.cimLyrDocs)
                    {
                        await QueuedTask.Run(() =>
                        {
                            // Create and add the layer
                            LayerCreationParams lcparams = new LayerCreationParams(a)
                            {
                                IsVisible = false,
                                MapMemberPosition = MapMemberPosition.AddToTop
                            };
                            var l = LayerFactory.Instance.CreateLayer<FeatureLayer>(lcparams, GL);

                            // Check for valid source Feature Class
                            bool bad_fc = false;

                            try
                            {
                                using (FeatureClass featureClass = l.GetFeatureClass())
                                {
                                    if (featureClass == null)
                                    {
                                        bad_fc = true;
                                    }
                                }
                            }
                            catch
                            {
                                bad_fc = true;
                            }

                            // If FC is invalid, leave
                            if (bad_fc)
                            {
                                GL.RemoveLayer(l);
                            }

                            // If FL is not a polygon feature layer, leave
                            else if (l.ShapeType != esriGeometryType.esriGeometryPolygon)
                            {
                                GL.RemoveLayer(l);
                            }

                            // If FL has null or unknown spatial reference, leave
                            else
                            {
                                var sr = l.GetSpatialReference();

                                if (sr == null || sr.IsUnknown)
                                {
                                    GL.RemoveLayer(l);
                                }
                                else
                                {
                                    // Finally, we have an OK feature layer!
                                    fls.Add(l);
                                }
                            }
                        });
                    }

                    lyrx_FLs.Add((o.lyrx_path, fls));
                }

                PRZH.CheckForCancellation(token);

                // Load the raster layers here
                foreach (var o in RL_CIMDocs)
                {
                    List<RasterLayer> rls = new List<RasterLayer>();

                    foreach (var a in o.cimLyrDocs)
                    {
                        await QueuedTask.Run(() =>
                        {
                            // Create and add the layer
                            LayerCreationParams lcparams = new LayerCreationParams(a)
                            {
                                IsVisible = false,
                                MapMemberPosition = MapMemberPosition.AddToTop
                            };

                            var l = LayerFactory.Instance.CreateLayer<RasterLayer>(lcparams, GL);

                            // Check for a valid Raster source
                            bool bad_raster = false;
                            int band_count = 0;

                            try
                            {
                                using (Raster raster = l.GetRaster())
                                {
                                    if (raster == null)
                                    {
                                        bad_raster = true;
                                    }

                                    band_count = raster.GetBandCount();
                                }
                            }
                            catch
                            {
                                bad_raster = true;
                            }

                            // If Raster is invalid, leave
                            if (bad_raster)
                            {
                                GL.RemoveLayer(l);
                            }

                            // If raster has more than 1 band, leave
                            else if (band_count > 1)
                            {
                                GL.RemoveLayer(l);
                            }

                            // If RL has null or unknown spatial reference, leave
                            else
                            {
                                var sr = l.GetSpatialReference();

                                if (sr == null || sr.IsUnknown)
                                {
                                    GL.RemoveLayer(l);
                                }
                                else
                                {
                                    // Finally, we have an OK raster layer!
                                    rls.Add(l);
                                }
                            }
                        });
                    }

                    lyrx_RLs.Add((o.lyrx_path, rls));
                }

                PRZH.CheckForCancellation(token);

                // ensure at least one layer is available
                bool fls_present = false;
                bool rls_present = false;

                foreach (var u in lyrx_FLs)
                {
                    if (u.FLs.Count > 0)
                    {
                        fls_present = true;
                        break;
                    }
                }

                foreach (var u in lyrx_RLs)
                {
                    if (u.RLs.Count > 0)
                    {
                        rls_present = true;
                        break;
                    }
                }

                if (!fls_present & !rls_present)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"No valid layers found within the regional {subFolderType} directory"), true, ++val);
                    await QueuedTask.Run(() => { _map.RemoveLayer(GL); });
                    return (true, "success");
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region PROCESS LAYERS TO EXTRACT ELEMENTS

                // Prepare list of Elements
                List<RegElement> regElements = new List<RegElement>();
                int id = 1;

                // Populate list with Feature Elements
                if (fls_present)
                {
                    for (int i = 0; i < lyrx_FLs.Count; i++)
                    {
                        (string lyrx_path, List<FeatureLayer> FLs) FLInfo = lyrx_FLs[i];

                        foreach (FeatureLayer FL in FLInfo.FLs)
                        {
                            // validate this feature layer for renderer stuff
                            await QueuedTask.Run(async () =>
                            {
                                // Get the renderer
                                var rend = FL.GetRenderer();

                                if (rend is CIMUniqueValueRenderer UVRend)
                                {
                                    // Get the field index plus the field type, for each of the 1, 2, or 3 fields in the UV Renderer
                                    Dictionary<int, FieldCategory> DICT_FieldIndex_and_category = new Dictionary<int, FieldCategory>();

                                    for (int b = 0; b < UVRend.Fields.Length; b++)
                                    {
                                        string uvrend_fieldname = UVRend.Fields[b];

                                        foreach (ArcGIS.Desktop.Mapping.FieldDescription fieldDescription in FL.GetFieldDescriptions())
                                        {
                                            if (uvrend_fieldname == fieldDescription.Name)
                                            {
                                                FieldCategory fcat = PRZH.GetFieldCategory(fieldDescription);

                                                if (fcat == FieldCategory.DATE)
                                                {
                                                    throw new Exception($"Layer: {FL.Name} >> Date Fields in UV Legends not supported.");
                                                }

                                                DICT_FieldIndex_and_category.Add(b, fcat);
                                            }
                                        }
                                    }

                                    // Make sure we picked up a DICT entry for each UVRend field name
                                    if (UVRend.Fields.Length != DICT_FieldIndex_and_category.Count)
                                    {
                                        throw new Exception($"Layer: {FL.Name} >> Not all renderer fields found within layer table.");
                                    }

                                    // Cycle through each Legend Group, retrieve Group heading...
                                    CIMUniqueValueGroup[] UVGroups = UVRend.Groups;

                                    if (UVGroups is null)
                                    {
                                        throw new Exception($"Layer: {FL.Name} >> UV Renderer has no groups.");
                                    }

                                    foreach (CIMUniqueValueGroup UVGroup in UVGroups)
                                    {
                                        // Get the group heading
                                        string group_heading = UVGroup.Heading;

                                        // Retrieve the Classes in this Group
                                        CIMUniqueValueClass[] UVClasses = UVGroup.Classes;

                                        // Each UVClass will become its own unique regional element
                                        foreach (CIMUniqueValueClass UVClass in UVClasses)
                                        {
                                            // Get the class label
                                            string class_label = UVClass.Label;

                                            // Retrieve the "Tuples" associated with this UVClass
                                            // A "Tuple" is a collection of 1, 2, or 3 specific field values (from the 1, 2 or 3 fields in the UV Renderer)
                                            // A UVClass can consist of 1 or more "Tuples".  By default, 1, but if the user groups together 2 or more classes into a single class,
                                            // the UVClass will consist of those 2 or more "Tuples".
                                            CIMUniqueValue[] UVClassTuples = UVClass.Values;

                                            string classClause = "";

                                            // For Each Tuple (could be 1 to many many)
                                            for (int tupIx = 0; tupIx < UVClassTuples.Length; tupIx++)
                                            {
                                                CIMUniqueValue tuple = UVClassTuples[tupIx];

                                                string tupleClause = "";

                                                // For each field value in the tuple (could be 1, 2, or 3)
                                                for (int fldIx = 0; fldIx < tuple.FieldValues.Length; fldIx++)
                                                {
                                                    string fieldValue = tuple.FieldValues[fldIx];
                                                    bool IsNull = fieldValue == "<Null>";

                                                    string Expression = "";

                                                    switch (DICT_FieldIndex_and_category[fldIx])
                                                    {
                                                        case FieldCategory.STRING:
                                                            Expression = (IsNull) ? "IS NULL" : "= '" + fieldValue.Replace("'", "''") + "'";
                                                            break;

                                                        case FieldCategory.NUMERIC:
                                                            Expression = (IsNull) ? "IS NULL" : "= " + fieldValue;
                                                            break;

                                                        default:
                                                            break;
                                                    }

                                                    // Assemble the unit clause - the building block of the main where clause
                                                    string unitClause = "(" + UVRend.Fields[fldIx] + " " + Expression + ")";

                                                    // Assemble the tuple clause - the where clause of this particular tuple
                                                    if (fldIx == 0)
                                                    {
                                                        tupleClause = unitClause;
                                                    }
                                                    else
                                                    {
                                                        tupleClause += " And " + unitClause;
                                                    }
                                                }

                                                tupleClause = "(" + tupleClause + ")";

                                                if (tupIx == 0)
                                                {
                                                    classClause = tupleClause;
                                                }
                                                else
                                                {
                                                    classClause += " Or " + tupleClause;
                                                }
                                            }

                                            classClause = "(" + classClause + ")";  // this is the where clause

                                            // Assemble the new element
                                            RegElement regElement = new RegElement();

                                            regElement.ElementID = id++;
                                            regElement.ElementName = FL.Name + " - " + group_heading + " - " + class_label;
                                            regElement.ElementType = (int)elementType;
                                            regElement.LayerObject = FL;
                                            regElement.LayerName = FL.Name;
                                            regElement.LayerType = (int)LayerType.FEATURE;
                                            regElement.LayerJson = await QueuedTask.Run(() => { return (FL.GetDefinition()).ToJson(); });
                                            regElement.LyrxPath = FLInfo.lyrx_path;
                                            regElement.WhereClause = classClause;
                                            regElement.LegendGroup = group_heading;
                                            regElement.LegendClass = class_label;
                                            regElement.ThemeID = (int)elementTheme;

                                            regElements.Add(regElement);
                                        }
                                    }
                                }
                                else
                                {
                                    RegElement regElement = new RegElement();

                                    regElement.ElementID = id++;
                                    regElement.ElementName = FL.Name;
                                    regElement.ElementType = (int)elementType;
                                    regElement.LayerObject = FL;
                                    regElement.LayerName = FL.Name;
                                    regElement.LayerType = (int)LayerType.FEATURE;
                                    regElement.LayerJson = await QueuedTask.Run(() => { return (FL.GetDefinition()).ToJson(); });
                                    regElement.LyrxPath = FLInfo.lyrx_path;
                                    regElement.ThemeID = (int)elementTheme;

                                    regElements.Add(regElement);
                                }
                            });
                        }

                        PRZH.CheckForCancellation(token);
                    }
                }

                // Populate list with Raster Elements
                if (rls_present)
                {
                    for (int i = 0; i < lyrx_RLs.Count; i++)
                    {
                        (string lyrx_path, List<RasterLayer> RLs) RLInfo = lyrx_RLs[i];

                        foreach (RasterLayer RL in RLInfo.RLs)
                        {
                            RegElement regElement = new RegElement();

                            regElement.ElementID = id++;
                            regElement.ElementName = RL.Name;
                            regElement.ElementType = (int)elementType;
                            regElement.LayerObject = RL;
                            regElement.LayerName = RL.Name;
                            regElement.LayerType = (int)LayerType.RASTER;
                            regElement.LayerJson = await QueuedTask.Run(() => { return (RL.GetDefinition()).ToJson(); });
                            regElement.LyrxPath = RLInfo.lyrx_path;
                            regElement.ThemeID = (int)elementTheme;

                            regElements.Add(regElement);
                        }

                        PRZH.CheckForCancellation(token);
                    }
                }

                #endregion

                #region WRITE ELEMENTS TO REG ELEMENT TABLE

                // Sort by element ID (just in case).  I think the objects got added to the list already ordered correctly...
                regElements.Sort((x, y) => x.ElementID.CompareTo(y.ElementID));

                // Ensure the table is present
                var tryex_regelem = await PRZH.TableExists_Project(PRZC.c_TABLE_REGPRJ_ELEMENTS);
                if (!tryex_regelem.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Unable to find the {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Unable to find the {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.");
                    return (false, "error finding table.");
                }

                // Populate the table
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Populating the {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.."), true, ++val);
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = PRZH.GetGDB_Project();

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_REGPRJ_ELEMENTS))
                    using (RowBuffer rowBuffer = table.CreateRowBuffer())
                    {
                        geodatabase.ApplyEdits(() =>
                        {
                            foreach (var elem in regElements)
                            {
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_ELEMENT_ID] = elem.ElementID;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_NAME] = elem.ElementName;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_TYPE] = elem.ElementType;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_STATUS] = (int)ElementStatus.Active;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_PRESENCE] = (int)ElementPresence.Absent;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_LYRXPATH] = elem.LyrxPath;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_LAYERNAME] = elem.LayerName;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_LAYERTYPE] = ((LayerType)elem.LayerType).ToString();
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_LAYERJSON] = elem.LayerJson;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_WHERECLAUSE] = elem.WhereClause;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_LEGENDGROUP] = elem.LegendGroup;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_LEGENDCLASS] = elem.LegendClass;
                                rowBuffer[PRZC.c_FLD_TAB_REGELEMENT_THEME_ID] = elem.ThemeID;

                                table.CreateRow(rowBuffer);
                            }
                        });
                    }
                });

                #endregion

                PRZH.CheckForCancellation(token);

                #region OVERLAY ELEMENTS WITH PLANNING UNITS

                // pu layer extent (in map SR)
                Envelope extent_pu_query = await QueuedTask.Run(() => { return pu_layer.QueryExtent(); });

                // Process each element
                foreach (var regElement in regElements)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Processing regional element {regElement.ElementID}: {regElement.ElementName}..."), true, ++val);

                    // Get the layer
                    Layer lyr = regElement.LayerObject;

                    // Set up geodatabase object names
                    string reg_raster_init = PRZC.c_RAS_REG_ELEM_PREFIX + regElement.ElementID.ToString() + PRZC.c_RAS_REG_ELEM_SUFFIX_ORIG;
                    string reg_raster_2 = PRZC.c_RAS_REG_ELEM_PREFIX + regElement.ElementID.ToString() + PRZC.c_RAS_REG_ELEM_SUFFIX_RECLASS;
                    string reg_raster_pureclass = PRZC.c_RAS_REG_ELEM_PREFIX + "pu_reclass_" + regElement.ElementID.ToString();

                    string zonal_stats_table = PRZC.c_TABLE_REG_ZONALSTATS_PREFIX + regElement.ElementID.ToString() + PRZC.c_TABLE_REG_ZONALSTATS_SUFFIX;

                    List<string> gdb_objects_delete = new List<string>();

                    #region PROJECTION AND GEOTRANSFORMATION INFORMATION

                    // layer extent (in map SR)
                    Envelope extent_lyr_query = await QueuedTask.Run(() => { return lyr.QueryExtent(); });

                    // intersection of pu and layer extents in map sr
                    EnvelopeBuilderEx builderEx = new EnvelopeBuilderEx(extent_pu_query);
                    bool hasoverlap = builderEx.Intersects(extent_lyr_query);
                    builderEx.Intersection(extent_lyr_query);
                    Envelope extent_intersection_query = (Envelope)builderEx.ToGeometry();

                    if (!hasoverlap)
                    {
                        // extents don't overlap, so no point continuing with this one.
                        continue;
                    }

                    // layer spatial reference
                    SpatialReference lyr_sr = await QueuedTask.Run(() => { return lyr.GetSpatialReference(); });
                    SpatialReference map_sr = _map.SpatialReference;

                    // Projection Transformations
                    // Forwards
                    ProjectionTransformation projTrans_pu_to_lyr = ProjectionTransformation.Create(PlanningUnitSR, lyr_sr);
                    ProjectionTransformation projTrans_pu_to_map = ProjectionTransformation.Create(PlanningUnitSR, map_sr);
                    ProjectionTransformation projTrans_lyr_to_map = ProjectionTransformation.Create(lyr_sr, map_sr);

                    // Inverse
                    ProjectionTransformation projTrans_lyr_to_pu = projTrans_pu_to_lyr.GetInverse();
                    ProjectionTransformation projTrans_map_to_pu = projTrans_pu_to_map.GetInverse();
                    ProjectionTransformation projTrans_map_to_lyr = projTrans_lyr_to_map.GetInverse();

                    // Geographic Transformations (may be null)
                    CompositeGeographicTransformation compGeoTrans_pu_to_lyr = (CompositeGeographicTransformation)projTrans_pu_to_lyr.Transformation;
                    CompositeGeographicTransformation compGeoTrans_lyr_to_pu = (CompositeGeographicTransformation)projTrans_lyr_to_pu.Transformation;
                    CompositeGeographicTransformation compGeoTrans_pu_to_map = (CompositeGeographicTransformation)projTrans_pu_to_map.Transformation;
                    CompositeGeographicTransformation compGeoTrans_map_to_pu = (CompositeGeographicTransformation)projTrans_map_to_pu.Transformation;
                    CompositeGeographicTransformation compGeoTrans_lyr_to_map = (CompositeGeographicTransformation)projTrans_lyr_to_map.Transformation;
                    CompositeGeographicTransformation compGeoTrans_map_to_lyr = (CompositeGeographicTransformation)projTrans_map_to_lyr.Transformation;

                    // Transformation Existences
                    bool geoTransExists_pu_to_lyr = false;
                    bool geoTransExists_lyr_to_pu = false;
                    bool geoTransExists_pu_to_map = false;
                    bool geoTransExists_map_to_pu = false;
                    bool geoTransExists_lyr_to_map = false;
                    bool geoTransExists_map_to_lyr = false;

                    // Transformations as semi-colon delimited strings
                    string geoTransNames_pu_to_lyr = "";
                    string geoTransNames_lyr_to_pu = "";
                    string geoTransNames_pu_to_map = "";
                    string geoTransNames_map_to_pu = "";
                    string geoTransNames_lyr_to_map = "";
                    string geoTransNames_map_to_lyr = "";

                    // Get the transformation strings

                    // PU to LYR
                    if (compGeoTrans_pu_to_lyr != null)
                    {
                        geoTransExists_pu_to_lyr = true;
                        var names = compGeoTrans_pu_to_lyr.Transformations.Select(t => t.Name).ToList();
                        geoTransNames_pu_to_lyr = string.Join(";", names);
                    }

                    // LYR to PU
                    if (compGeoTrans_lyr_to_pu != null)
                    {
                        geoTransExists_lyr_to_pu = true;
                        var names = compGeoTrans_lyr_to_pu.Transformations.Select(t => t.Name).ToList();
                        geoTransNames_lyr_to_pu = string.Join(";", names);
                    }

                    // PU to MAP
                    if (compGeoTrans_pu_to_map != null)
                    {
                        geoTransExists_pu_to_map = true;
                        var names = compGeoTrans_pu_to_map.Transformations.Select(t => t.Name).ToList();
                        geoTransNames_pu_to_map = string.Join(";", names);
                    }

                    // MAP to PU
                    if (compGeoTrans_map_to_pu != null)
                    {
                        geoTransExists_map_to_pu = true;
                        var names = compGeoTrans_map_to_pu.Transformations.Select(t => t.Name).ToList();
                        geoTransNames_map_to_pu = string.Join(";", names);
                    }

                    // LYR to MAP
                    if (compGeoTrans_lyr_to_map != null)
                    {
                        geoTransExists_lyr_to_map = true;
                        var names = compGeoTrans_lyr_to_map.Transformations.Select(t => t.Name).ToList();
                        geoTransNames_lyr_to_map = string.Join(";", names);
                    }

                    // MAP to LYR
                    if (compGeoTrans_map_to_lyr != null)
                    {
                        geoTransExists_map_to_lyr = true;
                        var names = compGeoTrans_map_to_lyr.Transformations.Select(t => t.Name).ToList();
                        geoTransNames_map_to_lyr = string.Join(";", names);
                    }

                    // Prepare the pu extent envelopes
                    Envelope pu_extent_in_lyr_sr = (Envelope)GeometryEngine.Instance.ProjectEx(extent_pu, projTrans_pu_to_lyr);
                    Envelope pu_extent_in_map_sr = (Envelope)GeometryEngine.Instance.ProjectEx(extent_pu, projTrans_pu_to_map);     // should be the same as pu_query_extent

                    #endregion

                    // Process varies based on element layer type
                    LayerType layerType = (LayerType)regElement.LayerType;
                    if (layerType == LayerType.FEATURE)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Source layer is of type: {layerType}"), true, ++val);

                        // Get the element feature layer
                        FeatureLayer regFL = (FeatureLayer)lyr;

                        // Apply selection if where clause was stored
                        if (!string.IsNullOrEmpty(regElement.WhereClause))
                        {
                            await QueuedTask.Run(() =>
                            {
                                QueryFilter queryFilter = new QueryFilter()
                                {
                                    WhereClause = regElement.WhereClause
                                };

                                regFL.Select(queryFilter, SelectionCombinationMethod.New);
                            });

                            if (regFL.SelectionCount == 0)
                            {
                                // no matching features - I can skip this element
                                PRZH.UpdateProgress(PM, PRZH.WriteLog($"No feature selection found.  Skipping layer."), true, ++val);
                                continue;
                            }
                        }
                        else
                        {
                            await QueuedTask.Run(() =>
                            {
                                regFL.ClearSelection();
                            });
                        }

                        // Get the object id field
                        string oidfield = await QueuedTask.Run(() =>
                        {
                            using (FeatureClass featureClass = regFL.GetFeatureClass())
                            using (FeatureClassDefinition fcDef = featureClass.GetDefinition())
                            {
                                return fcDef.GetObjectIDField();
                            }
                        });

                        // Convert feature layer to raster
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Converting {regElement.ElementName} to raster..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(regFL, oidfield, reg_raster_init, "CELL_CENTER", "NONE", poly_to_ras_cellsize, "BUILD");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            outputCoordinateSystem: lyr_sr,
                            extent: extent_pu_query);
                        toolOutput = await PRZH.RunGPTool("PolygonToRaster_conversion", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error converting polygons to raster.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error converting polygons to raster.");
                            return (false, "error converting polygons to raster.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Polygons converted to raster."), true, ++val);
                        }
                        gdb_objects_delete.Add(reg_raster_init);

                        PRZH.CheckForCancellation(token);

                        // Build Pyramids
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Building pyramids..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_init);
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await PRZH.RunGPTool("BuildPyramids_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error building pyramids.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error building pyramids.");
                            return (false, "error building pyramids.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"pyramids built successfully."), true, ++val);
                        }

                        PRZH.CheckForCancellation(token);

                        // Build Statistics
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Calculating Statistics..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_init, 1, 1, "");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await PRZH.RunGPTool("CalculateStatistics_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error calculating statistics.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error calculating statistics.");
                            return (false, "error calculating stats.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Statistics calculated."), true, ++val);
                        }

                        PRZH.CheckForCancellation(token);

                        // Reclass the new raster so that all non-null values are now equal to cell area
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Reclassing raster values..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_init, poly_to_ras_cellarea, reg_raster_2, "", "Value IS NOT NULL");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            outputCoordinateSystem: lyr_sr,
                            cellSize: poly_to_ras_cellsize);
                        toolOutput = await PRZH.RunGPTool("Con_sa", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error reclassing values.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error reclassing values.");
                            return (false, "error reclassing values.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Values reclassed."), true, ++val);
                        }
                        gdb_objects_delete.Add(reg_raster_2);

                        PRZH.CheckForCancellation(token);

                        // Build Pyramids
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Building pyramids..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_2);
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await PRZH.RunGPTool("BuildPyramids_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error building pyramids.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error building pyramids.");
                            return (false, "error building pyramids.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"pyramids built successfully."), true, ++val);
                        }

                        PRZH.CheckForCancellation(token);

                        // Build Statistics
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Calculating Statistics..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_2, 1, 1, "");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await PRZH.RunGPTool("CalculateStatistics_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error calculating statistics.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error calculating statistics.");
                            return (false, "error calculating stats.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Statistics calculated."), true, ++val);
                        }

                        PRZH.CheckForCancellation(token);

                        // resize pu raster to match cell size of layer raster
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Copying PU Raster..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(pu_layer, reg_raster_pureclass);
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            outputCoordinateSystem: PlanningUnitSR,
                            cellSize: poly_to_ras_cellsize,
                            resamplingMethod: "NEAREST");
                        toolOutput = await PRZH.RunGPTool("CopyRaster_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying raster.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error copying raster.");
                            return (false, "error copying raster");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"raster copied."), true, ++val);
                        }
                        gdb_objects_delete.Add(reg_raster_pureclass);

                        PRZH.CheckForCancellation(token);

                        // Snap raster for zonal stats
                        string snapraster = PRZH.GetPath_Project(reg_raster_pureclass).path;

                        // Zonal Statistics as Table
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Zonal Statistics as Table..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_pureclass, PRZC.c_FLD_RAS_PU_ID, reg_raster_2, zonal_stats_table, "DATA", "ALL");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            cellSize: poly_to_ras_cellsize,
                            snapRaster: snapraster,
                            outputCoordinateSystem: PlanningUnitSR,
                            geographicTransformations: geoTransNames_lyr_to_pu);
                        toolOutput = await PRZH.RunGPTool("ZonalStatisticsAsTable_sa", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error calculating zonal statistics as table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error calculating zonal statistics as table.");
                            return (false, "error calculating zonal statistics as table");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Zonal statistics calculated."), true, ++val);
                        }
                        gdb_objects_delete.Add(zonal_stats_table);
                    }
                    else if (layerType == LayerType.RASTER)
                    {
                        // Get the element raster layer
                        RasterLayer regRL = (RasterLayer)regElement.LayerObject;

                        // Get some raster details
                        int band_count = 0;
                        bool has_nodata_value = false;
                        List<double> nodata_values = new List<double>();
                        double celsiz = 0;
                        string interpolation_method = "";

                        await QueuedTask.Run(() =>
                        {
                            using (Raster raster = regRL.GetRaster())
                            {
                                // Determine interpolation method
                                interpolation_method = raster.IsInteger() ? "NEAREST" : "BILINEAR";

                                // Get Band Count
                                band_count = raster.GetBandCount();

                                // Get cell size
                                celsiz = raster.GetMeanCellSize().Item1;

                                // Get nodata values if they exist
                                var nodata_object = raster.GetNoDataValue();

                                if (nodata_object != null)
                                {
                                    // there is a nodata value
                                    has_nodata_value = true;

                                    // cast to ienumerable
                                    IEnumerable enumerable = nodata_object as IEnumerable;

                                    if (enumerable != null)
                                    {
                                        foreach (object o in enumerable)
                                        {
                                            try
                                            {
                                                double d = Convert.ToDouble(o);
                                                nodata_values.Add(d);
                                            }
                                            catch
                                            {
                                                nodata_values.Add(-9999);
                                            }
                                        }
                                    }
                                }
                            }
                        });

                        // Raster can only have a single band
                        if (band_count > 1)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Raster has more than 1 band.\n{regRL.Name}", LogMessageType.VALIDATION_ERROR), true, ++val);
                            ProMsgBox.Show($"Raster has more than 1 band.\n{regRL.Name}");
                            return (false, "error - raster has more than 1 band.");
                        }

                        // Prepare nodata string
                        string nodataval = "";
                        if (has_nodata_value)
                        {
                            nodataval = nodata_values[0].ToString();
                        }

                        if (celsiz > poly_to_ras_cellsize)
                        {
                            celsiz = poly_to_ras_cellsize;
                        }

                        // TODO: Ensure that nodata values get respected when a tif is used

                        // Copy Raster to temp raster
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Copy raster..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(regRL, reg_raster_init, "", "", "");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            outputCoordinateSystem: lyr_sr,
                            extent: extent_pu_query,
                            cellSize: celsiz,
                            rasterStatistics: "STATISTICS 1 1",
                            pyramid: "PYRAMIDS -1",
                            resamplingMethod: interpolation_method);
                        toolOutput = await PRZH.RunGPTool("CopyRaster_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying raster.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error copying raster.");
                            return (false, "error copying raster.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Raster copied."), true, ++val);
                        }
                        gdb_objects_delete.Add(reg_raster_init);

                        PRZH.CheckForCancellation(token);

                        // Reclass the new raster so that all zero values are set to NODATA
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Reclassing raster values..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_init, reg_raster_init, reg_raster_2, "Value = 0");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            outputCoordinateSystem: lyr_sr,
                            cellSize: celsiz);
                        toolOutput = await PRZH.RunGPTool("SetNull_sa", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error reclassing values.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error reclassing values.");
                            return (false, "error reclassing values.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Values reclassed."), true, ++val);
                        }
                        gdb_objects_delete.Add(reg_raster_2);

                        PRZH.CheckForCancellation(token);

                        // Build Pyramids
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Building pyramids..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_2);
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await PRZH.RunGPTool("BuildPyramids_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error building pyramids.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error building pyramids.");
                            return (false, "error building pyramids.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"pyramids built successfully."), true, ++val);
                        }

                        PRZH.CheckForCancellation(token);

                        // Build Statistics
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Calculating Statistics..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(reg_raster_2, 1, 1, "");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await PRZH.RunGPTool("CalculateStatistics_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error calculating statistics.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error calculating statistics.");
                            return (false, "error calculating stats.");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Statistics calculated."), true, ++val);
                        }

                        PRZH.CheckForCancellation(token);

                        // resize pu raster to match cell size of layer raster
                        string pu_raster_reclass = PRZC.c_RAS_PLANNING_UNITS_RECLASS + regElement.ElementID.ToString();
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Copying PU Raster..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(pu_layer, pu_raster_reclass);
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            outputCoordinateSystem: PlanningUnitSR,
                            cellSize: poly_to_ras_cellsize,
                            resamplingMethod: "NEAREST");
                        toolOutput = await PRZH.RunGPTool("CopyRaster_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error copying raster.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error copying raster.");
                            return (false, "error copying raster");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"raster copied."), true, ++val);
                        }
                        gdb_objects_delete.Add(pu_raster_reclass);

                        PRZH.CheckForCancellation(token);

                        // Snap raster for zonal stats
                        string snapraster = PRZH.GetPath_Project(pu_raster_reclass).path;

                        // Zonal Statistics as Table
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Zonal Statistics as Table..."), true, ++val);
                        toolParams = Geoprocessing.MakeValueArray(pu_raster_reclass, PRZC.c_FLD_RAS_PU_ID, reg_raster_2, zonal_stats_table, "DATA", "ALL");
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(
                            workspace: gdbpath,
                            overwriteoutput: true,
                            cellSize: poly_to_ras_cellsize,
                            snapRaster: snapraster,
                            outputCoordinateSystem: PlanningUnitSR,
                            geographicTransformations: geoTransNames_lyr_to_pu,
                            resamplingMethod: interpolation_method);
                        toolOutput = await PRZH.RunGPTool("ZonalStatisticsAsTable_sa", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error calculating zonal statistics as table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                            ProMsgBox.Show($"Error calculating zonal statistics as table.");
                            return (false, "error calculating zonal statistics as table");
                        }
                        else
                        {
                            PRZH.UpdateProgress(PM, PRZH.WriteLog($"Zonal statistics calculated."), true, ++val);
                        }
                        gdb_objects_delete.Add(zonal_stats_table);
                    }

                    // Retrieve dictionary from zonal stats table:   puid, sum
                    Dictionary<int, double> DICT_PUID_and_value_sum = new Dictionary<int, double>();

                    await QueuedTask.Run(() =>
                    {
                        var tryget_table = PRZH.GetTable_Project(zonal_stats_table);
                        if (!tryget_table.success)
                        {
                            throw new Exception($"Unable to retrieve the temp zonal stats table {zonal_stats_table}");
                        }

                        using (Table table = tryget_table.table)
                        using (RowCursor rowCursor = table.Search())
                        {
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    int puid = Convert.ToInt32(row[PRZC.c_FLD_RAS_PU_ID]);
                                    double sumval = Convert.ToDouble(row[PRZC.c_FLD_ZONALSTATS_SUM]);

                                    //if (puid > 0 && sumval > 0)
                                    if (puid > 0)
                                    {
                                        DICT_PUID_and_value_sum.Add(puid, sumval);
                                    }
                                }
                            }
                        }
                    });

                    PRZH.CheckForCancellation(token);

                    // Delete temp objects
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting temp objects..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(string.Join(";", gdb_objects_delete));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting temp objects.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error deleting temp objects.");
                        return (false, "error deleting temp objects.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Temp objects deleted successfully."), true, ++val);
                    }

                    PRZH.CheckForCancellation(token);

                    // if there are no entries in the zonal stats dict, move on to next regElement
                    if (DICT_PUID_and_value_sum.Count == 0)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"No data to store for regional element {regElement.ElementName}"), true, ++val);
                        continue;
                    }

                    // Create the regional element table
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Creating the {regElement.ElementTable} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(gdbpath, regElement.ElementTable, "", "", "Element " + regElement.ElementID.ToString(PRZH.CurrentGeodatabaseElementTableNameFormat));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await PRZH.RunGPTool("CreateTable_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error creating the {regElement.ElementTable} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error creating the {regElement.ElementTable} table.");
                        return (false, "error creating table.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Created the {regElement.ElementTable} table."), true, ++val);
                    }

                    PRZH.CheckForCancellation(token);

                    // Add fields to the table
                    string fldPUID = PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID + " LONG 'Planning Unit ID' # 0 #;";
                    string fldCellNum = PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER + " LONG 'Cell Number' # 0 #;";
                    string fldCellVal = PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_VALUE + " DOUBLE 'Cell Value' # 0 #;";

                    string flds = fldPUID + fldCellNum + fldCellVal;

                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Adding fields to the {regElement.ElementTable} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(regElement.ElementTable, flds);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await PRZH.RunGPTool("AddFields_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error adding fields to the {regElement.ElementTable} table.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error adding fields to the {regElement.ElementTable} table.");
                        return (false, "error adding fields.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Fields added successfully."), true, ++val);
                    }

                    // Populate the table...
                    await QueuedTask.Run(() =>
                    {
                        var tryget_gdb = PRZH.GetGDB_Project();

                        using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                        using (Table table = geodatabase.OpenDataset<Table>(regElement.ElementTable))
                        using (RowBuffer rowBuffer = table.CreateRowBuffer())
                        {
                            geodatabase.ApplyEdits(() =>
                            {
                                foreach (int puid in DICT_PUID_and_value_sum.Keys) // PUIDs showing up in this regElement's zonal stats table
                                {
                                    // store the puid
                                    rowBuffer[PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID] = puid;

                                    // store the value
                                    rowBuffer[PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_VALUE] = DICT_PUID_and_value_sum[puid];

                                    // store the cell number (if present)
                                    if (DICT_PUID_and_cellnumbers.ContainsKey(puid))
                                    {
                                        rowBuffer[PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER] = DICT_PUID_and_cellnumbers[puid];
                                    }

                                    table.CreateRow(rowBuffer);
                                }
                            });
                        }
                    });

                    PRZH.CheckForCancellation(token);

                    // index the PUID field
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Indexing {PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID} field in the {regElement.ElementTable} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(regElement.ElementTable, PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID, "ix" + PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID, "", "");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
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

                    // index the Cell Number field
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Indexing {PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER} field in the {regElement.ElementTable} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(regElement.ElementTable, PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER, "ix" + PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER, "", "");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
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

                    // Update the regElements table
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Updating the {PRZC.c_TABLE_REGPRJ_ELEMENTS} table.."), true, ++val);
                    await QueuedTask.Run(() =>
                    {
                        var tryget_gdb = PRZH.GetGDB_Project();

                        using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                        using (Table table = geodatabase.OpenDataset<Table>(PRZC.c_TABLE_REGPRJ_ELEMENTS))
                        {
                            QueryFilter queryFilter = new QueryFilter()
                            {
                                WhereClause = $"{PRZC.c_FLD_TAB_REGELEMENT_ELEMENT_ID} = {regElement.ElementID}"
                            };

                            using (RowCursor rowCursor = table.Search(queryFilter, false))
                            {
                                if (rowCursor.MoveNext())
                                {
                                    using (Row row = rowCursor.Current)
                                    {
                                        geodatabase.ApplyEdits(() =>
                                        {
                                            row[PRZC.c_FLD_TAB_REGELEMENT_PRESENCE] = (int)ElementPresence.Present;
                                            row.Store();
                                        });
                                    }
                                }
                            }
                        }
                    });
                }

                #endregion

                // Remove Group Layer
                await QueuedTask.Run(() => { _map.RemoveLayer(GL); });

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
        }*/

        /*private async Task<(bool success, string message)> GenerateSpatialDatasets(CancellationToken token)
        {
            int val = PM.Current;
            int max = PM.Max;

            try
            {
                #region GET REGIONAL ELEMENT INFOS

                // Get list of regional element tables (e.g. r00010)
                var tryget_LIST_elemtables_reg = await PRZH.GetRegionalElementTables();
                if (!tryget_LIST_elemtables_reg.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving list of regional element tables.\n{tryget_LIST_elemtables_reg.message}", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving list of regional element tables.\n{tryget_LIST_elemtables_reg.message}");
                    return (false, "error retrieving reg table list.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Found {tryget_LIST_elemtables_reg.tables.Count} regional element tables"), true, ++val);
                }

                List<string> LIST_RegElemTables = tryget_LIST_elemtables_reg.tables;

                // If no regional element tables are found, return.
                if (LIST_RegElemTables.Count == 0)
                {
                    // there are no regional tables, so there is no spatial data to process
                    return (true, "no regional tables to process (this is OK).");
                }

                // ASSEMBLE LISTS OF REGIONAL ELEMENTS
                // Get All Reg Elements where presence = yes
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieving all present regional elements..."), true, ++val);
                var tryget_all = await PRZH.GetRegionalElements(null, null, ElementPresence.Present);
                if (!tryget_all.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional elements.\n{tryget_all.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving regional elements.\n{tryget_all.message}");
                    return (false, "error retrieving elements.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"{tryget_all.elements.Count} regional element(s) retrieved."), true, ++val);
                }

                List<RegElement> LIST_RegElements = tryget_all.elements;

                // Ensure at least one element in list
                if (LIST_RegElements.Count == 0)
                {
                    return (true, "no regional elements to process (this is OK).");
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
                string base_fc_path = PRZH.GetPath_Project(base_fc, PRZC.c_FDS_REGIONAL_ELEMENTS).path;

                // Get the Planning Unit SR (from the PU FC)
                SpatialReference PlanningUnitSR = await QueuedTask.Run(() =>
                {
                    var tryget_fc = PRZH.GetFC_Project(PRZC.c_FC_PLANNING_UNITS);

                    using (FeatureClass featureClass = tryget_fc.featureclass)
                    using (FeatureClassDefinition fcDef = featureClass.GetDefinition())
                    {
                        return fcDef.GetSpatialReference();
                    }
                });

                // Copy the Planning Units FC into reg fds
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

                #region CYCLE THROUGH ALL REGIONAL ELEMENTS

                // LOOP THROUGH ELEMENTS
                for (int i = 0; i < LIST_RegElements.Count; i++)
                {
                    // Get the element
                    RegElement element = LIST_RegElements[i];

                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Processing Element #{element.ElementID}: {element.ElementName}"), true, ++val);

                    var tryget_elemtablename = PRZH.GetRegionalElementTableName(element.ElementID);
                    if (!tryget_elemtablename.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving table name for reg element {element.ElementID}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Error retrieving reg element table name.");
                        return (false, "error retrieving reg element table name.");
                    }
                    string table_name = tryget_elemtablename.table_name;

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
                    string elem_fc_path = PRZH.GetPath_Project(elem_fc_name, PRZC.c_FDS_REGIONAL_ELEMENTS).path;

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
                    toolParams = Geoprocessing.MakeValueArray(elem_fc_name, PRZC.c_FLD_FC_PU_ID, table_name, PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID);
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
                                table.DeleteRows(new QueryFilter { WhereClause = $"{PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID}_1 IS NULL" });
                            });
                        }
                    });
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting unjoined rows..."), true, ++val);

                    PRZH.CheckForCancellation(token);

                    // DELETE UNNECESSARY FIELD
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting extra id field..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(elem_fc_name, $"{PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID}_1", "DELETE_FIELDS");
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
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Indexing {PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER} field in {elem_fc_name} table..."), true, ++val);
                    toolParams = Geoprocessing.MakeValueArray(elem_fc_name, PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER, "ix" + PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER, "", "");
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
        }*/

        private async Task ValidateControls()
        {
            try
            {
                // Planning Units existence
                _pu_exists = (await PRZH.PUDataExists()).exists;

                if (_pu_exists)
                {
                    CompStat_Txt_PlanningUnits_Label = "Planning Units exist.";
                    CompStat_Img_PlanningUnits_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_Yes16.png";
                }
                else
                {
                    CompStat_Txt_PlanningUnits_Label = "Planning Units do not exist. Build them.";
                    CompStat_Img_PlanningUnits_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_No16.png";
                }

                // Regional Directory existence
                _regdb_exists = (await PRZH.GDBExists_Reg()).exists;

                if (_regdb_exists)
                {
                    CompStat_Txt_RegDB_Label = $"Regional Data Folder found at path: {PRZH.GetPath_RegionalDataFolder()}";
                    CompStat_Img_RegDB_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_Yes16.png";
                }
                else
                {
                    CompStat_Txt_RegDB_Label = $"Regional Data Folder not found at path: {PRZH.GetPath_RegionalDataFolder()}";
                    CompStat_Img_RegDB_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_No16.png";
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
            Operation_Cmd_IsEnabled = _pu_exists && _regdb_exists;
            OpStat_Img_Visibility = Visibility.Hidden;
            OpStat_Txt_Label = "Idle";
            _operationIsUnderway = false;
        }

        #endregion



    }
}
