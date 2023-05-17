using ArcGIS.Core.Data;
using ArcGIS.Core.Data.Raster;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ProMsgBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using PRZC = NCC.PRZTools.PRZConstants;
using PRZH = NCC.PRZTools.PRZHelper;
using YamlDotNet.Serialization;
using CsvHelper;
using CsvHelper.Configuration;
using System.Text;
using System.Collections.Concurrent;
using ArcGIS.Desktop.Internal.GeoProcessing;

namespace NCC.PRZTools
{
    public class ExportWTWVM : PropertyChangedBase
    {
        public ExportWTWVM()
        {
        }

        #region FIELDS

        private CancellationTokenSource _cts = null;
        private ProgressManager _pm = ProgressManager.CreateProgressManager(50);    // initialized to min=0, current=0, message=""
        private bool _operation_Cmd_IsEnabled;
        private bool _operationIsUnderway = false;
        private Cursor _proWindowCursor;

        private bool _pu_exists = false;
        private readonly SpatialReference Export_SR = SpatialReferences.WGS84;

        #region COMMANDS

        private ICommand _cmdExport;
        private ICommand _cmdCancel;
        private ICommand _cmdClearLog;

        #endregion

        #region COMPONENT STATUS INDICATORS

        // Planning Unit Dataset
        private string _compStat_Img_PlanningUnits_Path;
        private string _compStat_Txt_PlanningUnits_Label;

        // Boundary Lengths Table
        private string _compStat_Img_BoundaryLengths_Path;
        private string _compStat_Txt_BoundaryLengths_Label;

        #endregion

        #region OPERATION STATUS INDICATORS

        private Visibility _opStat_Img_Visibility = Visibility.Collapsed;
        private string _opStat_Txt_Label;

        #endregion

        #endregion

        #region PROPERTIES

        public ProgressManager PM
        {
            get => _pm;
            set => SetProperty(ref _pm, value, () => PM);
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

        // Boundary Lengths Table
        public string CompStat_Img_BoundaryLengths_Path
        {
            get => _compStat_Img_BoundaryLengths_Path;
            set => SetProperty(ref _compStat_Img_BoundaryLengths_Path, value, () => CompStat_Img_BoundaryLengths_Path);
        }

        public string CompStat_Txt_BoundaryLengths_Label
        {
            get => _compStat_Txt_BoundaryLengths_Label;
            set => SetProperty(ref _compStat_Txt_BoundaryLengths_Label, value, () => CompStat_Txt_BoundaryLengths_Label);
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

        #endregion

        #region COMMANDS

        public ICommand CmdClearLog => _cmdClearLog ?? (_cmdClearLog = new RelayCommand(() =>
        {
            PRZH.UpdateProgress(PM, "", false, 0, 1, 0);
        }, () => true, true, false));

        public ICommand CmdExport => _cmdExport ?? (_cmdExport = new RelayCommand(async () =>
        {
            // Change UI to Underway
            StartOpUI();

            // Start the operation
            using (_cts = new CancellationTokenSource())
            {
                await ExportWTWPackage(_cts.Token);
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

        #endregion

        #region METHODS

        public async Task OnProWinLoaded()
        {
            try
            {
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

        private async Task ExportWTWPackage(CancellationToken token)
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
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Unable to enabled editing for this ArcGIS Pro Project.", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show("Unable to enabled editing for this ArcGIS Pro Project.");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("ArcGIS Pro editing enabled."), true, ++val);
                    }
                }

                #endregion

                // Initialize a few objects and names
                string gdbpath = PRZH.GetPath_ProjectGDB();
                string export_folder_path = PRZH.GetPath_ExportWTWFolder();

                // Initialize ProgressBar and Progress Log
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Initializing the Where to Work Exporter..."), false, max, ++val);

                // Ensure the Project Geodatabase Exists
                var try_gdbexists = await PRZH.GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Project Geodatabase not found: '{gdbpath}'.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Project Geodatabase not found at {gdbpath}.");
                    return;
                }

                // Ensure the ExportWTW folder exists
                if (!PRZH.FolderExists_ExportWTW().exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"{PRZC.c_DIR_EXPORT_WTW} folder not found in project workspace.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"{PRZC.c_DIR_EXPORT_WTW} folder not found in project workspace.");
                    return;
                }

                // Ensure the Planning Units data exists
                var tryget_pudata = await PRZH.PUDataExists();
                if (!tryget_pudata.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Planning Units datasets not found.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Planning Units datasets not found.");
                    return;
                }

                #region VALIDATE NATIONAL AND REGIONAL ELEMENT DATA

                // Check for national element tables
                var tryread_national_element_tiles = PRZH.ReadBinary(PRZH.GetPath_ProjectNationalElementTilesMetadataPath());

                int nattables_present = tryread_national_element_tiles.success ? ((Dictionary<string, HashSet<int>>)tryread_national_element_tiles.obj).Count : 0;
                int regtables_present = 0; //TODO: Decide where to store reg element tables, update assignment accordingly

                if (nattables_present == 0 & regtables_present == 0)
                {
                    // there are no national or regional element tables, stop!
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"No national or regional elements intersect study area.  Unable to proceed.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"No national or regional elements intersect study area.  Please expand your study area or ensure your national or regional databases include data for your study area.");
                    return;
                }
                else
                {
                    string m = $"National element tables: {nattables_present}\nRegional element tables: {regtables_present}";
                    PRZH.UpdateProgress(PM, PRZH.WriteLog(m), true, ++val);
                }

                #endregion

                if (ProMsgBox.Show($"If you proceed, all files in the WTW folder will be overwritten." +
                    Environment.NewLine + Environment.NewLine +
                    $"Do you wish to proceed?",
                    "Overwrite WTW files?",
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

                PRZH.CheckForCancellation(token);

                #region DELETE OBJECTS

                // Delete all existing files within export dir
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting existing files..."), true, ++val);
                DirectoryInfo di = new DirectoryInfo(export_folder_path);

                try
                {
                    foreach (FileInfo fi in di.GetFiles())
                    {
                        fi.Delete();
                    }

                    foreach (DirectoryInfo sdi in di.GetDirectories())
                    {
                        sdi.Delete(true);
                    }
                }
                catch (Exception ex)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting files and folder from {export_folder_path}.\n{ex.Message}", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Unable to delete files & subfolders in the {export_folder_path} folder.\n{ex.Message}");
                    return;
                }
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Existing files deleted."), true, ++val);

                #endregion

                PRZH.CheckForCancellation(token);

                #region EXPORT SPATIAL DATA

                // Export the raster to TIFF
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Exporting raster spatial data."), true, ++val);
                var tryexport_raster = await ExportRasterToTiff(token);
                if (!tryexport_raster.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error exporting {PRZC.c_RAS_PLANNING_UNITS} raster dataset to TIFF.\n{tryexport_raster.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error exporting {PRZC.c_RAS_PLANNING_UNITS} raster dataset to TIFF.\n{tryexport_raster.message}");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done exporting spatial data."), true, ++val);
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region Build and Export Boundary Lengths Table

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Building and exporting boundary table"), true, ++val);
                var trybuild_boundary = await BuildBoundaryLengthsTable(token);
                if (!trybuild_boundary.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error building the boundary lengths table\n{trybuild_boundary.message}", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"error building the boundary lengths table.\n{trybuild_boundary.message}");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Boundary table export complete."), true, ++val);
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region Build and Export Attribute Table

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Building and exporting attribute table"), true, ++val);
                var trybuild_attribute = await BuildAttributeTable(token);
                if (!trybuild_attribute.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error building the attribute table\n{trybuild_attribute.message}", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"error building the attribute table.\n{trybuild_attribute.message}");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Attribute table export complete."), true, ++val);
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region Build and Export YAML File 

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Building and writing YAML file..."), true, ++val);

                #region GET NATIONAL TABLE CONTENTS

                // Prepare the empty lists (there may be no national data)
                List<NatTheme> nat_themes = new List<NatTheme>();
                List<NatElement> nat_goals = new List<NatElement>();
                List<NatElement> nat_weights = new List<NatElement>();
                List<NatElement> nat_includes = new List<NatElement>();
                List<NatElement> nat_excludes = new List<NatElement>();

                // If there's at least a single table, populate the lists
                if (nattables_present > 0)
                {
                    // Get the National Themes
                    var theme_outcome = await PRZH.GetNationalThemes(ElementPresence.Present);
                    if (!theme_outcome.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national themes.\n{theme_outcome.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving national themes.\n{theme_outcome.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {theme_outcome.themes.Count} national themes."), true, ++val);
                    }
                    nat_themes = theme_outcome.themes;

                    // Get the goals
                    var goal_outcome = await PRZH.GetNationalElements(ElementType.Goal, ElementStatus.Active, ElementPresence.Present);
                    if (!goal_outcome.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national {ElementType.Goal} elements.\n{goal_outcome.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving national {ElementType.Goal} elements.\n{goal_outcome.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {goal_outcome.elements.Count} national {ElementType.Goal} elements."), true, ++val);
                    }
                    nat_goals = goal_outcome.elements;

                    // Get the weights
                    var weight_outcome = await PRZH.GetNationalElements(ElementType.Weight, ElementStatus.Active, ElementPresence.Present);
                    if (!weight_outcome.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national {ElementType.Weight} elements.\n{weight_outcome.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving national {ElementType.Weight} elements.\n{weight_outcome.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {weight_outcome.elements.Count} national {ElementType.Weight} elements."), true, ++val);
                    }
                    nat_weights = weight_outcome.elements;

                    // Get the includes
                    var include_outcome = await PRZH.GetNationalElements(ElementType.Include, ElementStatus.Active, ElementPresence.Present);
                    if (!include_outcome.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national {ElementType.Include} elements.\n{include_outcome.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving national {ElementType.Include} elements.\n{include_outcome.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {include_outcome.elements.Count} national {ElementType.Include} elements."), true, ++val);
                    }
                    nat_includes = include_outcome.elements;

                    // Get the excludes
                    var exclude_outcome = await PRZH.GetNationalElements(ElementType.Exclude, ElementStatus.Active, ElementPresence.Present);
                    if (!exclude_outcome.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national {ElementType.Exclude} elements.\n{exclude_outcome.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving national {ElementType.Exclude} elements.\n{exclude_outcome.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {exclude_outcome.elements.Count} national {ElementType.Exclude} elements."), true, ++val);
                    }
                    nat_excludes = exclude_outcome.elements;
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region GET REGIONAL TABLE CONTENTS

                // Prepare the empty lists (there may be no regional data)
                Dictionary<int, string> DICT_RegThemes = new Dictionary<int, string>();
                List<RegElement> reg_goals = new List<RegElement>();
                List<RegElement> reg_weights = new List<RegElement>();
                List<RegElement> reg_includes = new List<RegElement>();
                List<RegElement> reg_excludes = new List<RegElement>();

                // If there's at least a single table, populate the lists
                if (regtables_present > 0)
                {
                    // Get the Regional Themes
                    var tryget_regThemes = await PRZH.GetRegionalThemesDomainKVPs();
                    if (!tryget_regThemes.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional themes.\n{tryget_regThemes.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving regional themes.\n{tryget_regThemes.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved regional themes."), true, ++val);
                    }
                    DICT_RegThemes = tryget_regThemes.dict;

                    // Get the goals
                    var tryget_reg_goals = await PRZH.GetRegionalElements(ElementType.Goal, ElementStatus.Active, ElementPresence.Present);
                    if (!tryget_reg_goals.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional {ElementType.Goal} elements.\n{tryget_reg_goals.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving regional {ElementType.Goal} elements.\n{tryget_reg_goals.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {tryget_reg_goals.elements.Count} regional {ElementType.Goal} elements."), true, ++val);
                    }
                    reg_goals = tryget_reg_goals.elements;

                    // Get the weights
                    var tryget_reg_weights = await PRZH.GetRegionalElements(ElementType.Weight, ElementStatus.Active, ElementPresence.Present);
                    if (!tryget_reg_weights.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional {ElementType.Weight} elements.\n{tryget_reg_weights.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving national {ElementType.Weight} elements.\n{tryget_reg_weights.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {tryget_reg_weights.elements.Count} regional {ElementType.Weight} elements."), true, ++val);
                    }
                    reg_weights = tryget_reg_weights.elements;

                    // Get the includes
                    var tryget_reg_includes = await PRZH.GetRegionalElements(ElementType.Include, ElementStatus.Active, ElementPresence.Present);
                    if (!tryget_reg_includes.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional {ElementType.Include} elements.\n{tryget_reg_includes.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving regional {ElementType.Include} elements.\n{tryget_reg_includes.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {tryget_reg_includes.elements.Count} regional {ElementType.Include} elements."), true, ++val);
                    }
                    reg_includes = tryget_reg_includes.elements;

                    // Get the excludes
                    var tryget_reg_excludes = await PRZH.GetRegionalElements(ElementType.Exclude, ElementStatus.Active, ElementPresence.Present);
                    if (!tryget_reg_excludes.success)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving regional {ElementType.Exclude} elements.\n{tryget_reg_excludes.message}", LogMessageType.ERROR), true, ++val);
                        ProMsgBox.Show($"Error retrieving regional {ElementType.Exclude} elements.\n{tryget_reg_excludes.message}");
                        return;
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {tryget_reg_excludes.elements.Count} regional {ElementType.Exclude} elements."), true, ++val);
                    }
                    reg_excludes = tryget_reg_excludes.elements;
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region Populate and write YAML Objects

                #region THEMES & GOALS

                SortedList<int, string> SLIST_Themes = new SortedList<int, string>();

                // Add the national themes
                for (int i = 0; i < nat_themes.Count; i++)
                {
                    NatTheme natTheme = nat_themes[i];

                    int theme_id = natTheme.ThemeID;
                    string theme_name = natTheme.ThemeName;

                    if (!SLIST_Themes.ContainsKey(theme_id))
                    {
                        SLIST_Themes.Add(theme_id, theme_name);
                    }
                }

                // Add the regional themes
                foreach (var regTheme in DICT_RegThemes)
                {
                    int theme_id = regTheme.Key;
                    string theme_name = regTheme.Value;

                    if (!SLIST_Themes.ContainsKey(theme_id))
                    {
                        SLIST_Themes.Add(theme_id, theme_name);
                    }
                }

                // Create the yamlTheme list
                List<YamlTheme> yamlThemes = new List<YamlTheme>();

                // Iterate through all themes (nat + reg)
//                for (int i = 0; i < nat_themes.Count; i++)
                foreach (var themeKVP in SLIST_Themes)
                {
                    // Get national goals with this theme
                    List<NatElement> nat_theme_goals = nat_goals.Where(g => g.ThemeID == themeKVP.Key).OrderBy(g => g.ElementID).ToList();

                    // Get all regional goals with this theme
                    List<RegElement> reg_theme_goals = reg_goals.Where(g => g.ThemeID == themeKVP.Key).OrderBy(g => g.ElementID).ToList();

                    if (nat_theme_goals.Count + reg_theme_goals.Count == 0)
                    {
                        // no goals (nat + reg) for this theme
                        continue;
                    }

                    // Assemble the yaml Goal list
                    List<YamlFeature> yamlGoals = new List<YamlFeature>();

                    // First, national goals
                    for (int j = 0; j < nat_theme_goals.Count; j++)
                    {
                        // Get the goal element
                        NatElement goal = nat_theme_goals[j];

                        // Build the Yaml Legend
                        YamlLegend yamlLegend = new YamlLegend();

                        List<Color> colors = new List<Color>()
                        {
                            Color.Transparent,
                            Color.DarkSeaGreen
                        };

                        yamlLegend.SetContinuousColors(colors);

                        // Build the Yaml Variable
                        YamlVariable yamlVariable = new YamlVariable();
                        yamlVariable.index = goal.ElementTable;
                        yamlVariable.units = goal.ElementUnit;
                        yamlVariable.provenance = WTWProvenanceType.national.ToString();
                        yamlVariable.legend = yamlLegend;

                        // Build the Yaml Goal
                        YamlFeature yamlGoal = new YamlFeature();
                        yamlGoal.name = goal.ElementName;
                        yamlGoal.status = true; // enabled or disabled
                        yamlGoal.visible = true;
                        yamlGoal.hidden = false;
                        yamlGoal.goal = 0.5;        // needs to be retrieved from somewhere, or just left to 0.5
                        yamlGoal.variable = yamlVariable;

                        // Add to list
                        yamlGoals.Add(yamlGoal);
                    }

                    // Next, regional goals
                    for (int j = 0; j < reg_theme_goals.Count; j++)
                    {
                        // Get the goal element
                        RegElement goal = reg_theme_goals[j];

                        // Build the Yaml Legend
                        YamlLegend yamlLegend = new YamlLegend();

                        List<Color> colors = new List<Color>()
                        {
                            Color.Transparent,
                            Color.MediumPurple
                        };

                        yamlLegend.SetContinuousColors(colors);

                        // Build the Yaml Variable
                        YamlVariable yamlVariable = new YamlVariable();
                        yamlVariable.index = goal.ElementTable;
                        yamlVariable.units = "reg";// goal.ElementUnit;
                        yamlVariable.provenance = WTWProvenanceType.regional.ToString();
                        yamlVariable.legend = yamlLegend;

                        // Build the Yaml Goal
                        YamlFeature yamlGoal = new YamlFeature();
                        yamlGoal.name = goal.ElementName;
                        yamlGoal.status = true; // enabled or disabled
                        yamlGoal.visible = true;
                        yamlGoal.hidden = false;
                        yamlGoal.goal = 0.5;        // needs to be retrieved from somewhere, or just left to 0.5
                        yamlGoal.variable = yamlVariable;

                        // Add to list
                        yamlGoals.Add(yamlGoal);
                    }

                    // Create the Yaml Theme
                    YamlTheme yamlTheme = new YamlTheme();

                    yamlTheme.name = themeKVP.Value;
                    yamlTheme.feature = yamlGoals.ToArray();

                    // Add to list
                    yamlThemes.Add(yamlTheme);
                }

                #endregion

                #region WEIGHTS

                // Create the yaml Weights list
                List<YamlWeight> yamlWeights = new List<YamlWeight>();

                // National Weights
                for (int i = 0; i < nat_weights.Count; i++)
                {
                    // Get the weight
                    NatElement weight = nat_weights[i];

                    // Build the Yaml Legend
                    YamlLegend yamlLegend = new YamlLegend();
                    List<Color> colors = new List<Color>()
                    {
                        Color.White,
                        Color.DarkOrchid
                    };

                    yamlLegend.SetContinuousColors(colors);

                    // Build the Yaml Variable
                    YamlVariable yamlVariable = new YamlVariable();
                    yamlVariable.index = weight.ElementTable;
                    yamlVariable.units = weight.ElementUnit;
                    yamlVariable.provenance = WTWProvenanceType.national.ToString();
                    yamlVariable.legend = yamlLegend;

                    // Build the Yaml Weight
                    YamlWeight yamlWeight = new YamlWeight();
                    yamlWeight.name = weight.ElementName;
                    yamlWeight.status = true; // enabled or disabled
                    yamlWeight.visible = true;
                    yamlWeight.hidden = false;
                    yamlWeight.factor = 0;                  // what's this?
                    yamlWeight.variable = yamlVariable;

                    // Add to list
                    yamlWeights.Add(yamlWeight);
                }

                // Regional Weights
                for (int i = 0; i < reg_weights.Count; i++)
                {
                    // Get the weight
                    RegElement weight = reg_weights[i];

                    // Build the Yaml Legend
                    YamlLegend yamlLegend = new YamlLegend();
                    List<Color> colors = new List<Color>()
                    {
                        Color.White,
                        Color.DarkOrchid
                    };

                    yamlLegend.SetContinuousColors(colors);

                    // Build the Yaml Variable
                    YamlVariable yamlVariable = new YamlVariable();
                    yamlVariable.index = weight.ElementTable;
                    yamlVariable.units = "reg";
                    yamlVariable.provenance = WTWProvenanceType.regional.ToString();
                    yamlVariable.legend = yamlLegend;

                    // Build the Yaml Weight
                    YamlWeight yamlWeight = new YamlWeight();
                    yamlWeight.name = weight.ElementName;
                    yamlWeight.status = true; // enabled or disabled
                    yamlWeight.visible = true;
                    yamlWeight.hidden = false;
                    yamlWeight.factor = 0;                  // what's this?
                    yamlWeight.variable = yamlVariable;

                    // Add to list
                    yamlWeights.Add(yamlWeight);
                }

                #endregion

                #region INCLUDES

                // Create the yaml Includes list
                List<YamlInclude> yamlIncludes = new List<YamlInclude>();

                // National Includes
                for (int i = 0; i < nat_includes.Count; i++)
                {
                    // Get the include
                    NatElement include = nat_includes[i];

                    // Build the Yaml Legend
                    YamlLegend yamlLegend = new YamlLegend();
                    List<(Color color, string label)> values = new List<(Color color, string label)>()
                        {
                            (Color.Transparent, "Do not include"),
                            (Color.Green, "Include")
                        };

                    yamlLegend.SetManualColors(values);

                    // Build the Yaml Variable
                    YamlVariable yamlVariable = new YamlVariable();
                    yamlVariable.index = include.ElementTable;
                    yamlVariable.units = "";//include.ElementUnit;
                    yamlVariable.provenance = WTWProvenanceType.national.ToString();
                    yamlVariable.legend = yamlLegend;

                    // Build the Yaml Include
                    YamlInclude yamlInclude = new YamlInclude();
                    yamlInclude.name = include.ElementName;
                    yamlInclude.mandatory = false;      // what's this
                    yamlInclude.status = true; // enabled or disabled
                    yamlInclude.visible = true;
                    yamlInclude.hidden = false;
                    yamlInclude.variable = yamlVariable;

                    // Add to list
                    yamlIncludes.Add(yamlInclude);
                }

                // Regional Includes
                for (int i = 0; i < reg_includes.Count; i++)
                {
                    // Get the include
                    RegElement include = reg_includes[i];

                    // Build the Yaml Legend
                    YamlLegend yamlLegend = new YamlLegend();
                    List<(Color color, string label)> values = new List<(Color color, string label)>()
                        {
                            (Color.Transparent, "Do not include"),
                            (Color.Green, "Include")
                        };

                    yamlLegend.SetManualColors(values);

                    // Build the Yaml Variable
                    YamlVariable yamlVariable = new YamlVariable();
                    yamlVariable.index = include.ElementTable;
                    yamlVariable.units = "reg";//include.ElementUnit;
                    yamlVariable.provenance = WTWProvenanceType.regional.ToString();
                    yamlVariable.legend = yamlLegend;

                    // Build the Yaml Include
                    YamlInclude yamlInclude = new YamlInclude();
                    yamlInclude.name = include.ElementName;
                    yamlInclude.mandatory = false;      // what's this
                    yamlInclude.status = true; // enabled or disabled
                    yamlInclude.visible = true;
                    yamlInclude.hidden = false;
                    yamlInclude.variable = yamlVariable;

                    // Add to list
                    yamlIncludes.Add(yamlInclude);
                }

                #endregion

                #region EXCLUDES

                // Create the yaml Excludes list
                List<YamlExclude> yamlExcludes = new List<YamlExclude>();

                // National Excludes
                for (int i = 0; i < nat_excludes.Count; i++)
                {
                    // Get the exclude
                    NatElement exclude = nat_excludes[i];

                    // Build the Yaml Legend
                    YamlLegend yamlLegend = new YamlLegend();       // default legend

                    // Build the Yaml Variable
                    YamlVariable yamlVariable = new YamlVariable();
                    yamlVariable.index = exclude.ElementTable;
                    yamlVariable.units = exclude.ElementUnit;
                    yamlVariable.provenance = WTWProvenanceType.national.ToString();
                    yamlVariable.legend = yamlLegend;

                    // Build the Yaml Exclude
                    YamlExclude yamlExclude = new YamlExclude();
                    yamlExclude.name = exclude.ElementName;
                    yamlExclude.mandatory = false;      // what's this
                    yamlExclude.status = true; // enabled or disabled
                    yamlExclude.visible = true;
                    yamlExclude.hidden = false;
                    yamlExclude.variable = yamlVariable;

                    // Add to list
                    yamlExcludes.Add(yamlExclude);
                }

                // Regional Excludes
                for (int i = 0; i < reg_excludes.Count; i++)
                {
                    // Get the exclude
                    RegElement exclude = reg_excludes[i];

                    // Build the Yaml Legend
                    YamlLegend yamlLegend = new YamlLegend();       // default legend

                    // Build the Yaml Variable
                    YamlVariable yamlVariable = new YamlVariable();
                    yamlVariable.index = exclude.ElementTable;
                    yamlVariable.units = "reg"; // exclude.ElementUnit;
                    yamlVariable.provenance = WTWProvenanceType.regional.ToString();
                    yamlVariable.legend = yamlLegend;

                    // Build the Yaml Exclude
                    YamlExclude yamlExclude = new YamlExclude();
                    yamlExclude.name = exclude.ElementName;
                    yamlExclude.mandatory = false;      // what's this
                    yamlExclude.status = true; // enabled or disabled
                    yamlExclude.visible = true;
                    yamlExclude.hidden = false;
                    yamlExclude.variable = yamlVariable;

                    // Add to list
                    yamlExcludes.Add(yamlExclude);
                }

                #endregion

                #region YAML PACKAGE

                YamlPackage yamlPackage = new YamlPackage();
                yamlPackage.name = System.IO.Path.GetFileNameWithoutExtension(CoreModule.CurrentProject.Name);
                yamlPackage.mode = WTWModeType.advanced.ToString();     // Which of these should I use?
                yamlPackage.themes = yamlThemes.ToArray();
                yamlPackage.weights = yamlWeights.ToArray();
                yamlPackage.includes = yamlIncludes.ToArray();
                // yamlPackage.excludes = yamlExcludes.ToArray();       // Excludes are not part of the yaml schema (yet)

                ISerializer builder = new SerializerBuilder().DisableAliases().Build();
                string the_yaml = builder.Serialize(yamlPackage);

                string yamlpath = Path.Combine(export_folder_path, PRZC.c_FILE_WTW_EXPORT_YAML);
                try
                {
                    File.WriteAllText(yamlpath, the_yaml);
                }
                catch (Exception ex)
                {
                    ProMsgBox.Show("Unable to write the Yaml Config File..." + Environment.NewLine + Environment.NewLine + ex.Message);
                    return;
                }

                #endregion

                #endregion

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done generating YAML file."), true, ++val);

                #endregion

                // End timer
                stopwatch.Stop();
                string message = PRZH.GetElapsedTimeMessage(stopwatch.Elapsed);
                PRZH.UpdateProgress(PM, PRZH.WriteLog("WTW export completed successfully!"), true, 1, 1);
                PRZH.UpdateProgress(PM, PRZH.WriteLog(message), true, 1, 1);
                if (yamlThemes.Count > 0)
                {
                    ProMsgBox.Show("WTW Export Completed Successfully!" + Environment.NewLine + Environment.NewLine + message);
                }
                else
                {
                    ProMsgBox.Show("No themes found for your study area! At least one theme is required to use the Where to Work application. Please check your study area and data sources." + Environment.NewLine + Environment.NewLine + message);
                }
            }
            catch (OperationCanceledException)
            {
                // Cancelled by user
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"ExportWTWPackage: cancelled by user.", LogMessageType.CANCELLATION), true, ++val);
                ProMsgBox.Show($"ExportWTWPackage: Cancelled by user.");
            }
            catch (Exception ex)
            {
                PRZH.UpdateProgress(PM, PRZH.WriteLog(ex.Message, LogMessageType.ERROR), true, ++val);
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return;
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

        private async Task<(bool success, string message)> ExportRasterToShapefile(CancellationToken token)
        {
            int val = 0;

            try
            {
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Creating shapefile..."), true, 30, ++val);

                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                string toolOutput;

                // Filenames and Paths
                string gdbpath = PRZH.GetPath_ProjectGDB();

                string export_folder_path = PRZH.GetPath_ExportWTWFolder();
                string export_shp_name = PRZC.c_FILE_WTW_EXPORT_SPATIAL + ".shp";
                string export_shp_path = Path.Combine(export_folder_path, export_shp_name);

                // Confirm that source raster is present
                if (!(await PRZH.RasterExists_Project(PRZC.c_RAS_PLANNING_UNITS)).exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"{PRZC.c_RAS_PLANNING_UNITS} raster dataset not found.", LogMessageType.ERROR), true, ++val);
                    return (false, $"{PRZC.c_RAS_PLANNING_UNITS} raster dataset not found");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"{PRZC.c_RAS_PLANNING_UNITS} raster found."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Convert source raster to temp polygon feature class
                string fldPUID_Temp = "gridcode";

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Converting {PRZC.c_RAS_PLANNING_UNITS} raster dataset to polygon feature class..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_RAS_PLANNING_UNITS, PRZC.c_FC_TEMP_WTW_FC1, "NO_SIMPLIFY", "VALUE", "SINGLE_OUTER_PART", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("RasterToPolygon_conversion", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error executing Raster To Polygon tool.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error executing Raster to Polygon tool.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Conversion successful."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Project temp polygon feature class
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Projecting {PRZC.c_FC_TEMP_WTW_FC1} feature class..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC1, PRZC.c_FC_TEMP_WTW_FC2, Export_SR, "", "", "NO_PRESERVE_SHAPE", "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("Project_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error projecting feature class.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error projecting feature class.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Projection successful."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Repair Geometry
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Repairing geometry..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                toolOutput = await PRZH.RunGPTool("RepairGeometry_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error repairing geometry.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error repairing geometry.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Geometry repaired."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Delete the unnecessary fields
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting extra fields..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2, fldPUID_Temp, "KEEP_FIELDS");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                toolOutput = await PRZH.RunGPTool("DeleteField_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting fields.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error deleting fields.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Fields deleted."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Calculate field
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Calculating {PRZC.c_FLD_FC_PU_ID} field..."), true, ++val);
                string expression = "!" + fldPUID_Temp + "!";
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2, PRZC.c_FLD_FC_PU_ID, expression, "PYTHON3", "", "LONG", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                toolOutput = await PRZH.RunGPTool("CalculateField_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error Calculating {PRZC.c_FLD_FC_PU_ID} field.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, $"Error calculating the new {PRZC.c_FLD_FC_PU_ID} field.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Field calculated successfully."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Export to Shapefile
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Export the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} shapefile..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2, export_shp_path);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("CopyFeatures_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error exporting the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} shapefile.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, $"Error exporting the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} shapefile.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Shapefile exported."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Index the new id field
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Indexing {PRZC.c_FLD_FC_PU_ID} field..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FILE_WTW_EXPORT_SPATIAL, new List<string>() { PRZC.c_FLD_FC_PU_ID }, "ix" + PRZC.c_FLD_FC_PU_ID, "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: export_folder_path);
                toolOutput = await PRZH.RunGPTool("AddIndex_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error indexing field.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error indexing field.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Field indexed."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Delete the unnecessary fields
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting extra fields (again)..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FILE_WTW_EXPORT_SPATIAL, PRZC.c_FLD_FC_PU_ID, "KEEP_FIELDS");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: export_folder_path);
                toolOutput = await PRZH.RunGPTool("DeleteField_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting fields.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error deleting fields.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Fields deleted."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Delete temp feature classes
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting {PRZC.c_FC_TEMP_WTW_FC1} and {PRZC.c_FC_TEMP_WTW_FC2} feature classes..."), true, ++val);

                if ((await PRZH.FCExists_Project(PRZC.c_FC_TEMP_WTW_FC1)).exists)
                {
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC1);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting {PRZC.c_FC_TEMP_WTW_FC1} feature class.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        return (false, $"Error deleting {PRZC.c_FC_TEMP_WTW_FC1} feature class.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Feature class deleted."), true, ++val);
                    }
                }

                PRZH.CheckForCancellation(token);

                if ((await PRZH.FCExists_Project(PRZC.c_FC_TEMP_WTW_FC2)).exists)
                {
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting {PRZC.c_FC_TEMP_WTW_FC2} feature class.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        return (false, $"Error deleting {PRZC.c_FC_TEMP_WTW_FC2} feature class.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Feature class deleted."), true, ++val);
                    }
                }

                return (true, "success");
            }
            catch (OperationCanceledException cancelex)
            {
                // Cancelled by user
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"ExportRasterToShapefile: cancelled by user.", LogMessageType.CANCELLATION), true, ++val);

                // Throw the cancellation error to the parent
                throw cancelex;
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private async Task<(bool success, string message)> ExportFeatureClassToShapefile(CancellationToken token)
        {
            int val = 0;

            try
            {
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Creating shapefile..."), true, ++val);

                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                string toolOutput;

                // Filenames and Paths
                string gdbpath = PRZH.GetPath_ProjectGDB();

                string export_folder_path = PRZH.GetPath_ExportWTWFolder();
                string export_shp_name = PRZC.c_FILE_WTW_EXPORT_SPATIAL + ".shp";
                string export_shp_path = Path.Combine(export_folder_path, export_shp_name);

                // Project feature class
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Projecting {PRZC.c_FC_PLANNING_UNITS} feature class..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_PLANNING_UNITS, PRZC.c_FC_TEMP_WTW_FC2, Export_SR, "", "", "NO_PRESERVE_SHAPE", "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(
                    workspace: gdbpath,
                    overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("Project_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error projecting feature class.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error projecting feature class.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Projection successful."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Repair Geometry
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Repairing geometry..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                toolOutput = await PRZH.RunGPTool("RepairGeometry_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error repairing geometry.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error repairing geometry.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Geometry repaired."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Delete the unnecessary fields
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting extra fields..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2, PRZC.c_FLD_FC_PU_ID, "KEEP_FIELDS");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                toolOutput = await PRZH.RunGPTool("DeleteField_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting fields.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error deleting fields.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Fields deleted."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Export to Shapefile
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Export the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} shapefile..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2, export_shp_path);
                toolEnvs = Geoprocessing.MakeEnvironmentArray(
                    workspace: gdbpath,
                    overwriteoutput: true,
                    outputCoordinateSystem: Export_SR);
                toolOutput = await PRZH.RunGPTool("CopyFeatures_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error exporting the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} shapefile.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, $"Error exporting the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} shapefile.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Shapefile exported."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Index the new id field
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Indexing {PRZC.c_FLD_FC_PU_ID} field..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FILE_WTW_EXPORT_SPATIAL, new List<string>() { PRZC.c_FLD_FC_PU_ID }, "ix" + PRZC.c_FLD_FC_PU_ID, "", "");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: export_folder_path);
                toolOutput = await PRZH.RunGPTool("AddIndex_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error indexing field.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error indexing field.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Field indexed."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Delete the unnecessary fields
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Deleting extra fields (again)..."), true, ++val);
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_FILE_WTW_EXPORT_SPATIAL, PRZC.c_FLD_FC_PU_ID, "KEEP_FIELDS");
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: export_folder_path);
                toolOutput = await PRZH.RunGPTool("DeleteField_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Error deleting fields.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, "Error deleting fields.");
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog("Fields deleted."), true, ++val);
                }

                PRZH.CheckForCancellation(token);

                // Delete temp feature class
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Deleting {PRZC.c_FC_TEMP_WTW_FC2} feature class..."), true, ++val);
                if ((await PRZH.FCExists_Project(PRZC.c_FC_TEMP_WTW_FC2)).exists)
                {
                    toolParams = Geoprocessing.MakeValueArray(PRZC.c_FC_TEMP_WTW_FC2);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath);
                    toolOutput = await PRZH.RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error deleting {PRZC.c_FC_TEMP_WTW_FC2} feature class.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                        return (false, $"Error deleting {PRZC.c_FC_TEMP_WTW_FC2} feature class.");
                    }
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("Feature class deleted."), true, ++val);
                    }
                }

                return (true, "success");
            }
            catch (OperationCanceledException cancelex)
            {
                // Cancelled by user
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"ExportFeaturesToShapefile: cancelled by user.", LogMessageType.CANCELLATION), true, ++val);

                // Throw the cancellation error to the parent
                throw cancelex;
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private async Task<(bool success, string message)> ExportRasterToTiff(CancellationToken token)
        {
            int val = 0;

            try
            {
                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                string toolOutput;

                // Filenames and Paths
                string gdbpath = PRZH.GetPath_ProjectGDB();

                string export_folder_path = PRZH.GetPath_ExportWTWFolder();
                string export_tif_name = PRZC.c_FILE_WTW_EXPORT_SPATIAL + ".tif";
                string export_tif_path = Path.Combine(export_folder_path, export_tif_name);

                // Confirm that source raster is present
                if (!(await PRZH.RasterExists_Project(PRZC.c_RAS_PLANNING_UNITS)).exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"{PRZC.c_RAS_PLANNING_UNITS} raster not found.", LogMessageType.ERROR), true, ++val);
                    return (false, $"{PRZC.c_RAS_PLANNING_UNITS} raster not found");
                }

                PRZH.CheckForCancellation(token);

                // Export to Tiff
                toolParams = Geoprocessing.MakeValueArray(PRZC.c_RAS_PLANNING_UNITS, export_tif_path); // TODO: update layer being updated if copied and modified
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await PRZH.RunGPTool("CopyRaster_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error exporting the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} raster.  GP Tool failed or was cancelled by user.", LogMessageType.ERROR), true, ++val);
                    return (false, $"Error exporting the {PRZC.c_FILE_WTW_EXPORT_SPATIAL} raster.");
                }

                // Clean up unnecessary files
                string[] spatial_files = Directory.GetFiles(export_folder_path, $"{PRZC.c_FILE_WTW_EXPORT_SPATIAL}*");
                foreach (string spatial_file in spatial_files)
                {
                    if (spatial_file != export_tif_path) File.Delete(spatial_file);
                }

                PRZH.CheckForCancellation(token);

                return (true, "success");
            }
            catch (OperationCanceledException cancelex)
            {
                // Cancelled by user
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"ExportRasterToTiff: cancelled by user.", LogMessageType.CANCELLATION), true, ++val);

                // Throw the cancellation error to the parent
                throw cancelex;
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        private async Task<(bool success, string message)> BuildBoundaryLengthsTable(CancellationToken token)
        {
            int val = PM.Current;
            int max = PM.Max;

            try
            {
                #region INITIALIZATION

                // Get the Planning Unit Perimeter (from the raster)
                double side_length = await QueuedTask.Run(() =>
                {
                    var tryget_ras = PRZH.GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);

                    using (RasterDataset RD = tryget_ras.rasterDataset)
                    using (Raster raster = RD.CreateFullRaster())
                    {
                        var cell_size_tuple = raster.GetMeanCellSize();
                        return Math.Round(cell_size_tuple.Item1, 2, MidpointRounding.AwayFromZero);
                    }
                });

                int grid_width = await QueuedTask.Run(() =>
                {
                    var tryget_ras = PRZH.GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);

                    using (RasterDataset RD = tryget_ras.rasterDataset)
                    using (Raster raster = RD.CreateFullRaster())
                    {
                        return raster.GetWidth();
                    }
                });

                #endregion

                PRZH.CheckForCancellation(token);

                #region Create PUID to boundary table ID lookup table

                var tryget_buiddict = await PRZH.GetPUIDsAndBUIDs();

                if (!tryget_buiddict.success)
                {
                    return (false, tryget_buiddict.message);
                }
                Dictionary<int, int> buid_dict = tryget_buiddict.dict;

                #endregion

                PRZH.CheckForCancellation(token);

                #region Find all neighbours

                // Find cells with valid neighours to the right and below
                // - For the right, this requires that the next PUID is also in the dictionary and the cell is not on the right edge (in which case the next valid PUID is on the next row)
                // - For below, this requires that the PUID one row below is also in the dictionary
                // - Keys are PUID, values are BUID
                Dictionary<int, int> buid_dict_right = buid_dict.Where(kv => buid_dict.ContainsKey(kv.Key + 1) & (kv.Key % grid_width != 0)).ToDictionary(kv => kv.Key, kv => kv.Value);
                Dictionary<int, int> buid_dict_below = buid_dict.Where(kv => buid_dict.ContainsKey(kv.Key + grid_width)).ToDictionary(kv => kv.Key, kv => kv.Value);

                // All cells need four side-lengths worth of edges, those without enough neighbours require self intersections to make up the difference
                // - Keys are BUID, values are side lengths
                Dictionary<int, double> buid_dict_self = new Dictionary<int, double>();
                foreach (int puid in buid_dict.Keys)
                {
                    // A cell can have at most 4 self intersections, one less for each valid neighbour
                    int edges = 4;

                    if (buid_dict_right.ContainsKey(puid)) edges--;              // Valid neighbour to right
                    if (buid_dict_right.ContainsKey(puid - 1)) edges--;          // Valid neighbour to left
                    if (buid_dict_below.ContainsKey(puid)) edges--;              // Valid neigbour below
                    if (buid_dict_below.ContainsKey(puid - grid_width)) edges--; // Valid neighbour above

                    // Skip cell if there is no self intersection
                    if (edges == 0) continue;

                    buid_dict_self.Add(buid_dict[puid], side_length * edges);
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region Write out boundary table

                // Initialize output as plain text file
                StringBuilder csv = new StringBuilder();

                // Add header
                csv.AppendLine("i,j,x");

                // Add boundaries to the right
                foreach (var entry in buid_dict_right)
                {
                    csv.AppendLine($"{entry.Value},{entry.Value + 1},{side_length}");
                }

                // Add boundaries below
                foreach (var entry in buid_dict_below)
                {
                    int puid_below = entry.Key + grid_width;
                    csv.AppendLine($"{entry.Value},{buid_dict[puid_below]},{side_length}");
                }

                // Add self intersections
                // - Note keys are BUID, values are side_lengths, unlike previous dictionaries
                foreach (var entry in buid_dict_self)
                {
                    csv.AppendLine($"{entry.Key},{entry.Key},{entry.Value}");
                }

                // Finally output
                string export_folder_path = PRZH.GetPath_ExportWTWFolder();
                string bndpath = Path.Combine(export_folder_path, PRZC.c_FILE_WTW_EXPORT_BND);

                File.WriteAllText(bndpath, csv.ToString());

                #endregion


                #region Compress Boundary CSV

                FileInfo bndfi = new FileInfo(bndpath);
                FileInfo bndgzipfi = new FileInfo(string.Concat(bndfi.FullName, ".gz"));

                using (FileStream fileToBeZippedAsStream = bndfi.OpenRead())
                using (FileStream gzipTargetAsStream = bndgzipfi.Create())
                using (GZipStream gzipStream = new GZipStream(gzipTargetAsStream, CompressionMode.Compress))
                {
                    fileToBeZippedAsStream.CopyTo(gzipStream);
                }

                File.Delete(bndpath);

                return (true, "success");

                #endregion
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
        private async Task<(bool success, string message)> BuildAttributeTable(CancellationToken token)
        {
            int val = PM.Current;
            int max = PM.Max;

            try
            {
                #region Initialization

                // Identify extracted elements and tiles
                var tryread_national_element_tiles = PRZH.ReadBinary(PRZH.GetPath_ProjectNationalElementTilesMetadataPath());
                if (!tryread_national_element_tiles.success)
                {
                    return (false, $"Could not read tile metadata for national data, please try re-extracting national data. Message: {tryread_national_element_tiles.message}");
                }

                Dictionary<string, HashSet<int>> national_element_tiles = (Dictionary<string, HashSet<int>>)tryread_national_element_tiles.obj;
                Dictionary<string, HashSet<int>> regional_element_tiles = new Dictionary<string, HashSet<int>>(); // TODO: Update to load regional data also

                Dictionary<string, HashSet<int>> element_tiles = national_element_tiles.Concat(regional_element_tiles).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                // Build dictionary of elements and corresponding dictionary locations
                Dictionary<string, string> element_location = new Dictionary<string, string>(element_tiles.Count);

                string natelementfolder = PRZH.GetPath_ProjectNationalElementsSubfolder();
                string regelementfolder = ""; // TODO: implement PRZH.GetPath_ProjectRegionalElementsSubfolder();

                foreach (string natelement in national_element_tiles.Keys) element_location.Add(natelement, natelementfolder);
                foreach (string regelement in regional_element_tiles.Keys) element_location.Add(regelement, regelementfolder);

                // Initialize dictionary
                //Dictionary<string, Dictionary<long, double>> attributes = new Dictionary<string, Dictionary<long, double>>(element_paths.Count());

                #endregion

                PRZH.CheckForCancellation(token);

                #region Create cell number to PUID lookup table split by tile

                var tryget_cndict = await PRZH.GetCellNumbersAndPUIDsbyTile();

                if (!tryget_cndict.success)
                {
                    return (false, tryget_cndict.message);
                }
                Dictionary<int, Dictionary<long, int>> cn_dict_by_tile = tryget_cndict.dict;

                #endregion

                PRZH.CheckForCancellation(token);

                #region Build and write attribute table

                // Finally output
                string export_folder_path = PRZH.GetPath_ExportWTWFolder();
                string atrpath = Path.Combine(export_folder_path, PRZC.c_FILE_WTW_EXPORT_ATTR);

                int progress = 0;

                // File is written Line-by-Line, data is loaded Tile-by-Tile
                await using (StreamWriter writetext = new StreamWriter(atrpath))
                {
                    // Header
                    foreach (string element_name in element_tiles.Keys)
                    {
                        await writetext.WriteAsync($"{element_name},");
                    }
                    await writetext.WriteLineAsync("_index");

                    // For each tile
                    // - Note: iterator Key is tile id, Value is dictionary of CNID to PUID
                    foreach (var tile_cells in cn_dict_by_tile)
                    {
                        // Load attributes in tile
                        Dictionary<string, Dictionary<long, double>> tile_attributes = new Dictionary<string, Dictionary<long, double>>(element_tiles.Count());

                        // Load attributes for tile asynchronously
                        // - Note: iterator Key is element id (string), Value is set of tiles element covers
                        await Parallel.ForEachAsync(element_tiles, async (element_tile, token) =>
                        {
                            // Skip element if it does not overlap current tile
                            if (!element_tile.Value.Contains(tile_cells.Key)) return;

                            // Read in element tile
                            string tile_path = Path.Combine(element_location[element_tile.Key], $"{element_tile.Key}-{tile_cells.Key}.bin");
                            var tryread = PRZH.ReadBinary(tile_path);
                            if (!tryread.success)
                            {
                                throw new Exception($"Could not read element binary file. Message: {tryread.message}");
                            }

                            // Add to attributes
                            tile_attributes.Add(element_tile.Key, (Dictionary<long, double>)tryread.obj);
                        });

                        PRZH.CheckForCancellation(token);

                        // Write to file line-by-line
                        foreach (var cn_puid in tile_cells.Value)
                        {
                            StringBuilder line = new StringBuilder();

                            foreach (string element in element_tiles.Keys)
                            {
                                // If the element does not overlap the tile, write zero
                                if (!tile_attributes.ContainsKey(element)) line.Append("0,");

                                // If the element does include the current cell, write zero
                                else if (!tile_attributes[element].ContainsKey(cn_puid.Key)) line.Append("0,");

                                // Otherwise write the relevant data
                                else line.Append($"{tile_attributes[element][cn_puid.Key]},");
                            }

                            // Finally write PUID and end line
                            line.Append($"{cn_puid.Value}");

                            await writetext.WriteLineAsync(line, token);
                        }

                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done writing {++progress} / {cn_dict_by_tile.Count()} tiles of attribute table."), true, ++val);

                        PRZH.CheckForCancellation(token);
                    }
                }

                #endregion

                PRZH.CheckForCancellation(token);

                #region Compress Boundary CSV

                FileInfo atrfi = new FileInfo(atrpath);
                FileInfo atrgzipfi = new FileInfo(string.Concat(atrfi.FullName, ".gz"));

                using (FileStream fileToBeZippedAsStream = atrfi.OpenRead())
                using (FileStream gzipTargetAsStream = atrgzipfi.Create())
                using (GZipStream gzipStream = new GZipStream(gzipTargetAsStream, CompressionMode.Compress))
                {
                    fileToBeZippedAsStream.CopyTo(gzipStream);
                }

                File.Delete(atrpath);

                return (true, "success");

                #endregion
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
            Operation_Cmd_IsEnabled = _pu_exists;
            OpStat_Img_Visibility = Visibility.Hidden;
            OpStat_Txt_Label = "Idle";
            _operationIsUnderway = false;
        }


        #endregion
    }
}
