using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Mapping;
using System.Reflection;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ProMsgBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using PRZC = NCC.PRZTools.PRZConstants;
using PRZH = NCC.PRZTools.PRZHelper;
using ArcGIS.Desktop.Core;
using System.Diagnostics;
using System.Collections.Generic;
using ArcGIS.Desktop.Core.Geoprocessing;
using System.IO;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Core.Data;
using System.Linq;

namespace NCC.PRZTools
{
    public class SerializeDatabaseVM : PropertyChangedBase
    {
        public SerializeDatabaseVM()
        {
        }

        #region FIELDS

        private CancellationTokenSource _cts = null;
        private ProgressManager _pm = ProgressManager.CreateProgressManager(50);
        private bool _operation_Cmd_Nat_IsEnabled;
        private bool _operation_Cmd_Reg_IsEnabled;
        private bool _operationIsUnderway = false;
        private Cursor _proWindowCursor;

        private bool _natdb_exists = false;
        private bool _regdb_exists = false;

        private Map _map;

        #region COMMANDS

        private ICommand _cmdSerializeNationalData;
        private ICommand _cmdSerializeRegionalData;
        private ICommand _cmdCancel;
        private ICommand _cmdClearLog;

        #endregion

        #region COMPONENT STATUS INDICATORS

        // National DB
        private string _compStat_Img_NatDB_Path;
        private string _compStat_Txt_NatDB_Label;

        // Regional DB
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

        public bool Operation_Cmd_Nat_IsEnabled
        {
            get => _operation_Cmd_Nat_IsEnabled;
            set => SetProperty(ref _operation_Cmd_Nat_IsEnabled, value, () => Operation_Cmd_Nat_IsEnabled);
        }

        public bool Operation_Cmd_Reg_IsEnabled
        {
            get => _operation_Cmd_Reg_IsEnabled;
            set => SetProperty(ref _operation_Cmd_Reg_IsEnabled, value, () => Operation_Cmd_Reg_IsEnabled);
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

        public ICommand CmdSerializeNationalData => _cmdSerializeNationalData ?? (_cmdSerializeNationalData = new RelayCommand(async () =>
        {
            // Change UI to Underway
            StartOpUI();

            // Start the operation
            using (_cts = new CancellationTokenSource())
            {
                await SerializeNationalData(_cts.Token);
            }

            // Set source to null (it's already disposed)
            _cts = null;

            // Validate controls
            await ValidateControls();

            // Reset UI to Idle
            ResetOpUI();

        }, () => true, true, false));

        public ICommand CmdSerializeRegionalData => _cmdSerializeRegionalData ?? (_cmdSerializeRegionalData = new RelayCommand(async () =>
        {
            // Change UI to Underway
            StartOpUI();

            // Start the operation
            using (_cts = new CancellationTokenSource())
            {
                await SerializeRegionalData(_cts.Token);
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

        private async Task SerializeNationalData(CancellationToken token)
        {
            bool edits_are_disabled = !Project.Current.IsEditingEnabled;
            int val = 0;
            int max = 5;

            try
            {
                #region INITIALIZATION

                // Initialize ProgressBar and Progress Log
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Initializing the Geodatabase Serializer..."), false, max, ++val);

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
                    else
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog("ArcGIS Pro editing enabled."), true, ++val);
                    }
                }

                #endregion

                // Ensure the National db exists
                string natpath = PRZH.GetPath_NatGDB();
                var tryexists_nat = await PRZH.GDBExists_Nat();
                if (!tryexists_nat.exists)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Valid National Geodatabase not found: '{natpath}'.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Valid National Geodatabase not found at {natpath}.");
                    return;
                }

                if (ProMsgBox.Show($"If you proceed, any existing serialized data in the National Geodatabase will be overwritten." +
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

                #region RETRIEVE AND PREPARE INFO FROM NATIONAL DATABASE

                int log_every = 250;
                int progress = 0;

                // Get the National Elements from the copied national elements table
                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieving national elements..."), true, ++val);
                var elem_outcome = await PRZH.GetNationalElements_Direct();
                if (!elem_outcome.success)
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Error retrieving national elements.\n{elem_outcome.message}", LogMessageType.ERROR), true, ++val);
                    ProMsgBox.Show($"Error retrieving national elements.\n{elem_outcome.message}");
                    return;
                }
                else
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Retrieved {elem_outcome.elements.Count} national elements."), true, max + (elem_outcome.elements.Count / log_every), ++val);
                }
                List<NatElement> elements = elem_outcome.elements;

                #endregion

                PRZH.CheckForCancellation(token);

                #region Serialize element tables

                PRZH.UpdateProgress(PM, PRZH.WriteLog("Serializing elements..."), true, ++val);

                // Refresh output and temp dirs
                String outputDir = PRZH.GetPath_NationalDatabaseElementsSubfolder();
                String metadataTempDir = Path.Combine(Path.GetDirectoryName(PRZH.GetPath_NationalDatabaseElementsSubfolder()), "temp");

                if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
                if (Directory.Exists(metadataTempDir)) Directory.Delete(metadataTempDir, true);

                Directory.CreateDirectory(outputDir);
                Directory.CreateDirectory(metadataTempDir);


                await Parallel.ForEachAsync(elements, async (element, token) =>
                {
                    progress++;
                    if (progress % log_every == 0)
                    {
                        PRZH.UpdateProgress(PM, PRZH.WriteLog($"Done serializing {progress} / {elements.Count} elements."), true, ++val);
                    }

                    // Construct dictionary of natgrid cells / table vbalues
                    Dictionary<long, double> element_dict = await QueuedTask.Run(() =>
                    {
                        var tryget = PRZH.GetTable_Nat(element.ElementTable);
                        if (!tryget.success)
                        {
                            throw new Exception("Error retrieving table.");
                        }

                        using (Table table = tryget.table)
                        using (RowCursor rowCursor = table.Search())
                        {
                            Dictionary<long, double> cells = new Dictionary<long, double>((int)table.GetCount());

                            // Fill dictionary
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    cells.TryAdd(Convert.ToInt64(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER]), Convert.ToDouble(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_VALUE]));
                                }
                            }

                            return cells;
                        }
                    });

                    PRZH.CheckForCancellation(token);

                    // Get tiles
                    Dictionary<int, HashSet<long>> tiles = NationalGrid.GetTilesFromCells(element_dict);

                    // Write tile metadata
                    PRZH.WriteBinary(new HashSet<int>(tiles.Keys), Path.Combine(metadataTempDir, $"{element.ElementTable}.bin"));

                    // Split and write by tile
                    foreach (KeyValuePair<int, HashSet<long>> tile in tiles)
                    {
                        Dictionary<long, double> tile_dict = element_dict.Where(cell => tile.Value.Contains(cell.Key)).ToDictionary(cell => cell.Key, cell => cell.Value);
                        string outputFile = $"{Path.Combine(outputDir, element.ElementTable)}-{tile.Key}.bin";
                        PRZH.WriteBinary(tile_dict, outputFile);
                    }

                    PRZH.CheckForCancellation(token);
                });

                #endregion

                #region WRAP UP

                PRZH.UpdateProgress(PM, PRZH.WriteLog($"Wrapping up..."), true, ++val);

                // Compile tile metadata, save, clear tempdir
                Dictionary<int, HashSet<int>> tileMetadata = new Dictionary<int, HashSet<int>>(elements.Count);
                foreach (var element in elements)
                {
                    tileMetadata.Add(Int32.Parse(element.ElementTable.Substring(1)), (HashSet<int>)PRZH.ReadBinary(Path.Combine(metadataTempDir, $"{element.ElementTable}.bin")).obj);
                }
                PRZH.WriteBinary(tileMetadata, Path.Combine(outputDir, PRZC.c_FILE_METADATA_TILES));
                Directory.Delete(metadataTempDir, true);

                // Final message
                stopwatch.Stop();
                string message = PRZH.GetElapsedTimeMessage(stopwatch.Elapsed);
                PRZH.UpdateProgress(PM, PRZH.WriteLog("National Geodatabase serialization completed successfully."), true, 1, 1);
                PRZH.UpdateProgress(PM, PRZH.WriteLog(message), true, 1, 1);

                ProMsgBox.Show("National Geodatabase serialization completed successfully!" + Environment.NewLine + Environment.NewLine + message);

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

        private async Task SerializeRegionalData(CancellationToken token)
        {
            bool edits_are_disabled = !Project.Current.IsEditingEnabled;
            int val = 0;
            int max = 50;

            try
            {
                #region INITIALIZATION

                // Initialize ProgressBar and Progress Log
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Initializing the Geodatabase Serializer..."), false, max, ++val);

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

                // Ensure the Regional Data Folder exists
                string regpath = PRZH.GetPath_RegionalDataFolder();
                if (!PRZH.FolderExists_RegionalData().exists) // TODO: Replace with standardized geodatabase check? 
                {
                    PRZH.UpdateProgress(PM, PRZH.WriteLog($"Valid Regional Geodatabase not found: '{regpath}'.", LogMessageType.VALIDATION_ERROR), true, ++val);
                    ProMsgBox.Show($"Valid Regional Geodatabase not found at {regpath}.");
                    return;
                }

                if (ProMsgBox.Show($"If you proceed, any existing serialized data in the Regional Geodatabase will be overwritten." +
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

                PRZH.UpdateProgress(PM, PRZH.WriteLog("Regional Geodatabase serializer not yet implemented! Not doing anything!"), false, max, ++val);

                #region WRAP UP

                // Final message
                stopwatch.Stop();
                string message = PRZH.GetElapsedTimeMessage(stopwatch.Elapsed);
                PRZH.UpdateProgress(PM, PRZH.WriteLog("Regional Geodatabase serialization completed successfully."), true, 1, 1);
                PRZH.UpdateProgress(PM, PRZH.WriteLog(message), true, 1, 1);

                ProMsgBox.Show("Regional Geodatabase serialization completed successfully!" + Environment.NewLine + Environment.NewLine + message);

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

        private async Task ValidateControls()
        {
            try
            {
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

                // Regional Database existence
                _regdb_exists = PRZH.FolderExists_RegionalData().exists; // TODO: replace with database existence check?

                if (_natdb_exists)
                {
                    CompStat_Txt_RegDB_Label = "Regional Database exists.";
                    CompStat_Img_RegDB_Path = "pack://application:,,,/PRZTools;component/ImagesWPF/ComponentStatus_Yes16.png";
                }
                else
                {
                    CompStat_Txt_RegDB_Label = "Regional Database does not exist or is invalid.";
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
            Operation_Cmd_Nat_IsEnabled = false;
            Operation_Cmd_Reg_IsEnabled = false;
            OpStat_Img_Visibility = Visibility.Visible;
            OpStat_Txt_Label = "Processing...";
            ProWindowCursor = Cursors.Wait;
        }

        private void ResetOpUI()
        {
            ProWindowCursor = Cursors.Arrow;
            Operation_Cmd_Nat_IsEnabled = _natdb_exists;
            Operation_Cmd_Reg_IsEnabled = _regdb_exists;
            OpStat_Img_Visibility = Visibility.Hidden;
            OpStat_Txt_Label = "Idle";
            _operationIsUnderway = false;
        }
        #endregion
    }
}
