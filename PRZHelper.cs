﻿using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.Raster;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ProMsgBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
//using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
using PRZC = NCC.PRZTools.PRZConstants;

namespace NCC.PRZTools
{
    public static class PRZHelper
    {
        #region LOGGING AND NOTIFICATIONS

        // Write to log
        public static string WriteLog(string message, LogMessageType type = LogMessageType.INFO, bool Append = true)
        {
            try
            {
                // Make sure we have a valid log file
                if (!FolderExists_Project().exists)
                {
                    return "";
                }

                string logpath = GetPath_ProjectLog();
                if (!ProjectLogExists().exists)
                {
                    using (FileStream fs = File.Create(logpath)) { }
                }

                // Create the message lines
                string lines = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss.ff tt") + " :: " + type.ToString() 
                                + Environment.NewLine + message 
                                + Environment.NewLine;

                if (Append)
                {
                    using (StreamWriter w = File.AppendText(logpath))
                    {
                        w.WriteLine(lines);
                        w.Flush();
                        w.Close();
                    }
                }
                else
                {
                    using (StreamWriter w = File.CreateText(logpath))
                    {
                        w.WriteLine(lines);
                        w.Flush();
                        w.Close();
                    }
                }

                return lines;
            }
            catch (Exception)
            {
                return "Unable to log message...";
            }
        }

        // Read from the log
        public static string ReadLog()
        {
            try
            {
                if (!ProjectLogExists().exists)
                {
                    return "";
                }

                string logpath = GetPath_ProjectLog();

                return File.ReadAllText(logpath);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }

        // User notifications
        public static void UpdateProgress(ProgressManager pm, string message, bool append)
        {
            try
            {
                DispatchProgress(pm, message, append, null, null, null);
            }
            catch (Exception)
            {
            }
        }

        public static void UpdateProgress(ProgressManager pm, string message, bool append, int current)
        {
            try
            {
                DispatchProgress(pm, message, append, null, null, current);
            }
            catch (Exception)
            {
            }
        }

        public static void UpdateProgress(ProgressManager pm, string message, bool append, int max, int current)
        {
            try
            {
                DispatchProgress(pm, message, append, null, max, current);
            }
            catch (Exception)
            {
            }
        }

        public static void UpdateProgress(ProgressManager pm, string message, bool append, int min, int max, int current)
        {
            try
            {
                DispatchProgress(pm, message, append, min, max, current);
            }
            catch (Exception)
            {
            }
        }

        private static void DispatchProgress(ProgressManager pm, string message, bool append, int? min, int? max, int? current)
        {
            try
            {
                if (System.Windows.Application.Current.Dispatcher.CheckAccess())
                {
                    ManageProgress(pm, message, append, min, max, current);
                }
                else
                {
                    ProApp.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)(() =>
                    {
                        ManageProgress(pm, message, append, min, max, current);
                    }));
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
        }

        private static void ManageProgress(ProgressManager pm, string message, bool append, int? min, int? max, int? current)
        {
            try
            {
                // Update the Message
                if (message == null)
                {
                    if (append == false)
                    {
                        pm.Message = "";
                    }
                }
                else
                {
                    pm.Message = append ? (pm.Message + Environment.NewLine + message) : message;
                }

                // Update the Min property
                if (min != null)
                {
                    pm.Min = (int)min;
                }

                // Update the Max property
                if (max != null)
                {
                    pm.Max = (int)max;
                }

                // Update the Value property
                if (current != null)
                {
                    pm.Current = (int)current;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
        }


        #endregion LOGGING

        #region PATHS

        #region FILE AND FOLDER PATHS

        // Project Folder
        public static string GetPath_ProjectFolder()
        {
            try
            {
                return Properties.Settings.Default.WORKSPACE_PATH;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // Project GDB
        public static string GetPath_ProjectGDB()
        {
            try
            {
                string wspath = GetPath_ProjectFolder();
                string gdbpath = Path.Combine(wspath, PRZC.c_FILE_PRZ_FGDB);

                return gdbpath;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // National GDB
        public static string GetPath_NatGDB()
        {
            try
            {
                return Properties.Settings.Default.NATDB_DBPATH;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // Raster Tools Scratch Geodatabase
        public static string GetPath_RTScratchGDB()
        {
            try
            {
                return Properties.Settings.Default.RT_SCRATCH_FGDB;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // Project Log File
        public static string GetPath_ProjectLog()
        {
            try
            {
                string ws = GetPath_ProjectFolder();
                string logpath = Path.Combine(ws, PRZC.c_FILE_PRZ_LOG);

                return logpath;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // Export WTW Subfolder
        public static string GetPath_ExportWTWFolder()
        {
            try
            {
                string wspath = GetPath_ProjectFolder();
                string exportwtwpath = Path.Combine(wspath, PRZC.c_DIR_EXPORT_WTW);

                return exportwtwpath;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // Regional Data Folder
        public static string GetPath_RegionalDataFolder()
        {
            try
            {
                return Properties.Settings.Default.REGIONAL_FOLDER_PATH;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // Regional Data Subfolder
        public static string GetPath_RegionalDataSubfolder(RegionalDataSubfolder subfolder)
        {
            try
            {
                // Get the folder path
                string regpath = GetPath_RegionalDataFolder();
                string subpath = "";

                switch (subfolder)
                {
                    case RegionalDataSubfolder.GOALS:
                        subpath = PRZC.c_DIR_REGDATA_GOALS;
                        break;

                    case RegionalDataSubfolder.WEIGHTS:
                        subpath = PRZC.c_DIR_REGDATA_WEIGHTS;
                        break;

                    case RegionalDataSubfolder.INCLUDES:
                        subpath = PRZC.c_DIR_REGDATA_INCLUDES;
                        break;
                    case RegionalDataSubfolder.EXCLUDES:
                        subpath = PRZC.c_DIR_REGDATA_EXCLUDES;
                        break;
                }

                return Path.Combine(regpath, subpath);

            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }


        #endregion

        #region GEODATABASE OBJECT PATHS

        /// <summary>
        /// Retrieve the path to a geodatabase object within the project geodatabase. If the optional
        /// second parameter is provided, the path to the object within a feature dataset is returned.
        /// Silent errors.
        /// </summary>
        /// <param name="gdb_obj_name"></param>
        /// <param name="fds_name"></param>
        /// <returns></returns>
        public static (bool success, string path, string message) GetPath_Project(string gdb_obj_name, string fds_name = "")
        {
            try
            {
                // Get the GDB Path
                string gdbpath = GetPath_ProjectGDB();

                // ensure gdbpath is valid
                if (string.IsNullOrEmpty(gdbpath))
                {
                    return (false, "", "geodatabase path is null");
                }

                // ensure object name is not empty
                if (string.IsNullOrEmpty(gdb_obj_name))
                {
                    return (false, "", "geodatabase object name not supplied");
                }

                if (string.IsNullOrEmpty(fds_name))
                {
                    return (true, Path.Combine(gdbpath, gdb_obj_name), "success");
                }
                else
                {
                    return (true, Path.Combine(gdbpath, fds_name, gdb_obj_name), "success");
                }
            }
            catch (Exception ex)
            {
                return (false, "", ex.Message);
            }
        }

        /// <summary>
        /// Retrieve the path to a geodatabase object within the national geodatabase
        /// (file or enterprise).  Silent errors.
        /// </summary>
        /// <param name="gdb_obj_name"></param>
        /// <returns></returns>
        public static (bool success, string path, string message) GetPath_Nat(string gdb_obj_name)
        {
            try
            {
                // Get the GDB Path
                string gdbpath = GetPath_NatGDB();

                if (string.IsNullOrEmpty(gdbpath))
                {
                    return (false, "", "geodatabase path is null");
                }

                return (true, Path.Combine(gdbpath, gdb_obj_name), "success");
            }
            catch (Exception ex)
            {
                return (false, "", ex.Message);
            }
        }

        /// <summary>
        /// Retrieve the path to a geodatabase object within the RT Scratch file
        /// geodatabase.  Silent errors.
        /// </summary>
        /// <param name="gdb_obj_name"></param>
        /// <returns></returns>
        public static (bool success, string path, string message) GetPath_RTScratch(string gdb_obj_name)
        {
            try
            {
                // Get the GDB Path
                string gdbpath = GetPath_RTScratchGDB();

                if (string.IsNullOrEmpty(gdbpath))
                {
                    return (false, "", "geodatabase path is null");
                }

                return (true, Path.Combine(gdbpath, gdb_obj_name), "success");
            }
            catch (Exception ex)
            {
                return (false, "", ex.Message);
            }
        }

        #endregion

        #endregion

        #region OBJECT EXISTENCE

        #region FOLDER EXISTENCE

        /// <summary>
        /// Establish the existence of the project folder.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static (bool exists, string message) FolderExists_Project()
        {
            try
            {
                string path = GetPath_ProjectFolder();
                bool exists = Directory.Exists(path);

                return (exists, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of the regional data folder.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static (bool exists, string message) FolderExists_RegionalData()
        {
            try
            {
                string path = GetPath_RegionalDataFolder();
                bool exists = Directory.Exists(path);

                return (exists, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of the specified regional data subfolder.  Silent errors.
        /// </summary>
        /// <param name="subfolder"></param>
        /// <returns></returns>
        public static (bool exists, string message) FolderExists_RegionalDataSubfolder(RegionalDataSubfolder subfolder)
        {
            try
            {
                string path = GetPath_RegionalDataSubfolder(subfolder);
                bool exists = Directory.Exists(path);

                return (exists, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }


        /// <summary>
        /// Establish the existence of the export where-to-work tool folder.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static (bool exists, string message) FolderExists_ExportWTW()
        {
            try
            {
                string path = GetPath_ExportWTWFolder();
                bool exists = Directory.Exists(path);

                return (exists, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region GDB EXISTENCE

        /// <summary>
        /// Establish the existence of a geodatabase (file or enterprise) from a path.  Silent errors.
        /// </summary>
        /// <param name="gdbpath"></param>
        /// <returns></returns>
        public static async Task<(bool exists, GeoDBType gdbType, string message)> GDBExists(string gdbpath)
        {
            try
            {
                // Run this on the worker thread
                return await QueuedTask.Run(() => 
                {
                    // Ensure a non-null and non-empty path
                    if (string.IsNullOrEmpty(gdbpath))
                    {
                        return (false, GeoDBType.Unknown, "Geodatabase path is null or empty.");
                    }

                    // Ensure a rooted path
                    if (!Path.IsPathRooted(gdbpath))
                    {
                        return (false, GeoDBType.Unknown, $"Path is not rooted: {gdbpath}");
                    }

                    // Create the Uri object
                    Uri uri = null;

                    try
                    {
                        uri = new Uri(gdbpath);
                    }
                    catch
                    {
                        return (false, GeoDBType.Unknown, $"Unable to create Uri from path: {gdbpath}");
                    }

                    // Determine if path is file geodatabase (.gdb) or database connection file (.sde)
                    if (Directory.Exists(gdbpath) && gdbpath.EndsWith(".gdb"))  // File Geodatabase (possibly)
                    {
                        // Create the Connection Path object
                        FileGeodatabaseConnectionPath conn = null;

                        try
                        {
                            conn = new FileGeodatabaseConnectionPath(uri);
                        }
                        catch
                        {
                            return (false, GeoDBType.Unknown, $"Unable to create file geodatabase connection path from path: {gdbpath}");
                        }

                        // Try to open the connection
                        try
                        {
                            using (Geodatabase gdb = new Geodatabase(conn)) { }
                        }
                        catch
                        {
                            return (false, GeoDBType.Unknown, $"File geodatabase could not be opened from path: {gdbpath}");
                        }

                        // If I get to this point, the file gdb exists and was successfully opened
                        return (true, GeoDBType.FileGDB, "success");
                    }
                    else if (File.Exists(gdbpath) && gdbpath.EndsWith(".sde"))    // It's a connection file (.sde)
                    {
                        // Create the Connection File object
                        DatabaseConnectionFile conn = null;

                        try
                        {
                            conn = new DatabaseConnectionFile(uri);
                        }
                        catch
                        {
                            return (false, GeoDBType.Unknown, $"Unable to create database connection file from path: {gdbpath}");
                        }

                        // try to open the connection
                        try
                        {
                            using (Geodatabase gdb = new Geodatabase(conn)) { }
                        }
                        catch
                        {
                            return (false, GeoDBType.Unknown, $"Enterprise geodatabase could not be opened from path: {gdbpath}");
                        }

                        // If I get to this point, the enterprise geodatabase exists and was successfully opened
                        return (true, GeoDBType.EnterpriseGDB, "success");
                    }
                    else
                    {
                        // something else, weird!
                        return (false, GeoDBType.Unknown, $"unable to process database path: {gdbpath}");
                    }
                });

            }
            catch (Exception ex)
            {
                return (false, GeoDBType.Unknown, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of the project geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> GDBExists_Project()
        {
            try
            {
                var tryexists = await GDBExists(GetPath_ProjectGDB());

                return (tryexists.exists, tryexists.message);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of the RT Scratch file geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> GDBExists_RTScratch()
        {
            try
            {
                var tryexists = await GDBExists(GetPath_RTScratchGDB());

                return (tryexists.exists, tryexists.message);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of the national geodatabase (file or enterprise).
        /// Geodatabase must exist and be valid (have the required tables).  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool exists, GeoDBType gdbType, string message)> GDBExists_Nat()
        {
            try
            {
                var tryex = await GDBExists(GetPath_NatGDB());

                if (!tryex.exists)
                {
                    return tryex;
                }

                // Ensure that national geodatabase is valid
                // TODO: This is not ideal.  Actual test should be done here rather than relying on a flag set somewhere else.
                if (!Properties.Settings.Default.NATDB_DBVALID)
                {
                    return (false, GeoDBType.Unknown, "National geodatabase exists but is invalid.");
                }
                else
                {
                    return tryex;
                }
            }
            catch (Exception ex)
            {
                return (false, GeoDBType.Unknown, ex.Message);
            }
        }

        #endregion

        #region FDS ANY GDB

        /// <summary>
        /// Establish the existence of a feature dataset by name within a specified geodatabase.
        /// Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="fds_name"></param>
        /// <returns></returns>
        public static (bool exists, string message) FDSExists(Geodatabase geodatabase, string fds_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Attempt to retrieve definition based on name
                using (FeatureDatasetDefinition fdsDef = geodatabase.GetDefinition<FeatureDatasetDefinition>(fds_name))
                {
                    // Error will be thrown by using statement above if FDS of the supplied name doesn't exist in GDB
                }

                return (true, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region FDS IN PROJECT GDB

        /// <summary>
        /// Establish the existence of a feature dataset by name from the project
        /// file geodatabase.  Silent errors.
        /// </summary>
        /// <param name="fds_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> FDSExists_Project(string fds_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_Project();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        return FDSExists(geodatabase, fds_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region FC/TABLE/RASTER ANY GDB

        /// <summary>
        /// Establish the existence of a feature class by name within a specified geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="fc_name"></param>
        /// <returns></returns>
        public static (bool exists, string message) FCExists(Geodatabase geodatabase, string fc_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Attempt to retrieve definition based on name
                using (FeatureClassDefinition fcDef = geodatabase.GetDefinition<FeatureClassDefinition>(fc_name))
                {
                    // Error will be thrown by using statement above if FC of the supplied name doesn't exist in GDB
                }

                return (true, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of a table by name within a specified geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static (bool exists, string message) TableExists(Geodatabase geodatabase, string table_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Attempt to retrieve definition based on name
                using (TableDefinition tabDef = geodatabase.GetDefinition<TableDefinition>(table_name))
                {
                    // Error will be thrown by using statement above if table of the supplied name doesn't exist in GDB
                }

                return (true, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of a raster dataset by name within a specified geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="raster_name"></param>
        /// <returns></returns>
        public static (bool exists, string message) RasterExists(Geodatabase geodatabase, string raster_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Attempt to retrieve definition based on name
                using (RasterDatasetDefinition rasDef = geodatabase.GetDefinition<RasterDatasetDefinition>(raster_name))
                {
                    // Error will be thrown by using statement above if rasterdataset of the supplied name doesn't exist in GDB
                }

                return (true, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region FC/TABLE/RASTER IN PROJECT GDB

        /// <summary>
        /// Establish the existence of a feature class by name from the project file geodatabase.  Silent errors.
        /// </summary>
        /// <param name="fc_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> FCExists_Project(string fc_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_Project();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        return FCExists(geodatabase, fc_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of a table by name from the project file geodatabase.  Silent errors.
        /// </summary>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> TableExists_Project(string table_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_Project();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        return TableExists(geodatabase, table_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of a raster dataset by name from the project file geodatabase.  Silent errors.
        /// </summary>
        /// <param name="raster_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> RasterExists_Project(string raster_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_Project();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        return RasterExists(geodatabase, raster_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region FC/TABLE/RASTER IN RT SCRATCH GDB

        /// <summary>
        /// Establish the existence of a feature class by name from the rt scratch file geodatabase.  Silent errors.
        /// </summary>
        /// <param name="fc_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> FCExists_RTScratch(string fc_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_RTScratch();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        return FCExists(geodatabase, fc_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of a table by name from the rt scratch file geodatabase.  Silent errors.
        /// </summary>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> TableExists_RTScratch(string table_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_RTScratch();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        return TableExists(geodatabase, table_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of a raster dataset by name from the rt scratch file geodatabase.  Silent errors.
        /// </summary>
        /// <param name="raster_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> RasterExists_RTScratch(string raster_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_RTScratch();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        return RasterExists(geodatabase, raster_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region FC/TABLE/RASTER IN NATIONAL GDB

        /// <summary>
        /// Establish the existence of a table by name from the national geodatabase.  Silent errors.
        /// </summary>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static async Task<(bool exists, string message)> TableExists_Nat(string table_name)
        {
            try
            {
                // Run this code on the worker thread
                return await QueuedTask.Run(() =>
                {
                    // Get geodatabase
                    var tryget_gdb = GetGDB_Nat();

                    if (!tryget_gdb.success)
                    {
                        return (false, tryget_gdb.message);
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        if (geodatabase == null)
                        {
                            return (false, "unable to access the geodatabase.");
                        }

                        // fully qualified table name as required
                        SQLSyntax syntax = geodatabase.GetSQLSyntax();

                        string db = Properties.Settings.Default.NATDB_DBNAME;
                        string schema = Properties.Settings.Default.NATDB_SCHEMANAME;

                        string qualified_table_name = syntax.QualifyTableName(db, schema, table_name);

                        return TableExists(geodatabase, qualified_table_name);
                    }
                });
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region MISCELLANEOUS

        /// <summary>
        /// Establish the existence of the Planning Units FC and Raster Datasets.
        /// Also identifies if Planning Units datasets are national-enabled.
        /// Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool exists, bool national_enabled, string message)> PUDataExists()
        {
            try
            {
                bool fc_exists = (await FCExists_Project(PRZC.c_FC_PLANNING_UNITS)).exists;
                bool ras_exists = (await RasterExists_Project(PRZC.c_RAS_PLANNING_UNITS)).exists;

                if(ras_exists) // if (fc_exists & ras_exists) // TODO: Clean up after testing without fc
                {
                    // Determine if dataset is national-enabled or not
                    return await QueuedTask.Run(() =>
                    {
                        var tryget_ras = GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);
                        using (RasterDataset rasterDataset = tryget_ras.rasterDataset)
                        using (Raster raster = rasterDataset.CreateFullRaster())
                        using (Table table = raster.GetAttributeTable())
                        using (TableDefinition rasDef= table.GetDefinition())
                        {
                            // Search for the cell number field
                            var a = rasDef.GetFields().Where(f => string.Equals(f.Name, PRZC.c_FLD_RAS_PU_NATGRID_CELL_NUMBER, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

                            if (a == null)
                            {
                                // no cellnumber field found
                                return (true, false, "cell number field not found");
                            }
                            else
                            {
                                // cell number field found
                                return (true, true, "cell number field found");
                            }
                        }
/*                        var tryget_fc = GetFC_Project(PRZC.c_FC_PLANNING_UNITS);
                        using (FeatureClass fc = tryget_fc.featureclass)
                        using (FeatureClassDefinition fcDef = fc.GetDefinition())
                        {
                            // Search for the cell number field
                            var a = fcDef.GetFields().Where(f => string.Equals(f.Name, PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

                            if (a == null)
                            {
                                // no cellnumber field found
                                return (true, false, "cell number field not found");
                            }
                            else
                            {
                                // cell number field found
                                return (true, true, "cell number field found");
                            }
                        }*/
                    });
                }
                else if (fc_exists)
                {
                    return (false, false, "raster planning units not found");
                }
                else if (ras_exists)
                {
                    return (false, false, "feature class planning units not found");
                }
                else
                {
                    return (false, false, "feature class and raster planning units not found");
                }
            }
            catch (Exception ex)
            {
                return (false, false, ex.Message);
            }
        }

        /// <summary>
        /// Establish the existence of a project log file.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static (bool exists, string message) ProjectLogExists()
        {
            try
            {
                string path = GetPath_ProjectLog();

                bool exists = File.Exists(path);

                return (exists, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #endregion

        #region GEODATABASE OBJECT RETRIEVAL

        #region GEODATABASES

        /// <summary>
        /// Retrieve a file geodatabase from a path.  Must be run on MCT. Silent errors.
        /// </summary>
        /// <param name="gdbpath"></param>
        /// <returns></returns>
        public static (bool success, Geodatabase geodatabase, string message) GetFileGDB(string gdbpath)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure a non-null and non-empty path
                if (string.IsNullOrEmpty(gdbpath))
                {
                    return (false, null, $"Path is null or empty.");
                }

                // Ensure a rooted path
                if (!Path.IsPathRooted(gdbpath))
                {
                    return (false, null, $"Path is not rooted: {gdbpath}");
                }

                // Ensure the path is an existing directory
                if (!Directory.Exists(gdbpath))
                {
                    return (false, null, $"Path is not a valid folder path.\n{gdbpath}");
                }

                // Create the Uri object
                Uri uri = null;

                try
                {
                    uri = new Uri(gdbpath);
                }
                catch
                {
                    return (false, null, $"Unable to create Uri from path: {gdbpath}");
                }

                // Create the Connection Path object
                FileGeodatabaseConnectionPath connpath = null;

                try
                {
                    connpath = new FileGeodatabaseConnectionPath(uri);
                }
                catch
                {
                    return (false, null, $"Unable to create file geodatabase connection path from path: {gdbpath}");
                }

                // Create the Geodatabase object
                Geodatabase geodatabase = null;

                // Try to open the geodatabase from the connection path
                try
                {
                    geodatabase = new Geodatabase(connpath);
                }
                catch (Exception ex)
                {
                    return (false, null, $"Error opening the geodatabase from the connection path.\n{ex.Message}");
                }

                // If we get to here, the geodatabase has been opened successfully!  Return it!
                return (true, geodatabase, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve an enterprise geodatabase from database connection file path.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="sdepath"></param>
        /// <returns></returns>
        public static (bool success, Geodatabase geodatabase, string message) GetEnterpriseGDB(string sdepath)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure a non-null and non-empty path
                if (string.IsNullOrEmpty(sdepath))
                {
                    return (false, null, "Geodatabase path is null or empty.");
                }

                // Ensure a rooted path
                if (!Path.IsPathRooted(sdepath))
                {
                    return (false, null, $"Path is not rooted: {sdepath}");
                }

                // Ensure path is a valid file
                if (!File.Exists(sdepath))
                {
                    return (false, null, $"Path is not a valid file: {sdepath}");
                }

                // Create the Uri object
                Uri uri = null;

                try
                {
                    uri = new Uri(sdepath);
                }
                catch
                {
                    return (false, null, $"Unable to create Uri from path: {sdepath}");
                }

                // Ensure the path is an existing sde connection file
                if (sdepath.EndsWith(".sde"))
                {
                    // Create the Connection File object
                    DatabaseConnectionFile conn = null;

                    try
                    {
                        conn = new DatabaseConnectionFile(uri);
                    }
                    catch
                    {
                        return (false, null, $"Unable to create database connection file from path: {sdepath}");
                    }

                    // try to open the connection
                    Geodatabase geodatabase = null;

                    try
                    {
                        geodatabase = new Geodatabase(conn);
                    }
                    catch
                    {
                        return (false, null, $"Enterprise geodatabase could not be opened from path: {sdepath}");
                    }

                    // If I get to this point, the enterprise geodatabase exists and was successfully opened
                    return (true, geodatabase, "success");
                }
                else
                {
                    return (false, null, "Unable to process database connection file.");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a file or enterprise geodatabase from a path.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static (bool success, Geodatabase geodatabase, GeoDBType gdbType, string message) GetGDB(string path)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure a non-null and non-empty path
                if (string.IsNullOrEmpty(path))
                {
                    return (false, null, GeoDBType.Unknown, "Geodatabase path is null or empty.");
                }

                // Ensure a rooted path
                if (!Path.IsPathRooted(path))
                {
                    return (false, null, GeoDBType.Unknown, $"Path is not rooted: {path}");
                }

                if (path.EndsWith(".gdb"))
                {
                    var tryget = GetFileGDB(path);
                    return (tryget.success, tryget.geodatabase, GeoDBType.FileGDB, tryget.message);
                }
                else if (path.EndsWith(".sde"))
                {
                    var tryget = GetEnterpriseGDB(path);
                    return (tryget.success, tryget.geodatabase, GeoDBType.EnterpriseGDB, tryget.message);
                }
                else
                {
                    return (false, null, GeoDBType.Unknown, "Invalid geodatabase path (not *.gdb or *.sde)");
                }
            }
            catch (Exception ex)
            {
                return (false, null, GeoDBType.Unknown, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve the project file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static (bool success, Geodatabase geodatabase, string message) GetGDB_Project()
        {
            try
            {
                // Ensure this is called on the worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Get the Project GDB Path
                string gdbpath = GetPath_ProjectGDB();

                // Retrieve the File Geodatabase
                return GetFileGDB(gdbpath);
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve the RT scratch file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static (bool success, Geodatabase geodatabase, string message) GetGDB_RTScratch()
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Get the RT Scratch Geodatabase Path
                string gdbpath = GetPath_RTScratchGDB();

                // Retrieve the File Geodatabase
                return GetFileGDB(gdbpath);
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve the national geodatabase (file or enterprise).  Geodatabase must be
        /// valid (e.g. have the required tables).  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static (bool success, Geodatabase geodatabase, GeoDBType gdbType, string message) GetGDB_Nat()
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Get the Nat Geodatabase Path
                string gdbpath = GetPath_NatGDB();

                // Get the Geodatabase
                var tryget = GetGDB(gdbpath);

                if (!tryget.success)
                {
                    return tryget;
                }

                // Ensure geodatabase is valid
                if (!Properties.Settings.Default.NATDB_DBVALID)
                {
                    return (false, null, GeoDBType.Unknown, "Geodatabase exists but is invalid.");
                }
                else
                {
                    return tryget;
                }
            }
            catch (Exception ex)
            {
                return (false, null, GeoDBType.Unknown, ex.Message);
            }
        }

        #endregion

        #region GENERIC GDB OBJECTS

        /// <summary>
        /// Retrieve a feature class by name from a specified geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="fc_name"></param>
        /// <returns></returns>
        public static (bool success, FeatureClass featureclass, string message) GetFC(Geodatabase geodatabase, string fc_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure geodatabase is not null
                if (geodatabase == null)
                {
                    return (false, null, "Null or invalid geodatabase.");
                }

                // Ensure feature class exists
                if (!FCExists(geodatabase, fc_name).exists)
                {
                    return (false, null, "Feature class not found in geodatabase");
                }

                // retrieve feature class
                FeatureClass featureClass = null;

                try
                {
                    featureClass = geodatabase.OpenDataset<FeatureClass>(fc_name);

                    return (true, featureClass, "success");
                }
                catch (Exception ex)
                {
                    return (false, null, $"Error opening {fc_name} feature class.\n{ex.Message}");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a table by name from a specified geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static (bool success, Table table, string message) GetTable(Geodatabase geodatabase, string table_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure geodatabase is not null
                if (geodatabase == null)
                {
                    return (false, null, "Null or invalid geodatabase.");
                }

                // Ensure table exists
                if (!TableExists(geodatabase, table_name).exists)
                {
                    return (false, null, "Table not found in geodatabase");
                }

                // Retrieve table
                Table table = null;

                try
                {
                    table = geodatabase.OpenDataset<Table>(table_name);

                    return (true, table, "success");
                }
                catch (Exception ex)
                {
                    return (false, null, $"Error opening {table_name} feature class.\n{ex.Message}");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a raster dataset by name from a specified geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="raster_name"></param>
        /// <returns></returns>
        public static (bool success, RasterDataset rasterDataset, string message) GetRaster(Geodatabase geodatabase, string raster_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure geodatabase is not null
                if (geodatabase == null)
                {
                    return (false, null, "Null or invalid geodatabase.");
                }

                // Ensure raster dataset exists
                if (!RasterExists(geodatabase, raster_name).exists)
                {
                    return (false, null, "Raster dataset not found in geodatabase");
                }

                // Retrieve raster dataset
                RasterDataset rasterDataset = null;

                try
                {
                    rasterDataset = geodatabase.OpenDataset<RasterDataset>(raster_name);

                    return (true, rasterDataset, "success");
                }
                catch (Exception ex)
                {
                    return (false, null, $"Error opening {raster_name} raster dataset.\n{ex.Message}");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a feature dataset by name from a specified geodatabase.
        /// Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="fds_name"></param>
        /// <returns></returns>
        public static (bool success, FeatureDataset featureDataset, string message) GetFDS(Geodatabase geodatabase, string fds_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure geodatabase is not null
                if (geodatabase == null)
                {
                    return (false, null, "Null or invalid geodatabase.");
                }

                // Ensure feature dataset exists
                if (!FDSExists(geodatabase, fds_name).exists)
                {
                    return (false, null, "Feature dataset not found in geodatabase");
                }

                // Retrieve feature dataset
                FeatureDataset featureDataset = null;

                try
                {
                    featureDataset = geodatabase.OpenDataset<FeatureDataset>(fds_name);

                    return (true, featureDataset, "success");
                }
                catch (Exception ex)
                {
                    return (false, null, $"Error opening {fds_name} feature dataset.\n{ex.Message}");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region PROJECT GDB OBJECTS

        /// <summary>
        /// Retrieve a feature class by name from the project file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="fc_name"></param>
        /// <returns></returns>
        public static (bool success, FeatureClass featureclass, string message) GetFC_Project(string fc_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_Project();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // ensure feature class exists
                    if (!FCExists(geodatabase, fc_name).exists)
                    {
                        return (false, null, "Feature class not found in geodatabase");
                    }

                    // get the feature class
                    FeatureClass featureClass = geodatabase.OpenDataset<FeatureClass>(fc_name);

                    return (true, featureClass, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a table by name from the project file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static (bool success, Table table, string message) GetTable_Project(string table_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_Project();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // ensure table exists
                    if (!TableExists(geodatabase, table_name).exists)
                    {
                        return (false, null, "Table not found in geodatabase");
                    }

                    // get the table
                    Table table = geodatabase.OpenDataset<Table>(table_name);

                    return (true, table, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a raster dataset by name from the project file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="raster_name"></param>
        /// <returns></returns>
        public static (bool success, RasterDataset rasterDataset, string message) GetRaster_Project(string raster_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_Project();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // ensure raster dataset exists
                    if (!RasterExists(geodatabase, raster_name).exists)
                    {
                        return (false, null, "Raster dataset not found in geodatabase");
                    }

                    // get the raster dataset
                    RasterDataset rasterDataset = geodatabase.OpenDataset<RasterDataset>(raster_name);

                    return (true, rasterDataset, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a feature dataset by name from the project file geodatabase.
        /// Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="fds_name"></param>
        /// <returns></returns>
        public static (bool success, FeatureDataset featureDataset, string message) GetFDS_Project(string fds_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_Project();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // ensure feature dataset exists
                    if (!FDSExists(geodatabase, fds_name).exists)
                    {
                        return (false, null, "Feature dataset not found in geodatabase");
                    }

                    // get the feature dataset
                    FeatureDataset featureDataset = geodatabase.OpenDataset<FeatureDataset>(fds_name);

                    return (true, featureDataset, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region RT SCRATCH GDB OBJECTS

        /// <summary>
        /// Retrieve a feature class by name from the RT Scratch file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="fc_name"></param>
        /// <returns></returns>
        public static (bool success, FeatureClass featureclass, string message) GetFC_RTScratch(string fc_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_RTScratch();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // ensure feature class exists
                    if (!FCExists(geodatabase, fc_name).exists)
                    {
                        return (false, null, "Feature class not found in geodatabase");
                    }

                    // get the feature class
                    FeatureClass featureClass = geodatabase.OpenDataset<FeatureClass>(fc_name);

                    return (true, featureClass, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a table by name from the RT Scratch file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static (bool success, Table table, string message) GetTable_RTScratch(string table_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_RTScratch();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // ensure table exists
                    if (!TableExists(geodatabase, table_name).exists)
                    {
                        return (false, null, "Table not found in geodatabase");
                    }

                    // get the table
                    Table table = geodatabase.OpenDataset<Table>(table_name);

                    return (true, table, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a raster dataset by name from the RT Scratch file geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="raster_name"></param>
        /// <returns></returns>
        public static (bool success, RasterDataset rasterDataset, string message) GetRaster_RTScratch(string raster_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_RTScratch();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // ensure raster dataset exists
                    if (!RasterExists(geodatabase, raster_name).exists)
                    {
                        return (false, null, "Raster dataset not found in geodatabase");
                    }

                    // get the raster dataset
                    RasterDataset rasterDataset = geodatabase.OpenDataset<RasterDataset>(raster_name);

                    return (true, rasterDataset, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region NAT GDB OBJECTS

        /// <summary>
        /// Retrieve a table by name from the national geodatabase (file or enterprise).
        /// Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="table_name"></param>
        /// <returns></returns>
        public static (bool success, Table table, string message) GetTable_Nat(string table_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_Nat();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    // fully qualified table name as required
                    SQLSyntax syntax = geodatabase.GetSQLSyntax();

                    string db = Properties.Settings.Default.NATDB_DBNAME;
                    string schema = Properties.Settings.Default.NATDB_SCHEMANAME;

                    string qualified_table_name = syntax.QualifyTableName(db, schema, table_name);

                    // ensure table exists
                    if (!TableExists(geodatabase, qualified_table_name).exists)
                    {
                        return (false, null, "Table not found in geodatabase");
                    }

                    // get the table
                    Table table = geodatabase.OpenDataset<Table>(qualified_table_name);

                    return (true, table, "success");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region FEATURE DATASET CONTENTS

        /// <summary>
        /// Retrieve the list of feature class names for the specified feature dataset
        /// in the provided geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="fds_name"></param>
        /// <returns></returns>
        public static (bool success, List<string> fc_names, string message) GetFCNamesFromFDS(Geodatabase geodatabase, string fds_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Ensure geodatabase is not null
                if (geodatabase == null)
                {
                    return (false, null, "Null or invalid geodatabase.");
                }

                // Ensure feature dataset exists
                if (!FDSExists(geodatabase, fds_name).exists)
                {
                    return (false, null, "Feature dataset not found in geodatabase");
                }

                // create the list of fc names
                List<string> fc_names = new List<string>();

                // populate the list
                using (FeatureDataset fds = geodatabase.OpenDataset<FeatureDataset>(fds_name))
                using (FeatureDatasetDefinition fdsDef = fds.GetDefinition())
                {
                    var fdDefs = geodatabase.GetRelatedDefinitions(fdsDef, DefinitionRelationshipType.DatasetInFeatureDataset).Where(d => d.DatasetType == DatasetType.FeatureClass);

                    foreach (var fcDef in fdDefs)
                    {
                        using (fcDef)
                        {
                            fc_names.Add(fcDef.GetName());
                        }
                    }
                }

                fc_names.Sort();

                return (true, fc_names, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of feature class names found within the supplied Feature Dataset,
        /// in the Project geodatabase.  Must be run on MCT.  Silent errors.
        /// </summary>
        /// <param name="fds_name"></param>
        /// <returns></returns>
        public static (bool success, List<string> fc_names, string message) GetFCNamesFromFDS_Project(string fds_name)
        {
            try
            {
                // Ensure this is called on worker thread
                if (!QueuedTask.OnWorker)
                {
                    throw new ArcGIS.Core.CalledOnWrongThreadException();
                }

                // Retrieve geodatabase
                var tryget = GetGDB_Project();
                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                // Get the fc names
                using (Geodatabase geodatabase = tryget.geodatabase)
                {
                    var getter = GetFCNamesFromFDS(geodatabase, fds_name);

                    if (!getter.success)
                    {
                        return (false, null, getter.message);
                    }
                    else
                    {
                        return (true, getter.fc_names, "success");
                    }
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region MISCELLANEOUS

        public static async Task<(bool success, string qualified_name, string message)> GetNatDBQualifiedName(string name)
        {
            try
            {
                // Ensure the national database is valid
                bool valid = Properties.Settings.Default.NATDB_DBVALID;

                if (!valid)
                {
                    return (false, "", "invalid national database");
                }

                // Construct the qualified name
                string db = Properties.Settings.Default.NATDB_DBNAME;
                string schema = Properties.Settings.Default.NATDB_SCHEMANAME;

                string qualified_name = "";

                await QueuedTask.Run(() =>
                {
                    // Get the nat gdb
                    var tryget = GetGDB_Nat();
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving valid national geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget.geodatabase)
                    {
                        SQLSyntax syntax = geodatabase.GetSQLSyntax();
                        qualified_name = syntax.QualifyTableName(db, schema, name);
                    }
                });

                return (true, qualified_name, "success");
            }
            catch (Exception ex)
            {
                return (false, "", ex.Message);
            }
        }

        #endregion

        #endregion

        #region PRZ LISTS AND DICTIONARIES

        #region NATIONAL TABLES

        /// <summary>
        /// Returns the national element table name for the supplied element id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static (bool success, string table_name, string message) GetNationalElementTableName(int element_id)
        {
            try
            {
                if (element_id > 99999 || element_id < 1)
                {
                    throw new Exception($"Element ID {element_id} is out of range (1 to 99999)");
                }
                else
                {
                    return (true, PRZC.c_TABLE_NATPRJ_PREFIX_ELEMENT + element_id.ToString("D5"), "success");
                }
            }
            catch (Exception ex)
            {
                return (false, "", ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of NatTheme objects from the project geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<NatTheme> themes, string message)> GetNationalThemes()
        {
            try
            {
                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, try_gdbexists.message);
                }

                // Check for existence of Theme table
                if (!(await TableExists_Project(PRZC.c_TABLE_NATPRJ_THEMES)).exists)
                {
                    return (false, null, $"{PRZC.c_TABLE_NATPRJ_THEMES} table not found in project geodatabase");
                }

                // Create the list
                List<NatTheme> themes = new List<NatTheme>();

                // Populate the list
                (bool success, string message) outcome = await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(PRZC.c_TABLE_NATPRJ_THEMES);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search())
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                int id = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATTHEME_THEME_ID]);
                                string name = (string)row[PRZC.c_FLD_TAB_NATTHEME_NAME];
                                string code = (string)row[PRZC.c_FLD_TAB_NATTHEME_CODE];
                                int theme_presence = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATTHEME_PRESENCE]);

                                if (id > 0 && !string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(code))
                                {
                                    NatTheme theme = new NatTheme()
                                    {
                                        ThemeID = id,
                                        ThemeName = name,
                                        ThemeCode = code,
                                        ThemePresence = theme_presence
                                    };

                                    themes.Add(theme);
                                }
                            }
                        }
                    }

                    return (true, "success");
                });

                if (outcome.success)
                {
                    // Sort the list by theme id
                    themes.Sort((a, b) => a.ThemeID.CompareTo(b.ThemeID));

                    return (true, themes, "success");
                }
                else
                {
                    return (false, null, outcome.message);
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of NatTheme objects from the project geodatabase, optionally filtered
        /// by the presence indicator.  Silent errors.
        /// </summary>
        /// <param name="presence"></param>
        /// <returns></returns>
        public static async Task<(bool success, List<NatTheme> themes, string message)> GetNationalThemes(ElementPresence? presence)
        {
            try
            {
                // Get the full Theme list
                var tryget = await GetNationalThemes();

                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                List<NatTheme> themes = tryget.themes;

                // Filter the list based on filter criteria:

                // By Presence
                IEnumerable<NatTheme> v = (presence != null) ? themes.Where(t => t.ThemePresence == ((int)presence)) : themes;

                // Sort by Theme ID
                IOrderedEnumerable<NatTheme> u = v.OrderBy(t => t.ThemeID);

                return (true, u.ToList(), "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of NatElement objects from the project geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<NatElement> elements, string message)> GetNationalElements()
        {
            try
            {
                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, try_gdbexists.message);
                }

                // Check for existence of Element table
                if (!(await TableExists_Project(PRZC.c_TABLE_NATPRJ_ELEMENTS)).exists)
                {
                    return (false, null, $"{PRZC.c_TABLE_NATPRJ_ELEMENTS} table not found in project geodatabase");
                }

                // Create list
                List<NatElement> elements = new List<NatElement>();

                // Populate the list
                await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(PRZC.c_TABLE_NATPRJ_ELEMENTS);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search())
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                int id = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_ELEMENT_ID]);
                                string name = (string)row[PRZC.c_FLD_TAB_NATELEMENT_NAME] ?? "";
                                int elem_type = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_TYPE]);
                                int elem_status = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_STATUS]);
                                string data_path = (string)row[PRZC.c_FLD_TAB_NATELEMENT_DATAPATH] ?? "";
                                int theme_id = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_THEME_ID]);
                                int elem_presence = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_PRESENCE]);
                                string unit = (string)row[PRZC.c_FLD_TAB_NATELEMENT_UNIT] ?? "";

                                if (id > 0 && elem_type > 0 && elem_status > 0 && theme_id > 0 && !string.IsNullOrEmpty(name))
                                {
                                    NatElement element = new NatElement()
                                    {
                                        ElementID = id,
                                        ElementName = name,
                                        ElementType = elem_type,
                                        ElementStatus = elem_status,
                                        ElementDataPath = data_path,
                                        ThemeID = theme_id,
                                        ElementPresence = elem_presence,
                                        ElementUnit = unit
                                    };

                                    elements.Add(element);
                                }
                            }
                        }
                    }
                });

                // Populate the Theme Information
                var theme_outcome = await GetNationalThemes();
                if (!theme_outcome.success)
                {
                    return (false, null, theme_outcome.message);
                }

                List<NatTheme> themes = theme_outcome.themes;

                foreach (NatElement element in elements)
                {
                    int theme_id = element.ThemeID;

                    if (theme_id < 1)
                    {
                        element.ThemeName = "INVALID THEME ID";
                        element.ThemeCode = "---";
                    }
                    else
                    {
                        NatTheme theme = themes.FirstOrDefault(t => t.ThemeID == theme_id);

                        if (theme != null)
                        {
                            element.ThemeName = theme.ThemeName;
                            element.ThemeCode = theme.ThemeCode;
                        }
                        else
                        {
                            element.ThemeName = "NO CORRESPONDING THEME";
                            element.ThemeCode = "???";
                        }
                    }
                }

                // Sort the list
                elements.Sort((a, b) => a.ElementID.CompareTo(b.ElementID));

                return (true, elements, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of NatElement objects from the project geodatabase, optionally filtered
        /// by type, status, or presence indicators.  Silent errors.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="status"></param>
        /// <param name="presence"></param>
        /// <returns></returns>
        public static async Task<(bool success, List<NatElement> elements, string message)> GetNationalElements(ElementType? type, ElementStatus? status, ElementPresence? presence)
        {
            try
            {
                // Get the full Elements list
                var tryget = await GetNationalElements();

                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                List<NatElement> elements = tryget.elements;

                // Filter the list based on filter criteria:

                // By Type
                IEnumerable<NatElement> v = (type != null) ? elements.Where(e => e.ElementType == ((int)type)) : elements;

                // By Status
                v = (status != null) ? v.Where(e => e.ElementStatus == ((int)status)) : v;

                // By Presence
                v = (presence != null) ? v.Where(e => e.ElementPresence == ((int)presence)) : v;

                // Sort by Element ID
                IOrderedEnumerable<NatElement> u = v.OrderBy(e => e.ElementID);

                return (true, u.ToList(), "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of national element table names (e.g. n00042) from the project geodatabase.
        /// Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<string> tables, string message)> GetNationalElementTables()
        {
            try
            {
                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, try_gdbexists.message);
                }

                // Create the list
                List<string> table_names = new List<string>();

                // Populate the list
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = GetGDB_Project();
                    if (!tryget_gdb.success)
                    {
                        throw new Exception("Error opening project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        var table_defs = geodatabase.GetDefinitions<TableDefinition>();

                        foreach (TableDefinition table_def in table_defs)
                        {
                            using (table_def)
                            {
                                string name = table_def.GetName();

                                if (name.Length == 6 & name.StartsWith(PRZC.c_TABLE_NATPRJ_PREFIX_ELEMENT))
                                {
                                    string numpart = name.Substring(1);

                                    if (int.TryParse(numpart, out int result))
                                    {
                                        // this is a national element table
                                        table_names.Add(name);
                                    }
                                }
                            }
                        }
                    }
                });

                return (true, table_names, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of national element feature class names (e.g. fc_n00042) from the
        /// project geodatabase, from the national feature dataset.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<string> feature_classes, string message)> GetNationalElementFCs()
        {
            try
            {
                // Define the new list
                List<string> matching_fc_names = new List<string>();

                // Retrieve names
                await QueuedTask.Run(() =>
                {
                    // Get FC Names from the FDS
                    var tryget_fcnames = GetFCNamesFromFDS_Project(PRZC.c_FDS_NATIONAL_ELEMENTS);
                    if (!tryget_fcnames.success)
                    {
                        throw new Exception(tryget_fcnames.message);
                    }
                    var all_fc_names = tryget_fcnames.fc_names;

                    // Select fc names matching pattern
                    foreach (string fc_name in all_fc_names)
                    {
                        if (fc_name.StartsWith($"fc_{PRZC.c_TABLE_NATPRJ_PREFIX_ELEMENT}", StringComparison.OrdinalIgnoreCase) & fc_name.Length == 9)
                        {
                            string numpart = fc_name.Substring(4);

                            if (int.TryParse(numpart, out int result))
                            {
                                // this is a national element table
                                matching_fc_names.Add(fc_name);
                            }
                        }
                    }
                });

                matching_fc_names.Sort();

                return (true, matching_fc_names, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a NatElement object from the national element table in the project GDB,
        /// based on the supplied element id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static async Task<(bool success, NatElement element, string message)> GetNationalElement(int element_id)
        {
            try
            {
                // Ensure valid element_id (>0)
                if (element_id <= 0)
                {
                    throw new Exception("element_id value <= zero");
                }

                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    throw new Exception("project geodatabase does not exist");
                }

                // Check for existence of Element table
                if (!(await TableExists_Project(PRZC.c_TABLE_NATPRJ_ELEMENTS)).exists)
                {
                    throw new Exception("nat element table not found");
                }

                // Create the new element
                NatElement element = null;

                // Fill out the element
                await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(PRZC.c_TABLE_NATPRJ_ELEMENTS);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    QueryFilter queryFilter = new QueryFilter
                    {
                        WhereClause = $"{PRZC.c_FLD_TAB_NATELEMENT_ELEMENT_ID} = {element_id}"
                    };

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search(queryFilter, false))
                    {
                        if (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                int id = element_id;
                                string name = (string)row[PRZC.c_FLD_TAB_NATELEMENT_NAME] ?? "";
                                int elem_type = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_TYPE]);
                                int elem_status = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_STATUS]);
                                string data_path = (string)row[PRZC.c_FLD_TAB_NATELEMENT_DATAPATH] ?? "";
                                int theme_id = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_THEME_ID]);
                                int elem_presence = Convert.ToInt32(row[PRZC.c_FLD_TAB_NATELEMENT_PRESENCE]);
                                string unit = (string)row[PRZC.c_FLD_TAB_NATELEMENT_UNIT] ?? "";

                                if (id > 0)
                                {
                                    element = new NatElement()
                                    {
                                        ElementID = id,
                                        ElementName = name,
                                        ElementType = elem_type,
                                        ElementStatus = elem_status,
                                        ElementDataPath = data_path,
                                        ThemeID = theme_id,
                                        ElementPresence = elem_presence,
                                        ElementUnit = unit
                                    };
                                }
                            }
                        }
                    }
                });

                if (element != null)
                {
                    return (true, element, "success");
                }
                else
                {
                    return (false, null, "element not found");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region REGIONAL TABLES

        /// <summary>
        /// Returns the regional element table name for the supplied element id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static (bool success, string table_name, string message) GetRegionalElementTableName(int element_id)
        {
            try
            {
                if (element_id > 99999 || element_id < 1)
                {
                    throw new Exception($"Element ID {element_id} is out of range (1 to 99999)");
                }
                else
                {
                    return (true, PRZC.c_TABLE_REGPRJ_PREFIX_ELEMENT + element_id.ToString("D5"), "success");
                }
            }
            catch (Exception ex)
            {
                return (false, "", ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of RegElement objects from the project geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<RegElement> elements, string message)> GetRegionalElements()
        {
            try
            {
                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, try_gdbexists.message);
                }

                // Check for existence of Element table
                if (!(await TableExists_Project(PRZC.c_TABLE_REGPRJ_ELEMENTS)).exists)
                {
                    return (false, null, $"{PRZC.c_TABLE_REGPRJ_ELEMENTS} table not found in project geodatabase");
                }

                // Create list
                List<RegElement> elements = new List<RegElement>();

                // Populate the list
                await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(PRZC.c_TABLE_REGPRJ_ELEMENTS);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search())
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                int elem_id = Convert.ToInt32(row[PRZC.c_FLD_TAB_REGELEMENT_ELEMENT_ID]);
                                string elem_name = (string)row[PRZC.c_FLD_TAB_REGELEMENT_NAME] ?? "";
                                int elem_type = Convert.ToInt32(row[PRZC.c_FLD_TAB_REGELEMENT_TYPE]);
                                int elem_status = Convert.ToInt32(row[PRZC.c_FLD_TAB_REGELEMENT_STATUS]);
                                int elem_presence = Convert.ToInt32(row[PRZC.c_FLD_TAB_REGELEMENT_PRESENCE]);
                                string elem_lyrxpath = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LYRXPATH] ?? "";
                                string elem_lyrxname = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LAYERNAME] ?? "";
                                int elem_lyrxtype = Convert.ToInt32(Enum.Parse(typeof(LayerType), (string)row[PRZC.c_FLD_TAB_REGELEMENT_LAYERTYPE], true));
                                string elem_lyrxjson = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LAYERJSON] ?? "";
                                string elem_whereclause = (string)row[PRZC.c_FLD_TAB_REGELEMENT_WHERECLAUSE] ?? "";
                                string elem_legendgroup = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LEGENDGROUP] ?? "";
                                string elem_legendclass = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LEGENDCLASS] ?? "";

                                object oThemeID = row[PRZC.c_FLD_TAB_REGELEMENT_THEME_ID] ?? -9999;
                                int elem_themeid = Convert.ToInt32(oThemeID);

                                if (elem_id > 0 && elem_type > 0 && elem_status > 0 && !string.IsNullOrEmpty(elem_name))
                                {
                                    RegElement regElement = new RegElement()
                                    {
                                        ElementID = elem_id,
                                        ElementName = elem_name,
                                        ElementType = elem_type,
                                        ElementStatus = elem_status,
                                        ElementPresence = elem_presence,
                                        LayerName = elem_lyrxname,
                                        LayerType = elem_lyrxtype,
                                        LayerJson = elem_lyrxjson,
                                        WhereClause = elem_whereclause,
                                        LegendGroup = elem_legendgroup,
                                        LegendClass = elem_legendclass,
                                        ThemeID = elem_themeid
                                    };

                                    elements.Add(regElement);
                                }
                            }
                        }
                    }
                });

                // Populate regional theme info
                var tryget_themes = await GetRegionalThemesDomainKVPs();

                if (!tryget_themes.success)
                {
                    return (false, null, tryget_themes.message);
                }

                Dictionary<int, string> themes = tryget_themes.dict;

                foreach (RegElement element in elements)
                {
                    int theme_id = element.ThemeID;

                    if (themes.ContainsKey(theme_id))
                    {
                        element.ThemeName = themes[theme_id];
                    }
                    else if (theme_id == -9999)
                    {
                        element.ThemeName = "Regional Theme Not Specified";
                    }
                    else
                    {
                        element.ThemeName = "Invalid Regional Theme";
                    }
                }

                // Sort the list
                elements.Sort((a, b) => a.ElementID.CompareTo(b.ElementID));

                return (true, elements, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of RegElement objects from the project geodatabase, optionally filtered
        /// by type, status, or presence indicators.  Silent errors.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="status"></param>
        /// <param name="presence"></param>
        /// <returns></returns>
        public static async Task<(bool success, List<RegElement> elements, string message)> GetRegionalElements(ElementType? type, ElementStatus? status, ElementPresence? presence)
        {
            try
            {
                // Get the full Elements list
                var tryget = await GetRegionalElements();

                if (!tryget.success)
                {
                    return (false, null, tryget.message);
                }

                List<RegElement> elements = tryget.elements;

                // Filter the list based on filter criteria:

                // By Type
                IEnumerable<RegElement> v = (type != null) ? elements.Where(e => e.ElementType == ((int)type)) : elements;

                // By Status
                v = (status != null) ? v.Where(e => e.ElementStatus == ((int)status)) : v;

                // By Presence
                v = (presence != null) ? v.Where(e => e.ElementPresence == ((int)presence)) : v;

                // Sort by Element ID
                IOrderedEnumerable<RegElement> u = v.OrderBy(e => e.ElementID);

                return (true, u.ToList(), "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of regional element table names (e.g. r00042) from the project geodatabase.
        /// Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<string> tables, string message)> GetRegionalElementTables()
        {
            try
            {
                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, try_gdbexists.message);
                }

                // Create the list
                List<string> table_names = new List<string>();

                // Populate the list
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = GetGDB_Project();
                    if (!tryget_gdb.success)
                    {
                        throw new Exception("Error opening project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        var table_defs = geodatabase.GetDefinitions<TableDefinition>();

                        foreach (TableDefinition table_def in table_defs)
                        {
                            using (table_def)
                            {
                                string name = table_def.GetName();

                                if (name.Length == 6 & name.StartsWith(PRZC.c_TABLE_REGPRJ_PREFIX_ELEMENT))
                                {
                                    string numpart = name.Substring(1);

                                    if (int.TryParse(numpart, out int result))
                                    {
                                        // this is a regional element table
                                        table_names.Add(name);
                                    }
                                }
                            }
                        }
                    }
                });

                return (true, table_names, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a list of regional element raster names.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<string> rasters, string message)> GetRegionalElementRasters()
        {
            try
            {
                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, try_gdbexists.message);
                }

                // Create the list
                List<string> raster_names = new List<string>();

                // Populate the list
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = GetGDB_Project();
                    if (!tryget_gdb.success)
                    {
                        throw new Exception("Error opening project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        var rasDefs = geodatabase.GetDefinitions<RasterDatasetDefinition>();

                        foreach (RasterDatasetDefinition rasDef in rasDefs)
                        {
                            using (rasDef)
                            {
                                string name = rasDef.GetName();

                                if (name.StartsWith(PRZC.c_RAS_REG_ELEM_PREFIX, StringComparison.OrdinalIgnoreCase))
                                {
                                    raster_names.Add(name);
                                }
                            }
                        }
                    }
                });

                return (true, raster_names, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        ///// <summary>
        ///// Retrieve a list of regional pu raster reclass raster names.  Silent errors.
        ///// </summary>
        ///// <returns></returns>
        //public static async Task<(bool success, List<string> rasters, string message)> GetRegionalPUReclassRasters()
        //{
        //    try
        //    {
        //        // Check for Project GDB
        //        var try_gdbexists = await GDBExists_Project();
        //        if (!try_gdbexists.exists)
        //        {
        //            return (false, null, try_gdbexists.message);
        //        }

        //        // Create the list
        //        List<string> raster_names = new List<string>();

        //        // Populate the list
        //        await QueuedTask.Run(() =>
        //        {
        //            var tryget_gdb = GetGDB_Project();
        //            if (!tryget_gdb.success)
        //            {
        //                throw new Exception("Error opening project geodatabase.");
        //            }

        //            using (Geodatabase geodatabase = tryget_gdb.geodatabase)
        //            {
        //                var rasDefs = geodatabase.GetDefinitions<RasterDatasetDefinition>();

        //                foreach (RasterDatasetDefinition rasDef in rasDefs)
        //                {
        //                    using (rasDef)
        //                    {
        //                        string name = rasDef.GetName();

        //                        if (name.StartsWith(PRZC.c_RAS_PLANNING_UNITS_RECLASS, StringComparison.OrdinalIgnoreCase))
        //                        {
        //                            raster_names.Add(name);
        //                        }
        //                    }
        //                }
        //            }
        //        });

        //        return (true, raster_names, "success");
        //    }
        //    catch (Exception ex)
        //    {
        //        return (false, null, ex.Message);
        //    }
        //}

        /// <summary>
        /// Retrieve a list of regional element feature class names (e.g. fc_r00042) from the
        /// project geodatabase, from the regional feature dataset.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<string> feature_classes, string message)> GetRegionalElementFCs()
        {
            try
            {
                // Define the new list
                List<string> matching_fc_names = new List<string>();

                // Retrieve names
                await QueuedTask.Run(() =>
                {
                    // Get FC Names from the FDS
                    var tryget_fcnames = GetFCNamesFromFDS_Project(PRZC.c_FDS_REGIONAL_ELEMENTS);
                    if (!tryget_fcnames.success)
                    {
                        throw new Exception(tryget_fcnames.message);
                    }
                    var all_fc_names = tryget_fcnames.fc_names;

                    // Select fc names matching pattern
                    foreach (string fc_name in all_fc_names)
                    {
                        if (fc_name.StartsWith($"fc_{PRZC.c_TABLE_REGPRJ_PREFIX_ELEMENT}", StringComparison.OrdinalIgnoreCase) & fc_name.Length == 9)
                        {
                            string numpart = fc_name.Substring(4);

                            if (int.TryParse(numpart, out int result))
                            {
                                // this is a regional element table
                                matching_fc_names.Add(fc_name);
                            }
                        }
                    }
                });

                matching_fc_names.Sort();

                return (true, matching_fc_names, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a RegElement object from the regional element table, based on supplied
        /// element id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static async Task<(bool success, RegElement element, string message)> GetRegionalElement(int element_id)
        {
            try
            {
                // Ensure valid element_id (>0)
                if (element_id <= 0)
                {
                    throw new Exception("element_id value <= zero");
                }

                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    throw new Exception("project geodatabase does not exist");
                }

                // Check for existence of Element table
                if (!(await TableExists_Project(PRZC.c_TABLE_REGPRJ_ELEMENTS)).exists)
                {
                    throw new Exception("reg element table not found");
                }

                // Create the new element
                RegElement element = null;

                // Fill out the element
                await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(PRZC.c_TABLE_REGPRJ_ELEMENTS);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    QueryFilter queryFilter = new QueryFilter
                    {
                        WhereClause = $"{PRZC.c_FLD_TAB_REGELEMENT_ELEMENT_ID} = {element_id}"
                    };

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search(queryFilter, false))
                    {
                        if (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                string elem_name = (string)row[PRZC.c_FLD_TAB_REGELEMENT_NAME] ?? "";
                                int elem_type = Convert.ToInt32(row[PRZC.c_FLD_TAB_REGELEMENT_TYPE]);
                                int elem_status = Convert.ToInt32(row[PRZC.c_FLD_TAB_REGELEMENT_STATUS]);
                                int elem_presence = Convert.ToInt32(row[PRZC.c_FLD_TAB_REGELEMENT_PRESENCE]);
                                string elem_lyrxpath = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LYRXPATH] ?? "";
                                string elem_lyrxname = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LAYERNAME] ?? "";
                                int elem_lyrxtype = Convert.ToInt32(Enum.Parse(typeof(LayerType), (string)row[PRZC.c_FLD_TAB_REGELEMENT_LAYERTYPE], true));
                                string elem_lyrxjson = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LAYERJSON] ?? "";
                                string elem_whereclause = (string)row[PRZC.c_FLD_TAB_REGELEMENT_WHERECLAUSE] ?? "";
                                string elem_legendgroup = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LEGENDGROUP] ?? "";
                                string elem_legendclass = (string)row[PRZC.c_FLD_TAB_REGELEMENT_LEGENDCLASS] ?? "";

                                object oThemeID = row[PRZC.c_FLD_TAB_REGELEMENT_THEME_ID] ?? -9999;
                                int elem_themeid = Convert.ToInt32(oThemeID);

                                element = new RegElement()
                                {
                                    ElementID = element_id,
                                    ElementName = elem_name,
                                    ElementType = elem_type,
                                    ElementStatus = elem_status,
                                    ElementPresence = elem_presence,
                                    LayerName = elem_lyrxname,
                                    LayerType = elem_lyrxtype,
                                    LayerJson = elem_lyrxjson,
                                    WhereClause = elem_whereclause,
                                    LegendGroup = elem_legendgroup,
                                    LegendClass = elem_legendclass,
                                    ThemeID = elem_themeid
                                };
                            }
                        }
                    }
                });

                if (element != null)
                {
                    return (true, element, "success");
                }
                else
                {
                    return (false, null, "element not found");
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region NATIONAL ELEMENT TABLE VALUES

        /// <summary>
        /// Retrieve a national grid value for a specified element and cell number.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <param name="cell_number"></param>
        /// <returns></returns>
        public static async Task<(bool success, double value, string message)> GetValueFromNatElementTable_CellNum(int element_id, long cell_number)
        {
            double value = -9999;

            try
            {
                // Ensure valid element id
                if (element_id < 1 || element_id > 99999)
                {
                    return (false, value, "Element ID out of range (1 - 99999)");
                }

                // Get element table name
                var trygettab = GetNationalElementTableName(element_id);
                if (!trygettab.success)
                {
                    return (false, value, "Unable to retrieve element table name");
                }

                string table_name = trygettab.table_name;

                // Check for Project GDB
                if (!(await GDBExists_Project()).exists)
                {
                    return (false, value, "Project GDB not found.");
                }

                // Verify that table exists in project GDB
                if (!(await TableExists_Project(table_name)).exists)
                {
                    return (false, value, $"Element table {table_name} not found in project geodatabase");
                }

                // retrieve the value for the cell number in the element table
                (bool success, string message) outcome = await QueuedTask.Run(() =>
                {
                    QueryFilter queryFilter = new QueryFilter
                    {
                        WhereClause = PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER + " = " + cell_number
                    };

                    var tryget = GetTable_Project(table_name);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    {
                        // Row Count
                        long rows = table.GetCount(queryFilter);

                        if (rows == 1)
                        {
                            using (RowCursor rowCursor = table.Search(queryFilter))
                            {
                                if (rowCursor.MoveNext())
                                {
                                    using (Row row = rowCursor.Current)
                                    {
                                        value = Convert.ToDouble(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_VALUE]);
                                        return (true, "success");
                                    }
                                }
                                else
                                {
                                    return (false, "no match found");
                                }
                            }
                        }
                        else if (rows == 0)
                        {
                            return (false, "no match found");
                        }
                        else if (rows > 1)
                        {
                            return (false, "more than one matching cell number found");
                        }
                        else
                        {
                            return (false, "there was a resounding kaboom");
                        }
                    }
                });

                if (outcome.success)
                {
                    return (true, value, "success");
                }
                else
                {
                    return (false, value, outcome.message);
                }
            }
            catch (Exception ex)
            {
                return (false, value, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a national grid value for a specified element and planning unit id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <param name="puid"></param>
        /// <returns></returns>
        public static async Task<(bool success, double value, string message)> GetValueFromNatElementTable_PUID(int element_id, int puid)
        {
            double value = -9999;

            try
            {
                // Ensure valid element id
                if (element_id < 1 || element_id > 99999)
                {
                    return (false, value, "Element ID out of range (1 - 99999)");
                }

                // Get element table name
                var trygetname = GetNationalElementTableName(element_id);

                if (!trygetname.success)
                {
                    return (false, value, "Unable to retrieve element table name");
                }

                string table_name = trygetname.table_name;

                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, value, "Project GDB not found.");
                }

                // Verify that table exists in project GDB
                if (!(await TableExists_Project(table_name)).exists)
                {
                    return (false, value, $"Element table {table_name} not found in project geodatabase");
                }

                // Retrieve the cell number for the provided puid
                var result = await GetCellNumberFromPUID(puid);
                if (!result.success)
                {
                    return (false, value, result.message);
                }

                long cell_number = result.cell_number;

                // retrieve the value for the cell number in the element table
                (bool success, string message) outcome = await QueuedTask.Run(() =>
                {
                    QueryFilter queryFilter = new QueryFilter
                    {
                        WhereClause = PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER + " = " + cell_number
                    };

                    var tryget = GetTable_Project(table_name);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving the table.");
                    }

                    using (Table table = tryget.table)
                    {
                        // Row Count
                        long rows = table.GetCount(queryFilter);

                        if (rows == 1)
                        {
                            using (RowCursor rowCursor = table.Search(queryFilter))
                            {
                                if (rowCursor.MoveNext())
                                {
                                    using (Row row = rowCursor.Current)
                                    {
                                        value = Convert.ToDouble(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_VALUE]);
                                        return (true, "success");
                                    }
                                }
                                else
                                {
                                    return (false, "no match found");
                                }
                            }
                        }
                        else if (rows == 0)
                        {
                            return (false, "no match found");
                        }
                        else if (rows > 1)
                        {
                            return (false, "more than one matching cell number found");
                        }
                        else
                        {
                            return (false, "there was a resounding kaboom");
                        }
                    }
                });

                if (outcome.success)
                {
                    return (true, value, "success");
                }
                else
                {
                    return (false, value, outcome.message);
                }
            }
            catch (Exception ex)
            {
                return (false, value, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a dictionary of cell numbers and associated element values from the
        /// project geodatabase, for the specified element.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<long, double> dict, string message)> GetValuesFromNatElementTable_CellNum(int element_id)
        {
            try
            {
                // Ensure valid element id
                if (element_id < 1 || element_id > 99999)
                {
                    return (false, null, "Element ID out of range (1 - 99999)");
                }

                // Get element table name
                var trygetname = GetNationalElementTableName(element_id);

                if (!trygetname.success)
                {
                    return (false, null, "Unable to retrieve element table name");
                }

                string table_name = trygetname.table_name;

                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, "Project GDB not found.");
                }

                // Verify that table exists in project GDB
                if (!(await TableExists_Project(table_name)).exists)
                {
                    return (false, null, $"Element table {table_name} not found in project geodatabase");
                }

                // Create the dictionary
                Dictionary<long, double> dict = new Dictionary<long, double>();

                // Populate the dictionary
                await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(table_name);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search())
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                long cellnum = Convert.ToInt64(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER]);
                                double cellval = Convert.ToDouble(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_VALUE]);

                                if (cellnum > 0 && !dict.ContainsKey(cellnum))
                                {
                                    dict.Add(cellnum, cellval);
                                }
                            }
                        }
                    }
                });

                return (true, dict, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a dictionary of planning unit ids and associated element values from the 
        /// project geodatabase, for the specified element id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<int, double> dict, string message)> GetValuesFromNatElementTable_PUID(int element_id)
        {
            try
            {
                // Ensure valid element id
                if (element_id < 1 || element_id > 99999)
                {
                    return (false, null, "Element ID out of range (1 - 99999)");
                }

                // Get element table name
                var trygetname = GetNationalElementTableName(element_id);

                if (!trygetname.success)
                {
                    return (false, null, "Unable to retrieve element table name");
                }

                string table_name = trygetname.table_name;

                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    return (false, null, "Project GDB not found.");
                }

                // Verify that table exists in project GDB
                if (!(await TableExists_Project(table_name)).exists)
                {
                    return (false, null, $"Element table {table_name} not found in project geodatabase");
                }

                // Get the dictionary of Cell Numbers > PUIDs
                var outcome = await GetCellNumbersAndPUIDs();
                if (!outcome.success)
                {
                    return (false, null, $"Unable to retrieve Cell Number dictionary\n{outcome.message}");
                }
                Dictionary<long, int> cellnumdict = outcome.dict;

                // Create the dictionary
                Dictionary<int, double> dict = new Dictionary<int, double>();

                // Populate the dictionary
                (bool success, string message) result = await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(table_name);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search())
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                long cellnum = Convert.ToInt64(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER]);
                                double cellval = Convert.ToDouble(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_VALUE]);

                                if (cellnum > 0)
                                {
                                    if (cellnumdict.ContainsKey(cellnum))
                                    {
                                        int puid = cellnumdict[cellnum];

                                        if (puid > 0 && !dict.ContainsKey(puid))
                                        {
                                            dict.Add(puid, cellval);
                                        }
                                    }
                                    else
                                    {
                                        return (false, $"No matching puid for cell number {cellnum}");
                                    }
                                }
                            }
                        }
                    }

                    return (true, "success");
                });

                if (result.success)
                {
                    return (true, dict, "success");
                }
                else
                {
                    return (false, null, result.message);
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a dictionary of cell numbers and associated element values from the
        /// national database, for the specified element and list of cell numbers.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <param name="cell_numbers"></param>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<long, double> dict, string message)> GetNatElementIntersection(int element_id, HashSet<long> cell_numbers)
        {
            try
            {
                // Ensure valid element id
                if (element_id < 1 || element_id > 99999)
                {
                    throw new Exception("Element ID out of range (1 - 99999)");
                }

                // Get element table name
                var trygetname = GetNationalElementTableName(element_id);
                if (!trygetname.success)
                {
                    throw new Exception("Unable to retrieve element table name");
                }
                string table_name = trygetname.table_name;  // unqualified table name

                // Create the dictionaries
                Dictionary<long, double> dict_final = new Dictionary<long, double>();
                Dictionary<long, double> dict_test = new Dictionary<long, double>();

                // Get the min and max cell numbers.  I can ignore all cell numbers outside this range
                long min_cell_number = cell_numbers.Min();
                long max_cell_number = cell_numbers.Max();

                // Populate dictionary
                await QueuedTask.Run(() =>
                {
                    // try getting the e0000n table
                    var trygettab = GetTable_Nat(table_name);

                    if (!trygettab.success)
                    {
                        throw new Exception("Unable to retrieve table.");
                    }

                    // Retrieve all element table KVPs within the min max range
                    using (Table table = trygettab.table)
                    {
                        QueryFilter queryFilter = new QueryFilter
                        {
                            SubFields = PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER + "," + PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_VALUE,
                            WhereClause = $"{PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER} BETWEEN {min_cell_number} AND {max_cell_number}"
                        };

                        using (RowCursor rowCursor = table.Search(queryFilter))
                        {
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    var cn = Convert.ToInt64(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_NUMBER]);
                                    var cv = Convert.ToDouble(row[PRZC.c_FLD_TAB_NAT_ELEMVAL_CELL_VALUE]);

                                    dict_test.Add(cn, cv);
                                }
                            }
                        }
                    }
                });

                // Populate the final dictionary
                foreach (long cellnum in cell_numbers)
                {
                    if (dict_test.ContainsKey(cellnum))
                    {
                        dict_final.Add(cellnum, dict_test[cellnum]);
                    }
                }

                return (true, dict_final, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #region REGIONAL ELEMENT TABLE VALUES

        /// <summary>
        /// Retrieve a dictionary of planning unit ids and associated element values from the 
        /// project geodatabase, for the specified regional element id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<int, double> dict, string message)> GetValuesFromRegElementTable_PUID(int element_id)
        {
            try
            {
                // Ensure valid element id
                if (element_id < 1 || element_id > 99999)
                {
                    throw new Exception("Element ID out of range (1-99999)");
                }

                // Get element table name
                var trygetname = GetRegionalElementTableName(element_id);

                if (!trygetname.success)
                {
                    throw new Exception("Unable to retrieve regional element table name.");
                }

                string table_name = trygetname.table_name;

                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    throw new Exception("Project geodatabase not found.");
                }

                // Verify that table exists in project GDB
                if (!(await TableExists_Project(table_name)).exists)
                {
                    throw new Exception("Element table not found.");
                }

                // Create the puid => value dictionary
                Dictionary<int, double> dict = new Dictionary<int, double>();

                // Populate the dictionary
                await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(table_name);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search())
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                int pu_id = Convert.ToInt32(row[PRZC.c_FLD_TAB_REG_ELEMVAL_PU_ID]);
                                double value = Convert.ToDouble(row[PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_VALUE]);

                                if (pu_id > 0 && !dict.ContainsKey(pu_id))
                                {
                                    dict.Add(pu_id, value);
                                }
                            }
                        }
                    }
                });

                return (true, dict, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a dictionary of cell numbers and associated element values from the
        /// project geodatabase, for the specified regional element id.  Silent errors.
        /// </summary>
        /// <param name="element_id"></param>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<long, double> dict, string message)> GetValuesFromRegElementTable_CellNum(int element_id)
        {
            try
            {
                // Ensure valid element id
                if (element_id < 1 || element_id > 99999)
                {
                    throw new Exception("Element ID out of range (1-99999)");
                }

                // Get element table name
                var trygetname = GetRegionalElementTableName(element_id);

                if (!trygetname.success)
                {
                    throw new Exception("Unable to retrieve regional element table name.");
                }

                string table_name = trygetname.table_name;

                // Check for Project GDB
                var try_gdbexists = await GDBExists_Project();
                if (!try_gdbexists.exists)
                {
                    throw new Exception("Project geodatabase not found.");
                }

                // Verify that table exists in project GDB
                if (!(await TableExists_Project(table_name)).exists)
                {
                    throw new Exception("Element table not found.");
                }

                // Create the puid => value dictionary
                Dictionary<long, double> dict = new Dictionary<long, double>();

                // Populate the dictionary
                await QueuedTask.Run(() =>
                {
                    var tryget = GetTable_Project(table_name);
                    if (!tryget.success)
                    {
                        throw new Exception("Error retrieving table.");
                    }

                    using (Table table = tryget.table)
                    using (RowCursor rowCursor = table.Search())
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                long cellnum = Convert.ToInt64(row[PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_NUMBER]);
                                double value = Convert.ToDouble(row[PRZC.c_FLD_TAB_REG_ELEMVAL_CELL_VALUE]);

                                if (cellnum > 0 && !dict.ContainsKey(cellnum))
                                {
                                    dict.Add(cellnum, value);
                                }
                            }
                        }
                    }
                });

                return (true, dict, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #endregion

        #region PUID AND CELL NUMBERS

        #region SINGLE VALUES

        /// <summary>
        /// Retrieve the national grid cell number associated with the specified
        /// planning unit id.  Silent errors.
        /// </summary>
        /// <param name="pu_id"></param>
        /// <returns></returns>
        public static async Task<(bool success, long cell_number, string message)> GetCellNumberFromPUID(int pu_id)
        {
            try
            {
                // Project GDB existence
                if (!(await GDBExists_Project()).exists)
                {
                    throw new Exception("Project GDB not found.");
                }

                // Planning Units existence
                var tryget_pudata = await PUDataExists();
                if (!tryget_pudata.exists)
                {
                    throw new Exception("Planning Units not found.");
                }
                else if (!tryget_pudata.national_enabled)
                {
                    throw new Exception("Planning Units not national enabled.");
                }

                bool success = false;
                string message = "";
                long cell_number = -9999;

                // Retrieve the Cell Number
                await QueuedTask.Run(() =>
                {
                    // Use the Planning Units Feature Class
                    var tryget_pufc = GetFC_Project(PRZC.c_FC_PLANNING_UNITS);

                    // Build query filter
                    QueryFilter queryFilter = new QueryFilter()
                    {
                        SubFields = PRZC.c_FLD_FC_PU_ID + "," + PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER,
                        WhereClause = $"{PRZC.c_FLD_FC_PU_ID} = {pu_id}"
                    };

                    using (Table table = tryget_pufc.featureclass)
                    using (RowCursor rowCursor = table.Search(queryFilter, false))
                    {
                        if (rowCursor.MoveNext())   // only check the first record
                        {
                            // Record found for pu_id
                            using (Row row = rowCursor.Current)
                            {
                                if (row[PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER] != null)
                                {
                                    cell_number = Convert.ToInt64(row[PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER]);
                                    success = true;
                                    message = "success";
                                }
                                else
                                {
                                    // null cell_number value (no need to do anything here)
                                    success = false;    // for clarity
                                    message = "null cell number value";
                                }
                            }
                        }
                        else
                        {
                            // no record found for pu_id (no need to do anything here)
                            success = false;        // for clarity
                            message = "no pu_id record found";
                        }
                    }
                });

                return (success, cell_number, message);
            }
            catch (Exception ex)
            {
                return (false, -9999, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve the planning unit id associated with the specified
        /// national grid cell number.  Silent errors.
        /// </summary>
        /// <param name="cell_number"></param>
        /// <returns></returns>
        public static async Task<(bool success, int pu_id, string message)> GetPUIDFromCellNumber(long cell_number)
        {
            try
            {
                // Project GDB existence
                if (!(await GDBExists_Project()).exists)
                {
                    throw new Exception("Project GDB not found.");
                }

                // Planning Units existence
                var tryget_pudata = await PUDataExists();
                if (!tryget_pudata.exists)
                {
                    throw new Exception("Planning Units not found.");
                }
                else if (!tryget_pudata.national_enabled)
                {
                    throw new Exception("Planning Units not national enabled.");
                }

                bool success = false;
                string message = "";
                int pu_id = -9999;

                // Retrieve the PU ID
                await QueuedTask.Run(() =>
                {
                    // Use the Planning Units Feature Class
                    var tryget_pufc = GetFC_Project(PRZC.c_FC_PLANNING_UNITS);

                    // Build query filter
                    QueryFilter queryFilter = new QueryFilter()
                    {
                        SubFields = PRZC.c_FLD_FC_PU_ID + "," + PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER,
                        WhereClause = $"{PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER} = {cell_number}"
                    };

                    using (Table table = tryget_pufc.featureclass)
                    using (RowCursor rowCursor = table.Search(queryFilter, false))
                    {
                        if (rowCursor.MoveNext())   // only check the first record
                        {
                            // Record found for pu_id
                            using (Row row = rowCursor.Current)
                            {
                                if (row[PRZC.c_FLD_FC_PU_ID] != null)
                                {
                                    pu_id = Convert.ToInt32(row[PRZC.c_FLD_FC_PU_ID]);
                                    success = true;
                                    message = "success";
                                }
                                else
                                {
                                    // null pu_id value (no need to do anything here)
                                    success = false;    // for clarity
                                    message = "null pu id value";
                                }
                            }
                        }
                        else
                        {
                            // no record found for cell number (no need to do anything here)
                            success = false;        // for clarity
                            message = "no cell number record found";
                        }
                    }
                });

                return (success, pu_id, message);
            }
            catch (Exception ex)
            {
                return (false, -9999, ex.Message);
            }
        }

        #endregion

        #region LISTS AND HASHSETS

        /// <summary>
        /// Retrieves a hashset of Planning Unit IDs from the planning unit feature class.
        /// Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, HashSet<int> puids, string message)> GetPUIDHashset()
        {
            try
            {
                // Project GDB existence
                if (!(await GDBExists_Project()).exists)
                {
                    throw new Exception("Project GDB not found.");
                }

                // Planning Units existence
                var tryget_pudata = await PUDataExists();
                if (!tryget_pudata.exists)
                {
                    throw new Exception("Planning Units not found.");
                }

                // Create the hashset
                HashSet<int> HASH_puid = new HashSet<int>();

                // Populate the hashset
                await QueuedTask.Run(() =>
                {
                    // Use the Planning Units Feature Class
                    var tryget_puras = GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);

                    // Build query filter
                    QueryFilter queryFilter = new QueryFilter()
                    {
                        SubFields = PRZC.c_FLD_RAS_PU_ID
                    };

                    using (RasterDataset rasterDataset = tryget_puras.rasterDataset)
                    using (Raster raster = rasterDataset.CreateFullRaster())
                    using (Table table = raster.GetAttributeTable())
                    using (RowCursor rowCursor = table.Search(queryFilter))
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                if (row[PRZC.c_FLD_RAS_PU_ID] != null)
                                {
                                    int pu_id = Convert.ToInt32(row[PRZC.c_FLD_RAS_PU_ID]);
                                    if (pu_id > 0)
                                    {
                                        // only keep values > 0
                                        HASH_puid.Add(pu_id);
                                    }
                                }
                            }
                        }
                    }
                });

                return (true, HASH_puid, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Returns a list of planning unit ids from the planning units feature class.
        /// Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<int> puids, string message)> GetPUIDList()
        {
            try
            {
                var outcome = await GetPUIDHashset();

                if (outcome.success)
                {
                    List<int> puids = outcome.puids.ToList();
                    puids.Sort();

                    return (true, puids, "success");
                }
                else
                {
                    return (false, null, outcome.message);
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieves a hashset of National Grid Cell Numbers from the planning unit feature class.
        /// Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, HashSet<long> cell_numbers, string message)> GetCellNumberHashset()
        {
            try
            {
                // Project GDB existence
                if (!(await GDBExists_Project()).exists)
                {
                    throw new Exception("Project GDB not found.");
                }

                // Planning Units existence
                var tryget_pudata = await PUDataExists();
                if (!tryget_pudata.exists)
                {
                    throw new Exception("Planning Units not found.");
                }
                else if (!tryget_pudata.national_enabled)
                {
                    throw new Exception("Planning Units not national enabled.");
                }

                // Create the hashset
                HashSet<long> HASH_cellnumbers = new HashSet<long>();

                // Populate the hashset
                await QueuedTask.Run(() =>
                {
                    // Use the Planning Units Feature Class
                    var tryget_pufc = GetFC_Project(PRZC.c_FC_PLANNING_UNITS);

                    // Build query filter
                    QueryFilter queryFilter = new QueryFilter()
                    {
                        SubFields = PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER
                    };

                    using (Table table = tryget_pufc.featureclass)
                    using (RowCursor rowCursor = table.Search(queryFilter))
                    {
                        while (rowCursor.MoveNext())
                        {
                            // Record found for cell_number
                            using (Row row = rowCursor.Current)
                            {
                                if (row[PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER] != null)
                                {
                                    long cell_number = Convert.ToInt64(row[PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER]);
                                    if (cell_number > 0)
                                    {
                                        // only keep values > 0
                                        HASH_cellnumbers.Add(cell_number);
                                    }
                                }
                            }
                        }
                    }
                });

                return (true, HASH_cellnumbers, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Returns a list of cell numbers from the planning unit feature class.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, List<long> cell_numbers, string message)> GetCellNumberList()
        {
            try
            {
                var outcome = await GetCellNumberHashset();

                if (outcome.success)
                {
                    List<long> cellnums = outcome.cell_numbers.ToList();
                    cellnums.Sort();

                    return (true, cellnums, "success");
                }
                else
                {
                    return (false, null, outcome.message);
                }
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #region DICTIONARIES

        /// <summary>
        /// Retrieve a dictionary of cell numbers and associated planning unit ids from the
        /// project geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<long, int> dict, string message)> GetCellNumbersAndPUIDs()
        {
            try
            {
                // Project GDB existence
                if (!(await GDBExists_Project()).exists)
                {
                    throw new Exception("Project GDB not found.");
                }

                // Planning Units existence
                var tryget_pudata = await PUDataExists();
                if (!tryget_pudata.exists)
                {
                    throw new Exception("Planning Units not found.");
                }
                else if (!tryget_pudata.national_enabled)
                {
                    throw new Exception("Planning Units not national enabled.");
                }

                // Create the dictionary
                Dictionary<long, int> DICT_CN_PUID = new Dictionary<long, int>();

                // Populate the dictionary
                await QueuedTask.Run(() =>
                {
                    // Use the Planning Units Feature Class
                    var tryget_puras = GetRaster_Project(PRZC.c_RAS_PLANNING_UNITS);

                    // Build query filter
                    QueryFilter queryFilter = new QueryFilter()
                    {
                        SubFields = PRZC.c_FLD_RAS_PU_ID + "," + PRZC.c_FLD_RAS_PU_NATGRID_CELL_NUMBER
                    };

                    using (RasterDataset rasterDataset = tryget_puras.rasterDataset)
                    using (Raster raster = rasterDataset.CreateFullRaster())
                    using (Table table = raster.GetAttributeTable())
                    using (RowCursor rowCursor = table.Search(queryFilter))
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                // both columns must have a value
                                if (row[PRZC.c_FLD_RAS_PU_NATGRID_CELL_NUMBER] != null & row[PRZC.c_FLD_RAS_PU_ID] != null)
                                {
                                    long cell_number = Convert.ToInt64(row[PRZC.c_FLD_RAS_PU_NATGRID_CELL_NUMBER]);
                                    int pu_id = Convert.ToInt32(row[PRZC.c_FLD_RAS_PU_ID]);

                                    if (cell_number > 0 & pu_id > 0 & !DICT_CN_PUID.ContainsKey(cell_number))
                                    {
                                        DICT_CN_PUID.Add(cell_number, pu_id);
                                    }
                                }
                            }
                        }
                    }
                });

                return (true, DICT_CN_PUID, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a dictionary of planning unit ids and associated cell numbers from the
        /// Project geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<int, long> dict, string message)> GetPUIDsAndCellNumbers()
        {
            try
            {
                // Project GDB existence
                if (!(await GDBExists_Project()).exists)
                {
                    throw new Exception("Project GDB not found.");
                }

                // Planning Units existence
                var tryget_pudata = await PUDataExists();
                if (!tryget_pudata.exists)
                {
                    throw new Exception("Planning Units not found.");
                }
                else if (!tryget_pudata.national_enabled)
                {
                    throw new Exception("Planning Units not national enabled.");
                }

                // Create the dictionary
                Dictionary<int, long> DICT_PUID_CN = new Dictionary<int, long>();

                // Populate the dictionary
                await QueuedTask.Run(() =>
                {
                    // Use the Planning Units Feature Class
                    var tryget_pufc = GetFC_Project(PRZC.c_FC_PLANNING_UNITS);

                    // Build query filter
                    QueryFilter queryFilter = new QueryFilter()
                    {
                        SubFields = PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER + "," + PRZC.c_FLD_FC_PU_ID
                    };

                    using (Table table = tryget_pufc.featureclass)
                    using (RowCursor rowCursor = table.Search(queryFilter))
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                // both columns must have a value
                                if (row[PRZC.c_FLD_FC_PU_ID] != null & row[PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER] != null)
                                {
                                    int pu_id = Convert.ToInt32(row[PRZC.c_FLD_FC_PU_ID]);
                                    long cell_number = Convert.ToInt64(row[PRZC.c_FLD_FC_PU_NATGRID_CELL_NUMBER]);

                                    if (pu_id > 0 & cell_number > 0 & !DICT_PUID_CN.ContainsKey(pu_id))
                                    {
                                        DICT_PUID_CN.Add(pu_id, cell_number);
                                    }
                                }
                            }
                        }
                    }
                });

                return (true, DICT_PUID_CN, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion

        #endregion

        #endregion

        #region LAYER EXISTENCE

        public static bool PRZLayerExists(Map map, PRZLayerNames layer_name)
        {
            try
            {
                switch (layer_name)
                {
                    case PRZLayerNames.MAIN:
                        return GroupLayerExists_MAIN(map);

                    case PRZLayerNames.SELRULES:
                        return GroupLayerExists_STATUS(map);

                    case PRZLayerNames.SELRULES_INCLUDE:
                        return GroupLayerExists_STATUS_INCLUDE(map);

                    case PRZLayerNames.SELRULES_EXCLUDE:
                        return GroupLayerExists_STATUS_EXCLUDE(map);

                    case PRZLayerNames.COST:
                        return GroupLayerExists_COST(map);

                    case PRZLayerNames.FEATURES:
                        return GroupLayerExists_FEATURE(map);

                    case PRZLayerNames.PU:
                        bool fle = FeatureLayerExists_PU(map);
                        bool rle = RasterLayerExists_PU(map);
                        return (fle | rle);

                    case PRZLayerNames.SA:
                        return FeatureLayerExists_SA(map);

                    case PRZLayerNames.SAB:
                        return FeatureLayerExists_SAB(map);

                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool GroupLayerExists_MAIN(Map map)
        {
            try
            {
                // map can't be null
                if (map == null)
                {
                    return false;
                }

                // Get list of map-level group layers having matching name
                List<Layer> LIST_layers = map.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_MAIN && (l is GroupLayer)).ToList();

                // If at least one match is found, return true.  Otherwise, false.
                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool GroupLayerExists_STATUS(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_SELRULES && (l is GroupLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool GroupLayerExists_STATUS_INCLUDE(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.SELRULES))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_SELRULES_INCLUDE && (l is GroupLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool GroupLayerExists_STATUS_EXCLUDE(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map , PRZLayerNames.SELRULES))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_SELRULES_EXCLUDE && (l is GroupLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool GroupLayerExists_COST(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_COST && (l is GroupLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool GroupLayerExists_FEATURE(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_FEATURES && (l is GroupLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool FeatureLayerExists_PU(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_PLANNING_UNITS_FC && (l is FeatureLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool FeatureLayerExists_SA(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_STUDY_AREA && (l is FeatureLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool FeatureLayerExists_SAB(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_STUDY_AREA_BUFFER && (l is FeatureLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool RasterLayerExists_PU(Map map)
        {
            try
            {
                if (map == null)
                {
                    return false;
                }

                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return false;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_PLANNING_UNITS_FC && (l is RasterLayer)).ToList();

                return LIST_layers.Count > 0;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        #endregion

        #region SINGLE LAYER RETRIEVAL

        public static Layer GetPRZLayer(Map map, PRZLayerNames layer_name)
        {
            try
            {
                switch (layer_name)
                {
                    case PRZLayerNames.MAIN:
                        return GetGroupLayer_MAIN(map);

                    case PRZLayerNames.SELRULES:
                        return GetGroupLayer_STATUS(map);

                    case PRZLayerNames.SELRULES_INCLUDE:
                        return GetGroupLayer_STATUS_INCLUDE(map);

                    case PRZLayerNames.SELRULES_EXCLUDE:
                        return GetGroupLayer_STATUS_EXCLUDE(map);

                    case PRZLayerNames.COST:
                        return GetGroupLayer_COST(map);

                    case PRZLayerNames.FEATURES:
                        return GetGroupLayer_FEATURE(map);

                    case PRZLayerNames.PU:
                        Layer fl = GetFeatureLayer_PU(map);
                        Layer rl = GetRasterLayer_PU(map);
                        return fl ?? rl;

                    case PRZLayerNames.SA:
                        return GetFeatureLayer_SA(map);

                    case PRZLayerNames.SAB:
                        return GetFeatureLayer_SAB(map);

                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static GroupLayer GetGroupLayer_MAIN(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.MAIN))
                {
                    return null;
                }

                List<Layer> LIST_layers = map.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_MAIN && (l is GroupLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as GroupLayer;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static GroupLayer GetGroupLayer_STATUS(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_SELRULES && (l is GroupLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as GroupLayer;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static GroupLayer GetGroupLayer_STATUS_INCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_INCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_SELRULES_INCLUDE && (l is GroupLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as GroupLayer;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static GroupLayer GetGroupLayer_STATUS_EXCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_EXCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_SELRULES_EXCLUDE && (l is GroupLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as GroupLayer;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static GroupLayer GetGroupLayer_COST(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.COST))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_COST && (l is GroupLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as GroupLayer;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static GroupLayer GetGroupLayer_FEATURE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.FEATURES))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_GROUPLAYER_FEATURES && (l is GroupLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as GroupLayer;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static FeatureLayer GetFeatureLayer_PU(Map map)
        {
            try
            {
                if (!FeatureLayerExists_PU(map))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_PLANNING_UNITS_FC && (l is FeatureLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as FeatureLayer;
                }

            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static FeatureLayer GetFeatureLayer_SA(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SA))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_STUDY_AREA && (l is FeatureLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as FeatureLayer;
                }

            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static FeatureLayer GetFeatureLayer_SAB(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SAB))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_STUDY_AREA_BUFFER && (l is FeatureLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as FeatureLayer;
                }

            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static RasterLayer GetRasterLayer_PU(Map map)
        {
            try
            {
                if (!RasterLayerExists_PU(map))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                List<Layer> LIST_layers = GL.Layers.Where(l => l.Name == PRZC.c_LAYER_PLANNING_UNITS_FC && (l is RasterLayer)).ToList();

                if (LIST_layers.Count == 0)
                {
                    return null;
                }
                else
                {
                    return LIST_layers[0] as RasterLayer;
                }

            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        #endregion

        #region LAYER COLLECTION RETRIEVAL

        public static List<Layer> GetPRZLayers(Map map, PRZLayerNames container, PRZLayerRetrievalType type)
        {
            try
            {
                // Proceed only for specific containers
                GroupLayer GL = null;

                if (container == PRZLayerNames.COST |
                    container == PRZLayerNames.FEATURES |
                    container == PRZLayerNames.SELRULES_INCLUDE |
                    container == PRZLayerNames.SELRULES_EXCLUDE)
                {
                    GL = (GroupLayer)GetPRZLayer(map, container);
                }
                else
                {
                    return null;
                }

                // If unable to retrieve container, leave
                if (GL == null)
                {
                    return null;
                }

                // Retrieve the layers in the container
                List<Layer> layers = new List<Layer>();

                switch (type)
                {
                    case PRZLayerRetrievalType.FEATURE:
                        layers = GL.Layers.Where(lyr => lyr is FeatureLayer).ToList();
                        break;
                    case PRZLayerRetrievalType.RASTER:
                        layers = GL.Layers.Where(lyr => lyr is RasterLayer).ToList();
                        break;
                    case PRZLayerRetrievalType.BOTH:
                        layers = GL.Layers.Where(lyr => lyr is FeatureLayer | lyr is RasterLayer).ToList();
                        break;


                    default:
                        return null;
                }

                return layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // STATUS INCLUDE
        public static List<FeatureLayer> GetFeatureLayers_STATUS_INCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_INCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES_INCLUDE);
                List<FeatureLayer> LIST_layers = GL.Layers.Where(l => l is FeatureLayer).Cast<FeatureLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<RasterLayer> GetRasterLayers_STATUS_INCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_INCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES_INCLUDE);
                List<RasterLayer> LIST_layers = GL.Layers.Where(l => l is RasterLayer).Cast<RasterLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<Layer> GetLayers_STATUS_INCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_INCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES_INCLUDE);
                List<Layer> LIST_layers = GL.Layers.Where(l => l is RasterLayer | l is FeatureLayer).ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // STATUS EXCLUDE
        public static List<FeatureLayer> GetFeatureLayers_STATUS_EXCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_EXCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES_EXCLUDE);
                List<FeatureLayer> LIST_layers = GL.Layers.Where(l => l is FeatureLayer).Cast<FeatureLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<RasterLayer> GetRasterLayers_STATUS_EXCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_EXCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES_EXCLUDE);
                List<RasterLayer> LIST_layers = GL.Layers.Where(l => l is RasterLayer).Cast<RasterLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<Layer> GetLayers_STATUS_EXCLUDE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.SELRULES_EXCLUDE))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.SELRULES_EXCLUDE);
                List<Layer> LIST_layers = GL.Layers.Where(l => l is RasterLayer | l is FeatureLayer).ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // COST
        public static List<FeatureLayer> GetFeatureLayers_COST(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.COST))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.COST);
                List<FeatureLayer> LIST_layers = GL.Layers.Where(l => l is FeatureLayer).Cast<FeatureLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<RasterLayer> GetRasterLayers_COST(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.COST))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.COST);
                List<RasterLayer> LIST_layers = GL.Layers.Where(l => l is RasterLayer).Cast<RasterLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<Layer> GetLayers_COST(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.COST))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.COST);
                List<Layer> LIST_layers = GL.Layers.Where(l => l is RasterLayer | l is FeatureLayer).ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        // FEATURE
        public static List<FeatureLayer> GetFeatureLayers_FEATURE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.FEATURES))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.FEATURES);
                List<FeatureLayer> LIST_layers = GL.Layers.Where(l => l is FeatureLayer).Cast<FeatureLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<RasterLayer> GetRasterLayers_FEATURE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.FEATURES))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.FEATURES);
                List<RasterLayer> LIST_layers = GL.Layers.Where(l => l is RasterLayer).Cast<RasterLayer>().ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static List<Layer> GetLayers_FEATURE(Map map)
        {
            try
            {
                if (!PRZLayerExists(map, PRZLayerNames.FEATURES))
                {
                    return null;
                }

                GroupLayer GL = (GroupLayer)GetPRZLayer(map, PRZLayerNames.FEATURES);
                List<Layer> LIST_layers = GL.Layers.Where(l => l is RasterLayer | l is FeatureLayer).ToList();

                return LIST_layers;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        #endregion

        #region GENERIC DATA METHODS

        /// <summary>
        /// Delete all contents of the project file geodatabase.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, string message)> DeleteProjectGDBContents()
        {
            try
            {
                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GPRefresh = GPExecuteToolFlags.RefreshProjectItems | GPExecuteToolFlags.GPThread;
                string toolOutput;

                // geodatabase path
                string gdbpath = GetPath_ProjectGDB();

                // Create the lists of object names
                List<string> relNames = new List<string>();
                List<string> fdsNames = new List<string>();
                List<string> rdsNames = new List<string>();
                List<string> fcNames = new List<string>();
                List<string> tabNames = new List<string>();
                List<string> domainNames = new List<string>();

                await QueuedTask.Run(() =>
                {
                    // Get the project gdb
                    var tryget_gdb = GetGDB_Project();
                    if (!tryget_gdb.success)
                    {
                        throw new Exception("Unable to retrieve geodatabase.");
                    }

                    // Populate the lists of existing objects
                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        // Get list of Relationship Classes
                        relNames = geodatabase.GetDefinitions<RelationshipClassDefinition>().Select(o => o.GetName()).ToList();
                        WriteLog($"{relNames.Count} Relationship Class(es) found in {gdbpath}...");

                        // Get list of Feature Dataset names
                        fdsNames = geodatabase.GetDefinitions<FeatureDatasetDefinition>().Select(o => o.GetName()).ToList();
                        WriteLog($"{fdsNames.Count} Feature Dataset(s) found in {gdbpath}...");

                        // Get list of Raster Dataset names
                        rdsNames = geodatabase.GetDefinitions<RasterDatasetDefinition>().Select(o => o.GetName()).ToList();
                        WriteLog($"{rdsNames.Count} Raster Dataset(s) found in {gdbpath}...");

                        // Get list of top-level Feature Classes
                        fcNames = geodatabase.GetDefinitions<FeatureClassDefinition>().Select(o => o.GetName()).ToList();
                        WriteLog($"{fcNames.Count} Feature Class(es) found in {gdbpath}...");

                        // Get list of tables
                        tabNames = geodatabase.GetDefinitions<TableDefinition>().Select(o => o.GetName()).ToList();
                        WriteLog($"{tabNames.Count} Table(s) found in {gdbpath}...");

                        // Get list of domains
                        domainNames = geodatabase.GetDomains().Select(o => o.GetName()).ToList();
                        WriteLog($"{domainNames.Count} domain(s) found in {gdbpath}...");
                    }
                });

                // Delete those objects using geoprocessing tools
                // Relationship Classes
                if (relNames.Count > 0)
                {
                    WriteLog($"Deleting {relNames.Count} relationship class(es)...");
                    toolParams = Geoprocessing.MakeValueArray(string.Join(";", relNames));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true, outputCoordinateSystem: GetSR_PRZCanadaAlbers());
                    toolOutput = await RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GPRefresh);
                    if (toolOutput == null)
                    {
                        WriteLog($"Error deleting relationship class(es). GP Tool failed or was cancelled by user", LogMessageType.ERROR);
                        return (false, "Error deleting relationship class(es).");
                    }
                    else
                    {
                        WriteLog($"Relationship class(es) deleted.");
                    }
                }

                // Feature Datasets
                if (fdsNames.Count > 0)
                {
                    WriteLog($"Deleting {fdsNames.Count} feature dataset(s)...");
                    toolParams = Geoprocessing.MakeValueArray(string.Join(";", fdsNames));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GPRefresh);
                    if (toolOutput == null)
                    {
                        WriteLog($"Error deleting feature dataset(s). GP Tool failed or was cancelled by user", LogMessageType.ERROR);
                        return (false, "Error deleting feature dataset(s).");
                    }
                    else
                    {
                        WriteLog($"Feature dataset(s) deleted.");
                    }
                }

                // Raster Datasets
                if (rdsNames.Count > 0)
                {
                    WriteLog($"Deleting {rdsNames.Count} raster dataset(s)...");
                    toolParams = Geoprocessing.MakeValueArray(string.Join(";", rdsNames));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GPRefresh);
                    if (toolOutput == null)
                    {
                        WriteLog($"Error deleting raster dataset(s). GP Tool failed or was cancelled by user", LogMessageType.ERROR);
                        return (false, "Error deleting raster dataset(s).");
                    }
                    else
                    {
                        WriteLog($"Raster dataset(s) deleted.");
                    }
                }

                // Feature Classes
                if (fcNames.Count > 0)
                {
                    WriteLog($"Deleting {fcNames.Count} feature class(es)...");
                    toolParams = Geoprocessing.MakeValueArray(string.Join(";", fcNames));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GPRefresh);
                    if (toolOutput == null)
                    {
                        WriteLog($"Error deleting feature class(es). GP Tool failed or was cancelled by user", LogMessageType.ERROR);
                        return (false, "Error deleting feature class(es).");
                    }
                    else
                    {
                        WriteLog($"Feature class(es) deleted.");
                    }
                }

                // Tables
                if (tabNames.Count > 0)
                {
                    WriteLog($"Deleting {tabNames.Count} table(s)...");
                    toolParams = Geoprocessing.MakeValueArray(string.Join(";", tabNames));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("Delete_management", toolParams, toolEnvs, toolFlags_GPRefresh);
                    if (toolOutput == null)
                    {
                        WriteLog($"Error deleting table(s). GP Tool failed or was cancelled by user", LogMessageType.ERROR);
                        return (false, "Error deleting table(s).");
                    }
                    else
                    {
                        WriteLog($"Table(s) deleted.");
                    }
                }

                // Domains
                if (domainNames.Count > 0)
                {
                    WriteLog($"Deleting domain(s)...");
                    foreach (string domainName in domainNames)
                    {
                        if (!string.Equals(domainName, PRZC.c_DOMAIN_PRESENCE, StringComparison.OrdinalIgnoreCase) &
                            !string.Equals(domainName, PRZC.c_DOMAIN_REG_STATUS, StringComparison.OrdinalIgnoreCase) &
                            !string.Equals(domainName, PRZC.c_DOMAIN_REG_TYPE, StringComparison.OrdinalIgnoreCase) &
                            !string.Equals(domainName, PRZC.c_DOMAIN_REG_THEME, StringComparison.OrdinalIgnoreCase))
                        {
                            WriteLog($"Deleting {domainName} domain...");
                            toolParams = Geoprocessing.MakeValueArray(gdbpath, domainName);
                            toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                            toolOutput = await RunGPTool("DeleteDomain_management", toolParams, toolEnvs, toolFlags_GPRefresh);
                            if (toolOutput == null)
                            {
                                WriteLog($"Error deleting {domainName} domain. GP Tool failed or was cancelled by user", LogMessageType.ERROR);
                                return (false, $"Error deleting {domainName} domain.");
                            }
                            else
                            {
                                WriteLog($"Domain deleted.");
                            }
                        }
                    }
                }

                // I've deleted everything.
                return (true, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Remove all core layers from the active mapview's Prioritization group layer.
        /// Silent errors.
        /// </summary>
        /// <param name="map"></param>
        /// <returns></returns>
        public static async Task<(bool success, string message)> RemovePRZItemsFromMap(Map map)
        {
            try
            {
                // Ensure that project gdb exists
                var tryexists = await GDBExists_Project();
                if (!tryexists.exists)
                {
                    return (false, tryexists.message);
                }

                // Get project gdb absolute path (forward slashes)
                Uri gdburi = new Uri(GetPath_ProjectGDB());
                string gdbpath = gdburi.AbsolutePath;

                // Create empty lists
                List<StandaloneTable> tables_to_remove = new List<StandaloneTable>();
                List<Layer> layers_to_remove = new List<Layer>();

                // Populate the lists of layers to remove, then remove them
                await QueuedTask.Run(() =>
                {
                    // Get layers & standalone tables to test
                    var standalone_tables = map.GetStandaloneTablesAsFlattenedList();
                    var layers = map.GetLayersAsFlattenedList().Where(l => l is RasterLayer | l is FeatureLayer);

                    // Standalone Tables
                    foreach (var standalone_table in standalone_tables)
                    {
                        var conn = standalone_table.GetDataConnection();
                        if (conn == null)
                        {
                            tables_to_remove.Add(standalone_table);
                        }
                        else if (conn is CIMStandardDataConnection dataconn)
                        {
                            string datapath = dataconn.WorkspaceConnectionString.Split('=')[1];
                            Uri uri = new Uri(datapath);
                            if (uri == null || string.Equals(uri.AbsolutePath, gdbpath, StringComparison.OrdinalIgnoreCase))
                            {
                                tables_to_remove.Add(standalone_table);
                            }
                        }
                    }

                    // Layers
                    foreach(var layer in layers)
                    {
                        // Connection Approach
                        var conn = layer.GetDataConnection();
                        if (conn == null)
                        {
                            layers_to_remove.Add(layer);
                        }
                        else if (conn is CIMStandardDataConnection dataconn)
                        {
                            string datapath = dataconn.WorkspaceConnectionString.Split('=')[1];
                            Uri uri = new Uri(datapath);
                            if (uri == null || string.Equals(uri.AbsolutePath, gdbpath, StringComparison.OrdinalIgnoreCase))
                            {
                                layers_to_remove.Add(layer);
                            }
                        }
                    }

                    standalone_tables = null;
                    layers = null;

                    // remove tables
                    if (tables_to_remove.Count > 0)
                    {
                        map.RemoveStandaloneTables(tables_to_remove);
                    }

                    // remove layers
                    if (layers_to_remove.Count > 0)
                    {
                        map.RemoveLayers(layers_to_remove);

                        MapView.Active.Redraw(true);
                    }
                });

                return (true, "success");
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message);
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Rebuild the active mapview's Prioritization group layer.  Silent errors.
        /// </summary>
        /// <param name="map"></param>
        /// <returns></returns>
        internal static async Task<(bool success, string message)> RedrawPRZLayers(Map map)
        {
            try
            {
                // Ensure that Project Workspace exists
                string project_path = GetPath_ProjectFolder();
                var trydirexists = FolderExists_Project();
                if (!trydirexists.exists)
                {
                    return (false, trydirexists.message);
                }

                // Check for Project GDB
                string gdb_path = GetPath_ProjectGDB();
                var trygdbexists = await GDBExists_Project();
                if (!trygdbexists.exists)
                {
                    return (false, trygdbexists.message);
                }

                var tryremove = await RemovePRZItemsFromMap(map);
                if (!tryremove.success)
                {
                    return tryremove;
                }

                // Process the layers
                await QueuedTask.Run(async () =>
                {
                    #region TOP-LEVEL LAYERS

                    // Main Group Layer
                    GroupLayer GL_MAIN = null;

                    // remove it if present...
                    if (PRZLayerExists(map, PRZLayerNames.MAIN))
                    {
                        GL_MAIN = (GroupLayer)GetPRZLayer(map, PRZLayerNames.MAIN);
                        map.RemoveLayer(GL_MAIN);
                    }

                    // ... and create a new one
                    GL_MAIN = LayerFactory.Instance.CreateGroupLayer(map, 0, PRZC.c_GROUPLAYER_MAIN);
                    GL_MAIN.SetVisibility(true);
                    GL_MAIN.SetExpanded(true);

                    // Add the Study Area Layer (MIGHT NOT EXIST)
                    if ((await FCExists_Project(PRZC.c_FC_STUDY_AREA_MAIN)).exists)
                    {
                        string fc_path = GetPath_Project(PRZC.c_FC_STUDY_AREA_MAIN).path;
                        Uri uri = new Uri(fc_path);
                        var layerParams = new FeatureLayerCreationParams(uri) // TODO: Check these are assigned correctly with Laurence
                        {
                            MapMemberIndex = 0,
                            Name = PRZC.c_LAYER_STUDY_AREA
                        };
                        FeatureLayer featureLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, GL_MAIN);
                        ApplyLegend_SA_Simple(featureLayer);
                        featureLayer.SetVisibility(true);
                    }

                    // Add the Study Area Buffer Layer (MIGHT NOT EXIST)
                    if ((await FCExists_Project(PRZC.c_FC_STUDY_AREA_MAIN_BUFFERED)).exists)
                    {
                        string fc_path = GetPath_Project(PRZC.c_FC_STUDY_AREA_MAIN_BUFFERED).path;
                        Uri uri = new Uri(fc_path);
                        var layerParams = new FeatureLayerCreationParams(uri) // TODO: Check these are assigned correctly with Laurence
                        {
                            MapMemberIndex = 1,
                            Name = PRZC.c_LAYER_STUDY_AREA_BUFFER
                        };
                        FeatureLayer featureLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, GL_MAIN);
                        ApplyLegend_SAB_Simple(featureLayer);
                        featureLayer.SetVisibility(true);
                    }

                    // Add the Planning Unit Feature Class (FC MIGHT NOT EXIST)
                    if ((await FCExists_Project(PRZC.c_FC_PLANNING_UNITS)).exists)
                    {
                        string fc_path = GetPath_Project(PRZC.c_FC_PLANNING_UNITS).path;
                        Uri uri = new Uri(fc_path);
                        var layerParams = new FeatureLayerCreationParams(uri) // TODO: Check these are assigned correctly with Laurence
                        {
                            MapMemberIndex = 2,
                            Name = PRZC.c_LAYER_PLANNING_UNITS_FC
                        };
                        FeatureLayer featureLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, GL_MAIN);
                        await ApplyLegend_PU_Basic(featureLayer);
                        featureLayer.SetVisibility(true);
                    }

                    // Add the National Elements Group Layer

                    GroupLayer GL_NAT = LayerFactory.Instance.CreateGroupLayer(GL_MAIN, 3, PRZC.c_GROUPLAYER_ELEMENTS_NAT);
                    GL_NAT.SetVisibility(false);
                    GL_NAT.SetExpanded(false);


                    // Add the Regional Elements Group Layer

                    GroupLayer GL_REG = LayerFactory.Instance.CreateGroupLayer(GL_MAIN, 4, PRZC.c_GROUPLAYER_ELEMENTS_REG);
                    GL_REG.SetVisibility(false);
                    GL_REG.SetExpanded(false);

                    // Add the Planning Unit Raster (RASTER DATASET MIGHT NOT EXIST)
                    if ((await RasterExists_Project(PRZC.c_RAS_PLANNING_UNITS)).exists)
                    {
                        string ras_path = GetPath_Project(PRZC.c_RAS_PLANNING_UNITS).path;
                        Uri uri = new Uri(ras_path);
                        var layerParams = new RasterLayerCreationParams(uri) // TODO: Check these are assigned correctly with Laurence
                        {
                            MapMemberIndex = 5,
                            Name = PRZC.c_LAYER_PLANNING_UNITS_RAS
                        };
                        RasterLayer rasterLayer = LayerFactory.Instance.CreateLayer<RasterLayer>(layerParams, GL_MAIN);
                        await ApplyLegend_PU_Basic(rasterLayer);
                        rasterLayer.SetVisibility(true);
                    }

                    #endregion

                    #region POPULATE NATIONAL GROUP LAYER

                    var tryex_natfds = await FDSExists_Project(PRZC.c_FDS_NATIONAL_ELEMENTS);
                    if (tryex_natfds.exists)
                    {
                        var tryget_fcs = await GetNationalElementFCs();
                        if (tryget_fcs.success)
                        {
                            foreach (var fc_name in tryget_fcs.feature_classes)
                            {
                                string fc_path = GetPath_Project(fc_name, PRZC.c_FDS_NATIONAL_ELEMENTS).path;
                                Uri uri = new Uri(fc_path);
                                var layerParams = new FeatureLayerCreationParams(uri)
                                {
                                    MapMemberPosition = MapMemberPosition.AddToBottom
                                };
                                FeatureLayer featureLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, GL_NAT);
                                await ApplyLegend_ElementFC_Basic(featureLayer, DataSource.National);
                                featureLayer.SetVisibility(false);
                                featureLayer.SetExpanded(false);

                                // Adjust the name to include element type code
                                string layer_name = featureLayer.Name;  // pattern: n00001: bla bla bla

                                string numpart = fc_name.Substring(4);
                                if (int.TryParse(numpart, out int element_id))
                                {
                                    var tryget_elem = await GetNationalElement(element_id);
                                    if (tryget_elem.success)
                                    {
                                        var element = tryget_elem.element;
                                        int type_int = element.ElementType;
                                        string txt = "";

                                        if (type_int == (int)ElementType.Goal)
                                        {
                                            txt = " G";
                                        }
                                        else if (type_int == (int)ElementType.Weight)
                                        {
                                            txt = " W";
                                        }
                                        else if (type_int == (int)ElementType.Include)
                                        {
                                            txt = " I";
                                        }
                                        else if (type_int == (int)ElementType.Exclude)
                                        {
                                            txt = " E";
                                        }
                                        else
                                        {
                                            txt = " ?";
                                        }

                                        featureLayer.SetName(layer_name.Insert(6, txt));
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    #region POPULATE REGIONAL GROUP LAYER

                    var tryex_regfds = await FDSExists_Project(PRZC.c_FDS_REGIONAL_ELEMENTS);
                    if (tryex_regfds.exists)
                    {
                        var tryget_fcs = await GetRegionalElementFCs();
                        if (tryget_fcs.success)
                        {
                            foreach (var fc_name in tryget_fcs.feature_classes)
                            {
                                string fc_path = GetPath_Project(fc_name, PRZC.c_FDS_REGIONAL_ELEMENTS).path;
                                Uri uri = new Uri(fc_path);
                                var layerParams = new FeatureLayerCreationParams(uri)
                                {
                                    MapMemberPosition = MapMemberPosition.AddToBottom
                                };
                                FeatureLayer featureLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, GL_REG);
                                await ApplyLegend_ElementFC_Basic(featureLayer, DataSource.Regional);
                                featureLayer.SetVisibility(false);
                                featureLayer.SetExpanded(false);

                                // Adjust the name to include element type code
                                string layer_name = featureLayer.Name;  // pattern: n00001: bla bla bla
                                string numpart = fc_name.Substring(4);

                                if (int.TryParse(numpart, out int element_id))
                                {
                                    var tryget_elem = await GetRegionalElement(element_id);
                                    if (tryget_elem.success)
                                    {
                                        var element = tryget_elem.element;
                                        int type_int = element.ElementType;
                                        string txt = "";

                                        if (type_int == (int)ElementType.Goal)
                                        {
                                            txt = " G";
                                        }
                                        else if (type_int == (int)ElementType.Weight)
                                        {
                                            txt = " W";
                                        }
                                        else if (type_int == (int)ElementType.Include)
                                        {
                                            txt = " I";
                                        }
                                        else if (type_int == (int)ElementType.Exclude)
                                        {
                                            txt = " E";
                                        }
                                        else
                                        {
                                            txt = " ?";
                                        }

                                        string new_name = layer_name.Insert(6, txt);
                                        featureLayer.SetName(new_name);
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    // Refresh map
                    await MapView.Active.RedrawAsync(true);
                });

                return (true, "success");
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message);
                return (false, ex.Message);
            }
        }

        #endregion GENERIC DATA METHODS

        #region EDIT OPERATIONS

        /// <summary>
        /// Retrieve a generic EditOperation object for editing tasks.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static EditOperation GetEditOperation(string name)
        {
            try
            {
                EditOperation editOp = new EditOperation();

                editOp.Name = name ?? "unnamed edit operation";
                editOp.ShowProgressor = false;
                editOp.ShowModalMessageAfterFailure = false;
                editOp.SelectNewFeatures = false;
                editOp.SelectModifiedFeatures = false;

                return editOp;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        #endregion

        #region SPATIAL REFERENCES

        /// <summary>
        /// Retrieves the NCC Prioritization Tools' specific Albers Equal Area projection.
        /// Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<SpatialReference> GetSR_PRZCanadaAlbers()
        {
            try
            {
                SpatialReference spatialReference = null;

                await QueuedTask.Run(() =>
                {
                    string wkt = PRZC.c_SR_WKT_WGS84_CanadaAlbers;  // Special PRZ WGS84 Canada Albers projection

                    using (SpatialReferenceBuilder builder = new SpatialReferenceBuilder(wkt))
                    {
                        builder.SetDefaultXYResolution();
                        builder.SetDefaultXYTolerance();

                        spatialReference = builder.ToSpatialReference();
                    }
                });

                return spatialReference;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        /// <summary>
        /// Establishes whether a specified spatial reference is equivalent to the custom
        /// prioritization Canada Albers projection.  The comparison ignores the PCS
        /// name since it is user-defined and might differ. XY Tolerances and XY Resolutions
        /// are not compared.  Silent errors.
        /// </summary>
        /// <param name="spatialReference"></param>
        /// <returns></returns>
        public static async Task<(bool match, string message)> SpatialReferenceIsPRZCanadaAlbers(SpatialReference spatialReference)
        {
            try
            {
                // Validate test sr
                if (spatialReference == null)
                {
                    throw new Exception("Test Spatial Reference is null");
                }
                else if (spatialReference.IsUnknown)
                {
                    throw new Exception("Test Spatial Reference is Unknown");
                }
                else if (spatialReference.IsGeographic)
                {
                    throw new Exception("Test Spatial Reference is Geographic");
                }

                // Get Nat SR
                SpatialReference NatSR = await GetSR_PRZCanadaAlbers();

                // Retrieve WKT for both SRs
                string wkt_test = spatialReference.Wkt;
                string wkt_nat = NatSR.Wkt;

                ProMsgBox.Show($"Test\n{wkt_test}\n\nNat\n{wkt_nat}");

                // Get PCS name for both SRs
                string pcs_name_test = spatialReference.Name;
                string pcs_name_nat = NatSR.Name;

                // Get swap strings for both
                string swap_test = @"PROJCS[""" + pcs_name_test;
                string swap_nat = @"PROJCS[""" + pcs_name_nat;

                ProMsgBox.Show($"{swap_test}\n{swap_nat}");

                // Swap out swap strings
                string swapped_test = wkt_test.Replace(swap_test, @"PROJCS[""");
                string swapped_nat = wkt_nat.Replace(swap_nat, @"PROJCS[""");

                ProMsgBox.Show($"Test\n{swapped_test}\n\nNat\n{swapped_nat}");

                // Compare swapped WKTs
                bool result = string.Equals(swapped_test, swapped_nat, StringComparison.OrdinalIgnoreCase);

                return result ? (true, "Spatial References are equivalent") : (false, "Spatial References are not equivalent");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        #endregion

        #region DOMAINS

        /// <summary>
        /// Creates several domains associated with Regional Data.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, string message)> CreateRegionalDomains()
        {
            try
            {
                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread;
                string toolOutput;

                // Flags for regional domain existence
                bool domain_exists_presence = false;
                bool domain_exists_status = false;
                bool domain_exists_type = false;
                bool domain_exists_theme = false;

                // Sorted Lists of domain values
                SortedList<object, string> sl_presence = null;
                SortedList<object, string> sl_status = null;
                SortedList<object, string> sl_type = null;
                SortedList<object, string> sl_theme = null;

                // project gdb path
                string gdbpath = GetPath_ProjectGDB();

                // Determine existence of each domain
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = GetGDB_Project();
                    if (!tryget_gdb.success)
                    {
                        throw new Exception("Error opening the project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        var domains = geodatabase.GetDomains();

                        foreach (var domain in domains)
                        {
                            using (domain)
                            {
                                if (domain is CodedValueDomain cvd)
                                {
                                    string domname = cvd.GetName();

                                    if (string.Equals(domname, PRZC.c_DOMAIN_PRESENCE, StringComparison.OrdinalIgnoreCase))
                                    {
                                        domain_exists_presence = true;
                                        sl_presence = cvd.GetCodedValuePairs();
                                    }
                                    else if (string.Equals(domname, PRZC.c_DOMAIN_REG_STATUS, StringComparison.OrdinalIgnoreCase))
                                    {
                                        domain_exists_status = true;
                                        sl_status = cvd.GetCodedValuePairs();
                                    }
                                    else if (string.Equals(domname, PRZC.c_DOMAIN_REG_TYPE, StringComparison.OrdinalIgnoreCase))
                                    {
                                        domain_exists_type = true;
                                        sl_type = cvd.GetCodedValuePairs();
                                    }
                                    else if (string.Equals(domname, PRZC.c_DOMAIN_REG_THEME, StringComparison.OrdinalIgnoreCase))
                                    {
                                        domain_exists_theme = true;
                                        sl_theme = cvd.GetCodedValuePairs();
                                    }
                                }
                            }
                        }
                    }
                });

                #region PRESENCE DOMAIN

                if (domain_exists_presence)
                {
                    if (sl_presence.Count > 0)
                    {
                        // Get the list of domain KVP key strings
                        var keyobjs = sl_presence.Keys.ToList();
                        List<string> keystrings = new List<string>();
                        foreach (object o in keyobjs)
                        {
                            int d = Convert.ToInt32(o);
                            keystrings.Add(d.ToString());
                        }

                        // delete the KVPs
                        toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_PRESENCE, string.Join(";", keystrings));
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await RunGPTool("DeleteCodedValueFromDomain_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            return (false, $"Error eliminating KVPs from {PRZC.c_DOMAIN_PRESENCE} domain.");
                        }
                    }
                    else
                    {
                        // domain exists but has no KVPs.
                    }
                }
                else
                {
                    // create domain
                    toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_PRESENCE, "", "SHORT", "CODED", "DEFAULT", "DEFAULT");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("CreateDomain_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        return (false, $"Error creating {PRZC.c_DOMAIN_PRESENCE}");
                    }
                }

                // Add coded value #1
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_PRESENCE, (int)ElementPresence.Present, ElementPresence.Present.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_PRESENCE} domain");
                }

                // Add coded value #2
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_PRESENCE, (int)ElementPresence.Absent, ElementPresence.Absent.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_PRESENCE} domain");
                }

                #endregion

                #region STATUS DOMAIN

                if (domain_exists_status)
                {
                    if (sl_status.Count > 0)
                    {
                        // Get the list of domain KVP key strings
                        var keyobjs = sl_status.Keys.ToList();
                        List<string> keystrings = new List<string>();
                        foreach (object o in keyobjs)
                        {
                            int d = Convert.ToInt32(o);
                            keystrings.Add(d.ToString());
                        }

                        // delete the KVPs
                        toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_STATUS, string.Join(";", keystrings));
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await RunGPTool("DeleteCodedValueFromDomain_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            return (false, $"Error eliminating KVPs from {PRZC.c_DOMAIN_REG_STATUS} domain.");
                        }
                    }
                    else
                    {
                        // domain exists but has no KVPs.
                    }
                }
                else
                {
                    // create domain
                    toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_STATUS, "", "SHORT", "CODED", "DEFAULT", "DEFAULT");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("CreateDomain_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        return (false, $"Error creating {PRZC.c_DOMAIN_REG_STATUS}");
                    }
                }

                // Add coded value #1
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_STATUS, (int)ElementStatus.Active, ElementStatus.Active.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_STATUS} domain");
                }

                // Add coded value #2
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_STATUS, (int)ElementStatus.Inactive, ElementStatus.Inactive.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_STATUS} domain");
                }

                #endregion

                #region TYPE DOMAIN

                if (domain_exists_type)
                {
                    if (sl_type.Count > 0)
                    {
                        // Get the list of domain KVP key strings
                        var keyobjs = sl_type.Keys.ToList();
                        List<string> keystrings = new List<string>();
                        foreach (object o in keyobjs)
                        {
                            int d = Convert.ToInt32(o);
                            keystrings.Add(d.ToString());
                        }

                        // delete the KVPs
                        toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_TYPE, string.Join(";", keystrings));
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await RunGPTool("DeleteCodedValueFromDomain_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            return (false, $"Error eliminating KVPs from {PRZC.c_DOMAIN_REG_TYPE} domain.");
                        }
                    }
                    else
                    {
                        // domain exists but has no KVPs.
                    }
                }
                else
                {
                    // create domain
                    toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_TYPE, "", "SHORT", "CODED", "DEFAULT", "DEFAULT");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("CreateDomain_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        return (false, $"Error creating {PRZC.c_DOMAIN_REG_TYPE}");
                    }
                }

                // Add coded value #1
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_TYPE, (int)ElementType.Goal, ElementType.Goal.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_TYPE} domain");
                }

                // Add coded value #2
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_TYPE, (int)ElementType.Weight, ElementType.Weight.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_TYPE} domain");
                }

                // Add coded value #3
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_TYPE, (int)ElementType.Include, ElementType.Include.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_TYPE} domain");
                }

                // Add coded value #4
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_TYPE, (int)ElementType.Exclude, ElementType.Exclude.ToString());
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_TYPE} domain");
                }

                #endregion

                #region THEME DOMAIN

                if (domain_exists_theme)
                {
                    if (sl_theme.Count > 0)
                    {
                        // Get the list of domain KVP key strings
                        var keyobjs = sl_theme.Keys.ToList();
                        List<string> keystrings = new List<string>();
                        foreach (object o in keyobjs)
                        {
                            int d = Convert.ToInt32(o);
                            keystrings.Add(d.ToString());
                        }

                        // delete the KVPs
                        toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, string.Join(";", keystrings));
                        toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                        toolOutput = await RunGPTool("DeleteCodedValueFromDomain_management", toolParams, toolEnvs, toolFlags_GP);
                        if (toolOutput == null)
                        {
                            return (false, $"Error eliminating KVPs from {PRZC.c_DOMAIN_REG_THEME} domain.");
                        }
                    }
                    else
                    {
                        // domain exists but has no KVPs.
                    }
                }
                else
                {
                    // create domain
                    toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, "", "SHORT", "CODED", "DEFAULT", "DEFAULT");
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("CreateDomain_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        return (false, $"Error creating {PRZC.c_DOMAIN_REG_THEME}");
                    }
                }

                // Add coded value #1
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, (int)ElementTheme.RegionalGoal, "Regional Goals");  // TODO: Better name management here!
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_THEME} domain");
                }

                // Add coded value #2
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, (int)ElementTheme.RegionalWeight, "Regional Weights"); // TODO: Better name management here!
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_THEME} domain");
                }

                // Add coded value #3
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, (int)ElementTheme.RegionalInclude, "Regional Includes"); // TODO: Better name management here!
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_THEME} domain");
                }

                // Add coded value #4
                toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, (int)ElementTheme.RegionalExclude, "Regional Excludes"); // TODO: Better name management here!
                toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                if (toolOutput == null)
                {
                    return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_THEME} domain");
                }

                #endregion

                return (true, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Removes and Re-adds National Theme KVPs to the Regional Themes domain.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, string message)> UpdateRegionalThemesDomain(Dictionary<int, string> national_values)
        {
            try
            {
                // stop if the dictionary parameter is null
                if (national_values == null)
                {
                    national_values = new Dictionary<int, string>();
                }

                // Declare some generic GP variables
                IReadOnlyList<string> toolParams;
                IReadOnlyList<KeyValuePair<string, string>> toolEnvs;
                GPExecuteToolFlags toolFlags_GP = GPExecuteToolFlags.GPThread | GPExecuteToolFlags.RefreshProjectItems;
                string toolOutput;

                // Max value for national themes
                int nat_max_value = 1000;

                // Flag for theme domain existence
                bool domain_exists_theme = false;

                // Sorted Lists of domain values
                SortedList<object, string> sl_theme = null;

                // project gdb path
                string gdbpath = GetPath_ProjectGDB();

                // Determine existence of domain
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = GetGDB_Project();
                    if (!tryget_gdb.success)
                    {
                        throw new Exception("Error opening the project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        var domains = geodatabase.GetDomains();

                        foreach (var domain in domains)
                        {
                            using (domain)
                            {
                                if (domain is CodedValueDomain cvd)
                                {
                                    string domname = cvd.GetName();

                                    if (string.Equals(domname, PRZC.c_DOMAIN_REG_THEME, StringComparison.OrdinalIgnoreCase))
                                    {
                                        domain_exists_theme = true;
                                        sl_theme = cvd.GetCodedValuePairs();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                });

                if (!domain_exists_theme)
                {
                    return (false, $"{PRZC.c_DOMAIN_REG_THEME} domain not found.  Please initialize the workspace");
                }

                // Get the list of domain codes (integers and string conversions)
                var keyobjs = sl_theme.Keys.ToList();
                List<int> keyints = new List<int>();
                List<string> keystrings = new List<string>();
                foreach (object o in keyobjs)
                {
                    int d = Convert.ToInt32(o);
                    keyints.Add(d);
                    keystrings.Add(d.ToString());
                }

                // Get the list of keys for kvps to remove
                List<string> remove_these_kvps = keyints.Where(key => key <= nat_max_value).Select(key => key.ToString()).ToList();

                // remove them
                if (remove_these_kvps.Count > 0)
                {
                    // delete the KVPs
                    toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, string.Join(";", remove_these_kvps));
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("DeleteCodedValueFromDomain_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        return (false, $"Error eliminating KVPs from {PRZC.c_DOMAIN_REG_THEME} domain.");
                    }
                }

                // Validate list of dictionary entries to add
                SortedList<int, string> valid_values = new SortedList<int, string>();
                foreach (var kvp in national_values)
                {
                    int key = kvp.Key;
                    string descr = kvp.Value;

                    if (key <= nat_max_value && descr.Trim().Length > 0 && !valid_values.ContainsKey(key))
                    {
                        valid_values.Add(key, descr);
                    }
                }

                // Now add the values
                foreach(var kvp in valid_values)
                {
                    toolParams = Geoprocessing.MakeValueArray(gdbpath, PRZC.c_DOMAIN_REG_THEME, kvp.Key, kvp.Value);
                    toolEnvs = Geoprocessing.MakeEnvironmentArray(workspace: gdbpath, overwriteoutput: true);
                    toolOutput = await RunGPTool("AddCodedValueToDomain_management", toolParams, toolEnvs, toolFlags_GP);
                    if (toolOutput == null)
                    {
                        return (false, $"Error adding coded value to {PRZC.c_DOMAIN_REG_THEME} domain");
                    }
                }

                return (true, "success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        /// <summary>
        /// Retrieve a dictionary of Regional Themes domain entries.  Silent errors.
        /// </summary>
        /// <returns></returns>
        public static async Task<(bool success, Dictionary<int, string> dict, string message)> GetRegionalThemesDomainKVPs()
        {
            try
            {
                // Max value for national themes
                int nat_max_value = 1000;

                // Flag for theme domain existence
                bool domain_exists = false;

                // Sorted Lists of domain values
                SortedList<object, string> sl_theme = null;

                // project gdb path
                string gdbpath = GetPath_ProjectGDB();

                // Determine existence of domain
                await QueuedTask.Run(() =>
                {
                    var tryget_gdb = GetGDB_Project();
                    if (!tryget_gdb.success)
                    {
                        throw new Exception("Error opening the project geodatabase.");
                    }

                    using (Geodatabase geodatabase = tryget_gdb.geodatabase)
                    {
                        var domains = geodatabase.GetDomains();

                        foreach (var domain in domains)
                        {
                            using (domain)
                            {
                                if (domain is CodedValueDomain cvd)
                                {
                                    string domname = cvd.GetName();

                                    if (string.Equals(domname, PRZC.c_DOMAIN_REG_THEME, StringComparison.OrdinalIgnoreCase))
                                    {
                                        domain_exists = true;
                                        sl_theme = cvd.GetCodedValuePairs();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                });

                if (!domain_exists)
                {
                    return (false, null, $"{PRZC.c_DOMAIN_REG_THEME} domain not found.");
                }

                Dictionary<int, string> dict = new Dictionary<int, string>();

                foreach (KeyValuePair<object, string> kvp in sl_theme)
                {
                    dict.Add(Convert.ToInt32(kvp.Key), kvp.Value);
                }

                return (true, dict, "success");
            }
            catch (Exception ex)
            {
                return (false, null, ex.Message);
            }
        }

        #endregion


        #region *** MAY NOT BE NECESSARY ANY MORE!!! ***

        public static async Task<Dictionary<int, double>> GetPlanningUnitIDsAndArea()
        {
            try
            {
                Dictionary<int, double> dict = new Dictionary<int, double>();

                if (!await QueuedTask.Run(() =>
                {
                    try
                    {
                        var tryget = GetFC_Project(PRZC.c_FC_PLANNING_UNITS);
                        if (!tryget.success)
                        {
                            throw new Exception("Error retrieving feature class.");
                        }

                        using (Table table = tryget.featureclass)
                        using (RowCursor rowCursor = table.Search())
                        {
                            while (rowCursor.MoveNext())
                            {
                                using (Row row = rowCursor.Current)
                                {
                                    int id = Convert.ToInt32(row[PRZC.c_FLD_FC_PU_ID]);
                                    double area_m2 = Convert.ToDouble(row[PRZC.c_FLD_FC_PU_AREA_M2]);

                                    dict.Add(id, area_m2);
                                }
                            }
                        }

                        return true;
                    }
                    catch (Exception ex)
                    {
                        ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                        return false;
                    }
                }))
                {
                    ProMsgBox.Show("Error retrieving dictionary of planning unit ids + area");
                    return null;
                }
                else
                {
                    return dict;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }


        #endregion

        #region GEOPROCESSING

        public static async Task<string> RunGPTool(string toolName, IReadOnlyList<string> toolParams, IReadOnlyList<KeyValuePair<string, string>> toolEnvs, GPExecuteToolFlags flags)
        {
            IGPResult gp_result = null;
            using (CancelableProgressorSource cps = new CancelableProgressorSource("Executing GP Tool: " + toolName, "Tool cancelled by user", false))
            {
                // Execute the Geoprocessing Tool
                try
                {
                    gp_result = await Geoprocessing.ExecuteToolAsync(toolName, toolParams, toolEnvs, cps.Progressor, flags);
                }
                catch (Exception ex)
                {
                    // handle error and leave
                    WriteLog("Error Executing GP Tool: " + toolName + " >>> " + ex.Message, LogMessageType.ERROR);
                    return null;
                }
            }

            // At this point, GP Tool has executed and either succeeded, failed, or been cancelled.  There's also a chance that the output IGpResult is null.
            ProcessGPMessages(gp_result, toolName);

            // Configure return value
            if (gp_result == null || gp_result.ReturnValue == null)
            {
                return null;
            }
            else
            {
                return gp_result.ReturnValue;
            }
        }

        private static void ProcessGPMessages(IGPResult gp_result, string toolName)
        {
            try
            {
                StringBuilder messageBuilder = new StringBuilder();
                messageBuilder.AppendLine("Executing GP Tool: " + toolName);

                // If GPTool execution (i.e. ExecuteToolAsync) didn't even run, we have a null IGpResult
                if (gp_result == null)
                {
                    messageBuilder.AppendLine(" > Failure Executing Tool. IGpResult is null...  Something fishy going on here...");
                    WriteLog(messageBuilder.ToString(), LogMessageType.ERROR);
                    return;
                }

                // I now have an existing IGpResult.

                // Assemble the IGpResult Messages into a single string
                if (gp_result.Messages.Count() > 0)
                {
                    foreach (var gp_message in gp_result.Messages)
                    {
                        string gpm = " > " + gp_message.Type.ToString() + " " + gp_message.ErrorCode.ToString() + ": " + gp_message.Text;
                        messageBuilder.AppendLine(gpm);
                    }
                }
                else
                {
                    // if no messages present, add my own message
                    messageBuilder.AppendLine(" > No messages generated...  Something fishy going on here... User might have cancelled");
                }

                // Now, provide some execution result info
                messageBuilder.AppendLine(" > Result Code (0 means success): " + gp_result.ErrorCode.ToString() + "   Execution Status: " + (gp_result.IsFailed ? "Failed or Cancelled" : "Succeeded"));
                messageBuilder.Append(" > Return Value: " + (gp_result.ReturnValue == null ? "null   --> definitely something fishy going on" : gp_result.ReturnValue));

                // Finally, log the message info and return
                if (gp_result.IsFailed)
                {
                    WriteLog(messageBuilder.ToString(), LogMessageType.ERROR);
                }
                else
                {
                    WriteLog(messageBuilder.ToString());
                }
            }
            catch
            {
            }
        }

        public static string GetElapsedTimeMessage(TimeSpan span)
        {
            try
            {
                int inthours = span.Hours;
                int intminutes = span.Minutes;
                int intseconds = span.Seconds;
                int intmilliseconds = span.Milliseconds;

                string hours = inthours.ToString() + ((inthours == 1) ? " hour" : " hours");
                string minutes = intminutes.ToString() + ((intminutes == 1) ? " minute" : " minutes");
                string seconds = intseconds.ToString() + ((intseconds == 1) ? " second" : " seconds");
                string milliseconds = intmilliseconds.ToString() + ((intmilliseconds == 1) ? " millisecond" : " milliseconds");

                string elapsedmessage = "";

                if (inthours == 0 & intminutes == 0)
                {
                    elapsedmessage = seconds;
                }
                else if (inthours == 0)
                {
                    elapsedmessage = minutes + " and " + seconds;
                }
                else
                {
                    elapsedmessage = hours + ", " + minutes + ", " + seconds;
                }

                return "Elapsed Time: " + elapsedmessage;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return "<error calculating duration>";
            }
        }

        public static string GetElapsedTimeInSeconds(TimeSpan span)
        {
            try
            {
                double sec = span.TotalSeconds;

                string message = $"Elapsed Time: {sec:N3}";

                return message;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return "<error calculating duration>";
            }
        }

        #endregion

        #region RENDERERS

        public static async Task<bool> ApplyLegend_ElementFC_Basic(FeatureLayer featureLayer, DataSource dataSource)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Colors
                    CIMColor outlineColor_Nat = GetRGBColor(60,60,60);
                    CIMColor fillColor_Nat = GetNamedColor(Color.Khaki);

                    CIMColor outlineColor_Reg = GetRGBColor(60, 60, 60);
                    CIMColor fillColor_Reg = GetNamedColor(Color.PaleVioletRed);

                    // Symbols
                    CIMStroke outlineSym_Nat = SymbolFactory.Instance.ConstructStroke(outlineColor_Nat, 1, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillSym_Nat = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor_Nat, SimpleFillStyle.Solid, outlineSym_Nat);

                    CIMStroke outlineSym_Reg = SymbolFactory.Instance.ConstructStroke(outlineColor_Reg, 1, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillSym_Reg = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor_Reg, SimpleFillStyle.Solid, outlineSym_Reg);

                    // Renderer Definitions
                    SimpleRendererDefinition rendDef_Nat = new SimpleRendererDefinition
                    {
                        SymbolTemplate = fillSym_Nat.MakeSymbolReference()
                    };

                    SimpleRendererDefinition rendDef_Reg = new SimpleRendererDefinition
                    {
                        SymbolTemplate = fillSym_Reg.MakeSymbolReference()
                    };

                    if (dataSource == DataSource.National)
                    {
                        CIMSimpleRenderer rend = (CIMSimpleRenderer)featureLayer.CreateRenderer(rendDef_Nat);
                        rend.Patch = PatchShape.AreaRectangle;
                        featureLayer.SetRenderer(rend);
                    }
                    else if (dataSource == DataSource.Regional)
                    {
                        CIMSimpleRenderer rend = (CIMSimpleRenderer)featureLayer.CreateRenderer(rendDef_Reg);
                        rend.Patch = PatchShape.AreaRectangle;
                        featureLayer.SetRenderer(rend);
                    }
                });

                MapView.Active.Redraw(true);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static async Task<bool> ApplyLegend_PU_Basic(FeatureLayer featureLayer)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Colors
                    CIMColor outlineColor = GetRGBColor(0, 112, 255); // Blue-ish
                    CIMColor fillColor = CIMColor.NoColor();

                    // Symbols
                    CIMStroke outlineSym = SymbolFactory.Instance.ConstructStroke(outlineColor, 1, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillSym = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor, SimpleFillStyle.Solid, outlineSym);

                    // Create a new Renderer Definition
                    SimpleRendererDefinition rendDef = new SimpleRendererDefinition
                    {
                        SymbolTemplate = fillSym.MakeSymbolReference()
                    };

                    CIMSimpleRenderer rend = (CIMSimpleRenderer)featureLayer.CreateRenderer(rendDef);
                    rend.Patch = PatchShape.AreaSquare;
                    featureLayer.SetRenderer(rend);
                });

                MapView.Active.Redraw(false);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static async Task<bool> ApplyLegend_PU_Basic(RasterLayer rasterLayer)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Color #1
                    var ramp1 = new CIMPolarContinuousColorRamp();
                    ramp1.FromColor = GetNamedColor(Color.LightGreen);
                    ramp1.ToColor = GetNamedColor(Color.MediumSeaGreen);
                    ramp1.PolarDirection = PolarDirection.Clockwise;

                    //// Color #2
                    //var ramp2 = new CIMLinearContinuousColorRamp();
                    //ramp2.FromColor = GetNamedColor(Color.Yellow);
                    //ramp2.ToColor = GetNamedColor(Color.YellowGreen);

                    //// Color #3
                    //var ramp3 = new CIMLinearContinuousColorRamp();
                    //ramp3.FromColor = GetNamedColor(Color.YellowGreen);
                    //ramp3.ToColor = GetNamedColor(Color.Green);

                    //// Create multipart color ramp
                    //var multiPartRamp = new CIMMultipartColorRamp
                    //{
                    //    Weights = new double[3] { 1, 1, 1 },
                    //    ColorRamps = new CIMLinearContinuousColorRamp[3] { ramp1, ramp2, ramp3 }
                    //};

                    // Create the colorizer definition
                    var colDef = new StretchColorizerDefinition(0, RasterStretchType.StandardDeviations, 1, ramp1)
                    {
                        StandardDeviationsParam = 3
                    };

                    // Create the colorizer
                    var colorizer = (CIMRasterStretchColorizer)rasterLayer.CreateColorizer(colDef);

                    // Apply the colorizer
                    rasterLayer.SetColorizer(colorizer);
                });

                MapView.Active.Redraw(true);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static async Task<bool> ApplyLegend_PU_SelRules(FeatureLayer FL)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // COLORS
                    CIMColor outlineColor = GetNamedColor(Color.Gray); // outline color for all 3 poly symbols
                    CIMColor fillColor_Available = GetNamedColor(Color.Bisque);
                    CIMColor fillColor_Include = GetNamedColor(Color.GreenYellow);
                    CIMColor fillColor_Exclude = GetNamedColor(Color.OrangeRed);

                    // SYMBOLS
                    CIMStroke outlineSym = SymbolFactory.Instance.ConstructStroke(outlineColor, 0.3, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillSym_Available = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor_Available, SimpleFillStyle.Solid, outlineSym);
                    CIMPolygonSymbol fillSym_Include = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor_Include, SimpleFillStyle.Solid, outlineSym);
                    CIMPolygonSymbol fillSym_Exclude = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor_Exclude, SimpleFillStyle.Solid, outlineSym);

                    // CIM UNIQUE VALUES
                    CIMUniqueValue uv_Available = new CIMUniqueValue { FieldValues = new string[] { "<Null>" } };
                    CIMUniqueValue uv_Include = new CIMUniqueValue { FieldValues = new string[] { SelectionRuleType.INCLUDE.ToString() } };
                    CIMUniqueValue uv_Exclude = new CIMUniqueValue { FieldValues = new string[] { SelectionRuleType.EXCLUDE.ToString() } };

                    // CIM UNIQUE VALUE CLASSES
                    CIMUniqueValueClass uvcAvailable = new CIMUniqueValueClass
                    {
                        Editable = true,
                        Label = "Available",
                        Symbol = fillSym_Available.MakeSymbolReference(),
                        Description = "",
                        Visible = true,
                        Values = new CIMUniqueValue[] { uv_Available }
                    };
                    CIMUniqueValueClass uvcInclude = new CIMUniqueValueClass
                    {
                        Editable = true,
                        Label = "Included",
                        Symbol = fillSym_Include.MakeSymbolReference(),
                        Description = "",
                        Visible = true,
                        Values = new CIMUniqueValue[] { uv_Include }
                    };
                    CIMUniqueValueClass uvcExclude = new CIMUniqueValueClass
                    {
                        Editable = true,
                        Label = "Excluded",
                        Symbol = fillSym_Exclude.MakeSymbolReference(),
                        Description = "",
                        Visible = true,
                        Values = new CIMUniqueValue[] { uv_Exclude }
                    };

                    // CIM UNIQUE VALUE GROUP
                    CIMUniqueValueGroup uvgMain = new CIMUniqueValueGroup
                    {
                        Classes = new CIMUniqueValueClass[] { uvcInclude, uvcExclude, uvcAvailable },
                        Heading = "Effective Selection Rule"
                    };

                    // UV RENDERER
                    CIMUniqueValueRenderer UVRend = new CIMUniqueValueRenderer
                    {
                        UseDefaultSymbol = false,
                        Fields = new string[] { PRZC.c_FLD_FC_PU_EFFECTIVE_RULE },
                        Groups = new CIMUniqueValueGroup[] { uvgMain },
                        DefaultSymbolPatch = PatchShape.AreaRoundedRectangle
                    };

                    FL.SetRenderer(UVRend);
                });

                MapView.Active.Redraw(false);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static async Task<bool> ApplyLegend_PU_SelRuleConflicts(FeatureLayer FL)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // COLORS
                    CIMColor outlineColor = GetNamedColor(Color.Gray); // outline color for all 3 poly symbols
                    CIMColor fillColor_Conflict = GetNamedColor(Color.Magenta);
                    CIMColor fillColor_NoConflict = GetNamedColor(Color.LightGray);

                    // SYMBOLS
                    CIMStroke outlineSym = SymbolFactory.Instance.ConstructStroke(outlineColor, 0.1, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillSym_Conflict = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor_Conflict, SimpleFillStyle.Solid, outlineSym);
                    CIMPolygonSymbol fillSym_NoConflict = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor_NoConflict, SimpleFillStyle.Solid, outlineSym);

                    // CIM UNIQUE VALUES
                    CIMUniqueValue uv_Conflict = new CIMUniqueValue { FieldValues = new string[] { "1" } };
                    CIMUniqueValue uv_NoConflict = new CIMUniqueValue { FieldValues = new string[] { "0" } };

                    // CIM UNIQUE VALUE CLASSES
                    CIMUniqueValueClass uvcConflict = new CIMUniqueValueClass
                    {
                        Editable = true,
                        Label = "Conflict",
                        Symbol = fillSym_Conflict.MakeSymbolReference(),
                        Description = "",
                        Visible = true,
                        Values = new CIMUniqueValue[] { uv_Conflict }
                    };
                    CIMUniqueValueClass uvcNoConflict = new CIMUniqueValueClass
                    {
                        Editable = true,
                        Label = "OK",
                        Symbol = fillSym_NoConflict.MakeSymbolReference(),
                        Description = "",
                        Visible = true,
                        Values = new CIMUniqueValue[] { uv_NoConflict }
                    };

                    // CIM UNIQUE VALUE GROUP
                    CIMUniqueValueGroup uvgMain = new CIMUniqueValueGroup
                    {
                        Classes = new CIMUniqueValueClass[] { uvcConflict, uvcNoConflict },
                        Heading = "Selection Rule Conflicts"
                    };

                    // UV RENDERER
                    CIMUniqueValueRenderer UVRend = new CIMUniqueValueRenderer
                    {
                        UseDefaultSymbol = false,
                        Fields = new string[] { PRZC.c_FLD_FC_PU_CONFLICT },
                        Groups = new CIMUniqueValueGroup[] { uvgMain },
                        DefaultSymbolPatch = PatchShape.AreaRoundedRectangle
                    };

                    FL.SetRenderer(UVRend);
                });

                MapView.Active.Redraw(false);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static async Task<bool> ApplyLegend_PU_Cost(FeatureLayer FL)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // get the lowest and highest cost values in PUCF
                    double minCost = 0;
                    double maxCost = 0;
                    bool seeded = false;

                    using (Table table = FL.GetFeatureClass())
                    using (RowCursor rowCursor = table.Search(null, false))
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                double cost = Convert.ToDouble(row[PRZC.c_FLD_FC_PU_COST]);

                                if (!seeded)
                                {
                                    minCost = cost;
                                    maxCost = cost;

                                    seeded = true;
                                }
                                else
                                {
                                    if (cost > maxCost)
                                    {
                                        maxCost = cost;
                                    }

                                    if (cost < minCost)
                                    {
                                        minCost = cost;
                                    }
                                }

                            }
                        }
                    }

                    // Create the polygon fill template
                    CIMStroke outline = SymbolFactory.Instance.ConstructStroke(GetNamedColor(Color.Gray), 0, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillWithOutline = SymbolFactory.Instance.ConstructPolygonSymbol(GetNamedColor(Color.White), SimpleFillStyle.Solid, outline);

                    // Create the color ramp
                    CIMLinearContinuousColorRamp ramp = new CIMLinearContinuousColorRamp
                    {
                        FromColor = GetNamedColor(Color.LightGray),
                        ToColor = GetNamedColor(Color.Red)
                    };

                    // Create the Unclassed Renderer
                    UnclassedColorsRendererDefinition ucDef = new UnclassedColorsRendererDefinition();

                    ucDef.Field = PRZC.c_FLD_FC_PU_COST;
                    ucDef.ColorRamp = ramp;
                    ucDef.LowerColorStop = minCost;
                    ucDef.LowerLabel = minCost.ToString();
                    ucDef.UpperColorStop = maxCost;
                    ucDef.UpperLabel = maxCost.ToString();
                    ucDef.SymbolTemplate = fillWithOutline.MakeSymbolReference();

                    CIMClassBreaksRenderer rend = (CIMClassBreaksRenderer)FL.CreateRenderer(ucDef);
                    FL.SetRenderer(rend);
                });

                await MapView.Active.RedrawAsync(false);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static async Task<bool> ApplyLegend_PU_CFCount(FeatureLayer FL)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // What's the highest number of CF within a single planning unit?
                    int maxCF = 0;

                    using (Table table = FL.GetFeatureClass())
                    using (RowCursor rowCursor = table.Search(null, false))
                    {
                        while (rowCursor.MoveNext())
                        {
                            using (Row row = rowCursor.Current)
                            {
                                int max = Convert.ToInt32(row[PRZC.c_FLD_FC_PU_FEATURECOUNT]);

                                if (max > maxCF)
                                {
                                    maxCF = max;
                                }
                            }
                        }
                    }

                    // Create the polygon fill template
                    CIMStroke outline = SymbolFactory.Instance.ConstructStroke(GetNamedColor(Color.Gray), 0, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillWithOutline = SymbolFactory.Instance.ConstructPolygonSymbol(GetNamedColor(Color.White), SimpleFillStyle.Solid, outline);

                    // Create the color ramp
                    CIMLinearContinuousColorRamp ramp = new CIMLinearContinuousColorRamp
                    {
                        FromColor = GetNamedColor(Color.LightGray),
                        ToColor = GetNamedColor(Color.ForestGreen)
                    };

                    // Create the Unclassed Renderer
                    UnclassedColorsRendererDefinition ucDef = new UnclassedColorsRendererDefinition();

                    ucDef.Field = PRZC.c_FLD_FC_PU_FEATURECOUNT;
                    ucDef.ColorRamp = ramp;
                    ucDef.LowerColorStop = 0;
                    ucDef.LowerLabel = "0";
                    ucDef.UpperColorStop = maxCF;
                    ucDef.UpperLabel = maxCF.ToString();
                    ucDef.SymbolTemplate = fillWithOutline.MakeSymbolReference();

                    CIMClassBreaksRenderer rend = (CIMClassBreaksRenderer)FL.CreateRenderer(ucDef);
                    FL.SetRenderer(rend);
                });

                await MapView.Active.RedrawAsync(false);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static async Task<bool> ApplyLegend_PU_Boundary(FeatureLayer FL)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Colors
                    CIMColor colorOutline = GetNamedColor(Color.Gray);
                    CIMColor colorEdge = GetNamedColor(Color.Magenta);
                    CIMColor colorNonEdge = GetNamedColor(Color.LightGray);

                    // Symbols
                    CIMStroke outlineSym = SymbolFactory.Instance.ConstructStroke(colorOutline, 1, SimpleLineStyle.Solid);
                    CIMPolygonSymbol fillEdge = SymbolFactory.Instance.ConstructPolygonSymbol(colorEdge, SimpleFillStyle.Solid, outlineSym);
                    CIMPolygonSymbol fillNonEdge = SymbolFactory.Instance.ConstructPolygonSymbol(colorNonEdge, SimpleFillStyle.Solid, outlineSym);

                    // fields array
                    string[] fields = new string[] { PRZC.c_FLD_FC_PU_HAS_UNSHARED_PERIM };

                    // CIM Unique Values
                    CIMUniqueValue uvEdge = new CIMUniqueValue { FieldValues = new string[] { "1" } };
                    CIMUniqueValue uvNonEdge = new CIMUniqueValue { FieldValues = new string[] { "0" } };

                    // CIM Unique Value Classes
                    CIMUniqueValueClass uvcEdge = new CIMUniqueValueClass
                    {
                        Editable = true,
                        Label = "Edge",
                        Symbol = fillEdge.MakeSymbolReference(),
                        Description = "",
                        Visible = true,
                        Values = new CIMUniqueValue[] { uvEdge }
                    };

                    CIMUniqueValueClass uvcNonEdge = new CIMUniqueValueClass
                    {
                        Editable = true,
                        Label = "Non Edge",
                        Symbol = fillNonEdge.MakeSymbolReference(),
                        Description = "",
                        Visible = true,
                        Values = new CIMUniqueValue[] { uvNonEdge }
                    };

                    // CIM Unique Value Group
                    CIMUniqueValueGroup uvgMain = new CIMUniqueValueGroup
                    {
                        Classes = new CIMUniqueValueClass[] { uvcEdge, uvcNonEdge },
                        Heading = "Has Unshared Perimeter"                        
                    };


                    // Unique Values Renderer
                    CIMUniqueValueRenderer UVRend = new CIMUniqueValueRenderer
                    {
                        UseDefaultSymbol = false,
                        Fields = fields,
                        Groups = new CIMUniqueValueGroup[] { uvgMain },
                        DefaultSymbolPatch = PatchShape.AreaSquare
                    };

                    FL.SetRenderer(UVRend);
                });

                await MapView.Active.RedrawAsync(false);

                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool ApplyLegend_SAB_Simple(FeatureLayer FL)
        {
            try
            {
                // Colors
                CIMColor outlineColor = GetNamedColor(Color.Black);
                CIMColor fillColor = CIMColor.NoColor();

                CIMStroke outlineSym = SymbolFactory.Instance.ConstructStroke(outlineColor, 1, SimpleLineStyle.Solid);
                CIMPolygonSymbol fillSym = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor, SimpleFillStyle.Solid, outlineSym);
                CIMSimpleRenderer rend = FL.GetRenderer() as CIMSimpleRenderer;
                rend.Symbol = fillSym.MakeSymbolReference();
                //rend.Label = "";
                FL.SetRenderer(rend);
                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        public static bool ApplyLegend_SA_Simple(FeatureLayer FL)
        {
            try
            {
                // Colors
                CIMColor outlineColor = GetNamedColor(Color.Black);
                CIMColor fillColor = CIMColor.NoColor();

                CIMStroke outlineSym = SymbolFactory.Instance.ConstructStroke(outlineColor, 2, SimpleLineStyle.Solid);
                CIMPolygonSymbol fillSym = SymbolFactory.Instance.ConstructPolygonSymbol(fillColor, SimpleFillStyle.Solid, outlineSym);
                CIMSimpleRenderer rend = FL.GetRenderer() as CIMSimpleRenderer;
                rend.Symbol = fillSym.MakeSymbolReference();
                //rend.Label = "";
                FL.SetRenderer(rend);
                return true;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        #endregion

        #region COLORS AND SYMBOLS

        public static CIMColor GetRGBColor(byte r, byte g, byte b, byte a = 100)
        {
            try
            {
                return ColorFactory.Instance.CreateRGBColor(r, g, b, a);
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        public static CIMColor GetNamedColor(Color color)
        {
            try
            {
                return ColorFactory.Instance.CreateRGBColor(color.R, color.G, color.B, color.A);
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return null;
            }
        }

        //internal static ISimpleFillSymbol ReturnSimpleFillSymbol(IColor FillColor, IColor OutlineColor, double OutlineWidth, esriSimpleFillStyle FillStyle)
        //{
        //    try
        //    {
        //        ISimpleLineSymbol Outline = new SimpleLineSymbolClass();
        //        Outline.Color = OutlineColor;
        //        Outline.Style = esriSimpleLineStyle.esriSLSSolid;
        //        Outline.Width = OutlineWidth;

        //        ISimpleFillSymbol FillSymbol = new SimpleFillSymbolClass();
        //        FillSymbol.Color = FillColor;
        //        FillSymbol.Outline = Outline;
        //        FillSymbol.Style = FillStyle;

        //        return FillSymbol;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        return null;
        //    }
        //}

        //internal static IMarkerSymbol ReturnMarkerSymbol(string SymbolName, string StyleName, IColor SymbolColor, int SymbolSize)
        //{
        //    try
        //    {
        //        Type t = Type.GetTypeFromProgID("esriFramework.StyleGallery");
        //        System.Object obj = Activator.CreateInstance(t);
        //        IStyleGallery Gallery = obj as IStyleGallery;
        //        IStyleGalleryStorage GalleryStorage = (IStyleGalleryStorage)Gallery;
        //        string StylePath = GalleryStorage.DefaultStylePath + StyleName;

        //        bool StyleFound = false;

        //        for (int i = 0; i < GalleryStorage.FileCount; i++)
        //        {
        //            if (GalleryStorage.get_File(i).ToUpper() == StyleName.ToUpper())
        //            {
        //                StyleFound = true;
        //                break;
        //            }
        //        }

        //        if (!StyleFound)
        //            GalleryStorage.AddFile(StylePath);

        //        IEnumStyleGalleryItem EnumGalleryItem = Gallery.get_Items("Marker Symbols", StyleName, "DEFAULT");
        //        IStyleGalleryItem GalleryItem = EnumGalleryItem.Next();

        //        while (GalleryItem != null)
        //        {
        //            if (GalleryItem.Name == SymbolName)
        //            {
        //                IClone SourceClone = (IClone)GalleryItem.Item;
        //                IClone DestClone = SourceClone.Clone();
        //                IMarkerSymbol MarkerSymbol = (IMarkerSymbol)DestClone;
        //                MarkerSymbol.Color = SymbolColor;
        //                MarkerSymbol.Size = SymbolSize;
        //                return MarkerSymbol;
        //            }
        //            GalleryItem = EnumGalleryItem.Next();
        //        }

        //        MessageBox.Show("Unable to locate Marker Symbol '" + SymbolName + "' in style '" + StyleName + "'.");
        //        return null;

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        return null;
        //    }
        //}

        //internal static ISimpleMarkerSymbol ReturnSimpleMarkerSymbol(IColor MarkerColor, IColor OutlineColor, esriSimpleMarkerStyle MarkerStyle, double MarkerSize, double OutlineSize)
        //{
        //    try
        //    {
        //        ISimpleMarkerSymbol MarkerSymbol = new SimpleMarkerSymbolClass();
        //        MarkerSymbol.Color = MarkerColor;
        //        MarkerSymbol.Style = MarkerStyle;
        //        MarkerSymbol.Size = MarkerSize;
        //        MarkerSymbol.Outline = true;
        //        MarkerSymbol.OutlineSize = OutlineSize;
        //        MarkerSymbol.OutlineColor = OutlineColor;

        //        return MarkerSymbol;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        return null;
        //    }
        //}

        //internal static IPictureMarkerSymbol ReturnPictureMarkerSymbol(Bitmap SourceBitmap, IColor TransparentColor, double MarkerSize)
        //{
        //    try
        //    {
        //        IPictureMarkerSymbol PictureSymbol = new PictureMarkerSymbolClass();
        //        PictureSymbol.Picture = (IPictureDisp)OLE.GetIPictureDispFromBitmap(SourceBitmap);
        //        PictureSymbol.Size = MarkerSize;
        //        PictureSymbol.BitmapTransparencyColor = TransparentColor;
        //        return PictureSymbol;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        return null;
        //    }

        //}


        #endregion

        #region TABLES AND FIELDS

        public static FieldCategory GetFieldCategory(ArcGIS.Desktop.Mapping.FieldDescription fieldDescription)
        {
            try
            {
                FieldCategory fc;

                switch (fieldDescription.Type)
                {
                    // Field values require single quotes
                    case FieldType.String:
                    case FieldType.GUID:
                    case FieldType.GlobalID:
                        fc = FieldCategory.STRING;
                        break;

                    // Field values require datestamp ''
                    case FieldType.Date:
                        fc = FieldCategory.DATE;
                        break;

                    // Field values require nothing, just the value
                    case FieldType.Double:
                    case FieldType.Integer:
                    case FieldType.OID:
                    case FieldType.Single:
                    case FieldType.SmallInteger:
                        fc = FieldCategory.NUMERIC;
                        break;

                    // Everything else...
                    default:
                        fc = FieldCategory.OTHER;
                        break;
                }

                return fc;
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return FieldCategory.UNKNOWN;
            }
        }

        #endregion

        #region GEOMETRIES

        public static async Task<string> GetNationalGridInfo(Polygon cellPolygon)
        {
            try
            {

                return "hi";
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return "hi";
            }
        }

        #endregion

        #region MISCELLANEOUS

        public static (bool ValueFound, int Value, string AdjustedString) ExtractValueFromString(string string_to_search, string regex_pattern)
        {
            (bool, int, string) errTuple = (false, 0, "");

            try
            {
                Regex regex = new Regex(regex_pattern);
                Match match = regex.Match(string_to_search);

                if (match.Success)
                {
                    string matched_pattern = match.Value;                                                   // match.Value is the [n], [nn], or [nnn] substring includng the square brackets
                    string string_adjusted = string_to_search.Replace(matched_pattern, "").Trim();          // string to search minus the [n], [nn], or [nnn] substring, then trim
                    string value_string = matched_pattern.Replace("[", "").Replace("]", "").Replace("{", "").Replace("}", "");    // leaves just the 1, 2, or 3 numeric digits, no more brackets

                    // convert text value to int
                    if (!int.TryParse(value_string, out int value_int))
                    {
                        return errTuple;
                    }

                    return (true, value_int, string_adjusted);
                }
                else
                {
                    return errTuple;
                }
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
                return errTuple;
            }
        }

        public static string GetUser()
        {
            string[] fulluser = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split(new Char[] { '\\' });
            return fulluser[fulluser.Length - 1];
        }


        /// <summary>
        /// Examines a Cancellation Token and determines if the token has been cancelled.
        /// If it has been cancelled, the method will either throw an OperationCancelledException
        /// (if the throw_error parameter is true) or return a true (cancelled) or false
        /// (not cancelled).
        /// </summary>
        /// <param name="token"></param>
        /// <param name="throw_error"></param>
        /// <returns></returns>
        public static bool CheckForCancellation(CancellationToken token, bool throw_error = true)
        {
            try
            {
                // Check if the token has been cancelled
                if (token.IsCancellationRequested)
                {
                    // throw the error if requested
                    if (throw_error)
                    {
                        token.ThrowIfCancellationRequested();
                    }
                    else
                    {
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

    }
}

/*

        internal static void ExportDataViewToExcel(DataView DV, string Title, string SheetName)
        {
            IProgressDialog2 PD = null;
            Excel.Application xlApp = null;
            Excel.Workbook xlWb = null;
            Excel.Worksheet xlWs = null;
            Excel.Range TempRange = null;

            try
            {
                if (DV == null)
                {
                    return;
                }
                else if (DV.Count == 0)
                {
                    return;
                }

                //create progress dialog
                PD = CreateProgressDialog("EXPORT TO EXCEL", 10, IApp.hWnd, esriProgressAnimationTypes.esriProgressGlobe);
                IStepProgressor Stepper = (IStepProgressor)PD;
                Stepper.Message = "Creating Excel Document...";
                Stepper.Step();

                //Make Sure Excel is present and will start
                xlApp = new Excel.ApplicationClass();
                if (xlApp == null)
                {
                    MessageBox.Show("EXCEL could not be started.  Verify your MS Office Installation and/or project references...");
                    return;
                }

                xlWb = xlApp.Workbooks.Add(Missing.Value);
                xlWs = (Excel.Worksheet)xlWb.Worksheets.get_Item(1);
                if (xlWs == null)
                {
                    MessageBox.Show("Worksheet could not be created.  Verify your MS Office Installation and/or project references...");
                    return;
                }
                xlWs.Name = SheetName;

                //Insert Title Information
                xlWs.Cells[1, 1] = Title;
                TempRange = (Excel.Range)xlWs.Cells[1, 1];
                TempRange.Font.Bold = false;
                TempRange.Font.Size = 15;

                xlWs.Cells[2, 1] = "Exported on " + DateTime.Now.ToLongDateString() + "  by " + GetUser();
                TempRange = (Excel.Range)xlWs.Cells[2, 1];
                TempRange.Font.Bold = false;
                TempRange.Font.Italic = true;
                TempRange.Font.Size = 10;

                TempRange = xlWs.get_Range("A1", "W2");
                TempRange.Borders.Color = ColorTranslator.ToOle(Color.LightYellow);
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = ColorTranslator.ToOle(Color.Black);
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Color = ColorTranslator.ToOle(Color.Black);

                //Add Column Headers
                DataTable DT = DV.ToTable("sptab");
                int ColCount = DT.Columns.Count;
                int RowCount = DV.Count;

                for (int i = 1; i <= ColCount; i++)
                {
                    xlWs.Cells[3, i] = DT.Columns[i - 1].ColumnName;
                }

                TempRange = (Excel.Range)xlWs.Rows[3, Missing.Value];
                TempRange.Font.Size = 10;
                TempRange.Font.Bold = true;
                TempRange.Interior.Color = ColorTranslator.ToOle(Color.Gold);
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Color = ColorTranslator.ToOle(Color.Black);
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = ColorTranslator.ToOle(Color.Black);
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
                TempRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Color = ColorTranslator.ToOle(Color.Black);

                //Now Add DataView information
                Stepper.Message = "Loading Excel Sheet...";
                Stepper.MaxRange = RowCount;

                object[,] rowdata = new object[RowCount, ColCount];

                for (int i = 0; i < RowCount; i++)
                {
                    for (int j = 0; j < ColCount; j++)
                    {
                        if (DT.Rows[i].ItemArray.GetValue(j) is string)
                        {
                            string fullvalue = DT.Rows[i].ItemArray.GetValue(j).ToString();
                            int stringlen = fullvalue.Length;
                            if (stringlen > 911)
                            {
                                rowdata[i, j] = fullvalue.Substring(0, 911);
                            }
                            else
                            {
                                rowdata[i, j] = fullvalue;
                            }
                        }
                        else
                            rowdata[i, j] = DT.Rows[i].ItemArray.GetValue(j);
                    }
                    Stepper.Step();
                }

                TempRange = xlWs.get_Range("A4", Missing.Value);
                TempRange = TempRange.get_Resize(RowCount, ColCount);
                TempRange.set_Value(Missing.Value, rowdata);

                //autofit columns
                TempRange = xlWs.get_Range("A3", Missing.Value);
                TempRange = TempRange.get_Resize(RowCount + 1, ColCount);
                TempRange.Columns.AutoFit();

                xlApp.Visible = true;
                xlApp.UserControl = true;


            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                if (xlApp != null)
                {
                    xlApp.DisplayAlerts = false;
                    xlApp.Quit();
                    Marshal.FinalReleaseComObject(xlApp);
                }
            }

            finally
            {
                PD.HideDialog();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
 
 
 
 */
