using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using System;
using System.Reflection;
using ProMsgBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using PRZH = NCC.PRZTools.PRZHelper;

namespace NCC.PRZTools
{
    internal class Button_Serialize_Database : Button
    {
        protected override async void OnClick()
        {
            try
            {
                #region SHOW DIALOG

                SerializeDatabase dlg = new SerializeDatabase();                    // View
                SerializeDatabaseVM vm = (SerializeDatabaseVM)dlg.DataContext;      // View Model

                dlg.Owner = FrameworkApplication.Current.MainWindow;

                // Closing event handler
                dlg.Closing += (o, e) =>
                {
                    // Event handler for Dialog closing event
                    if (vm.OperationIsUnderway)
                    {
                        ProMsgBox.Show("Operation is underway.  Please cancel the operation before closing this window.");
                        e.Cancel = true;
                    }
                };

                // Closed Event Handler
                dlg.Closed += (o, e) =>
                {
                    // Event Handler for Dialog close in case I need to do things...
                    // ProMsgBox.Show("Closed...");
                    // System.Diagnostics.Debug.WriteLine("Pro Window Dialog Closed";)
                };

                // Loaded Event Handler
                dlg.Loaded += async (sender, e) =>
                {
                    if (vm != null)
                    {
                        await vm.OnProWinLoaded();
                    }
                };

                var result = dlg.ShowDialog();

                // Take whatever action required here once the dialog is close (true or false)
                // do stuff here!

                #endregion
            }
            catch (Exception ex)
            {
                ProMsgBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
        }
    }
}
