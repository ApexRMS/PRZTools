using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Mapping;
using System;
using System.Reflection;
using PRZH = NCC.PRZTools.PRZHelper;

namespace NCC.PRZTools
{
    internal class Button_LoadPRZLayers : Button
    {
        protected override async void OnClick()
        {
            try
            {
                Map map = MapView.Active.Map;

                if (!(await PRZH.RedrawPRZLayers(map)).success)
                {
                    MessageBox.Show("Unable to redraw the PRZ layers");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + "Error in method: " + MethodBase.GetCurrentMethod().Name);
            }
        }
    }
}
