using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;

namespace NCC.PRZTools
{
    internal class Module_Main : Module
    {
        private static Module_Main _this = null;

        /// <summary>
        /// Retrieve the singleton instance to this module here
        /// </summary>
        public static Module_Main Current
        {
            get
            {
                return _this ?? (_this = (Module_Main)FrameworkApplication.FindModule("prz_module_main"));
            }
        }

        internal Module_Main()
        {
            try
            {
                // disable state upon which Daml condition of unwanted UI buttons is based
                FrameworkApplication.State.Deactivate("prz_disabled_state");
            }
            catch
            {

            }
        }

        #region Overrides
        /// <summary>
        /// Called by Framework when ArcGIS Pro is closing
        /// </summary>
        /// <returns>False to prevent Pro from closing, otherwise True</returns>
        protected override bool CanUnload()
        {
            //TODO - add your business logic
            //return false to ~cancel~ Application close
            return true;
        }

        #endregion Overrides

    }
}
