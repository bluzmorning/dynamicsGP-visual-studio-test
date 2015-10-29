using System;
using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.DynamicsDictionary;
//using Microsoft.Office.Interop.Excel;

namespace GP2015
{

    public class GPAddIn : IDexterityAddIn
    {
        // Add a local variable to the PM Vendor Maintenance
        private PmVendorMaintenanceForm pmTrxForm = Dynamics.Forms.PmVendorMaintenance;

        public void Initialize()
        {
            // Create a subscription to the event of the PM Vendor window opening
            this.pmTrxForm.PmVendorMaintenance.OpenAfterOriginal += new EventHandler(OnPMVendornOpened);
        }

        // Create a method that is called when the PM Vendor window is opened
        private void OnPMVendornOpened(object sender, EventArgs e)
        {
            // Add a default value to the open PmVendorMaintenance fields
            this.pmTrxForm.PmVendorMaintenance.City.Value = "Barcelona";
            this.pmTrxForm.PmVendorMaintenance.Country.Value = "Spain";
        }
    }
}