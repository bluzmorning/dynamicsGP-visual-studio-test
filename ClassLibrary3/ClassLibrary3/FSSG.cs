using System;
using Microsoft.Dexterity.Bridge;
using Microsoft.Dexterity.Applications;
using Microsoft.Dexterity.Applications.FssgDictionary;
using Excel = Microsoft.Office.Interop.Excel;

namespace GP2015
{

    public class GPAddIn : IDexterityAddIn
    {
        // Add a local variable to the PM Vendor Maintenance
        private PmVendorMaintenanceForm pmTrxForm = Fssg.Forms.PmVendorMaintenance;

        public void Initialize()
        {
            // Create a subscription to the event of the PM Vendor window opening
            this.pmTrxForm.PmVendorMaintenance.TestExcel.ClickAfterOriginal += new EventHandler(CreateExcel);
            this.pmTrxForm.PmVendorMaintenance.OpenAfterOriginal += new EventHandler(OnPMVendornOpened);
        }

        // Create a method that is called when the PM Vendor window is opened
        private void OnPMVendornOpened(object sender, EventArgs e)
        {
            // Add a default value to the open PmVendorMaintenance fields
            this.pmTrxForm.PmVendorMaintenance.City.Value = "Barcelona";
            this.pmTrxForm.PmVendorMaintenance.Country.Value = "Spain";
        }
        // Create a method that is called when the testExcel button is pushed
        static void CreateExcel(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            Excel.Workbook wb = excel.Workbooks.Open("testdata.xlsx");
            Excel.Worksheet sh = wb.Sheets.Add();
            sh.Name = "TestSheet";
            sh.Cells[1, "A"].Value2 = "SNO";
            sh.Cells[2, "B"].Value2 = "A";
            sh.Cells[2, "C"].Value2 = "1122";
            wb.Close(true);
            excel.Quit();
        }
    }
}