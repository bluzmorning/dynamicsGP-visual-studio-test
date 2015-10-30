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
        private void CreateExcel(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            Excel.Workbook wb = excel.Workbooks.Add();
            Excel.Worksheet sh = wb.Sheets.Add();
            sh.Name = "JPMorgan";
            sh.Cells[1, "A"].Value = "Vendor Name";
            sh.Cells[1, "B"].Value = "City";
            sh.Cells[1, "C"].Value = "Zip Code";
            sh.Cells[1, "D"].Value = "Tax ID";
            sh.Cells[1, "E"].Value = "Phone";
            sh.Cells[1, "F"].Value = "Email";
            sh.Cells[1, "G"].Value = "Currency ID";
            sh.Cells[2, "A"].Value = this.pmTrxForm.PmVendorMaintenance.VendorName.Value;
            sh.Cells[2, "B"].Value = this.pmTrxForm.PmVendorMaintenance.City.Value;
            sh.Cells[2, "C"].Value = this.pmTrxForm.PmVendorMaintenance.ZipCode.Value;
            sh.Cells[2, "D"].Value = this.pmTrxForm.PmVendorMaintenanceAdditionalInformation.TaxIdNumber.Value;
            sh.Cells[2, "E"].Value = this.pmTrxForm.PmVendorMaintenance.PhoneNumber1.Value;
            sh.Cells[2, "F"].Value = this.pmTrxForm.PmVendorMaintenance.Comment1.Value;
            sh.Cells[2, "G"].Value = this.pmTrxForm.PmVendorMaintenanceAdditionalInformation.CurrencyId.Value;
            wb.SaveAs(@"D:\Excel\testdata.xlsx");
            wb.Close(true);
            excel.Quit();
        }
    }
}