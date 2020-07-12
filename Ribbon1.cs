using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace JobApplications_Excel_Add_in
{
    public partial class JobsTracker
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnNewSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet worksheet = activeWorkbook.ActiveSheet;

            try { worksheet.Name = "Job Applications"; }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }

            worksheet.Range["A1"].Value2 = "Company Name";
            worksheet.Range["B1"].Value2 = "Position";
            worksheet.Range["C1"].Value2 = "Date Applied";
            worksheet.Range["D1"].Value2 = "Response Date";
            worksheet.Range["E1"].Value2 = "Offer/Reject";
            worksheet.Range["F1"].Value2 = "Primary Contact/Recruiter";
            worksheet.Range["G1"].Value2 = "Referral Name";
            int i = 1;
            for (char ch = 'A'; ch <= 'H'; ch++)
            {
                worksheet.Range[ch+"1"].Font.Bold = true;
                worksheet.Columns[i].ColumnWidth = 25;
                i++;
            }
        }

        private void btnJobdetails_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet worksheet = activeWorkbook.ActiveSheet;
            Excel.Range rng = worksheet.Application.ActiveCell;
            int row = rng.Row;
            object cellvalue = rng.Value;
            if(cellvalue == null && row>1)
            {
                NewJob obj = new NewJob(row);
                obj.AppliedDate();
                obj.ShowDialog();
            }
            else if(row>1)
            {
                NewJob obj = new NewJob(row);
                obj.ShowDialog();
            }
        }
    }
}
