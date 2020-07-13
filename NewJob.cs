using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace JobApplications_Excel_Add_in
{
    public partial class NewJob : Form
    {
        public int row;
        public NewJob(int row)
        {
            InitializeComponent();
            this.row = row;
            this.rdB1.Parent = panel1;
            this.rdB2.Parent = panel1;
            InitialSetup(row);
        }

        private void InitialSetup(int row)
        {
            this.row = row;
            Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = (Excel.Worksheet)activeWorkbook.Worksheets["Job Applications"];
            
            tbCompanyName.Text = Convert.ToString(worksheet.Range["A" + row].Value);
            tbPosition.Text = Convert.ToString(worksheet.Range["B" + row].Value);
            tbAppliedDate.Text = Convert.ToString(worksheet.Range["C" + row].Value);
            tbResponseDate.Text = Convert.ToString(worksheet.Range["D" + row].Value);
            tbPrimaryContact.Text = Convert.ToString(worksheet.Range["E" + row].Value);
            tbReferralName.Text = Convert.ToString(worksheet.Range["F" + row].Value);
            if (worksheet.Range["G" + row].Value2 == "Yes")
                rdB1.Checked = true;
            else if(worksheet.Range["G"+row].Value2 == "No")
                rdB2.Checked = true;
        }


        private void btnNext_Click(object sender, EventArgs e)
        {
            Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = (Excel.Worksheet)activeWorkbook.Worksheets["Job Applications"];
            if (tbCompanyName.Text != "")
            {
                SaveData();
            }
            if (row > 1)
                InitialSetup(row+1);
            if(worksheet.Range["G" + row].Value2 == null)
            {
                this.rdB1.Checked = false;
                this.rdB2.Checked = false;
            }
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (row > 2)
                InitialSetup(row - 1);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (tbCompanyName.Text == "" || tbPosition.Text == "")
            {
                DialogResult d = MessageBox.Show("Please Enter Company name & Position", "", MessageBoxButtons.OKCancel);
                if (d == DialogResult.OK)
                {
                    this.Show();
                }
                else
                    this.Close();
            }
            else
            {
                SaveData();
                this.Close();
            }
        }
        public void AppliedDate()
        {
            DateTime dt = DateTime.Now;
            this.tbAppliedDate.Text = Convert.ToString(dt.Month + "/" + dt.Day + "/" + dt.Year);
        }
        public void SaveData()
        {
            try
            {
                Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                Excel.Worksheet worksheet = activeWorkbook.ActiveSheet;
                if (tbCompanyName.Text != "")
                {
                    worksheet.Range["A" + row].Value = tbCompanyName.Text;
                }
                if (tbPosition.Text != "")
                {
                    worksheet.Range["B" + row].Value = tbPosition.Text;
                }
                if (tbAppliedDate.Text != "")
                {
                    worksheet.Range["C" + row].Value = tbAppliedDate.Text;
                }
                if (tbResponseDate.Text != "")
                {
                    worksheet.Range["D" + row].Value = tbResponseDate.Text;
                }
                if (tbPrimaryContact.Text != "")
                {
                    worksheet.Range["E" + row].Value = tbPrimaryContact.Text;
                }
                if (tbReferralName.Text != "")
                {
                    worksheet.Range["F" + row].Value = tbReferralName.Text;
                }
                if (rdB1.Checked)
                    worksheet.Range["G" + row].Value = "Yes";
                else if (rdB2.Checked)
                    worksheet.Range["G" + row].Value = "No";
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDiscard_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
