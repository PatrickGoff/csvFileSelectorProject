using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace csvFileSelectorProject
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //Excel objects
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorksheet;
        string sFileName;
        int iRow, iCol = 1;
        BigInteger sum = 0;
       
        



        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "CSV files | *.csv";
            dialog.Multiselect = false;
            if(dialog.ShowDialog() == DialogResult.OK)
            {
                sFileName = dialog.FileName; 

                if(sFileName.Trim() != "")
                {
                    ReadExcel(sFileName);
                }

                lblFileName.Text = sFileName;
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void ReadExcel(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(sFile);
            xlWorksheet = xlApp.ActiveWorkbook.Sheets[1];
         
            Excel.Range xlRng = xlWorksheet.Range["A:A"];
           

            for (iRow = 1; iRow <= xlWorksheet.UsedRange.Rows.Count; iRow++)  // START FROM THE FIRST ROW.
            {
                
                string excelValue = xlWorksheet.Cells[iRow, iCol].value == null
                 ? MessageBox.Show("The first cell in column A is empty.")
                 : xlWorksheet.Cells[iRow, iCol].value.ToString();
                  excelValue = excelValue.Replace("'", " ");



                sum += BigInteger.TryParse(excelValue, out BigInteger _)
                ? BigInteger.Parse(excelValue)
                : 0;





            }

            String digits = sum.ToString();
           string digit = digits.Substring(Math.Max(0, digits.Length - 10));



                txtResults.Text = digit;
                xlWorkbook.Close();
                xlApp.Quit();

            //}
        }
    
    
    
    
    
    }
}
