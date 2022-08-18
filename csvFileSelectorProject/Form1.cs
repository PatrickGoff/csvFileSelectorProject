using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
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
        double sum = 0;
        



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
            }
        }
   
    
        private void ReadExcel(string sFileName)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(sFileName);
            xlWorksheet = xlApp.Worksheets["ExampleCSV"];

            for (iRow = 1; iRow <= xlWorksheet.Rows.Count; iRow++)  // START FROM THE FIRST ROW.
            {
                if (xlWorksheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    
                    txtResults.Text = sum.ToString();
                }
            }

            xlWorkbook.Close();
            xlApp.Quit();


        }
    
    
    
    
    
    }
}
