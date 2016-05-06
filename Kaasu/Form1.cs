using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace Kaasu
{
    

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public class raw_data_struct
        {
            public string type_of_transcation;
            public DateTime transaction_date;
            public DateTime posting_date;
            public string given_description;
            public string modified_description;
            public float cost;
        };

        public List<raw_data_struct> fromExcel = new List<raw_data_struct>();


        private void button1_Click(object sender, EventArgs e)
        {
            string excel_path = @"C:\Users\sanand2\Dropbox\Finance\Chase.CSV";

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(excel_path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //Clear all formatting in the blank cells
            xlWorkSheet.Columns.ClearFormats();
            xlWorkSheet.Rows.ClearFormats();

            range = xlWorkSheet.UsedRange;

            //int rCnt = xlWorkSheet.UsedRange.Rows.Count;
            //int cCnt = xlWorkSheet.UsedRange.Columns.Count;
            string csv_row_data;
            
            for (int rCnt = 2; rCnt <= xlWorkSheet.UsedRange.Rows.Count; rCnt++)
            {
                for (int cCnt = 1; cCnt <= xlWorkSheet.UsedRange.Columns.Count; cCnt++)
                {
                    //Get entire row data
                    csv_row_data = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    string[] row_data = csv_row_data.Split(',');
                    
                    if(row_data[0].Trim() == "" || row_data[1].Trim() == "" || row_data[2].Trim() == "" || row_data[3].Trim() == "" || row_data[4].Trim() == "")
                    {
                        //TODO; Print warning
                        continue;
                    }

                    fromExcel.Add(new raw_data_struct
                    {
                        type_of_transcation     = row_data[0],
                        transaction_date        = DateTime.Parse(row_data[1]),
                        posting_date            = DateTime.Parse(row_data[2]),
                        given_description       = row_data[3],
                        cost                    = float.Parse(row_data[4]),
                        modified_description    = ""
                    });
                }
            }

            //At this point we have all the excel raw data in fromExcel

            //Sort according to posting date
            fromExcel.Sort((x, y) => x.transaction_date.CompareTo(y.transaction_date));




            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
