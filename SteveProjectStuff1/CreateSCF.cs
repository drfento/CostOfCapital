using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CostOfCapital
{
    class CreateSCF
    {

        public void GenerateFile(CapitalFrm cform)
        {
            // Test paths
            //string archivefileloc = @"\\cnjfile00\OperCommon\Operations\EQUITIES\Cost of Capital\Temp\";
            //string tempfileloc = @"\\cnjfile00\OperCommon\Operations\EQUITIES\Cost of Capital\Temp\";
            //string finalfileloc = @"C:\Temp\";

            string archivefileloc = @"\\cnjfile00\OperCommon\Operations\EQUITIES\Cost of Capital\";
            string tempfileloc = @"\\cnjfile00\opercommon\gdk_divadjustment\MurexFiles\Temp\";
            string finalfileloc = @"\\cnjfile00\opercommon\gdk_divadjustment\MurexFiles\";

            string finalfilename = DateTime.Now.ToString("yyyyMMdd") + DateTime.Now.ToString("HHmmss") + "_COSTCAPSCF_new.csv";
            string archivefilename = DateTime.Now.ToString("MMMM") + "_" + DateTime.Now.ToString("yyyy") + "_COSTCAPSCF.xls";

            Connection connection = new Connection();
            GetSQL sql = new GetSQL();

            //loop through feed view and use merge to update values
            for (int i = 0; i < cform.dgvFeedView.Rows.Count; ++i)
            {
               connection.SendQuery(sql.Final_Sent_Merge(cform, i));
            }

            // Create and save CSV
            try
            {
                // CreateCSV
                CreateCSV(cform.dgvFinalView, Path.Combine(tempfileloc, finalfilename));
                File.Move(Path.Combine(tempfileloc, finalfilename), Path.Combine(finalfileloc, finalfilename));
                //Archive Data   
                CreateExcel(cform, archivefileloc, archivefilename);
                //Clear final view to force action
                cform.dgvFinalView.Rows.Clear();
                cform.dgvFinalView.Refresh();
                //Select tab
                cform.CapitalTabs.SelectTab(1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception occurred while creating files " + ex.Message);
            }
                
        }

        public void CreateCSV(DataGridView gridIn, string outputFile)
        {
            string value = "";
            DataGridViewRow dr = new DataGridViewRow();
            StreamWriter swOut = new StreamWriter(outputFile);

            //write header rows to csv
            for (int i = 0; i <= gridIn.Columns.Count - 1; i++)
            {
                if (i > 0)
                {
                    swOut.Write(",");
                }
                swOut.Write(gridIn.Columns[i].HeaderText);
            }

            swOut.WriteLine();

            //write DataGridView rows to csv
            for (int j = 0; j <= gridIn.Rows.Count - 1; j++)
            {
                if (j > 0)
                {
                    swOut.WriteLine();
                }

                dr = gridIn.Rows[j];

                for (int i = 0; i <= gridIn.Columns.Count - 1; i++)
                {
                    if (i > 0)
                    {
                        swOut.Write(",");
                    }

                    value = dr.Cells[i].Value.ToString();
                    //replace comma's with spaces
                    value = value.Replace(',', ' ');
                    //replace embedded newlines with spaces
                    value = value.Replace(Environment.NewLine, " ");

                    swOut.Write(value);
                }
            }
            swOut.Close();
        }

        public void CreateExcel(CapitalFrm cform, string filepath, string filename)
        {
            object misValue = System.Reflection.Missing.Value;
            Excel.Application xlexcel = new Excel.Application()
            { DisplayAlerts = false};

            //Add workbook
            Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);

            // Loop through DGVS
            for (int i = 1; i < 4; ++i)
            {
                // Copy DataGridView results to clipboard
                switch (i)
                {
                    case 1:
                        //cform.dgvCostView.RowHeadersVisible = false; //No hidden row
                        CopyAlltoClipboard(cform.dgvCostView);
                        break;
                    case 2:
                        CopyAlltoClipboard(cform.dgvFeedView);
                        break;
                    case 3:
                        CopyAlltoClipboard(cform.dgvFinalView);
                        break;
                }



                //select sheet
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(i);
                // Paste clipboard results to worksheet range
                xlWorkSheet.Activate();
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                xlWorkSheet.get_Range("A2").Select();

                ReleaseObject(xlWorkSheet);
                ReleaseObject(CR);
            }

                // Save the excel file
                xlWorkBook.SaveAs(Path.Combine(filepath,filename), Excel.XlFileFormat.xlWorkbookNormal, 
                    misValue, misValue, misValue, misValue, 
                    Excel.XlSaveAsAccessMode.xlExclusive, 
                    misValue, misValue, misValue, misValue, misValue);

                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                cform.dgvFinalView.ClearSelection();
                cform.dgvFeedView.ClearSelection();
                cform.dgvCostView.ClearSelection();

        }

        private void CopyAlltoClipboard(DataGridView gridIn)
        {
            gridIn.SelectAll();
            DataObject dataObj = gridIn.GetClipboardContent();
            if (dataObj != null)
            { Clipboard.SetDataObject(dataObj); }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
