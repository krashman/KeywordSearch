using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 


namespace WindowsFormsApplication2
{
    class XLSFileParser : Parsers
    {
        //string filePath;

        public string gettext(string filePath)
        {
           // Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str = "";
            int rCnt = 0;
            int cCnt = 0;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            int totalSheets = (int)xlWorkBook.Sheets.Count;
            for (int scnt = 1; scnt <= totalSheets; scnt++)
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(scnt);

                range = xlWorkSheet.UsedRange;

                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        str += (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2 + "  ";
                    }
                }
                releaseObject(xlWorkSheet);
            }
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            //MessageBox.Show(str);
            return str;
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
