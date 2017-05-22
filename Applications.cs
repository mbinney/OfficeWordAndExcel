using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace Office
{
    public partial class Applications : Form
    {
        public Applications()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            build_excel_file();
            MessageBox.Show("Finished");
            

        }

        private void build_excel_file()
        {

            Microsoft.Office.Interop.Excel.Application objApp;                
            Workbook objWB;
            Worksheet objWS;
            string filename = "c:\\temp\\interop.xlsx";
            object misValue = System.Reflection.Missing.Value;

            //objApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            objApp = new Microsoft.Office.Interop.Excel.Application();

            objWB = objApp.Workbooks.Add();
            objWS = (Worksheet)objWB.Sheets[1];

            objWS.Select();

            objWS.Cells[1, 1] = "First Name";
            objWS.Cells[1, 2] = "Last Name";
            objWS.Cells[1, 3] = "Full Name";
            objWS.Cells[1, 4] = "Salary";
                       
            objApp.Visible = false;

            objWB.Close(true, filename, misValue);
            //objWB.Close(0); -- no save changes.

            //http://stackoverflow.com/questions/8977571/excel-process-remains-open-after-interop-traditional-method-not-working
            //These will close the excel object and force garbage collection.  Watch Task manager. Keep on eye on the Excel process.
            objApp.Quit();
            if (objWB != null) { Marshal.ReleaseComObject(objWB); } //release each workbook like this
            if (objWS != null) { Marshal.ReleaseComObject(objWS); } //release each worksheet like this
            if (objApp != null) { Marshal.ReleaseComObject(objApp); } //release the Excel application
            objWB = null; //set each memory reference to null.
            objWS = null;
            objApp = null;
            GC.Collect();


            

        }

        public void build_word_doc() {

            //http://www.c-sharpcorner.com/article/word-automation-using-C-Sharp/
            object misValue = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application objApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document objDoc = new Microsoft.Office.Interop.Word.Document();

            objApp.Documents.Add(misValue, misValue, misValue, misValue);
            
            objApp.Visible = true;

            
            if (objApp != null) { Marshal.ReleaseComObject(objApp); }
            if (objDoc != null) { Marshal.ReleaseComObject(objDoc); }

            //http://stackoverflow.com/questions/8977571/excel-process-remains-open-after-interop-traditional-method-not-working
            //These will close the excel object and force garbage collection.  Watch Task manager. Keep on eye on the Excel process.
            objApp = null;
            objDoc = null;
            GC.Collect();


        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            build_word_doc();
        }
    }
}
