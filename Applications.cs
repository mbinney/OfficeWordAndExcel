using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

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

        public void Create_PDF_from_DOC(string DocFileName, string PDFFileName) {

            try
            {            

                //http://www.c-sharpcorner.com/article/word-automation-using-C-Sharp/
                object misValue = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Word.Application objApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document objDoc = new Microsoft.Office.Interop.Word.Document();
            
                objDoc = objApp.Documents.Open(DocFileName);
                objDoc.ExportAsFixedFormat(PDFFileName, WdExportFormat.wdExportFormatPDF);

                objDoc.Close();
                objApp.Quit();
           
                if (objDoc != null) { Marshal.ReleaseComObject(objDoc); }
                if (objApp != null) { Marshal.ReleaseComObject(objApp); }

                //http://stackoverflow.com/questions/8977571/excel-process-remains-open-after-interop-traditional-method-not-working
                //These will close the excel object and force garbage collection.  Watch Task manager. Keep on eye on the Excel process.
                objApp = null;
                objDoc = null;
                GC.Collect();

            }
            catch (Exception ex)
            {

                throw new System.ArgumentException(ex.Message.ToString());
            }


        }


        private void createMergedDoc(string DOCFileName)
        { 
            try
                {

                    //http://www.c-sharpcorner.com/article/word-automation-using-C-Sharp/
                    object misValue = System.Reflection.Missing.Value;

                    Microsoft.Office.Interop.Word.Application objApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document objDoc = new Microsoft.Office.Interop.Word.Document();

                    objDoc = objApp.Documents.Open(DOCFileName);

                    foreach (Microsoft.Office.Interop.Word.Field myMergedFields in objDoc.Fields)
                    {
                        //Microsoft.Office.Interop.Word.Range  rngFieldCode = myMergedFields
                        Microsoft.Office.Interop.Word.Range rngFieldCode = myMergedFields.Code;
                        String fieldText = rngFieldCode.Text;

                        Debug.Print(fieldText.ToString());                        

                        if (fieldText.StartsWith(" MERGEFIELD"))
                        {
                            Int32 endMerge = fieldText.IndexOf("\\");
                            Int32 fieldNameLength = fieldText.Length - endMerge;
                            String fieldName = fieldText.Substring(11, endMerge - 11);
                            fieldName = fieldName.Trim();

                            if (fieldName == "Title")
                            {
                                myMergedFields.Select();                                
                                objApp.Selection.TypeText("Hello World");
                            }

                        Debug.Print(fieldName);
                        }

                    }

                    objDoc.Close();
                    objApp.Quit();

                    if (objDoc != null) { Marshal.ReleaseComObject(objDoc); }
                    if (objApp != null) { Marshal.ReleaseComObject(objApp); }

                    //http://stackoverflow.com/questions/8977571/excel-process-remains-open-after-interop-traditional-method-not-working
                    //These will close the excel object and force garbage collection.  Watch Task manager. Keep on eye on the Excel process.
                    objApp = null;
                    objDoc = null;
                    GC.Collect();

                }
                catch (Exception ex)
                {

                    throw new System.ArgumentException(ex.Message.ToString());
                }
            }       
        

        private void btnWord_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            string PDFFileName = @"c:\temp\32321";
            string DOCFileName = @"c:\temp\32321.doc";

            for (Int32 i = 0; i < 30; i++)
            {                
                Create_PDF_from_DOC(DOCFileName, PDFFileName + i + ".pdf");
            }
            
            MessageBox.Show("Finished");
            Cursor.Current = Cursors.Default;

        }

        private void btnMerge_Click(object sender, EventArgs e)
        {
            string DOCFileName = @"c:\temp\MailMergeDemo.docx";            
            createMergedDoc(DOCFileName);
        }
    }
}
