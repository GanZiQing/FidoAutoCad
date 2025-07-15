using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using System.Xml.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDB = Autodesk.AutoCAD.DatabaseServices;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;



namespace FidoAutoCad.Forms
{
    public partial class FidoAutoCadMain : Form
    {
        AutoCADCommands parent;
        public FidoAutoCadMain(AutoCADCommands parent)
        {
            InitializeComponent();
            this.parent = parent;
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
        }

        Excel.Application excelApp = null;
        private void button1_Click(object sender, EventArgs e)
        {
            //string msg = textBox1.Text;
            //parent.PrintText(msg);
            TestFunction();
        }

        #region Excel Attachment
        private void launchExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelApp != null)
                {
                    excelApp.Visible = true;
                    UpdateExcelStatus(true, "NA");
                    DialogResult res = MessageBox.Show("Excel already attached, close previous instance of excel?", "Warning", MessageBoxButtons.YesNo);
                    if (res == DialogResult.Yes)
                    {
                        excelApp.Quit();
                        InternalReleaseExcel();
                    }
                    else
                    {
                        throw new Exception("Instance of excel is already attached.");
                    }
                }

                AttachNewExcelInstance();
            }
            catch (Exception ex) when (ex.HResult == -2147023174) // excelApp not set to null, but not found
            {
                InternalReleaseExcel(false);
                AttachNewExcelInstance();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message + ex.HResult, "Error"); }
        }
        private void detachExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelApp == null) { UpdateExcelStatus(false, "NA"); MessageBox.Show("Successfully detached"); return; }

                excelApp.Visible = true;
                //if (excelApp.Workbooks.Count > 0) { throw new Exception("Worbook is open in excel, close workbook first"); }

                DialogResult res = MessageBox.Show("Initiate close Execl?", "Warning", MessageBoxButtons.YesNoCancel);
                if (res == DialogResult.Yes) { excelApp.Quit(); }
                else if (res == DialogResult.No) { } // Don't do anything, just release later
                else { throw new Exception("Operation terminated by user."); }

                //Marshal.FinalReleaseComObject(excelApp);
                InternalReleaseExcel();
            }
            catch (Exception ex) when (ex.HResult == -2147023174)
            {
                InternalReleaseExcel();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private void attachRunningExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelApp != null)
                {
                    DialogResult res = MessageBox.Show("Instance of excel already attached. Detach before proceeding?", "Warning", MessageBoxButtons.YesNo);
                    if (res == DialogResult.Yes) { detachExcel_Click(sender, e); }
                    else { throw new Exception("Operation terminated by user."); }
                }

                AttachExistingExcelInstance();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        #region Helper Functions
        private void SubscribeToExcelStatus()
        {
            if (excelApp == null) { return; }
            excelApp.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(RefreshActiveWB);
            excelApp.WorkbookBeforeClose += new AppEvents_WorkbookBeforeCloseEventHandler(DetachIfFinalWB);
        }
        private void UnsubscribeToExcelStatus()
        {
            if (excelApp != null)
            {
                try { excelApp.WorkbookActivate -= new AppEvents_WorkbookActivateEventHandler(RefreshActiveWB); }
                catch { } //Ignore if event doesn't exist 
                try { excelApp.WorkbookBeforeClose -= new AppEvents_WorkbookBeforeCloseEventHandler(DetachIfFinalWB); }
                catch { } //Ignore if event doesn't exist 
            }
        }
        private void RefreshActiveWB(Workbook wb)
        {
            try { UpdateExcelStatus(true, wb.Name); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private void DetachIfFinalWB(Workbook Wb, ref bool Cancel)
        {
            try
            {
                if (excelApp.Workbooks.Count - 1 == 0)
                {
                    excelApp.Quit();
                    InternalReleaseExcel(false);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private void InternalReleaseExcel(bool showMsgBox = true)
        {
            UnsubscribeToExcelStatus();
            excelApp = null;
            UpdateExcelStatus(false, "NA");
            if (showMsgBox) { MessageBox.Show("Successfully detached"); }
        }

        private void UpdateExcelStatus(bool appStatus, string wbName)
        {
            DispExcelStatus.Invoke(new System.Action(() =>
            {
                DispExcelStatus.Text = $"Application attached: {appStatus.ToString()}\nActive Workbook: {wbName}";
            }));
        }
        private void AttachNewExcelInstance()
        {
            excelApp = new Excel.Application();
            excelApp.Visible = true;
            UpdateExcelStatus(true, "NA");
            SubscribeToExcelStatus();

            BringToFront();
            Activate();
            MessageBox.Show("Excel Launched Successfully");
        }
        private void AttachExistingExcelInstance()
        {
            bool excelExist = false;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                excelApp.Visible = true;
                excelExist = true;
                SubscribeToExcelStatus();
                UpdateExcelStatus(true, "NA");
                BringToFront();
                MessageBox.Show("Sucessfully attached to excel");
            }
            catch (Exception ex)
            {
                if (!excelExist)
                {
                    MessageBox.Show("No instance of excel found to attach");
                }
                else { MessageBox.Show($"{ex.Message}", "Error"); }
            }
        }
        #endregion
        #endregion
        #region Line Functions

        public void TestFunction()
        {
            //string msgText = "Test Form Print";
            try
            {
                // Get the current document and database, and start a transaction
                Document acDoc = AcadApp.DocumentManager.MdiActiveDocument;
                Editor editor = acDoc.Editor;
                Database acCurDb = acDoc.Database;

                using (acDoc.LockDocument())
                {
                    // Starts a new transaction with the Transaction Manager
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        PromptSelectionResult selectionResult = editor.GetSelection();

                        if (selectionResult.Status == PromptStatus.OK)
                        {
                            SelectionSet selectedObjs = selectionResult.Value;
                            foreach (SelectedObject selObj in selectedObjs)
                            {
                                //// Process each object...
                                Entity ent = acTrans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                                if (ent is AcadDB.Line line)
                                {
                                    double length = line.Length;
                                    MessageBox.Show($"Line found: length = {length}");
                                }
                                else
                                {
                                    MessageBox.Show($"Not a line");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }

        }

        #endregion
    }
}
