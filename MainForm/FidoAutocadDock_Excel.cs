using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDB = Autodesk.AutoCAD.DatabaseServices;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;

namespace FidoAutoCad.Forms
{
    public partial class FidoAutocadDock : UserControl
    {
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
        public void detachExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelApp == null) { UpdateExcelStatus(false, "NA"); MessageBox.Show("Successfully detached"); return; }

                excelApp.Visible = true;
                InternalReleaseExcel();
            }
            catch (Exception ex) when (ex.HResult == -2147023174)
            {
                InternalReleaseExcel();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        public void detachExcel_Auto(object sender, EventArgs e)
        {
            // To be called upon autocad shutdown
            try
            {
                if (excelApp == null) { return; }
                InternalReleaseExcel(false);
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

        private AppEvents_WorkbookActivateEventHandler workbookActivateHandler;
        private AppEvents_WorkbookAfterSaveEventHandler workbookAfterSaveHandler;
        private AppEvents_WorkbookBeforeCloseEventHandler workbookBeforeCloseHandler;
        private void SubscribeToExcelStatus()
        {
            UnsubscribeToExcelStatus();
            if (excelApp == null) { return; }

            workbookActivateHandler = new AppEvents_WorkbookActivateEventHandler((Workbook wb) => RefreshActiveWB());
            workbookAfterSaveHandler = new AppEvents_WorkbookAfterSaveEventHandler((Workbook wb, bool success) => RefreshActiveWB());
            workbookBeforeCloseHandler = new AppEvents_WorkbookBeforeCloseEventHandler(DetachIfFinalWB);

            excelApp.WorkbookActivate += workbookActivateHandler;
            excelApp.WorkbookAfterSave += workbookAfterSaveHandler;
            excelApp.WorkbookBeforeClose += workbookBeforeCloseHandler;
        }
        private void UnsubscribeToExcelStatus()
        {
            try { excelApp.WorkbookActivate -= workbookActivateHandler; } catch { }
            try { excelApp.WorkbookAfterSave -= workbookAfterSaveHandler; } catch { }
            try { excelApp.WorkbookBeforeClose -= workbookBeforeCloseHandler; } catch { }
        }
        private void RefreshActiveWB(bool showError = true)
        {
            try
            {
                Workbook wb = excelApp.ActiveWorkbook;
                UpdateExcelStatus(true, wb.Name);
            }
            catch (Exception ex) { if (showError) { MessageBox.Show(ex.Message, "Error"); } }
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
            try { Marshal.ReleaseComObject(excelApp); }
            catch { }
            
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
                RefreshActiveWB(false);
                //MessageBox.Show("Sucessfully attached to excel");
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

        #region Checks
        private bool CheckIfExcelIsAttached(bool throwError = true)
        {
            if (excelApp == null)
            {
                if (throwError) { throw new Exception("No instance of excel attached"); }
                return false;
            }
            else { return true; }
        }
        #endregion


    }
}
